using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Папка восстановления — перенос чтения (Get-PurchaseRecoveryRoot,
/// Get-RecoverySnapshotManifest, Get-PurchaseRecoverySnapshots) и записи
/// (New-PurchaseRecoverySnapshot, Remove-PurchaseRecoverySnapshot) из
/// app/modules/DataSafety.ps1. Восстановление из полного архива — через
/// StageRestore + ApplyPendingRestore: архив распаковывается сразу, а замена
/// базы выполняется при следующем запуске приложения, пока файлы не заняты.
/// </summary>
public static class RecoveryStore
{
    /// <summary>Аналог Get-PurchaseRecoveryRoot: {AppRoot}/data/recovery.</summary>
    public static string RecoveryRoot => Path.Combine(AppPaths.DataRoot, "recovery");

    /// <summary>Маркер отложенного восстановления (JSON со путём распакованной копии).</summary>
    public static string PendingRestoreMarkerPath => Path.Combine(RecoveryRoot, "restore_pending.json");

    /// <summary>Готово ли восстановление, ожидающее перезапуска приложения.</summary>
    public static bool HasPendingRestore => File.Exists(PendingRestoreMarkerPath);

    /// <summary>Аналог Get-PurchaseRecoveryFolder: папки trash и backups.</summary>
    public static string GetFolder(string kind)
        => Path.Combine(RecoveryRoot, kind == "Trash" ? "trash" : "backups");

    /// <summary>
    /// Аналог Get-PurchaseRecoverySnapshots: все *.zip из обеих папок с
    /// манифестом; повреждённые архивы помечаются отдельно. Сортировка по
    /// дате создания по убыванию, как в оригинале (строка 158).
    /// </summary>
    public static List<RecoverySnapshot> GetSnapshots()
    {
        var items = new List<RecoverySnapshot>();
        foreach (var kind in new[] { "Trash", "Backup" })
        {
            var folder = GetFolder(kind);
            if (!Directory.Exists(folder))
            {
                continue;
            }

            foreach (var file in new DirectoryInfo(folder).EnumerateFiles("*.zip", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    var manifest = ReadManifest(file.FullName);
                    items.Add(new RecoverySnapshot
                    {
                        Path = file.FullName,
                        Name = file.Name,
                        Kind = manifest.Kind ?? string.Empty,
                        Reason = manifest.Reason ?? string.Empty,
                        CreatedAt = DateTime.Parse(manifest.CreatedAt ?? string.Empty, null, System.Globalization.DateTimeStyles.RoundtripKind),
                        SizeBytes = file.Length,
                    });
                }
                catch
                {
                    items.Add(new RecoverySnapshot
                    {
                        Path = file.FullName,
                        Name = file.Name,
                        Kind = kind,
                        Reason = "Поврежденный или устаревший архив",
                        CreatedAt = file.LastWriteTime,
                        SizeBytes = file.Length,
                    });
                }
            }
        }

        return items.OrderByDescending(s => s.CreatedAt).ToList();
    }

    /// <summary>
    /// Аналог Refresh-RecoverySnapshots: записи корзины из базы и резервные копии.
    /// Отличие от оригинала (строка 3573): список единый и сортируется по дате
    /// по убыванию — иначе новые копии терялись в конце за десятками строк корзины.
    /// </summary>
    public static List<RecoveryRow> GetRecoveryRows(IReadOnlyList<TrashItem> trashItems)
    {
        var rows = new List<RecoveryRow>();
        foreach (var item in trashItems)
        {
            rows.Add(new RecoveryRow
            {
                RecoveryType = "Trash",
                TrashId = item.Id,
                KindText = "Корзина",
                CreatedText = FormatTrashDate(item.DeletedAt),
                CreatedAtSort = ParseTrashDate(item.DeletedAt) ?? DateTime.MinValue,
                Reason = item.Title,
                SizeText = item.HasFiles ? "с файлами" : "запись",
            });
        }

        foreach (var snapshot in GetSnapshots().Where(s => s.Kind == "Backup"))
        {
            var localCreated = snapshot.CreatedAt.ToLocalTime();
            rows.Add(new RecoveryRow
            {
                RecoveryType = "Backup",
                SnapshotPath = snapshot.Path,
                KindText = "Резервная копия",
                CreatedText = localCreated.ToString("dd.MM.yyyy HH:mm"),
                CreatedAtSort = localCreated,
                Reason = snapshot.Reason,
                SizeText = snapshot.SizeText,
            });
        }

        return rows.OrderByDescending(r => r.CreatedAtSort).ToList();
    }

    /// <summary>
    /// Порт New-PurchaseRecoverySnapshot (DataSafety.ps1, строки 64-115): в архив входят база (через временную копию),
    /// папки purchase_control/files, components/files и notes плюс manifest.json.
    /// Хранится не более 20 последних резервных копий.
    /// </summary>
    public static string CreateSnapshot(string kind = "Backup", string reason = "Резервная копия")
    {
        // Аналог Initialize-PurchaseRecoveryStore (строки 15-19).
        Directory.CreateDirectory(GetFolder("Backup"));
        Directory.CreateDirectory(GetFolder("Trash"));

        var createdAt = DateTime.Now;
        var stamp = createdAt.ToString("yyyyMMdd_HHmmss_fff");
        var safeReason = Regex.Replace(reason, @"[^0-9A-Za-zА-Яа-я_-]+", "_").Trim('_');
        if (string.IsNullOrWhiteSpace(safeReason))
        {
            safeReason = "snapshot";
        }
        if (safeReason.Length > 45)
        {
            safeReason = safeReason[..45];
        }

        var prefix = kind == "Trash" ? "trash" : "backup";
        var archivePath = Path.Combine(GetFolder(kind), $"{prefix}_{stamp}_{safeReason}.zip");
        var stage = Path.Combine(Path.GetTempPath(), "procurement-recovery-" + Guid.NewGuid().ToString("N"));

        try
        {
            var snapshotDbDir = Path.Combine(stage, "purchase_control");
            Directory.CreateDirectory(snapshotDbDir);

            // База кладётся из временной прочекпоинченной копии — живой файл и его WAL не трогаем.
            using (var snapshot = PurchaseSnapshot.Create())
            {
                var sourceDb = snapshot.SnapshotDatabasePath;
                if (sourceDb is not null)
                {
                    CopyLockedFile(sourceDb, Path.Combine(snapshotDbDir, "purchase_control.sqlite"));
                }
                else
                {
                    // ReadOnly-фолбэк: копируем живой файл напрямую.
                    CopyLockedFile(AppPaths.PurchaseDatabasePath, Path.Combine(snapshotDbDir, "purchase_control.sqlite"));
                }
            }

            CopyDirectory(Path.Combine(AppPaths.PurchaseDataDirectory, "files"), Path.Combine(snapshotDbDir, "files"));
            CopyDirectory(Path.Combine(AppPaths.DataRoot, "components", "files"), Path.Combine(stage, "components", "files"));
            CopyDirectory(AppPaths.NotesDirectory, Path.Combine(stage, "notes"));

            var manifest = new Dictionary<string, object>
            {
                ["Format"] = 1,
                ["Kind"] = kind,
                ["Reason"] = reason,
                ["CreatedAt"] = createdAt.ToString("o"),
                ["Includes"] = new[] { "purchase_control.sqlite", "purchase_control/files", "components/files", "notes" },
                ["Note"] = "Внешняя папка документов не входит в резервную копию.",
            };
            File.WriteAllText(Path.Combine(stage, "manifest.json"), JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            ZipFile.CreateFromDirectory(stage, archivePath, CompressionLevel.Optimal, includeBaseDirectory: false);
        }
        finally
        {
            try
            {
                if (Directory.Exists(stage))
                {
                    Directory.Delete(stage, recursive: true);
                }
            }
            catch
            {
                // Временная папка может удаляться с задержкой — не критично.
            }
        }

        if (!File.Exists(archivePath))
        {
            throw new InvalidOperationException("Не удалось создать резервную копию.");
        }

        if (kind == "Backup")
        {
            // Как в оригинале: старые копии удаляются, но не более 20 последних остаются.
            var oldBackups = GetSnapshots().Where(s => s.Kind == "Backup").OrderByDescending(s => s.CreatedAt).Skip(20);
            foreach (var oldBackup in oldBackups)
            {
                try
                {
                    RemoveSnapshot(oldBackup.Path);
                }
                catch
                {
                    // Очистка старых копий не должна ломать создание новой.
                }
            }
        }

        return archivePath;
    }

    /// <summary>
    /// Порт Remove-PurchaseRecoverySnapshot + Test-PurchaseRecoveryArchivePath
    /// (DataSafety.ps1, строки 161-171, 211-216): удаляется только архив из папки восстановления.
    /// </summary>
    public static void RemoveSnapshot(string path)
    {
        if (!IsRecoveryArchivePath(path))
        {
            throw new InvalidOperationException("Можно удалить только архив из списка восстановления.");
        }

        File.Delete(path);
    }

    /// <summary>Аналог Test-PurchaseRecoveryArchivePath.</summary>
    private static bool IsRecoveryArchivePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return false;
        }

        try
        {
            var fullPath = Path.GetFullPath(path);
            var root = Path.GetFullPath(RecoveryRoot);
            if (!root.EndsWith(Path.DirectorySeparatorChar))
            {
                root += Path.DirectorySeparatorChar;
            }

            return fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Готовит восстановление из резервной копии: архив распаковывается в папку
    /// восстановления и создаётся маркер. Сама замена данных выполняется
    /// ApplyPendingRestore при следующем запуске приложения (оно закрывается
    /// сразу после подтверждения), пока живая база никем не открыта.
    /// </summary>
    public static void StageRestore(string snapshotPath)
    {
        if (!IsRecoveryArchivePath(snapshotPath))
        {
            throw new InvalidOperationException("Восстановить можно только архив из списка восстановления.");
        }

        ReadManifest(snapshotPath);

        var stage = Path.Combine(RecoveryRoot, "restore_stage_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
        if (Directory.Exists(stage))
        {
            Directory.Delete(stage, recursive: true);
        }

        ZipFile.ExtractToDirectory(snapshotPath, stage, overwriteFiles: true);

        // Архив без базы восстанавливать нельзя — остались бы только файлы.
        if (!File.Exists(Path.Combine(stage, "purchase_control", "purchase_control.sqlite")))
        {
            try { Directory.Delete(stage, recursive: true); } catch { }
            throw new InvalidOperationException("В архиве нет файла базы purchase_control.sqlite.");
        }

        var marker = JsonSerializer.Serialize(new Dictionary<string, object>
        {
            ["Format"] = 1,
            ["Stage"] = stage,
            ["Source"] = snapshotPath,
            ["RequestedAt"] = DateTime.Now.ToString("o"),
        });
        File.WriteAllText(PendingRestoreMarkerPath, marker, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    /// <summary>
    /// Выполняет отложенное восстановление при запуске (вызывается из App.OnStartup
    /// до любых подключений к базе): подменяет базу, папки документов и заметки
    /// содержимым распакованного архива. Перед заменой делается страховочная
    /// копия текущих данных «Перед восстановлением»; её сбой не отменяет восстановление.
    /// </summary>
    public static void ApplyPendingRestore()
    {
        if (!HasPendingRestore)
        {
            return;
        }

        var marker = JsonSerializer.Deserialize<PendingRestoreMarker>(
            File.ReadAllText(PendingRestoreMarkerPath), ManifestOptions);
        if (marker?.Stage is null)
        {
            throw new InvalidOperationException("Маркер восстановления повреждён.");
        }

        var stage = marker.Stage;
        if (!Directory.Exists(stage))
        {
            throw new InvalidOperationException("Распакованная копия не найдена: " + stage);
        }

        try
        {
            CreateSnapshot("Backup", "Перед восстановлением");
        }
        catch (Exception exception)
        {
            ErrorLog.Append("restore", "Страховочная копия перед восстановлением не создана: " + exception.Message);
        }

        var dataPurchaseDirectory = AppPaths.PurchaseDataDirectory;
        Directory.CreateDirectory(dataPurchaseDirectory);

        // База: старый файл и его WAL/SHM удаляются, из архива кладётся копия.
        foreach (var suffix in new[] { string.Empty, "-wal", "-shm" })
        {
            var liveFile = AppPaths.PurchaseDatabasePath + suffix;
            if (File.Exists(liveFile))
            {
                File.Delete(liveFile);
            }
        }

        File.Copy(Path.Combine(stage, "purchase_control", "purchase_control.sqlite"), AppPaths.PurchaseDatabasePath);

        // Папки файлов и заметок замещаются целиком; папки, которых в архиве
        // нет (старая копия), остаются без изменений.
        ReplaceDirectory(Path.Combine(stage, "purchase_control", "files"), Path.Combine(dataPurchaseDirectory, "files"));
        ReplaceDirectory(Path.Combine(stage, "components", "files"), Path.Combine(AppPaths.DataRoot, "components", "files"));
        ReplaceDirectory(Path.Combine(stage, "notes"), AppPaths.NotesDirectory);

        File.Delete(PendingRestoreMarkerPath);
        try { Directory.Delete(stage, recursive: true); } catch { }

        ErrorLog.Append("restore", "Данные восстановлены из резервной копии: " + (marker.Source ?? stage));
    }

    /// <summary>Замещает папку назначения содержимым папки из распакованной копии.</summary>
    private static void ReplaceDirectory(string source, string destination)
    {
        if (!Directory.Exists(source))
        {
            return;
        }

        if (Directory.Exists(destination))
        {
            Directory.Delete(destination, recursive: true);
        }

        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        Directory.Move(source, destination);
    }

    /// <summary>Аналог Copy-RecoveryDirectory (DataSafety.ps1, строки 30-38).</summary>
    private static void CopyDirectory(string source, string destination)
    {
        if (!Directory.Exists(source))
        {
            return;
        }

        Directory.CreateDirectory(destination);
        foreach (var item in new DirectoryInfo(source).EnumerateFileSystemInfos("*", SearchOption.AllDirectories))
        {
            var target = Path.Combine(destination, Path.GetRelativePath(source, item.FullName));
            if (item is DirectoryInfo)
            {
                Directory.CreateDirectory(target);
            }
            else
            {
                Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                File.Copy(item.FullName, target, overwrite: true);
            }
        }
    }

    /// <summary>Копирует файл, который может быть открыт другим процессом (как в PurchaseSnapshot).</summary>
    private static void CopyLockedFile(string sourcePath, string destPath)
    {
        using var input = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var output = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None);
        input.CopyTo(output);
    }

    /// <summary>Аналог текста подсказки под сеткой (строка 3582).</summary>
    public static string Hint(int count)
        => $"Записей и копий: {count}. Корзина хранит только удалённые записи и их файлы; полные копии — не более 20 последних.";

    /// <summary>
    /// Аналог $btnOpenRecoveryFolder: открывает папку копий в проводнике.
    /// Папки создаются, как это делал оригинальный Initialize-PurchaseRecoveryStore.
    /// </summary>
    public static void OpenFolder()
    {
        Directory.CreateDirectory(GetFolder("Backup"));
        Directory.CreateDirectory(GetFolder("Trash"));

        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = "\"" + RecoveryRoot + "\"",
            UseShellExecute = true,
        });
    }

    /// <summary>Аналог [datetime]::Parse для deleted_at из базы.</summary>
    private static string FormatTrashDate(string deletedAt)
    {
        var parsed = ParseTrashDate(deletedAt);
        return parsed is null ? deletedAt : parsed.Value.ToString("dd.MM.yyyy HH:mm");
    }

    private static DateTime? ParseTrashDate(string deletedAt)
        => DateTime.TryParse(deletedAt, out var parsed) ? parsed : null;

    /// <summary>Аналог Get-RecoverySnapshotManifest: manifest.json из zip.</summary>
    private static RecoveryManifest ReadManifest(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry("manifest.json");
        if (entry is null)
        {
            throw new InvalidOperationException("В архиве нет manifest.json.");
        }

        using var stream = entry.Open();
        var manifest = JsonSerializer.Deserialize<RecoveryManifest>(stream, ManifestOptions);
        if (manifest is null)
        {
            throw new InvalidOperationException("Пустой manifest.json.");
        }

        return manifest;
    }

    private static readonly JsonSerializerOptions ManifestOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    private sealed class RecoveryManifest
    {
        public string? Kind { get; set; }

        public string? Reason { get; set; }

        public string? CreatedAt { get; set; }
    }

    private sealed class PendingRestoreMarker
    {
        public string? Stage { get; set; }

        public string? Source { get; set; }
    }
}
