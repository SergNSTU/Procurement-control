using System.IO;
using System.IO.Compression;
using System.Text.Json;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Read-only чтение папки восстановления — перенос Get-PurchaseRecoveryRoot,
/// Get-RecoverySnapshotManifest и Get-PurchaseRecoverySnapshots
/// (app/modules/DataSafety.ps1). Каталог не создаётся, файлы не пишутся;
/// изменение (создание/восстановление/удаление копий) в порт не входит.
/// </summary>
public static class RecoveryStore
{
    /// <summary>Аналог Get-PurchaseRecoveryRoot: {AppRoot}/data/recovery.</summary>
    public static string RecoveryRoot => Path.Combine(AppPaths.DataRoot, "recovery");

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
    /// Аналог Refresh-RecoverySnapshots: сначала записи корзины из базы,
    /// затем резервные копии (только Kind == «Backup», как строка 3573 оригинала).
    /// </summary>
    public static List<RecoveryRow> GetRecoveryRows(IReadOnlyList<TrashItem> trashItems)
    {
        var rows = new List<RecoveryRow>();
        foreach (var item in trashItems)
        {
            rows.Add(new RecoveryRow
            {
                KindText = "Корзина",
                CreatedText = FormatTrashDate(item.DeletedAt),
                Reason = item.Title,
                SizeText = item.HasFiles ? "с файлами" : "запись",
            });
        }

        foreach (var snapshot in GetSnapshots().Where(s => s.Kind == "Backup"))
        {
            rows.Add(new RecoveryRow
            {
                KindText = "Резервная копия",
                CreatedText = snapshot.CreatedAt.ToLocalTime().ToString("dd.MM.yyyy HH:mm"),
                Reason = snapshot.Reason,
                SizeText = snapshot.SizeText,
            });
        }

        return rows;
    }

    /// <summary>Аналог текста подсказки под сеткой (строка 3582).</summary>
    public static string Hint(int count)
        => $"Записей и копий: {count}. Корзина хранит только удалённые записи и их файлы; полные копии — не более 20 последних.";

    /// <summary>
    /// Аналог $btnOpenRecoveryFolder: открывает папку копий в проводнике.
    /// Каталог не создаётся — если папки нет, бросается ошибка для показа в интерфейсе.
    /// </summary>
    public static void OpenFolder()
    {
        if (!Directory.Exists(RecoveryRoot))
        {
            throw new InvalidOperationException("Папка копий не найдена.");
        }

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
        if (DateTime.TryParse(deletedAt, out var parsed))
        {
            return parsed.ToString("dd.MM.yyyy HH:mm");
        }

        return deletedAt;
    }

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
}
