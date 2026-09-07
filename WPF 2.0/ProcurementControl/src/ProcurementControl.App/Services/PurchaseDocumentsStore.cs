using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Документы сделки — перенос блока документов из app/modules/PurchaseStore.ps1:
/// Get-PurchaseDocuments (строки 1942-1963), Resolve-PurchaseDocumentFile
/// (строки 1995-2070), Add-PurchaseDocument (строки 2141-2192) и
/// Delete-PurchaseDocument (строки 2194-2216). Пути хранятся портативно
/// относительно корня приложения, удаление всегда идёт через корзину с файлом.
/// </summary>
public static class PurchaseDocumentsStore
{
    /// <summary>Типы документов из диалога оригинала (Show-DocumentTypeDialog).</summary>
    public static readonly string[] DocumentTypes =
    {
        "PI/Invoice", "Общий PI", "RFQ", "RRFQ", "PO", "ERP", "Other",
    };

    /// <summary>Корень документов по умолчанию: {данные закупок}\files.</summary>
    private static string DefaultDocumentsRoot => Path.Combine(AppPaths.PurchaseDataDirectory, "files");

    /// <summary>Дерево для починки переехавшей папки (как в оригинале).</summary>
    private static string DocumentsPortableTree => Path.Combine("data", "purchase_control", "files");

    /// <summary>
    /// Порт Get-PurchaseDocuments (строки 1942-1963): документы сделки с именем
    /// поставщика, новые сверху.
    /// </summary>
    public static List<PurchaseDocumentRow> GetDocuments(long dealId)
    {
        var result = new List<PurchaseDocumentRow>();
        if (dealId <= 0)
        {
            return result;
        }

        using var connection = OpenReadConnection();
        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT
    doc.id,
    doc.document_type,
    IFNULL(ds.supplier, '') AS supplier,
    doc.original_name,
    doc.stored_path,
    IFNULL(doc.file_hash, '') AS file_hash,
    doc.uploaded_at
FROM documents doc
LEFT JOIN deal_suppliers ds ON ds.id = doc.supplier_id
WHERE doc.deal_id = @deal_id
ORDER BY doc.uploaded_at DESC, doc.id DESC;";
        PurchaseWriteRepository.Add(command, "@deal_id", dealId);

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new PurchaseDocumentRow
            {
                Id = reader.GetInt64(0),
                DocumentType = reader.GetString(1),
                Supplier = reader.GetString(2),
                OriginalName = reader.GetString(3),
                StoredPath = reader.GetString(4),
                FileHash = reader.GetString(5),
                UploadedAt = reader.IsDBNull(6) ? string.Empty : reader.GetString(6),
            });
        }
        return result;
    }

    /// <summary>
    /// Порт Get-PurchaseDocumentsRoot (строки 135-146): настройка
    /// 'documents_root' с дефолтом и поддержкой переехавшего дерева.
    /// </summary>
    public static string GetDocumentsRoot()
    {
        using var connection = OpenReadConnection();
        return ResolveDocumentsRoot(connection, persistDefault: false);
    }

    /// <summary>
    /// Порт Add-PurchaseDocument (строки 2141-2192): копия файла кладётся в
    /// {корень}\{номер сделки}\{папка}, где папка = переданное имя, тип
    /// документа уровня сделки или имя поставщика. Возвращает путь копии.
    /// </summary>
    public static string AddDocument(long dealId, long supplierId, string documentType, string sourcePath, string folderName = "")
    {
        if (dealId <= 0)
        {
            throw new InvalidOperationException("Выберите сделку.");
        }
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException("Файл не найден: " + sourcePath, sourcePath);
        }

        var type = ConvertDocumentType(documentType);
        if (IsDealLevelDocumentType(type))
        {
            supplierId = 0;
        }

        return PurchaseWriteSession.Execute(connection =>
        {
            EnsureFileHashColumn(connection);
            var dealNumber = ScalarString(connection, "SELECT deal_number FROM deals WHERE id = @id", dealId);
            if (string.IsNullOrWhiteSpace(dealNumber))
            {
                throw new InvalidOperationException("Сделка не найдена.");
            }

            var supplier = string.Empty;
            if (supplierId > 0)
            {
                supplier = ScalarString(connection, "SELECT supplier FROM deal_suppliers WHERE id = @id", supplierId) ?? string.Empty;
            }

            var root = ResolveDocumentsRoot(connection, persistDefault: true);
            var folderPart = !string.IsNullOrWhiteSpace(folderName)
                ? folderName
                : IsDealLevelDocumentType(type) ? type : supplier;
            var targetDir = Path.Combine(root, SafePathPart(dealNumber), SafePathPart(folderPart));
            Directory.CreateDirectory(targetDir);

            var originalName = Path.GetFileName(sourcePath);
            var targetPath = GetUniqueDocumentPath(targetDir, originalName);
            File.Copy(sourcePath, targetPath);

            using var insert = connection.CreateCommand();
            insert.CommandText = @"
INSERT INTO documents(deal_id, supplier_id, document_type, original_name, stored_path, file_hash, uploaded_at)
VALUES(@deal_id, @supplier_id, @document_type, @original_name, @stored_path, @file_hash, @uploaded_at);";
            PurchaseWriteRepository.Add(insert, "@deal_id", dealId);
            PurchaseWriteRepository.Add(insert, "@supplier_id", supplierId > 0 ? supplierId : null);
            PurchaseWriteRepository.Add(insert, "@document_type", type);
            PurchaseWriteRepository.Add(insert, "@original_name", originalName);
            PurchaseWriteRepository.Add(insert, "@stored_path", ToPortablePath(targetPath));
            PurchaseWriteRepository.Add(insert, "@file_hash", GetFileHash(targetPath));
            PurchaseWriteRepository.Add(insert, "@uploaded_at", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            insert.ExecuteNonQuery();

            // PI/Invoice всегда загружается в контексте конкретного поставщика.
            // После успешного сохранения файла автоматически отмечаем, что
            // инвойс от этого поставщика получен. Общий PI к поставщику не
            // привязан и на его флаги не влияет.
            if (type == "PI/Invoice" && supplierId > 0)
            {
                using var markInvoiceReceived = connection.CreateCommand();
                markInvoiceReceived.CommandText = @"
UPDATE deal_suppliers
SET invoice_received = 1, updated_at = @updated_at
WHERE id = @supplier_id AND IFNULL(invoice_received, 0) = 0;";
                PurchaseWriteRepository.Add(markInvoiceReceived, "@updated_at", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                PurchaseWriteRepository.Add(markInvoiceReceived, "@supplier_id", supplierId);
                if (markInvoiceReceived.ExecuteNonQuery() > 0)
                {
                    PurchaseWriteRepository.WriteActivity(connection, "supplier", supplierId,
                        "Инвойс получен", "Отметка установлена при загрузке PI/Invoice.", dealId, supplierId);
                }
            }

            TouchDeal(connection, dealId);
            PurchaseWriteRepository.WriteActivity(connection, "document", 0, "Загружен документ",
                $"{type}: {originalName}", dealId, supplierId);
            return targetPath;
        });
    }

    /// <summary>
    /// Порт Delete-PurchaseDocument (строки 2194-2216): файл переезжает в
    /// корзину, строка документа сохраняется в payload и удаляется.
    /// </summary>
    public static void DeleteDocument(long documentId)
    {
        if (documentId <= 0)
        {
            throw new InvalidOperationException("Выберите документ.");
        }

        PurchaseWriteSession.Execute(connection =>
        {
            var rows = PurchaseWriteRepository.SelectRows(connection,
                "SELECT * FROM documents WHERE id = @id",
                command => PurchaseWriteRepository.Add(command, "@id", documentId));
            if (rows.Count == 0)
            {
                throw new InvalidOperationException("Документ не найден.");
            }

            var row = rows[0];
            var dealId = PurchaseWriteRepository.ToLong(row.GetValueOrDefault("deal_id"));
            var originalName = AsString(row.GetValueOrDefault("original_name"));
            var resolved = ResolveDocumentFile(connection, documentId,
                AsString(row.GetValueOrDefault("stored_path")),
                originalName,
                AsString(row.GetValueOrDefault("file_hash")));

            PurchaseWriteRepository.InsertTrashItem(connection, "document", documentId,
                "Документ: " + originalName,
                new Dictionary<string, object> { ["documents"] = rows },
                new[] { resolved.Path },
                string.Empty);

            using var delete = connection.CreateCommand();
            delete.CommandText = "DELETE FROM documents WHERE id = @id;";
            PurchaseWriteRepository.Add(delete, "@id", documentId);
            delete.ExecuteNonQuery();

            TouchDeal(connection, dealId);
            PurchaseWriteRepository.WriteActivity(connection, "document", documentId, "Удален документ",
                resolved.Path, dealId);
            return 0;
        });
    }

    /// <summary>
    /// Резолвит файл документа (с починкой ссылки в базе, как оригинал) и
    /// открывает его в проводнике с выделением. Возвращает успех операции.
    /// </summary>
    public static bool OpenInExplorer(long documentId, string storedPath, string originalName, string fileHash)
    {
        var resolved = PurchaseWriteSession.Execute(connection =>
            ResolveDocumentFile(connection, documentId, storedPath, originalName, fileHash));
        if (!resolved.Found || !File.Exists(resolved.Path))
        {
            return false;
        }

        Process.Start(new ProcessStartInfo("explorer.exe", "/select, \"" + resolved.Path + "\"")
        {
            UseShellExecute = true,
        });
        return true;
    }

    /// <summary>Открывает документ двойным щелчком в приложении, назначенном для его типа.</summary>
    public static bool OpenDocument(long documentId, string storedPath, string originalName, string fileHash)
    {
        var resolved = PurchaseWriteSession.Execute(connection =>
            ResolveDocumentFile(connection, documentId, storedPath, originalName, fileHash));
        if (!resolved.Found || !File.Exists(resolved.Path))
        {
            return false;
        }

        Process.Start(new ProcessStartInfo(resolved.Path)
        {
            UseShellExecute = true,
        });
        return true;
    }

    /// <summary>
    /// Порт Resolve-PurchaseDocumentFile (строки 1995-2070): сначала прямой путь;
    /// если файла нет — поиск в той же папке по уникальному хэшу, затем по
    /// уникальному расширению среди незанятых файлов, затем единственный свободный.
    /// Найденная ссылка сохраняется обратно в базу.
    /// </summary>
    internal static ResolvedDocumentFile ResolveDocumentFile(
        SqliteConnection connection,
        long documentId,
        string storedPath,
        string originalName,
        string fileHash)
    {
        var path = ResolveStoredPath(storedPath);
        var name = string.IsNullOrWhiteSpace(originalName) ? Path.GetFileName(path) : originalName;
        var hash = fileHash ?? string.Empty;

        if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
        {
            var currentName = Path.GetFileName(path);
            if (string.IsNullOrWhiteSpace(hash))
            {
                hash = GetFileHash(path);
            }
            var portablePath = ToPortablePath(path);
            if (documentId > 0 && (currentName != name || string.IsNullOrWhiteSpace(fileHash) || portablePath != storedPath))
            {
                UpdateDocumentFileReference(connection, documentId, path, hash);
            }
            return new ResolvedDocumentFile(documentId, path, currentName, hash, true,
                currentName != name || portablePath != storedPath);
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            return new ResolvedDocumentFile(documentId, path, name, hash, false, false);
        }

        string? directory = null;
        try
        {
            directory = Path.GetDirectoryName(path);
        }
        catch (ArgumentException)
        {
        }
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
        {
            return new ResolvedDocumentFile(documentId, path, name, hash, false, false);
        }

        var candidates = new DirectoryInfo(directory).EnumerateFiles().ToList();
        string? chosen = null;
        if (!string.IsNullOrWhiteSpace(hash))
        {
            var hashMatches = candidates
                .Where(candidate => GetFileHash(candidate.FullName) == hash)
                .ToList();
            if (hashMatches.Count == 1)
            {
                chosen = hashMatches[0].FullName;
            }
        }

        if (chosen is null)
        {
            var knownPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var otherRow in PurchaseWriteRepository.SelectRows(connection,
                         "SELECT stored_path FROM documents WHERE id <> @id",
                         command => PurchaseWriteRepository.Add(command, "@id", documentId)))
            {
                var otherPath = ResolveStoredPath(AsString(otherRow.GetValueOrDefault("stored_path")));
                if (string.IsNullOrWhiteSpace(otherPath))
                {
                    continue;
                }
                try
                {
                    knownPaths.Add(Path.GetFullPath(otherPath));
                }
                catch (ArgumentException)
                {
                }
            }

            var available = candidates
                .Where(candidate =>
                {
                    try
                    {
                        return !knownPaths.Contains(Path.GetFullPath(candidate.FullName));
                    }
                    catch (ArgumentException)
                    {
                        return false;
                    }
                })
                .ToList();

            var extension = Path.GetExtension(path);
            if (!string.IsNullOrWhiteSpace(extension))
            {
                var sameExtension = available
                    .Where(candidate => Path.GetExtension(candidate.FullName).Equals(extension, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                if (sameExtension.Count == 1)
                {
                    chosen = sameExtension[0].FullName;
                }
            }
            if (chosen is null && available.Count == 1)
            {
                chosen = available[0].FullName;
            }
        }

        if (chosen is not null)
        {
            var newHash = GetFileHash(chosen);
            UpdateDocumentFileReference(connection, documentId, chosen, newHash);
            return new ResolvedDocumentFile(documentId, chosen, Path.GetFileName(chosen), newHash, true, true);
        }

        return new ResolvedDocumentFile(documentId, path, name, hash, false, false);
    }

    /// <summary>Порт Convert-PurchaseDocumentType (строки 1376-1382).</summary>
    public static string ConvertDocumentType(string? documentType)
    {
        var text = string.IsNullOrWhiteSpace(documentType) ? "Other" : documentType.Trim();
        if (text is "ERP supplier" or "ERP Roger")
        {
            return "ERP";
        }
        return text;
    }

    /// <summary>Порт Test-DealLevelDocumentType (строки 1384-1389).</summary>
    public static bool IsDealLevelDocumentType(string? documentType)
    {
        var text = string.IsNullOrWhiteSpace(documentType) ? string.Empty : documentType.Trim();
        return text is "Общий PI" or "RFQ" or "RRFQ" or "PO" or "Other";
    }

    // ----- вспомогательные порты путей и хэшей ------------------------------

    /// <summary>Порт Resolve-PurchaseStoredPath (строки 102-117) для файла.</summary>
    private static string ResolveStoredPath(string? storedPath)
    {
        if (string.IsNullOrWhiteSpace(storedPath))
        {
            return string.Empty;
        }
        var path = PurchaseWriteRepository.ResolvePortable(storedPath);
        if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
        {
            try
            {
                return Path.GetFullPath(path);
            }
            catch (ArgumentException)
            {
                return path;
            }
        }

        var candidate = GetMovedTreeCandidate(storedPath, DocumentsPortableTree);
        if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate))
        {
            try
            {
                return Path.GetFullPath(candidate);
            }
            catch (ArgumentException)
            {
                return candidate;
            }
        }
        return path;
    }

    /// <summary>Порт Get-MovedPortableTreeCandidate (строки 81-100).</summary>
    internal static string GetMovedTreeCandidate(string storedPath, string portableTree)
    {
        if (string.IsNullOrWhiteSpace(storedPath) || string.IsNullOrWhiteSpace(portableTree))
        {
            return string.Empty;
        }

        var separator = Path.DirectorySeparatorChar;
        var normalizedPath = storedPath.Replace(Path.AltDirectorySeparatorChar, separator);
        var normalizedTree = portableTree
            .Trim(separator, Path.AltDirectorySeparatorChar)
            .Replace(Path.AltDirectorySeparatorChar, separator);
        if (string.IsNullOrWhiteSpace(normalizedTree))
        {
            return string.Empty;
        }

        var marker = normalizedTree + separator;
        var index = normalizedPath.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index >= 0)
        {
            return Path.Combine(AppPaths.AppRoot, normalizedPath[index..]);
        }
        if (normalizedPath.EndsWith(normalizedTree, StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine(AppPaths.AppRoot, normalizedTree);
        }
        return string.Empty;
    }

    /// <summary>Порт Convert-ToPortablePath (строки 65-79).</summary>
    internal static string ToPortablePath(string path)
    {
        try
        {
            var full = Path.GetFullPath(PurchaseWriteRepository.ResolvePortable(path));
            var root = Path.GetFullPath(AppPaths.AppRoot);
            if (!root.EndsWith(Path.DirectorySeparatorChar))
            {
                root += Path.DirectorySeparatorChar;
            }
            if (full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            {
                return full[root.Length..];
            }
        }
        catch (ArgumentException)
        {
        }
        return path;
    }

    /// <summary>Порт Get-SafePathPart (строки 1412-1420).</summary>
    internal static string SafePathPart(string value)
    {
        var text = string.IsNullOrWhiteSpace(value) ? "empty" : value.Trim();
        foreach (var invalidChar in Path.GetInvalidFileNameChars())
        {
            text = text.Replace(invalidChar, '_');
        }
        return Regex.Replace(text, @"\s+", "_");
    }

    /// <summary>Порт Get-UniquePurchaseDocumentPath (строки 1391-1410).</summary>
    private static string GetUniqueDocumentPath(string directory, string fileName)
    {
        var safeName = string.IsNullOrWhiteSpace(fileName) ? "document" : fileName.Trim();
        foreach (var invalidChar in Path.GetInvalidFileNameChars())
        {
            safeName = safeName.Replace(invalidChar, '_');
        }
        if (string.IsNullOrWhiteSpace(safeName))
        {
            safeName = "document";
        }

        var candidate = Path.Combine(directory, safeName);
        if (!File.Exists(candidate))
        {
            return candidate;
        }

        var stem = Path.GetFileNameWithoutExtension(safeName);
        var extension = Path.GetExtension(safeName);
        for (var index = 2; index < 1000; index++)
        {
            candidate = Path.Combine(directory, $"{stem} ({index}){extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidOperationException("Не удалось подобрать свободное имя файла: " + fileName);
    }

    /// <summary>Порт Get-PurchaseDocumentFileHash (строки 1965-1974): SHA256 в верхнем регистре.</summary>
    private static string GetFileHash(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return string.Empty;
        }
        try
        {
            using var stream = File.OpenRead(path);
            return Convert.ToHexString(SHA256.HashData(stream));
        }
        catch (IOException)
        {
            return string.Empty;
        }
    }

    /// <summary>Порт Update-PurchaseDocumentFileReference (строки 1976-1993).</summary>
    private static void UpdateDocumentFileReference(SqliteConnection connection, long documentId, string path, string fileHash)
    {
        if (documentId <= 0 || string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        var resolvedPath = ResolveStoredPath(path);
        if (string.IsNullOrWhiteSpace(resolvedPath))
        {
            resolvedPath = path;
        }
        var name = Path.GetFileName(resolvedPath);
        if (string.IsNullOrWhiteSpace(fileHash))
        {
            fileHash = GetFileHash(resolvedPath);
        }

        using var command = connection.CreateCommand();
        command.CommandText = @"
UPDATE documents
SET stored_path = @stored_path,
    original_name = @original_name,
    file_hash = CASE WHEN @file_hash <> '' THEN @file_hash ELSE file_hash END
WHERE id = @id;";
        PurchaseWriteRepository.Add(command, "@stored_path", ToPortablePath(resolvedPath));
        PurchaseWriteRepository.Add(command, "@original_name", name);
        PurchaseWriteRepository.Add(command, "@file_hash", fileHash);
        PurchaseWriteRepository.Add(command, "@id", documentId);
        command.ExecuteNonQuery();
    }

    /// <summary>Читает 'documents_root' и при необходимости сохраняет дефолт.</summary>
    private static string ResolveDocumentsRoot(SqliteConnection connection, bool persistDefault)
    {
        var root = ScalarString(connection, "SELECT value FROM settings WHERE key = 'documents_root'");
        if (string.IsNullOrWhiteSpace(root))
        {
            root = ToPortablePath(DefaultDocumentsRoot);
            if (persistDefault)
            {
                using var upsert = connection.CreateCommand();
                upsert.CommandText = @"
INSERT INTO settings(key, value) VALUES('documents_root', @value)
ON CONFLICT(key) DO UPDATE SET value = excluded.value;";
                PurchaseWriteRepository.Add(upsert, "@value", root);
                upsert.ExecuteNonQuery();
            }
        }

        var movedRoot = GetMovedTreeCandidate(root, DocumentsPortableTree);
        var resolved = !string.IsNullOrWhiteSpace(movedRoot)
            ? movedRoot
            : PurchaseWriteRepository.ResolvePortable(root);
        return string.IsNullOrWhiteSpace(resolved) ? DefaultDocumentsRoot : resolved;
    }

    /// <summary>Порт Touch-PurchaseDeal (строки 1615-1623).</summary>
    private static void TouchDeal(SqliteConnection connection, long dealId)
    {
        if (dealId <= 0)
        {
            return;
        }
        using var command = connection.CreateCommand();
        command.CommandText = "UPDATE deals SET updated_at = @updated_at WHERE id = @id;";
        PurchaseWriteRepository.Add(command, "@updated_at", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        PurchaseWriteRepository.Add(command, "@id", dealId);
        command.ExecuteNonQuery();
    }

    /// <summary>
    /// Как оригинальный Ensure-PurchaseColumn 'documents' 'file_hash'
    /// (PurchaseStore.ps1, строка 621) — дописывает колонку в старых базах.
    /// </summary>
    private static void EnsureFileHashColumn(SqliteConnection connection)
    {
        using var pragma = connection.CreateCommand();
        pragma.CommandText = "PRAGMA table_info(documents);";
        using var reader = pragma.ExecuteReader();
        while (reader.Read())
        {
            if (string.Equals(reader.GetString(1), "file_hash", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
        }

        using var alter = connection.CreateCommand();
        alter.CommandText = "ALTER TABLE documents ADD COLUMN file_hash TEXT;";
        alter.ExecuteNonQuery();
    }

    private static SqliteConnection OpenReadConnection()
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = AppPaths.PurchaseDatabasePath,
            Mode = SqliteOpenMode.ReadOnly,
            Pooling = false,
            DefaultTimeout = 10,
        };
        var connection = new SqliteConnection(builder.ConnectionString);
        connection.Open();
        return connection;
    }

    private static string? ScalarString(SqliteConnection connection, string sql, long? id = null)
    {
        using var command = connection.CreateCommand();
        command.CommandText = sql + ";";
        if (id is not null)
        {
            PurchaseWriteRepository.Add(command, "@id", id.Value);
        }
        var value = command.ExecuteScalar();
        return value is null or DBNull ? null : Convert.ToString(value);
    }

    private static string AsString(object? value)
        => value is null or DBNull ? string.Empty : Convert.ToString(value) ?? string.Empty;
}
