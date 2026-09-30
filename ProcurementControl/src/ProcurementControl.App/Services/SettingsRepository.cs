using System.IO;
using Microsoft.Data.Sqlite;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Read-only запросы раздела «Настройки» — перенос Get-PurchaseSetting,
/// Get-PurchaseBooleanSetting, Get-PurchaseTrashItems и Get-PurchaseDocumentsRoot
/// из app/modules/PurchaseStore.ps1. Только SELECT: Set-PurchaseSetting и
/// Set-PurchaseDocumentsRoot в порт не входят.
/// </summary>
public sealed class SettingsRepository
{
    private readonly SqliteConnection _connection;

    public SettingsRepository(SqliteConnection connection)
    {
        _connection = connection;
    }

    /// <summary>Аналог Get-PurchaseSetting: значение ключа из таблицы settings.</summary>
    public string GetSetting(string key)
    {
        using var command = _connection.CreateCommand();
        command.CommandText = "SELECT value FROM settings WHERE key = @key";
        command.Parameters.AddWithValue("@key", key ?? string.Empty);
        var value = command.ExecuteScalar();
        return value is null or DBNull ? string.Empty : Convert.ToString(value) ?? string.Empty;
    }

    /// <summary>
    /// Аналог Get-PurchaseBooleanSetting без дописывания дефолта: оригинал
    /// сохраняет значение при отсутствии ключа, порт работает только на чтение.
    /// </summary>
    public bool GetBooleanSetting(string key, bool defaultValue)
    {
        var raw = GetSetting(key).Trim().ToLowerInvariant();
        if (raw is "1" or "true" or "yes")
        {
            return true;
        }

        if (raw is "0" or "false" or "no")
        {
            return false;
        }

        return defaultValue;
    }

    /// <summary>Аналог Get-PurchaseTrashItems: записи корзины, свежие сверху.</summary>
    public List<TrashItem> GetTrashItems()
    {
        const string sql = @"
SELECT id, entity_type, entity_id, title, deleted_at,
       CASE WHEN IFNULL(files_json, '') = '' THEN 0 ELSE 1 END AS has_files
FROM trash_items
ORDER BY deleted_at DESC, id DESC";

        var result = new List<TrashItem>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new TrashItem
            {
                Id = reader.GetInt64(0),
                EntityType = Text(reader, 1),
                EntityId = reader.IsDBNull(2) ? 0 : Convert.ToInt64(reader.GetValue(2)),
                Title = Text(reader, 3),
                DeletedAt = Text(reader, 4),
                HasFiles = !reader.IsDBNull(5) && Convert.ToInt64(reader.GetValue(5)) != 0,
            });
        }
        return result;
    }

    /// <summary>
    /// Аналог Get-PurchaseDocumentsRoot без записи: чтение ключа
    /// «documents_root», резолв портативного пути, фолбэк на папку по умолчанию.
    /// </summary>
    public string GetDocumentsRoot()
    {
        var root = GetSetting("documents_root");
        if (string.IsNullOrWhiteSpace(root))
        {
            // Оригинал здесь сохранял дефолт в базу; порт только читает.
            root = DefaultDocumentsRoot;
        }

        var movedRoot = GetMovedPortableTreeCandidate(root, @"data\purchase_control\files");
        var resolved = string.IsNullOrWhiteSpace(movedRoot) ? ResolvePortablePath(root) : movedRoot;
        if (string.IsNullOrWhiteSpace(resolved))
        {
            resolved = DefaultDocumentsRoot;
        }

        return resolved;
    }

    /// <summary>Аналог Get-DefaultPurchaseDocumentsRoot.</summary>
    private static string DefaultDocumentsRoot => Path.Combine(AppPaths.DataRoot, "purchase_control", "files");

    /// <summary>Аналог Resolve-PortablePath (PurchaseStore.ps1, строки 52-63).</summary>
    private static string ResolvePortablePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        var text = path.Trim();
        if (Path.IsPathRooted(text))
        {
            return text;
        }

        try
        {
            return Path.GetFullPath(Path.Combine(AppPaths.AppRoot, text));
        }
        catch
        {
            return Path.Combine(AppPaths.AppRoot, text);
        }
    }

    /// <summary>
    /// Аналог Get-MovedPortableTreeCandidate (PurchaseStore.ps1, строки 81-99):
    /// если сохранённый путь содержит переносимое дерево, путь строится от
    /// текущего корня приложения.
    /// </summary>
    private static string GetMovedPortableTreeCandidate(string storedPath, string portableTree)
    {
        if (string.IsNullOrWhiteSpace(storedPath) || string.IsNullOrWhiteSpace(portableTree))
        {
            return string.Empty;
        }

        var separator = Path.DirectorySeparatorChar;
        var normalizedPath = storedPath.Replace(Path.AltDirectorySeparatorChar, separator);
        var normalizedTree = portableTree
            .Trim(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            .Replace(Path.AltDirectorySeparatorChar, separator);
        if (string.IsNullOrWhiteSpace(normalizedTree))
        {
            return string.Empty;
        }

        var marker = normalizedTree + separator;
        var index = normalizedPath.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index >= 0)
        {
            var relative = normalizedPath[index..];
            return Path.Combine(AppPaths.AppRoot, relative);
        }

        if (normalizedPath.EndsWith(normalizedTree, StringComparison.OrdinalIgnoreCase))
        {
            return Path.Combine(AppPaths.AppRoot, normalizedTree);
        }

        return string.Empty;
    }

    private static string Text(SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? string.Empty : Convert.ToString(reader.GetValue(ordinal)) ?? string.Empty;
}
