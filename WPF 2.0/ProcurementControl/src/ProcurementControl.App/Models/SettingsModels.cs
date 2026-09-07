namespace ProcurementControl.Models;

/// <summary>
/// Запись корзины — перенос строк запроса Get-PurchaseTrashItems
/// (app/modules/PurchaseStore.ps1, строки 749-756). Только чтение.
/// </summary>
public sealed class TrashItem
{
    public long Id { get; init; }

    public string EntityType { get; init; } = string.Empty;

    public long EntityId { get; init; }

    public string Title { get; init; } = string.Empty;

    public string DeletedAt { get; init; } = string.Empty;

    public bool HasFiles { get; init; }
}

/// <summary>
/// Манифест резервной копии из zip-архива — перенос
/// Get-RecoverySnapshotManifest (app/modules/DataSafety.ps1, строки 117-131).
/// </summary>
public sealed class RecoverySnapshot
{
    public string Path { get; init; } = string.Empty;

    public string Name { get; init; } = string.Empty;

    /// <summary>«Backup» или «Trash» (аналог поля Kind манифеста).</summary>
    public string Kind { get; init; } = string.Empty;

    public string Reason { get; init; } = string.Empty;

    public DateTime CreatedAt { get; init; }

    public long SizeBytes { get; init; }

    public string SizeText => RecoveryRow.ReadableFileSize(SizeBytes);
}

/// <summary>
/// Строка сетки «Корзина и резервные копии» — перенос строк
/// Refresh-RecoverySnapshots (RRFQComparer.ps1, строки 3562-3583).
/// </summary>
public sealed class RecoveryRow
{
    /// <summary>«Trash» или «Backup» — аналог поля Type из Get-SelectedRecoveryItem оригинала.</summary>
    public string RecoveryType { get; init; } = string.Empty;

    /// <summary>Для записи корзины — её id в таблице trash_items.</summary>
    public long TrashId { get; init; }

    /// <summary>Для резервной копии — путь к zip-архиву.</summary>
    public string SnapshotPath { get; init; } = string.Empty;

    /// <summary>«Корзина» или «Резервная копия».</summary>
    public string KindText { get; init; } = string.Empty;

    /// <summary>Дата в формате «dd.MM.yyyy HH:mm», как в оригинале.</summary>
    public string CreatedText { get; init; } = string.Empty;

    /// <summary>Для мусора — название записи, для копии — причина из манифеста.</summary>
    public string Reason { get; init; } = string.Empty;

    /// <summary>Для мусора — «с файлами»/«запись», для копии — читаемый размер.</summary>
    public string SizeText { get; init; } = string.Empty;

    /// <summary>Аналог Get-ReadableFileSize (app/modules/DataSafety.ps1, строки 21-28).</summary>
    public static string ReadableFileSize(long bytes)
    {
        if (bytes >= 1024L * 1024 * 1024)
        {
            return string.Format(System.Globalization.CultureInfo.CurrentCulture, "{0:N1} GB", bytes / (1024d * 1024 * 1024));
        }

        if (bytes >= 1024L * 1024)
        {
            return string.Format(System.Globalization.CultureInfo.CurrentCulture, "{0:N1} MB", bytes / (1024d * 1024));
        }

        if (bytes >= 1024)
        {
            return string.Format(System.Globalization.CultureInfo.CurrentCulture, "{0:N1} KB", bytes / 1024d);
        }

        return $"{bytes} B";
    }
}
