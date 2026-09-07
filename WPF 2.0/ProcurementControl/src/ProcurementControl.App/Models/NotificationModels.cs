namespace ProcurementControl.Models;

/// <summary>
/// Одна сработавшая дата уведомления — перенос элементов Due из
/// Get-DueNotificationCandidates (PurchaseStore.ps1, строки 1065-1186).
/// </summary>
public sealed class DueEntry
{
    /// <summary>reminder / deadline / manual.</summary>
    public string Kind { get; set; } = string.Empty;

    /// <summary>Напоминание / Дедлайн.</summary>
    public string Label { get; set; } = string.Empty;

    public DateTime Date { get; set; }

    /// <summary>Текст даты для ключа состояния (yyyy-MM-dd).</summary>
    public string DateText { get; set; } = string.Empty;
}

/// <summary>
/// Кандидат на карточку уведомления. Соответствует объектам, которые строит
/// Get-DueNotificationCandidates по трём источникам: задачи (component),
/// напоминания сделок (deal), ручные напоминания (reminder).
/// </summary>
public sealed class DueNotification
{
    /// <summary>component / deal / reminder.</summary>
    public string Source { get; set; } = string.Empty;

    public long SourceId { get; set; }
    public string Title { get; set; } = string.Empty;

    /// <summary>Только для задач: описание.</summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>Только для задач: следующее действие.</summary>
    public string NextAction { get; set; } = string.Empty;

    public string Detail { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Stage { get; set; } = string.Empty;
    public long DealId { get; set; }
    public long SupplierId { get; set; }

    public List<DueEntry> Due { get; } = new();
}
