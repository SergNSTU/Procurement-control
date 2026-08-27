namespace ProcurementControl.Models;

/// <summary>
/// Строка раздела «Задачи» (таблица component_deals). Перенос полей запроса
/// Get-ComponentDeals из app/modules/PurchaseStore.ps1.
/// </summary>
public sealed class TaskRow
{
    public long Id { get; set; }
    public string EntryDate { get; set; } = string.Empty;
    public string DealNumber { get; set; } = string.Empty;
    public string Status { get; set; } = "В работе";
    public string Stage { get; set; } = "Запросил поставщиков";
    public string Description { get; set; } = string.Empty;
    public string NextAction { get; set; } = string.Empty;
    public string ReminderDate { get; set; } = string.Empty;
    public string DeadlineDate { get; set; } = string.Empty;
    public string Priority { get; set; } = "3";
    public string Period { get; set; } = string.Empty;
    public string Notes { get; set; } = string.Empty;
    public string UpdatedAt { get; set; } = string.Empty;
}
