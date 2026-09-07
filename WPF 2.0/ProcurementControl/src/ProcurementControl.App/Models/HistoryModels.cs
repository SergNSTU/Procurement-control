namespace ProcurementControl.Models;

/// <summary>
/// Строка журнала действий — выборка из activity_log
/// (аналог строки результата Get-ActivityLog из PurchaseStore.ps1, строки 880-902).
/// </summary>
public sealed class ActivityLogRow
{
    public long Id { get; init; }

    public string CreatedAt { get; init; } = string.Empty;

    public string Action { get; init; } = string.Empty;

    public string Details { get; init; } = string.Empty;

    public string EntityType { get; init; } = string.Empty;

    public long? DealId { get; init; }
}
