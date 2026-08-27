namespace ProcurementControl.Models;

/// <summary>Метрики по одному направлению (сборки / компоненты).</summary>
public sealed class DirectionMetrics
{
    public int AddedThisMonth { get; init; }
    public int OrderedThisMonth { get; init; }
    public double? CurrentConversion { get; init; }
    public double? MatureConversion { get; init; }
    public int CurrentCohortCreated { get; init; }
    public int MatureCohortCreated { get; init; }
}

/// <summary>Строка месячной истории дашборда.</summary>
public sealed class MonthHistoryRow
{
    public DateTime MonthStart { get; init; }
    public string Label { get; init; } = string.Empty;
    public int AssemblyAdded { get; init; }
    public int ComponentAdded { get; init; }
    public int? AssemblyOrdered { get; init; }
    public int? ComponentOrdered { get; init; }
    public bool HasOrderData { get; init; }
}

/// <summary>Полный набор данных дашборда (аналог Get-PurchaseDashboard).</summary>
public sealed class DashboardData
{
    public string MonthLabel { get; init; } = string.Empty;
    public DateTime? TrackingStartedAt { get; init; }
    public DirectionMetrics Assemblies { get; init; } = new();
    public DirectionMetrics Components { get; init; } = new();
    public IReadOnlyList<MonthHistoryRow> History { get; init; } = Array.Empty<MonthHistoryRow>();
}
