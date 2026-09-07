namespace ProcurementControl.Models;

public enum DashboardPeriodMode { Month, Quarter }

public sealed record DashboardPeriodModeOption(DashboardPeriodMode Value, string Label);

public sealed record DashboardPeriodOption(int Value, string Label);

/// <summary>Выбранный на дашборде календарный интервал.</summary>
public sealed record DashboardPeriod(int Year, DashboardPeriodMode Mode, int Value)
{
    public DateTime Start => Mode == DashboardPeriodMode.Month
        ? new DateTime(Year, Value, 1)
        : new DateTime(Year, (Value - 1) * 3 + 1, 1);

    public DateTime End => Mode == DashboardPeriodMode.Month
        ? Start.AddMonths(1)
        : Start.AddMonths(3);

    public string UnitLabel => Mode == DashboardPeriodMode.Month ? "месяц" : "квартал";
}

/// <summary>Метрики по одному направлению (сборки / компоненты).</summary>
public sealed class DirectionMetrics
{
    public int AddedThisMonth { get; init; }
    public int OrderedThisMonth { get; init; }
    public double? CurrentConversion { get; init; }
    public double? MatureConversion { get; init; }
    public int CurrentCohortCreated { get; init; }
    public int MatureCohortCreated { get; init; }
    public decimal OrderAmountUsd { get; init; }
}

/// <summary>Сумма поставщика, привязанная к конкретной сделке.</summary>
public sealed record DashboardSupplierAmountRow(long DealId, string AmountUsd);

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
    public string PeriodLabel { get; init; } = string.Empty;
    public string PeriodUnitLabel { get; init; } = string.Empty;
    public DateTime? TrackingStartedAt { get; init; }
    public DirectionMetrics Assemblies { get; init; } = new();
    public DirectionMetrics Components { get; init; } = new();
    public IReadOnlyList<MonthHistoryRow> History { get; init; } = Array.Empty<MonthHistoryRow>();
}
