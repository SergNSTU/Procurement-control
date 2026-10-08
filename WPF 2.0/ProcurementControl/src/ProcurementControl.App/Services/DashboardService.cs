using System.Globalization;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Расчёт метрик дашборда. Переносит app/modules/PurchaseDashboard.ps1
/// (Get-PurchaseDashboard, Get-PurchaseDashboardDirection, Get-PurchaseDashboardHistory).
/// Работает только через read-only снапшот — исходная БД не изменяется.
/// Имена таблиц и колонок захардкожены (whitelist) — SQL-инъекции исключены.
/// </summary>
public sealed class DashboardService
{
    private const string DateFormat = "yyyy-MM-dd HH:mm:ss";
    private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");

    public IReadOnlyList<int> GetAvailableYears()
    {
        using var snapshot = PurchaseSnapshot.Create();
        var repo = new PurchaseRepository(snapshot.Connection);
        return repo.GetDashboardYears().Append(DateTime.Now.Year).Distinct().OrderByDescending(year => year).ToList();
    }

    /// <summary>Метрики для выбранного календарного месяца или квартала.</summary>
    public DashboardData GetDashboard(DashboardPeriod period)
    {
        using var snapshot = PurchaseSnapshot.Create();
        var repo = new PurchaseRepository(snapshot.Connection);

        var now = DateTime.Now;
        var trackingStartedAt = GetTrackingStartedAt(repo) ?? now;

        return new DashboardData
        {
            PeriodLabel = FormatPeriodLabel(period),
            PeriodUnitLabel = period.UnitLabel,
            TrackingStartedAt = trackingStartedAt,
            Assemblies = GetDirection(repo, "deals", period.Start, period.End, trackingStartedAt, now,
                SumOrderedDealAmounts(repo, period.Start, period.End)),
            Components = GetDirection(repo, "component_deals", period.Start, period.End, trackingStartedAt, now,
                repo.GetOrderedComponentAmounts(period.Start, period.End).Sum(MoneyText.ParseUsd)),
            History = GetHistory(repo, period.End.AddMonths(-1), trackingStartedAt),
            MonthlyAmounts = GetMonthlyAmounts(repo, period.Year, trackingStartedAt)
        };
    }

    /// <summary>Аналог Get-PurchaseOrderTrackingStartedAt.</summary>
    private static DateTime? GetTrackingStartedAt(PurchaseRepository repo)
    {
        var raw = repo.GetSetting("dashboard.orders_tracking_started_at");
        return !string.IsNullOrWhiteSpace(raw) && DateTime.TryParse(raw, out var date) ? date : null;
    }

    /// <summary>
    /// Сборки с местом сборки «Китай» исключены из всех показателей дашборда
    /// (просьба 2026-10-05); для component_deals фильтр пустой.
    /// </summary>
    private static string NotChinaFilter(string table)
        => table == "deals" ? " AND IFNULL(TRIM(assembly_location), '') <> 'Китай'" : string.Empty;

    /// <summary>Аналог Get-PurchaseDashboardCount.</summary>
    private static int GetCount(PurchaseRepository repo, string table, string dateColumn, DateTime start, DateTime end)
    {
        var orderFilter = dateColumn == "ordered_at" ? " AND IFNULL(ordered_at, '') <> ''" : string.Empty;
        return repo.ScalarInt(
            $"SELECT COUNT(*) FROM {table} WHERE {dateColumn} >= @start AND {dateColumn} < @end{orderFilter}{NotChinaFilter(table)}",
            new Dictionary<string, object?>
            {
                ["@start"] = start.ToString(DateFormat),
                ["@end"] = end.ToString(DateFormat)
            });
    }

    /// <summary>Аналог Get-PurchaseDashboardConversion.</summary>
    private static double? GetConversion(int ordered, int created)
    {
        if (created <= 0)
        {
            return null;
        }
        return Math.Round(100.0 * ordered / created, 1);
    }

    /// <summary>Аналог Get-PurchaseDashboardDirection.</summary>
    private static DirectionMetrics GetDirection(
        PurchaseRepository repo, string table,
        DateTime periodStart, DateTime periodEnd,
        DateTime trackingStartedAt, DateTime now, decimal orderAmountUsd)
    {
        var currentCohortStart = trackingStartedAt > periodStart ? trackingStartedAt : periodStart;
        var matureCutoff = now.Date.AddDays(-60);

        var createdThisMonth = GetCount(repo, table, "created_at", periodStart, periodEnd);
        var orderedThisMonth = GetCount(repo, table, "ordered_at", periodStart, periodEnd);

        var notChina = NotChinaFilter(table);
        var currentCohortCreated = repo.ScalarInt(
            $"SELECT COUNT(*) FROM {table} WHERE created_at >= @s AND created_at < @e{notChina}",
            P(currentCohortStart, periodEnd));
        var currentCohortOrdered = repo.ScalarInt(
            $"SELECT COUNT(*) FROM {table} WHERE created_at >= @s AND created_at < @e AND IFNULL(ordered_at, '') <> ''{notChina}",
            P(currentCohortStart, periodEnd));
        var matureCohortEnd = matureCutoff < periodEnd ? matureCutoff : periodEnd;
        var matureCohortCreated = repo.ScalarInt(
            $"SELECT COUNT(*) FROM {table} WHERE created_at >= @s AND created_at < @e{notChina}",
            P(currentCohortStart, matureCohortEnd));
        var matureCohortOrdered = repo.ScalarInt(
            $"SELECT COUNT(*) FROM {table} WHERE created_at >= @s AND created_at < @e AND IFNULL(ordered_at, '') <> ''{notChina}",
            P(currentCohortStart, matureCohortEnd));

        return new DirectionMetrics
        {
            AddedThisMonth = createdThisMonth,
            OrderedThisMonth = orderedThisMonth,
            CurrentConversion = GetConversion(currentCohortOrdered, currentCohortCreated),
            MatureConversion = GetConversion(matureCohortOrdered, matureCohortCreated),
            CurrentCohortCreated = currentCohortCreated,
            MatureCohortCreated = matureCohortCreated,
            OrderAmountUsd = orderAmountUsd
        };
    }

    /// <summary>Сначала суммируем поставщиков в рамках каждой сделки, затем сделки периода.</summary>
    private static decimal SumOrderedDealAmounts(PurchaseRepository repo, DateTime start, DateTime end)
        => repo.GetOrderedDealSupplierAmounts(start, end)
            .GroupBy(row => row.DealId)
            .Sum(deal => deal.Sum(row => MoneyText.ParseUsd(row.AmountUsd)));

    /// <summary>Аналог Get-PurchaseDashboardHistory (6 месяцев).</summary>
    private static IReadOnlyList<MonthHistoryRow> GetHistory(
        PurchaseRepository repo, DateTime monthStart, DateTime trackingStartedAt)
    {
        var rows = new List<MonthHistoryRow>();
        for (var offset = 5; offset >= 0; offset--)
        {
            var start = monthStart.AddMonths(-offset);
            var end = start.AddMonths(1);
            var hasOrderData = end > trackingStartedAt;

            rows.Add(new MonthHistoryRow
            {
                MonthStart = start,
                Label = start.ToString("MMM yy", Ru),
                AssemblyAdded = GetCount(repo, "deals", "created_at", start, end),
                ComponentAdded = GetCount(repo, "component_deals", "created_at", start, end),
                AssemblyOrdered = hasOrderData ? GetCount(repo, "deals", "ordered_at", start, end) : null,
                ComponentOrdered = hasOrderData ? GetCount(repo, "component_deals", "ordered_at", start, end) : null,
                HasOrderData = hasOrderData
            });
        }
        return rows;
    }

    /// <summary>Суммы заказов по месяцам выбранного года — данные для компактных графиков.</summary>
    private static IReadOnlyList<DashboardMonthlyAmountRow> GetMonthlyAmounts(
        PurchaseRepository repo, int year, DateTime trackingStartedAt)
    {
        var rows = new List<DashboardMonthlyAmountRow>();
        for (var month = 1; month <= 12; month++)
        {
            var start = new DateTime(year, month, 1);
            var end = start.AddMonths(1);
            var hasOrderData = end > trackingStartedAt;
            rows.Add(new DashboardMonthlyAmountRow
            {
                MonthStart = start,
                Month = month,
                Label = ShortMonths[month - 1],
                HasOrderData = hasOrderData,
                AssemblyAmountUsd = hasOrderData ? SumOrderedDealAmounts(repo, start, end) : 0m,
                ComponentAmountUsd = hasOrderData ? repo.GetOrderedComponentAmounts(start, end).Sum(MoneyText.ParseUsd) : 0m
            });
        }
        return rows;
    }

    private static readonly string[] ShortMonths =
        { "янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек" };

    private static string FormatPeriodLabel(DashboardPeriod period) => period.Mode == DashboardPeriodMode.Month
        ? period.Start.ToString("MMMM yyyy", Ru)
        : $"{period.Value} квартал {period.Year}";

    private static Dictionary<string, object?> P(DateTime s, DateTime e) => new()
    {
        ["@s"] = s.ToString(DateFormat),
        ["@e"] = e.ToString(DateFormat)
    };

}
