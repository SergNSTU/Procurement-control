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

    /// <summary>Аналог Get-PurchaseDashboard.</summary>
    public DashboardData GetDashboard()
    {
        using var snapshot = PurchaseSnapshot.Create();
        var repo = new PurchaseRepository(snapshot.Connection);

        var now = DateTime.Now;
        var monthStart = new DateTime(now.Year, now.Month, 1);
        var monthEnd = monthStart.AddMonths(1);
        var trackingStartedAt = GetTrackingStartedAt(repo) ?? now;

        return new DashboardData
        {
            MonthLabel = monthStart.ToString("MMMM yyyy", Ru),
            TrackingStartedAt = trackingStartedAt,
            Assemblies = GetDirection(repo, "deals", monthStart, monthEnd, trackingStartedAt, now),
            Components = GetDirection(repo, "component_deals", monthStart, monthEnd, trackingStartedAt, now),
            History = GetHistory(repo, monthStart, trackingStartedAt)
        };
    }

    /// <summary>Аналог Get-PurchaseOrderTrackingStartedAt.</summary>
    private static DateTime? GetTrackingStartedAt(PurchaseRepository repo)
    {
        var raw = repo.GetSetting("dashboard.orders_tracking_started_at");
        return !string.IsNullOrWhiteSpace(raw) && DateTime.TryParse(raw, out var date) ? date : null;
    }

    /// <summary>Аналог Get-PurchaseDashboardCount.</summary>
    private static int GetCount(PurchaseRepository repo, string table, string dateColumn, DateTime start, DateTime end)
    {
        var orderFilter = dateColumn == "ordered_at" ? " AND IFNULL(ordered_at, '') <> ''" : string.Empty;
        return repo.ScalarInt(
            $"SELECT COUNT(*) FROM {table} WHERE {dateColumn} >= @start AND {dateColumn} < @end{orderFilter}",
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
        DateTime monthStart, DateTime monthEnd,
        DateTime trackingStartedAt, DateTime now)
    {
        var currentCohortStart = trackingStartedAt > monthStart ? trackingStartedAt : monthStart;
        var matureCutoff = now.Date.AddDays(-60);

        var createdThisMonth = GetCount(repo, table, "created_at", monthStart, monthEnd);
        var orderedThisMonth = GetCount(repo, table, "ordered_at", monthStart, monthEnd);

        var currentCohortCreated = repo.ScalarInt(
            $"SELECT COUNT(*) FROM {table} WHERE created_at >= @s AND created_at < @e",
            P(currentCohortStart, monthEnd));
        var currentCohortOrdered = repo.ScalarInt(
            $"SELECT COUNT(*) FROM {table} WHERE created_at >= @s AND created_at < @e AND IFNULL(ordered_at, '') <> ''",
            P(currentCohortStart, monthEnd));
        var matureCohortCreated = repo.ScalarInt(
            $"SELECT COUNT(*) FROM {table} WHERE created_at >= @t AND created_at < @m",
            T(trackingStartedAt, matureCutoff));
        var matureCohortOrdered = repo.ScalarInt(
            $"SELECT COUNT(*) FROM {table} WHERE created_at >= @t AND created_at < @m AND IFNULL(ordered_at, '') <> ''",
            T(trackingStartedAt, matureCutoff));

        return new DirectionMetrics
        {
            AddedThisMonth = createdThisMonth,
            OrderedThisMonth = orderedThisMonth,
            CurrentConversion = GetConversion(currentCohortOrdered, currentCohortCreated),
            MatureConversion = GetConversion(matureCohortOrdered, matureCohortCreated),
            CurrentCohortCreated = currentCohortCreated,
            MatureCohortCreated = matureCohortCreated
        };
    }

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

    private static Dictionary<string, object?> P(DateTime s, DateTime e) => new()
    {
        ["@s"] = s.ToString(DateFormat),
        ["@e"] = e.ToString(DateFormat)
    };

    private static Dictionary<string, object?> T(DateTime t, DateTime m) => new()
    {
        ["@t"] = t.ToString(DateFormat),
        ["@m"] = m.ToString(DateFormat)
    };
}
