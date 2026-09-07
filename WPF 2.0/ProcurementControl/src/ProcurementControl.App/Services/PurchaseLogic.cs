using System.Text.RegularExpressions;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Чистая логика раздела «Контроль закупки»: фильтры, поиск и метрики.
/// Перенос условий WHERE из Get-PurchaseDeals, Update-PurchaseCockpitMetrics
/// и Get-ActiveCockpitDealCount из оригинала.
/// </summary>
public static class PurchaseLogic
{
    private static readonly Regex DoneStatusRegex = new("Done|Completed", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>Перенос Normalize-Key из RrfqEngine.ps1: NBSP->пробел, регистр вверх, сжатие пробелов.</summary>
    public static string NormalizeKey(string? value)
    {
        var text = (value ?? string.Empty).Replace('\u00A0', ' ').Trim().ToUpperInvariant();
        text = Regex.Replace(text, @"[\r\n\t]+", " ");
        text = Regex.Replace(text, @"\s+", " ");
        return text;
    }

    /// <summary>Перенос условий WHERE Get-PurchaseDeals для каждого кода фильтра.</summary>
    public static List<PurchaseDealRow> FilterDeals(IReadOnlyList<PurchaseDealRow> rows, DealFilter filter)
    {
        var result = new List<PurchaseDealRow>();
        foreach (var row in rows)
        {
            bool keep;
            switch (filter)
            {
                case DealFilter.Archived:
                    keep = row.Archived == 1;
                    break;
                case DealFilter.AllRecords:
                    keep = true;
                    break;
                case DealFilter.China:
                    keep = row.AssemblyLocation.Trim() == "Китай";
                    break;
                default:
                    keep = row.Archived == 0;
                    if (keep && filter == DealFilter.Active)
                    {
                        var status = row.Status.Trim();
                        keep = row.Executor.Trim().Length == 0
                            && (status is "RFQ" or "PO" or "Закупка" or "В работе"
                                || (status == "Заказано"
                                    && (row.SupplierCount == 0 || row.InvoiceConfirmedCount < row.SupplierCount)));
                    }
                    break;
            }

            // Правило «не Китай» оригинала: действует для всех видов кроме
            // China, Tracking и AllRecords.
            if (keep && filter is not (DealFilter.China or DealFilter.Tracking or DealFilter.AllRecords))
            {
                keep = row.AssemblyLocation.Trim() != "Китай";
            }

            if (keep)
            {
                keep = filter switch
                {
                    DealFilter.Tracking => row.Executor.Trim().Length > 0,
                    DealFilter.Pause => row.Status.Trim() is "RRFQ" or "PI",
                    DealFilter.Done => row.Status.Trim() == "Заказано"
                        && row.SupplierCount > 0
                        && row.InvoiceConfirmedCount == row.SupplierCount,
                    _ => true
                };
            }

            if (keep)
            {
                result.Add(row);
            }
        }
        return result;
    }

    /// <summary>Аналог LIKE-поиска оригинала по номеру, названию и остальным полям строки.</summary>
    public static List<PurchaseDealRow> SearchDeals(IReadOnlyList<PurchaseDealRow> rows, string search)
    {
        var needle = search.Trim();
        if (needle.Length == 0)
        {
            return new List<PurchaseDealRow>(rows);
        }

        var result = new List<PurchaseDealRow>();
        foreach (var row in rows)
        {
            if (Contains(row.DealNumber, needle)
                || Contains(row.Title ?? string.Empty, needle)
                || Contains(row.Client, needle)
                || Contains(row.Status, needle)
                || Contains(row.TrackingStatus, needle)
                || Contains(row.Period, needle)
                || Contains(row.Executor, needle)
                || Contains(row.AssemblyLocation, needle)
                || Contains(row.BoardCount, needle))
            {
                result.Add(row);
            }
        }
        return result;
    }

    private static bool Contains(string value, string needle)
        => value.Contains(needle, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Метрики карточек: перенос Update-PurchaseCockpitMetrics и
    /// Get-ActiveCockpitDealCount (дедупликация по ключу сделки).
    /// </summary>
    public static (int Active, int Overdue, int WorkAndPaymentTasks, int Attention) ComputeMetrics(
        IReadOnlyList<PurchaseDealRow> rows,
        IReadOnlyList<CockpitTaskKey> taskKeys)
    {
        var keys = new HashSet<string>();
        foreach (var row in rows)
        {
            if (row.Archived == 1)
            {
                continue;
            }

            var status = row.Status.Trim();
            var location = row.AssemblyLocation.Trim();
            var executor = row.Executor.Trim();
            var isOperationalStage = status is "PO" or "RFQ" or "Закупка";
            var isActiveFilterDeal =
                location != "Китай"
                && executor.Length == 0
                && (status is "RFQ" or "PO" or "Закупка" or "В работе"
                    || (status == "Заказано"
                        && (row.SupplierCount == 0 || row.InvoiceConfirmedCount < row.SupplierCount)));
            var isTrackingFilterDeal = executor.Length > 0 && isOperationalStage;
            var isChinaFilterDeal = location == "Китай" && isOperationalStage;

            if (!isActiveFilterDeal && !isTrackingFilterDeal && !isChinaFilterDeal)
            {
                continue;
            }

            keys.Add(DealKey(row.DealNumber, row.Id, "purchase"));
        }

        foreach (var task in taskKeys)
        {
            keys.Add(DealKey(task.DealNumber, task.Id, "task"));
        }

        var today = DateTime.Today;
        var overdue = 0;
        var attention = 0;
        foreach (var row in rows)
        {
            var isDone = DoneStatusRegex.IsMatch(row.Status);
            var receipt = PurchaseFormatting.ParseDate(row.CompletionReceiptDate);
            if (receipt is not null && row.Archived != 1 && !isDone)
            {
                if (receipt.Value.Date < today)
                {
                    overdue++;
                }
            }
            if (row.Archived != 1 && row.SupplierCount > 0 && row.InvoiceConfirmedCount < row.SupplierCount)
            {
                attention++;
            }
        }

        return (keys.Count, overdue, taskKeys.Count, attention);
    }

    private static string DealKey(string dealNumber, long id, string prefix)
    {
        var number = dealNumber.Trim();
        return number.Length == 0 ? prefix + ":" + id : "deal:" + NormalizeKey(number);
    }

    // ------------------------------------------------------------------
    // Раздел «Напоминания» — перенос Get-PurchaseActionItems
    // (PurchaseStore.ps1, строки 1221-1338).
    // ------------------------------------------------------------------

    private static readonly Regex ErpNotRequiredRegex = new(
        @"^(компэл|compel)(_\d+)?$|^(чид|chid)(_\d+)?$|^(промэлектроника|promelec|promelektronika)(_\d+)?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Перенос Test-PurchaseSupplierErpNotRequired: поставщики из списка не требуют
    /// отправки ERP-документов. Имя приводится к нижнему регистру без пробелов.
    /// </summary>
    public static bool IsErpNotRequired(string supplier)
    {
        var name = Regex.Replace(supplier.Trim().ToLowerInvariant(), @"\s+", string.Empty);
        return name.Length > 0 && ErpNotRequiredRegex.IsMatch(name);
    }

    /// <summary>
    /// Перенос тела Get-PurchaseActionItems: строит авто-пункты по поставщикам,
    /// датные пункты по напоминаниям сделок и ручные напоминания.
    /// </summary>
    public static List<ActionItemRow> BuildActionItems(
        IReadOnlyList<SupplierActionRow> suppliers,
        IReadOnlyList<DealReminderRow> dealReminders,
        IReadOnlyList<ManualReminderRow> manuals,
        DateTime today,
        IReadOnlySet<string>? suppressedAutomaticKeys = null)
    {
        var items = new List<ActionItemRow>();

        foreach (var row in suppliers)
        {
            var prefix = row.DealNumber + " / " + row.Supplier;

            if (string.IsNullOrWhiteSpace(row.PiAmountUsd))
            {
                    AddAutoItem(items, row, "Warn", prefix + ": не указана сумма PI в USD", "", suppressedAutomaticKeys);
            }

            if (!row.InvoiceReceived)
            {
                AddAutoItem(items, row, "Warn", prefix + ": нет PI", "", suppressedAutomaticKeys);
                continue;
            }

            if (!row.PaymentSubmitted)
            {
                AddAutoItem(items, row, "Info", prefix + ": PI есть, но не подан в оплату", "", suppressedAutomaticKeys);
            }
            else if (!row.Paid)
            {
                AddAutoItem(items, row, "Attention", prefix + ": подан в оплату, но не оплачен", "", suppressedAutomaticKeys);
            }

            if (string.IsNullOrWhiteSpace(row.InvoiceConfirmedDate))
            {
                AddAutoItem(items, row, "Warn", prefix + ": нет даты подтверждения инвойса", "", suppressedAutomaticKeys);
            }

            if (!IsErpNotRequired(row.Supplier))
            {
                if (!row.ErpSupplierSent)
                {
                    AddAutoItem(items, row, "Info", prefix + ": Не отправлен ERP поставщику", "", suppressedAutomaticKeys);
                }
                if (!row.ErpRogerSent)
                {
                    AddAutoItem(items, row, "Info", prefix + ": Не отправлен ERP заказ", "", suppressedAutomaticKeys);
                }
            }

            // Пункты поступления: только если факт пуст, а плановая дата есть.
            var actualReceiptDate = PurchaseFormatting.ParseDate(row.ActualReceiptDate);
            var receiptDate = PurchaseFormatting.ParseDate(row.ComponentsReceiptDate);
            if (actualReceiptDate is null && receiptDate is not null)
            {
                var days = (int)(receiptDate.Value.Date - today).TotalDays;
                if (days < 0)
                {
                    AddAutoItem(items, row, "Danger", prefix + ": поступление просрочено", receiptDate.Value.ToString("dd.MM.yyyy"), suppressedAutomaticKeys);
                }
                else if (days <= 7)
                {
                    AddAutoItem(items, row, "Attention", prefix + ": скоро поступление", receiptDate.Value.ToString("dd.MM.yyyy"), suppressedAutomaticKeys);
                }
            }
        }

        foreach (var row in dealReminders)
        {
            var due = PurchaseFormatting.DateOrRaw(row.ReminderDate);
            items.Add(new ActionItemRow
            {
                DealId = row.Id,
                SupplierId = 0,
                ReminderId = 0,
                Severity = DueSeverity(due, today),
                Title = row.DealNumber,
                DueDate = due,
                Source = "deal",
            });
        }

        foreach (var row in manuals)
        {
            var due = PurchaseFormatting.DateOrRaw(row.DueDate);
            items.Add(new ActionItemRow
            {
                DealId = row.DealId,
                SupplierId = row.SupplierId,
                ReminderId = row.Id,
                Severity = DueSeverity(due, today),
                Title = row.Title,
                DueDate = due,
                Source = row.Source,
            });
        }

        return items;
    }

    private static void AddAutoItem(List<ActionItemRow> items, SupplierActionRow row, string severity, string title, string dueDate, IReadOnlySet<string>? suppressedKeys)
    {
        var key = AutomaticReminderKey(row, title);
        if (suppressedKeys?.Contains(key) == true) return;
        items.Add(AutoItem(row, severity, title, dueDate, key));
    }

    public static string AutomaticReminderKey(long dealId, long supplierId, string title)
        => $"auto:{dealId}:{supplierId}:{title}";

    private static string AutomaticReminderKey(SupplierActionRow row, string title)
        => AutomaticReminderKey(row.DealId, row.SupplierId, title);

    private static ActionItemRow AutoItem(SupplierActionRow row, string severity, string title, string dueDate, string key)
        => new()
        {
            DealId = row.DealId,
            SupplierId = row.SupplierId,
            Severity = severity,
            Title = title,
            DueDate = dueDate,
            Source = "auto",
            SuppressionKey = key,
        };

    /// <summary>Важность датного пункта: не парсится -> Warn, прошло -> Danger, <=7 дней -> Attention.</summary>
    private static string DueSeverity(string due, DateTime today)
    {
        var date = PurchaseFormatting.ParseDate(due);
        if (date is null)
        {
            return "Warn";
        }
        if (date.Value.Date < today)
        {
            return "Danger";
        }
        return (date.Value.Date - today).TotalDays <= 7 ? "Attention" : "Info";
    }

    /// <summary>Поиск по заголовку и тексту сделки без учёта регистра (строки 4294-4300 оригинала).</summary>
    public static List<ActionItemRow> SearchItems(IReadOnlyList<ActionItemRow> items, string search)
    {
        var needle = search.Trim();
        if (needle.Length == 0)
        {
            return new List<ActionItemRow>(items);
        }

        var result = new List<ActionItemRow>();
        foreach (var item in items)
        {
            if (item.Title.Contains(needle, StringComparison.OrdinalIgnoreCase)
                || item.DealId.ToString().Contains(needle, StringComparison.OrdinalIgnoreCase))
            {
                result.Add(item);
            }
        }
        return result;
    }
}
