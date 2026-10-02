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
    /// Попадание сделки в плитку «Активные» (обновление метрики и список по клику).
    /// reason — почему сделка активна: «активная», «отслеживание» или «Китай».
    /// </summary>
    public static bool IsCockpitActiveDeal(PurchaseDealRow row, out string reason)
    {
        reason = string.Empty;
        if (row.Archived == 1)
        {
            return false;
        }

        var status = row.Status.Trim();
        var location = row.AssemblyLocation.Trim();
        var executor = row.Executor.Trim();
        var isOperationalStage = status is "PO" or "RFQ" or "Закупка";
        if (executor.Length > 0 && isOperationalStage)
        {
            reason = "отслеживание";
            return true;
        }

        if (location == "Китай" && isOperationalStage)
        {
            reason = "Китай";
            return true;
        }

        var isActiveFilterDeal =
            location != "Китай"
            && executor.Length == 0
            && (status is "RFQ" or "PO" or "Закупка" or "В работе"
                || (status == "Заказано"
                    && (row.SupplierCount == 0 || row.InvoiceConfirmedCount < row.SupplierCount)));
        if (isActiveFilterDeal)
        {
            reason = "активная";
            return true;
        }

        return false;
    }

    /// <summary>
    /// Метрики карточек: перенос Update-PurchaseCockpitMetrics и
    /// Get-ActiveCockpitDealCount (дедупликация по ключу сделки).
    /// «Ожидаем поступления» — сделки, где у поставщиков есть план, а факта
    /// ещё нет; «Просрочено» — из них те, чей план уже прошёл.
    /// </summary>
    public static (int Active, int Overdue, int WorkAndPaymentTasks, int ExpectingReceipt) ComputeMetrics(
        IReadOnlyList<PurchaseDealRow> rows,
        IReadOnlyList<CockpitTaskKey> taskKeys,
        IReadOnlyList<SupplierPlanRow> planRows)
    {
        var keys = new HashSet<string>();
        foreach (var row in rows)
        {
            if (!IsCockpitActiveDeal(row, out _))
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
        var expectingDeals = new HashSet<long>();
        var overdueDeals = new HashSet<long>();
        foreach (var plan in planRows)
        {
            if (DoneStatusRegex.IsMatch(plan.Status) ||
                PurchaseFormatting.ParseDate(plan.ComponentsReceiptDate) is null ||
                PurchaseFormatting.ParseDate(plan.ActualReceiptDate) is not null)
            {
                continue;
            }

            expectingDeals.Add(plan.DealId);
            if (PurchaseFormatting.IsReceiptOverdue(plan.ComponentsReceiptDate, plan.ActualReceiptDate, today))
            {
                overdueDeals.Add(plan.DealId);
            }
        }

        return (keys.Count, overdueDeals.Count, taskKeys.Count, expectingDeals.Count);
    }

    /// <summary>
    /// Списки под плитками метрик: клик по плитке показывает, что в неё входит,
    /// двойной клик по строке переходит к сделке. Активные дедуплицируются
    /// по номеру так же, как в ComputeMetrics.
    /// </summary>
    public static (string Title, List<TileLineRow> Lines) BuildTileLines(
        string tileKey,
        IReadOnlyList<PurchaseDealRow> rows,
        IReadOnlyList<CockpitTaskKey> taskKeys,
        IReadOnlyList<SupplierPlanRow> planRows,
        IReadOnlyList<DealReminderRow> dealReminders,
        IReadOnlyList<ManualReminderRow> manualReminders,
        DateTime today)
    {
        switch (tileKey)
        {
            case "active":
            {
                var seen = new HashSet<string>();
                var lines = new List<TileLineRow>();
                foreach (var row in rows)
                {
                    if (!IsCockpitActiveDeal(row, out var reason) ||
                        !seen.Add(DealKey(row.DealNumber, row.Id, "purchase")))
                    {
                        continue;
                    }

                    lines.Add(new TileLineRow
                    {
                        Text = row.DealNumber + " — " + reason + ClientSuffix(row.Client),
                        DealId = row.Id,
                        DealNumber = row.DealNumber,
                    });
                }

                foreach (var task in taskKeys)
                {
                    if (!seen.Add(DealKey(task.DealNumber, task.Id, "task")))
                    {
                        continue;
                    }

                    lines.Add(new TileLineRow
                    {
                        Text = task.DealNumber + " — задача, этап: " + task.Status,
                        DealNumber = task.DealNumber,
                    });
                }

                return ("Активные сделки", lines);
            }

            case "expecting":
            {
                // Одна строка на сделку: у поставщиков есть план, факта ещё нет.
                var lines = new List<TileLineRow>();
                foreach (var group in planRows
                    .Where(plan => !DoneStatusRegex.IsMatch(plan.Status) &&
                                   PurchaseFormatting.ParseDate(plan.ComponentsReceiptDate) is not null &&
                                   PurchaseFormatting.ParseDate(plan.ActualReceiptDate) is null)
                    .GroupBy(plan => plan.DealId))
                {
                    var first = group.First();
                    var details = string.Join(", ", group.Select(plan =>
                    {
                        var planText = PurchaseFormatting.ParseDate(plan.ComponentsReceiptDate) is { } planDate
                            ? planDate.ToString("dd.MM.yyyy")
                            : PurchaseFormatting.DateOrRaw(plan.ComponentsReceiptDate);
                        return plan.Supplier + " (план " + planText + ")";
                    }));

                    lines.Add(new TileLineRow
                    {
                        Text = first.DealNumber + " — ждём: " + details,
                        DealId = first.DealId,
                        DealNumber = first.DealNumber,
                    });
                }

                return ("Ожидаем поступления", lines);
            }

            case "overdue":
            {
                // Одна строка на сделку — чтобы число строк совпадало с плиткой;
                // в строке перечислены поставщики, чей план прошёл без факта.
                var lines = new List<TileLineRow>();
                foreach (var group in planRows
                    .Where(plan => !DoneStatusRegex.IsMatch(plan.Status) &&
                                   PurchaseFormatting.IsReceiptOverdue(
                                       plan.ComponentsReceiptDate, plan.ActualReceiptDate, today))
                    .GroupBy(plan => plan.DealId))
                {
                    var first = group.First();
                    var details = string.Join(", ", group.Select(plan =>
                    {
                        var planText = PurchaseFormatting.ParseDate(plan.ComponentsReceiptDate) is { } planDate
                            ? planDate.ToString("dd.MM.yyyy")
                            : PurchaseFormatting.DateOrRaw(plan.ComponentsReceiptDate);
                        return plan.Supplier + " (план " + planText + ")";
                    }));

                    lines.Add(new TileLineRow
                    {
                        Text = first.DealNumber + " — не поступило: " + details,
                        DealId = first.DealId,
                        DealNumber = first.DealNumber,
                    });
                }

                return ("Просрочено", lines);
            }

            case "tasks":
                return ("Задачи в работе", taskKeys.Select(task => new TileLineRow
                {
                    Text = task.DealNumber + " — этап: " + task.Status,
                    DealNumber = task.DealNumber,
                }).ToList());

            case "reminders":
            {
                var lines = new List<TileLineRow>();
                foreach (var reminder in dealReminders)
                {
                    lines.Add(new TileLineRow
                    {
                        Text = reminder.DealNumber + " — напоминание " +
                               PurchaseFormatting.DateOrRaw(reminder.ReminderDate),
                        DealId = reminder.Id,
                        DealNumber = reminder.DealNumber,
                    });
                }

                foreach (var reminder in manualReminders)
                {
                    lines.Add(new TileLineRow
                    {
                        Text = reminder.Title + (reminder.DueDate.Length > 0
                            ? " — " + PurchaseFormatting.DateOrRaw(reminder.DueDate)
                            : string.Empty),
                        DealId = reminder.DealId,
                    });
                }

                return ("Напоминания", lines);
            }

            default:
                return (string.Empty, new List<TileLineRow>());
        }
    }

    private static string ClientSuffix(string client)
        => string.IsNullOrWhiteSpace(client) ? string.Empty : " · " + client.Trim();

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
