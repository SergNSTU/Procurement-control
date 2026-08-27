using System.Globalization;

namespace ProcurementControl.Models;

/// <summary>Фильтры списка сделок — коды из оригинального $cmbDealFilter.</summary>
public enum DealFilter
{
    Active,
    Tracking,
    China,
    Purchase,
    Pause,
    WaitingReceipt,
    Done,
    Archived,
    AllRecords,
}

/// <summary>Русские подписи фильтров (пункты комбобокса оригинала).</summary>
public static class DealFilterLabels
{
    public static string Label(this DealFilter filter) => filter switch
    {
        DealFilter.Active => "Активные",
        DealFilter.Tracking => "На отслеживании",
        DealFilter.China => "Китай",
        DealFilter.Purchase => "Закупка",
        DealFilter.Pause => "Пауза",
        DealFilter.WaitingReceipt => "Ожидается поступление",
        DealFilter.Done => "Готово",
        DealFilter.Archived => "Архив",
        DealFilter.AllRecords => "Все включая архив",
        _ => filter.ToString()
    };
}

/// <summary>Фон строки сделки — перенос правил раскраски $dealsGrid из оригинала.</summary>
public enum DealRowState
{
    Danger,
    Pause,
    Archived,
    Success,
}

/// <summary>
/// Строка списка «Контроль закупки» (запрос Get-PurchaseDeals со статистикой
/// поставщиков). Вычисляемые свойства повторяют форматирование оригинала.
/// </summary>
public sealed class PurchaseDealRow
{
    public long Id { get; set; }
    public string DealNumber { get; set; } = string.Empty;
    public string BoardCount { get; set; } = string.Empty;
    public string Client { get; set; } = string.Empty;
    public string Period { get; set; } = string.Empty;
    public string Executor { get; set; } = string.Empty;
    public string ReminderDate { get; set; } = string.Empty;
    public string AssemblyLocation { get; set; } = string.Empty;
    public string Priority { get; set; } = "3";
    public string TrackingStatus { get; set; } = "Ожидание";
    public string Status { get; set; } = "RFQ";
    public int Masks { get; set; }
    public int Archived { get; set; }
    public string? Title { get; set; }
    public string? Comment { get; set; }
    public string UpdatedAt { get; set; } = string.Empty;

    // Агрегаты по поставщикам (подзапрос stats).
    public int SupplierCount { get; set; }
    public int InvoiceCount { get; set; }
    public int PaymentSubmittedCount { get; set; }
    public int PaidCount { get; set; }
    public int DoneCount { get; set; }
    public int InvoiceConfirmedCount { get; set; }
    public int ReceiptCount { get; set; }
    public string CompletionReceiptDate { get; set; } = string.Empty;

    /// <summary>Колонка «Оплачено»: "оплачено/поставщики".</summary>
    public string PaymentText => PaidCount + "/" + SupplierCount;

    /// <summary>Колонка «Готово»: "готово/поставщики".</summary>
    public string DoneText => DoneCount + "/" + SupplierCount;

    /// <summary>Колонка «Маски»: значение 2 показывается пустой строкой.</summary>
    public string MasksText => Masks == 2 ? string.Empty : (Masks == 1 ? "Да" : "Нет");

    public string ReminderDateText => PurchaseFormatting.Date(ReminderDate);

    public string CompletionReceiptText => PurchaseFormatting.Date(CompletionReceiptDate);

    public bool ReminderOverdue
    {
        get
        {
            var date = PurchaseFormatting.ParseDate(ReminderDate);
            return date is not null && date.Value.Date <= DateTime.Today;
        }
    }

    /// <summary>Правила фона строки (оригинал, строки 3812-3823).</summary>
    public DealRowState RowState
    {
        get
        {
            if (Status is "RRFQ" or "PI")
            {
                return DealRowState.Pause;
            }
            if (Archived == 1)
            {
                return DealRowState.Archived;
            }
            if (AssemblyLocation.Trim() == "Китай" && Status == "Заказано")
            {
                return DealRowState.Success;
            }
            if (Status == "Заказано" && SupplierCount > 0 && InvoiceConfirmedCount == SupplierCount)
            {
                return DealRowState.Success;
            }
            return DealRowState.Danger;
        }
    }
}

/// <summary>Поставщик сделки — перенос Get-PurchaseSuppliers с форматированием.</summary>
public sealed class SupplierRow
{
    public long Id { get; set; }
    public long DealId { get; set; }
    public string Supplier { get; set; } = string.Empty;
    public string PiAmountUsd { get; set; } = string.Empty;
    public string PiAmountCny { get; set; } = string.Empty;
    public string PiAmountRub { get; set; } = string.Empty;
    public string PaidAmount { get; set; } = string.Empty;
    public string DeliveryWeeks { get; set; } = string.Empty;
    public bool PaymentSubmitted { get; set; }
    public bool Paid { get; set; }
    public bool InvoiceReceived { get; set; }
    public bool InvoiceConfirmed { get; set; }
    public bool SupplierOrderCreated { get; set; }
    public bool ErpSupplierSent { get; set; }
    public bool ErpRogerSent { get; set; }
    public string InvoiceConfirmedDate { get; set; } = string.Empty;
    public string ComponentsReceiptDate { get; set; } = string.Empty;
    public string ActualReceiptDate { get; set; } = string.Empty;
    public string? Comment { get; set; }
}

/// <summary>Задача для счётчика «Активные» (ключ сделки + статус).</summary>
public sealed class CockpitTaskKey
{
    public long Id { get; set; }
    public string DealNumber { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
}

/// <summary>
/// Форматирование дат и сумм — перенос Format-PurchaseDate /
/// Format-PurchaseAmount / Convert-PurchaseDateText из RrfqEngine.ps1.
/// </summary>
public static class PurchaseFormatting
{
    private static readonly string[] DateFormats =
    {
        "dd.MM.yyyy HH:mm:ss",
        "dd.MM.yyyy HH:mm",
        "dd.MM.yyyy",
        "d.M.yyyy",
        "yyyy-MM-dd HH:mm:ss",
        "yyyy-MM-dd",
        "dd/MM/yyyy",
        "d/M/yyyy",
    };

    public static DateTime? ParseDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var text = value.Trim();
        foreach (var culture in new[] { CultureInfo.GetCultureInfo("ru-RU"), CultureInfo.InvariantCulture })
        {
            if (DateTime.TryParseExact(text, DateFormats, culture, DateTimeStyles.None, out var parsed))
            {
                return parsed;
            }
        }

        // Последний запасной вариант оригинального Convert-PurchaseDateText.
        if (DateTime.TryParse(text, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.None, out var loose))
        {
            return loose;
        }
        return null;
    }

    /// <summary>Аналог Format-PurchaseDate: не парсится — вернуть исходный текст.</summary>
    public static string DateOrRaw(string? value)
    {
        if (value is null)
        {
            return string.Empty;
        }
        var parsed = ParseDate(value);
        return parsed is null ? value.Trim() : parsed.Value.ToString("dd.MM.yyyy");
    }

    public static string Date(string? value)
    {
        var parsed = ParseDate(value);
        return parsed is null ? string.Empty : parsed.Value.ToString("dd.MM.yyyy");
    }

    public static string Amount(double? value)
        => value is null ? string.Empty : value.Value.ToString("#,##0.00", CultureInfo.InvariantCulture);
}
