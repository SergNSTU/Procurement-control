namespace ProcurementControl.Models;

/// <summary>
/// Пункт списка «Напоминания» — перенос элементов, которые строит
/// Get-PurchaseActionItems (PurchaseStore.ps1). Severity определяет цвет строки.
/// </summary>
public sealed class ActionItemRow
{
    public long DealId { get; set; }
    public long SupplierId { get; set; }
    public long ReminderId { get; set; }

    /// <summary>Danger / Attention / Warn / Info (фон строки, строки 4315-4321 оригинала).</summary>
    public string Severity { get; set; } = "Info";

    public string Title { get; set; } = string.Empty;
    public string DueDate { get; set; } = string.Empty;

    /// <summary>auto / deal / значение source из таблицы reminders.</summary>
    public string Source { get; set; } = string.Empty;
}

/// <summary>Строка запроса поставщиков для авто-пунктов (строки 1232-1252 оригинала).</summary>
public sealed class SupplierActionRow
{
    public long DealId { get; set; }
    public string DealNumber { get; set; } = string.Empty;
    public long SupplierId { get; set; }
    public string Supplier { get; set; } = string.Empty;
    public bool InvoiceReceived { get; set; }
    public bool InvoiceConfirmed { get; set; }
    public bool SupplierOrderCreated { get; set; }
    public bool ErpSupplierSent { get; set; }
    public bool ErpRogerSent { get; set; }
    public bool PaymentSubmitted { get; set; }
    public bool Paid { get; set; }
    public string InvoiceConfirmedDate { get; set; } = string.Empty;
    public string ComponentsReceiptDate { get; set; } = string.Empty;
    public string ActualReceiptDate { get; set; } = string.Empty;
    public string DeliveryWeeks { get; set; } = string.Empty;
}

/// <summary>Напоминание сделки (поле reminder_date) для датных пунктов.</summary>
public sealed class DealReminderRow
{
    public long Id { get; set; }
    public string DealNumber { get; set; } = string.Empty;
    public string ReminderDate { get; set; } = string.Empty;
}

/// <summary>Ручное напоминание из таблицы reminders.</summary>
public sealed class ManualReminderRow
{
    public long Id { get; set; }
    public long DealId { get; set; }
    public long SupplierId { get; set; }
    public string Title { get; set; } = string.Empty;
    public string DueDate { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;
}
