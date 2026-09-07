namespace ProcurementControl.Models;

/// <summary>
/// Данные для создания сделки. Повторяет параметры New-PurchaseDeal из
/// PurchaseStore.ps1, но не зависит от PowerShell-приложения.
/// </summary>
public sealed class PurchaseDealDraft
{
    public string DealNumber { get; init; } = string.Empty;

    public string Title { get; init; } = string.Empty;

    public string Comment { get; init; } = string.Empty;

    public string Client { get; init; } = string.Empty;

    public string Status { get; init; } = "RFQ";
}

/// <summary>
/// Поля редактирования существующей сделки. Null в необязательных полях
/// означает «не менять», пустая строка — «очистить значение».
/// </summary>
public sealed class PurchaseDealUpdate
{
    public string Client { get; init; } = string.Empty;

    public string Status { get; init; } = "RFQ";

    public string Comment { get; init; } = string.Empty;

    public string? DealNumber { get; init; }

    public string? BoardCount { get; init; }

    public string? Period { get; init; }

    public string? Priority { get; init; }

    public string? TrackingStatus { get; init; }

    public string? Executor { get; init; }

    public string? ReminderDate { get; init; }

    public string? AssemblyLocation { get; init; }
}

/// <summary>Полная редактируемая карточка поставщика из deal_suppliers.</summary>
public sealed class PurchaseSupplierUpdate
{
    public bool InvoiceReceived { get; init; }
    public bool InvoiceConfirmed { get; init; }
    public bool SupplierOrderCreated { get; init; }
    public bool ErpSupplierSent { get; init; }
    public bool ErpRogerSent { get; init; }
    public string PiAmountUsd { get; init; } = string.Empty;
    public string PiAmountCny { get; init; } = string.Empty;
    public string PiAmountRub { get; init; } = string.Empty;
    public string PaidAmount { get; init; } = string.Empty;
    public string DeliveryWeeks { get; init; } = string.Empty;
    public bool PaymentSubmitted { get; init; }
    public bool Paid { get; init; }
    public string InvoiceConfirmedDate { get; init; } = string.Empty;
    public string ComponentsReceiptDate { get; init; } = string.Empty;
    public string ActualReceiptDate { get; init; } = string.Empty;
    public string Comment { get; init; } = string.Empty;
}

/// <summary>Изменение фактической даты поступления одного поставщика из карточки сделки.</summary>
public sealed record SupplierReceiptDateUpdate(long SupplierId, DateTime? ActualReceiptDate);
