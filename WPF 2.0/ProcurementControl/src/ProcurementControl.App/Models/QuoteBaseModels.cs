using System.Globalization;

namespace ProcurementControl.Models;

/// <summary>
/// Строка вкладки «Наши квоты» — выборка из quote_history + quote_batches
/// (аналог строки результата Get-QuoteHistory из QuoteHistory.ps1).
/// </summary>
public sealed class QuoteHistoryRow
{
    public long Id { get; init; }

    public string QuoteDate { get; init; } = string.Empty;

    public string RfqValue { get; init; } = string.Empty;

    public string PN { get; init; } = string.Empty;

    public string Supplier { get; init; } = string.Empty;

    public double? UnitPrice { get; init; }

    public string LeadTime { get; init; } = string.Empty;

    public string LeadTimeTotal { get; init; } = string.Empty;

    public string Mfg { get; init; } = string.Empty;

    public bool IsWinner { get; init; }

    public string WinnerReason { get; init; } = string.Empty;

    public string Warning { get; init; } = string.Empty;

    public string RfqPath { get; init; } = string.Empty;

    public string Priority { get; init; } = string.Empty;

    /// <summary>Цена в формате оригинального Format-Price (N6).</summary>
    public string PriceText => UnitPrice is null
        ? string.Empty
        : UnitPrice.Value.ToString("N6", CultureInfo.GetCultureInfo("ru-RU"));

    /// <summary>Колонка «Побед.» в сетке (аналог Convert-DbBool → 'Да').</summary>
    public string WinnerMark => IsWinner ? "Да" : string.Empty;

    /// <summary>Признак для подсветки ячейки «Предупреждение» (аналог проверки в Refresh-QuoteBase).</summary>
    public bool HasWarning => !string.IsNullOrWhiteSpace(Warning);
}

/// <summary>
/// Строка вкладки «Globalist» — выборка из globalist_quotes
/// (аналог строки результата Get-GlobalistQuotes из QuoteHistory.ps1).
/// </summary>
public sealed class GlobalistRow
{
    public long Id { get; init; }

    public string ImportedAt { get; init; } = string.Empty;

    public string Factory { get; init; } = string.Empty;

    public string PN { get; init; } = string.Empty;

    public string Comment { get; init; } = string.Empty;

    public string PiNumber { get; init; } = string.Empty;

    public string Replacement { get; init; } = string.Empty;

    public string ChineseRemark { get; init; } = string.Empty;

    public string Package { get; init; } = string.Empty;

    public string Brand { get; init; } = string.Empty;

    public string Datacode { get; init; } = string.Empty;

    public string Moq { get; init; } = string.Empty;

    public string Qty { get; init; } = string.Empty;

    public string Stock { get; init; } = string.Empty;

    public string NeedSpq { get; init; } = string.Empty;

    public string Spq { get; init; } = string.Empty;

    public double? UnitPrice { get; init; }

    public double? TotalAmount { get; init; }

    public string LeadTime { get; init; } = string.Empty;

    public string Weight { get; init; } = string.Empty;

    public string Target { get; init; } = string.Empty;

    public string SupplierQuoteId { get; init; } = string.Empty;

    public string SheetName { get; init; } = string.Empty;

    /// <summary>Цена в формате Format-Price (N6), пусто при отсутствии.</summary>
    public string UnitPriceText => UnitPrice is null
        ? string.Empty
        : UnitPrice.Value.ToString("N6", CultureInfo.GetCultureInfo("ru-RU"));

    /// <summary>Сумма в формате Format-Price (N6), пусто при отсутствии.</summary>
    public string TotalAmountText => TotalAmount is null
        ? string.Empty
        : TotalAmount.Value.ToString("N6", CultureInfo.GetCultureInfo("ru-RU"));
}
