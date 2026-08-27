namespace ProcurementControl.Models;

/// <summary>
/// Позиция разбора HTML-файла расчёта Компэла — перенос строк,
/// которые строит Parse-CompelHtml (CompelParser.ps1, строки 262-275).
/// </summary>
public sealed class CompelRow
{
    public int? Number { get; set; }
    public string SourceName { get; set; } = string.Empty;
    public string MatchedPart { get; set; } = string.Empty;
    public string Manufacturer { get; set; } = string.Empty;
    public string Package { get; set; } = string.Empty;
    public int? StockCount { get; set; }
    public string LeadTime { get; set; } = string.Empty;
    public int? Quantity { get; set; }
    public double? UnitPrice { get; set; }
    public string Currency { get; set; } = string.Empty;
    public double? Total { get; set; }
    public string TotalCurrency { get; set; } = string.Empty;

    /// <summary>Нет подобранной суммы — строка подсвечивается жёлтым (строки 2971-2973 оригинала).</summary>
    public bool HasTotal => Total is not null;

    /// <summary>Аналог Format-CompelNumber c 5 знаками (колонка «Цена за шт»).</summary>
    public string UnitPriceText => FormatNumber(UnitPrice, 5);

    /// <summary>Аналог Format-CompelNumber c 2 знаками (колонка «Сумма»).</summary>
    public string TotalText => FormatNumber(Total, 2);

    internal static string FormatNumber(double? value, int digits)
    {
        if (value is null)
        {
            return string.Empty;
        }
        return value.Value.ToString("N" + digits, System.Globalization.CultureInfo.GetCultureInfo("ru-RU"));
    }
}

/// <summary>Сводка разбора — перенос New-CompelSummary (строки 154-215).</summary>
public sealed class CompelSummary
{
    public int TotalRows { get; set; }
    public int PricedRows { get; set; }
    public List<int> MissingRows { get; set; } = new();
    public double SelectedTotalUsd { get; set; }
    public int? LongestLeadDays { get; set; }
    public string LongestLeadTime { get; set; } = string.Empty;
    public double? CheapTotalUsd { get; set; }
    public string CheapLeadTime { get; set; } = string.Empty;
    public double? OptimalTotalUsd { get; set; }
    public string OptimalLeadTime { get; set; } = string.Empty;
}

/// <summary>Результат разбора: позиции + сводка.</summary>
public sealed class CompelParseResult
{
    public List<CompelRow> Rows { get; set; } = new();
    public CompelSummary Summary { get; set; } = new();
}
