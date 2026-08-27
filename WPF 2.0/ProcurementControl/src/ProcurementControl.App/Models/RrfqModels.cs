using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;

namespace ProcurementControl.Models;

/// <summary>Поставщик из конфигурации (аналог элементов $script:Config.Suppliers).</summary>
public sealed class SupplierInfo
{
    public string Name { get; set; } = string.Empty;

    public List<string> Aliases { get; set; } = new();
}

/// <summary>Конфигурация раздела Сравнение RRFQ (аналог Get-AppConfig из AppConfig.ps1).</summary>
public sealed class RrfqConfig
{
    public List<SupplierInfo> Suppliers { get; set; } = new();

    public string OtherSupplier { get; set; } = "Др.";

    public double VatDivisor { get; set; } = 1.22;
}

/// <summary>Строка списка файлов поставщиков (аналог строки $supplierGrid).</summary>
public partial class SupplierFileEntry : ObservableObject
{
    [ObservableProperty]
    private string _path = string.Empty;

    [ObservableProperty]
    private string _supplier = string.Empty;
}

/// <summary>
/// Слепок листа книги: значения UsedRange и скрытые строки/столбцы
/// (аналог New-SheetData из RrfqEngine.ps1).
/// </summary>
public sealed class SheetData
{
    public string Name { get; init; } = string.Empty;

    public int Index { get; init; }

    public int FirstRow { get; init; }

    public int FirstCol { get; init; }

    public int LastRow { get; init; }

    public int LastCol { get; init; }

    /// <summary>Значения Value2, нормализованные в 0-базовый массив.</summary>
    public object?[,]? Values { get; init; }

    /// <summary>Скалярное значение для UsedRange из одной ячейки.</summary>
    public object? ScalarValue { get; init; }

    public HashSet<int> HiddenRows { get; init; } = new();

    public HashSet<int> HiddenCols { get; init; } = new();

    /// <summary>Аналог Get-SheetCellValue.</summary>
    public object? GetValue(int row, int column)
    {
        if (row < FirstRow || row > LastRow || column < FirstCol || column > LastCol)
        {
            return null;
        }

        if (Values is null)
        {
            return ScalarValue;
        }

        return Values[row - FirstRow, column - FirstCol];
    }

    /// <summary>Аналог Get-SheetCellText для слепка.</summary>
    public string GetCellText(int row, int column)
    {
        var value = GetValue(row, column);
        return value is null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture)!.Trim();
    }

    /// <summary>Аналог Get-SheetCellNumber: значение → число через Convert-ToNumber.</summary>
    public double? GetCellNumber(int row, int column)
    {
        if (column <= 0)
        {
            return null;
        }

        return Services.RrfqEngine.ConvertToNumber(GetValue(row, column));
    }

    public bool IsRowHidden(int row) => HiddenRows.Contains(row);

    public bool IsColumnHidden(int column) => HiddenCols.Contains(column);
}

/// <summary>Распознанная шапка таблицы (аналог результата Build-HeaderMapForRow).</summary>
public sealed class HeaderMap
{
    public string Type { get; init; } = string.Empty;

    public int HeaderRow { get; init; }

    public Dictionary<string, int> Fields { get; init; } = new(StringComparer.Ordinal);

    public int Score { get; init; }
}

/// <summary>Строка RFQ (аналог New-RfqRow).</summary>
public sealed class RfqRow
{
    public string Id { get; init; } = string.Empty;

    public string WorkbookPath { get; init; } = string.Empty;

    public string SheetName { get; init; } = string.Empty;

    public int SheetIndex { get; init; }

    public string Format { get; init; } = string.Empty;

    public int HeaderRow { get; init; }

    public int Row { get; init; }

    public string Key { get; init; } = string.Empty;

    public string KeyNorm { get; init; } = string.Empty;

    public string RussianRemark { get; init; } = string.Empty;

    public string ChinaRemark { get; init; } = string.Empty;

    public string PN { get; init; } = string.Empty;

    public string DC { get; init; } = string.Empty;

    public string MfgRussia { get; init; } = string.Empty;

    public string MfgChina { get; init; } = string.Empty;

    public string QtyPacking { get; init; } = string.Empty;

    public string QtyToBuy { get; init; } = string.Empty;

    /// <summary>Ручной выбор победителя (заполняется в разделе, аналог свойства в New-RfqRow).</summary>
    public string ManualWinnerId { get; set; } = string.Empty;
}

/// <summary>Квота поставщика (аналог New-Quote / New-ManualQuote).</summary>
public sealed class Quote
{
    public string Id { get; init; } = Guid.NewGuid().ToString("N");

    public string WorkbookPath { get; init; } = string.Empty;

    public string FileName { get; init; } = string.Empty;

    public string SheetName { get; init; } = string.Empty;

    public int SheetIndex { get; init; }

    public string Format { get; init; } = string.Empty;

    public int Row { get; init; }

    public string Supplier { get; init; } = string.Empty;

    public string Key { get; init; } = string.Empty;

    public string KeyNorm { get; init; } = string.Empty;

    public string RussianRemark { get; init; } = string.Empty;

    public string ChinaRemark { get; set; } = string.Empty;

    public string PN { get; set; } = string.Empty;

    public string DC { get; set; } = string.Empty;

    public string MfgRussia { get; set; } = string.Empty;

    public string MfgChina { get; set; } = string.Empty;

    public string QtyPacking { get; set; } = string.Empty;

    public string QtyToBuy { get; set; } = string.Empty;

    public double? UnitPrice { get; set; }

    public double? UnitPriceVat { get; set; }

    public string LeadTime { get; init; } = string.Empty;

    public double? LeadDays { get; init; }

    public string LeadTimeTotal { get; init; } = string.Empty;

    public string Currency { get; init; } = string.Empty;

    public bool IsPriceComparable { get; init; } = true;

    public string Warning { get; set; } = string.Empty;

    public string? TargetId { get; set; }

    public string MatchStatus { get; set; } = "Unmatched";

    public List<string> CandidateIds { get; set; } = new();
}

/// <summary>
/// Строка решения по позиции RFQ (аналог элемента из Build-Decisions).
/// Include — редактируемый чекбокс сетки, обновляется привязкой напрямую.
/// </summary>
public partial class Decision : ObservableObject
{
    [ObservableProperty]
    private bool _include;

    public string Id { get; init; } = string.Empty;

    public string Status { get; init; } = string.Empty;

    public string SheetName { get; init; } = string.Empty;

    public int Row { get; init; }

    public string Key { get; init; } = string.Empty;

    /// <summary>Отображаемое значение позиции (в оригинале — Value = Key).</summary>
    public string Value { get; init; } = string.Empty;

    public RfqRow RfqRow { get; init; } = null!;

    public List<Quote> Quotes { get; init; } = new();

    public Quote? Winner { get; init; }

    public string ManualWinnerId { get; init; } = string.Empty;

    public bool IsManualWinner { get; init; }

    public string WinnerSupplier { get; init; } = string.Empty;

    public string WinnerReason { get; init; } = string.Empty;

    public double? WinnerPrice { get; init; }

    public string WinnerLead { get; init; } = string.Empty;

    public string WinnerLeadTotal { get; init; } = string.Empty;

    public string WinnerPN { get; init; } = string.Empty;

    public string RussianRemark { get; init; } = string.Empty;

    public string ChinaRemark { get; init; } = string.Empty;

    public string DC { get; init; } = string.Empty;

    public string MfgRussia { get; init; } = string.Empty;

    public string MfgChina { get; init; } = string.Empty;

    public string QtyPacking { get; init; } = string.Empty;

    public string QtyToBuy { get; init; } = string.Empty;

    public string QuotesSummary { get; init; } = string.Empty;

    public string Warning { get; init; } = string.Empty;

    /// <summary>Цена победителя в формате оригинального Format-Price (N6).</summary>
    public string WinnerPriceText => WinnerPrice is null
        ? string.Empty
        : WinnerPrice.Value.ToString("N6", CultureInfo.GetCultureInfo("ru-RU"));
}

/// <summary>Строка нижней панели «Квоты выбранной позиции».</summary>
public sealed class QuoteDetail
{
    public Quote Quote { get; init; } = null!;

    public bool IsWinner { get; init; }

    public string WinnerMark => IsWinner ? "Да" : string.Empty;

    public string PriceText => Quote.UnitPrice is null
        ? string.Empty
        : Quote.UnitPrice.Value.ToString("N6", CultureInfo.GetCultureInfo("ru-RU"));
}

/// <summary>Результат анализа (аналог $script:LastAnalysis из Invoke-Analysis, без записи в БД).</summary>
public sealed class RrfqAnalysis
{
    public string RfqPath { get; init; } = string.Empty;

    public List<SupplierFileEntry> Suppliers { get; init; } = new();

    public List<RfqRow> RfqRows { get; init; } = new();

    public List<Quote> Quotes { get; init; } = new();

    public List<Decision> Decisions { get; set; } = new();

    /// <summary>Приоритет победителя: "Price" или "LeadTime".</summary>
    public string Priority { get; init; } = "Price";
}
