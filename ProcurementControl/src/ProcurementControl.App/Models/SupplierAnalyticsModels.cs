using System.Globalization;
using System.IO;

namespace ProcurementControl.Models;

public sealed record AnalysisSaveResult(long BatchId, int SavedQuotes, bool IsDuplicate, bool IsWinnerOnly = false,
    long PreviousSummaryDocumentId = 0);

public sealed record QuoteImportSaveResult(long BatchId, int Added, int Updated, int Skipped, int NeedsReview)
{
    public int Changed => Added + Updated;
}

public sealed class SupplierRecommendationRow
{
    public string Supplier { get; init; } = string.Empty;
    public double Score { get; init; }
    public string ScoreText => Score.ToString("0.0", CultureInfo.GetCultureInfo("ru-RU"));
    public string Confidence { get; init; } = string.Empty;
    public int Comparisons { get; init; }
    public string ComparisonsText => Comparisons.ToString(CultureInfo.InvariantCulture);
    public double PriceCompetitiveness { get; init; }
    public string PriceCompetitivenessText => (PriceCompetitiveness * 100).ToString("0.0", CultureInfo.GetCultureInfo("ru-RU")) + "%";
    public double BestPriceRate { get; init; }
    public string BestPriceRateText => (BestPriceRate * 100).ToString("0.0", CultureInfo.GetCultureInfo("ru-RU")) + "%";
    public double ResponseRate { get; init; }
    public string ResponseRateText => (ResponseRate * 100).ToString("0.0", CultureInfo.GetCultureInfo("ru-RU")) + "%";
    public double? AverageLeadDays { get; init; }
    public string AverageLeadText => AverageLeadDays is null ? "нет данных" : AverageLeadDays.Value.ToString("0.#", CultureInfo.GetCultureInfo("ru-RU")) + " дн.";
    public string Explanation { get; init; } = string.Empty;
    public bool IsFallback { get; init; }
}

public sealed class AnalyticsManufacturerOption
{
    public long Id { get; init; }
    public string Name { get; init; } = string.Empty;
    public string Alias { get; init; } = string.Empty;
    public string DisplayName => string.IsNullOrWhiteSpace(Alias) ? Name : Name + " ← " + Alias;
}

public sealed class ManufacturerAliasRow
{
    public long Id { get; init; }
    public long ManufacturerId { get; init; }
    public string Manufacturer { get; init; } = string.Empty;
    public string Alias { get; init; } = string.Empty;
    public bool IsUnresolved => ManufacturerId <= 0;
    public string Status => IsUnresolved ? "Не разобрано" : "Привязан";
}

public sealed class AnalyticsBatchRow
{
    public long Id { get; init; }
    public string CreatedAt { get; init; } = string.Empty;
    public string SourceKind { get; init; } = string.Empty;
    public string DataQuality { get; init; } = string.Empty;
    public string RfqPath { get; init; } = string.Empty;
    public string ResultPath { get; init; } = string.Empty;
    public string DealNumber { get; init; } = string.Empty;
    public string DealDisplay => string.IsNullOrWhiteSpace(DealNumber) ? "Сделка не определена" : DealNumber;
    public int PositionCount { get; init; }
    public int QuoteCount { get; init; }
    public string FileName => string.IsNullOrWhiteSpace(RfqPath) ? "Без файла" : Path.GetFileName(RfqPath);
}

public sealed class AnalyticsBatchQuoteRow
{
    public long BatchId { get; init; }
    public string CreatedAt { get; init; } = string.Empty;
    public string RfqValue { get; init; } = string.Empty;
    public string PN { get; init; } = string.Empty;
    public string Manufacturer { get; init; } = string.Empty;
    public string Supplier { get; init; } = string.Empty;
    public double? UnitPrice { get; init; }
    public string PriceText => UnitPrice is null ? string.Empty : UnitPrice.Value.ToString("N6", CultureInfo.GetCultureInfo("ru-RU"));
    public string LeadTime { get; init; } = string.Empty;
    public bool IsWinner { get; init; }
    public string WinnerText => IsWinner ? "Да" : string.Empty;
    public string DataQuality { get; init; } = string.Empty;
}

public sealed class AnalyticsMonthlyRow
{
    public string Month { get; init; } = string.Empty;
    public string Supplier { get; init; } = string.Empty;
    public int Comparisons { get; init; }
    public string ComparisonsText => Comparisons.ToString(CultureInfo.InvariantCulture);
    public double PriceCompetitiveness { get; init; }
    public string PriceText => (PriceCompetitiveness * 100).ToString("0.0", CultureInfo.GetCultureInfo("ru-RU")) + "%";
    public double BestPriceRate { get; init; }
    public string BestText => (BestPriceRate * 100).ToString("0.0", CultureInfo.GetCultureInfo("ru-RU")) + "%";
}

public sealed record AnalyticsFilter(string DateFrom, string DateTo, string Manufacturer, string Supplier, string DealNumber);

/// <summary>Прозрачные показатели поставщика. Дроби сохраняются рядом с процентами для проверки выборки.</summary>
public sealed class SupplierOverviewRow
{
    public string Supplier { get; init; } = string.Empty;
    public int Cohorts { get; init; }
    public int Positions { get; init; }
    public int UniquePn { get; init; }
    public int ResponseNumerator { get; init; }
    public int ResponseDenominator { get; init; }
    public int Wins { get; init; }
    public int WinDenominator { get; init; }
    public int BestPriceCount { get; init; }
    public int PriceDenominator { get; init; }
    public double? MeanPriceDeviation { get; init; }
    public double? MedianPriceDeviation { get; init; }
    public double? MeanLeadDays { get; init; }
    public double? MedianLeadDays { get; init; }
    public int BestLeadCount { get; init; }
    public int LeadDenominator { get; init; }
    public string CohortsText => Cohorts.ToString(CultureInfo.InvariantCulture);
    public string PositionsText => Positions.ToString(CultureInfo.InvariantCulture);
    public string UniquePnText => UniquePn.ToString(CultureInfo.InvariantCulture);
    public string ResponseText => RatioText(ResponseNumerator, ResponseDenominator);
    public string WinsText => RatioText(Wins, WinDenominator);
    public string BestPriceText => RatioText(BestPriceCount, PriceDenominator);
    public string MeanPriceDeviationText => PercentText(MeanPriceDeviation);
    public string MedianPriceDeviationText => PercentText(MedianPriceDeviation);
    public string MeanLeadText => DaysText(MeanLeadDays);
    public string MedianLeadText => DaysText(MedianLeadDays);
    public string BestLeadText => RatioText(BestLeadCount, LeadDenominator);
    public string CoverageText => $"Цена {PriceDenominator}; победы {WinDenominator}; ответы {ResponseDenominator}";
    private static string RatioText(int value, int total) => total == 0 ? "—" : $"{value / (double)total:P1} ({value}/{total})";
    private static string PercentText(double? value) => value is null ? "—" : value.Value.ToString("+0.0%;-0.0%;0.0%", CultureInfo.GetCultureInfo("ru-RU"));
    private static string DaysText(double? value) => value is null ? "—" : value.Value.ToString("0.#", CultureInfo.GetCultureInfo("ru-RU")) + " дн.";
}

public sealed class SupplierOverviewDetailRow
{
    public string DealNumber { get; init; } = string.Empty;
    public string PositionKey { get; init; } = string.Empty;
    public string PN { get; init; } = string.Empty;
    public string Manufacturer { get; init; } = string.Empty;
    public double? Price { get; init; }
    public string PriceText => Price is null ? string.Empty : Price.Value.ToString("N6", CultureInfo.GetCultureInfo("ru-RU"));
    public double? LeadDays { get; init; }
    public string LeadText => LeadDays is null ? string.Empty : LeadDays.Value.ToString("0.#", CultureInfo.GetCultureInfo("ru-RU"));
    public bool IsWinner { get; init; }
    public string SourceKind { get; init; } = string.Empty;
}

/// <summary>Строка аудита знаменателя частоты ответа выбранного поставщика.</summary>
public sealed class SupplierResponseAuditRow
{
    public string DealNumber { get; init; } = string.Empty;
    public long BatchId { get; init; }
    public string SourceKind { get; init; } = string.Empty;
    public string PositionKey { get; init; } = string.Empty;
    public string PN { get; init; } = string.Empty;
    public string Manufacturer { get; init; } = string.Empty;
    public bool HasQuote { get; init; }
    public string HasQuoteText => HasQuote ? "Да" : "Нет";
    public string CalculationBasis { get; init; } = string.Empty;
}
