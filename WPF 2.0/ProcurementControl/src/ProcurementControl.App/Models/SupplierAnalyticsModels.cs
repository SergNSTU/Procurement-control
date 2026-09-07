using System.Globalization;
using System.IO;

namespace ProcurementControl.Models;

public sealed record AnalysisSaveResult(long BatchId, int SavedQuotes, bool IsDuplicate, bool IsWinnerOnly = false);

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
