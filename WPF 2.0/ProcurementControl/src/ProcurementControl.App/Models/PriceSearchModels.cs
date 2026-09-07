using System.Globalization;

namespace ProcurementControl.Models;

/// <summary>
/// Строка сетки «Поиск цен» — аналог объекта New-PriceSearchResult из
/// app/modules/PriceSearch.ps1 (строки 12-15).
/// </summary>
public sealed class PriceSearchRow
{
    public string Source { get; init; } = string.Empty;

    public string PN { get; init; } = string.Empty;

    public string Manufacturer { get; init; } = string.Empty;

    public double? Price { get; init; }

    public string Currency { get; init; } = string.Empty;

    public string Moq { get; init; } = string.Empty;

    public string Stock { get; init; } = string.Empty;

    public string LeadTime { get; init; } = string.Empty;

    public string Date { get; init; } = string.Empty;

    public string Link { get; init; } = string.Empty;

    /// <summary>Цена в формате оригинального Format-Price (N6), пусто при отсутствии.</summary>
    public string PriceText => Price is null
        ? string.Empty
        : Price.Value.ToString("N6", CultureInfo.GetCultureInfo("ru-RU"));

    /// <summary>Ссылка пригодна для открытия (аналог проверки '^https?://' в обработчике сетки).</summary>
    public bool HasLink => !string.IsNullOrWhiteSpace(Link)
        && (Link.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || Link.StartsWith("https://", StringComparison.OrdinalIgnoreCase));
}
