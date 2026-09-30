using ProcurementControl.Services;
using Xunit;

namespace ProcurementControl.Tests;

/// <summary>
/// Тесты чистых хелперов аналитики поставщиков: нормализация производителей
/// и разбор сроков в дни.
/// </summary>
public class SupplierAnalyticsServiceTests
{
    [Theory]
    [InlineData("  Texas Instruments  ", "TEXASINSTRUMENTS")]
    [InlineData("МосЦентех-ЛТД", "МОСЦЕНТЕХЛТД")]    // кириллическая ЛТД не срезается (только LTD/ООО/…)
    [InlineData("Ромашка ООО", "РОМАШКА")]           // юрформа срезается только в конце
    [InlineData("ADI CORP.", "ADI")]
    [InlineData("Segment INC", "SEGMENT")]
    [InlineData(null, "")]
    [InlineData("", "")]
    public void NormalizeManufacturer_TrimsSeparatorsAndLegalSuffixes(string? value, string expected)
        => Assert.Equal(expected, SupplierAnalyticsService.NormalizeManufacturer(value));

    [Theory]
    [InlineData("stock", 0.0)]
    [InlineData("На складе", 0.0)]
    [InlineData("2 weeks", 14.0)]
    [InlineData("8-12 weeks", 84.0)]                 // берётся максимум диапазона
    [InlineData("10 days", 10.0)]
    [InlineData("10 дн", 10.0)]
    [InlineData("1,5 weeks", 10.5)]
    [InlineData("", null)]
    [InlineData("скоро", null)]                      // без чисел — не парсится
    public void ParseLeadTimeTotalDays_ParsesWeeksAndDays(string? value, double? expected)
        => Assert.Equal(expected, SupplierAnalyticsService.ParseLeadTimeTotalDays(value));
}
