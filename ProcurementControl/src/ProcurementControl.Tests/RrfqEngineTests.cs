using ProcurementControl.Models;
using ProcurementControl.Services;
using Xunit;

namespace ProcurementControl.Tests;

/// <summary>
/// Тесты чистых функций движка RRFQ: нормализация значений, распознавание
/// сроков и ядро выбора победителя (BuildDecisions) с сопоставлением целей.
/// </summary>
public class RrfqEngineTests
{
    // ----- Нормализация и разбор значений -----

    [Theory]
    [InlineData("  Value ", "value")]
    [InlineData("Цена\u00A0за\u00A0шт", "цена за шт")]
    [InlineData(null, "")]
    public void NormalizeHeader_TrimsLowercasesAndCollapses(object? value, string expected)
        => Assert.Equal(expected, RrfqEngine.NormalizeHeader(value));

    [Theory]
    [InlineData(" lm35 ", "LM35")]
    [InlineData("a\r\nb", "A B")]
    public void NormalizeKey_UppercasesAndCollapses(object? value, string expected)
        => Assert.Equal(expected, RrfqEngine.NormalizeKey(value));

    [Theory]
    [InlineData("1 234,56 $", 1234.56)]     // валютный мусор, запятая как разделитель
    [InlineData("$1,234.56", 1234.56)]      // тысячные через запятую
    [InlineData("12 usd", 12.0)]
    [InlineData("3.5", 3.5)]
    [InlineData("", null)]
    [InlineData("нет цены", null)]
    public void ConvertToNumber_ParsesPricesAndQuantities(object? value, double? expected)
        => Assert.Equal(expected, RrfqEngine.ConvertToNumber(value));

    // ----- ConvertLeadTime / FormatWeeks -----

    [Fact]
    public void ConvertLeadTime_Empty_ReturnsEmptyWithoutWarning()
    {
        var lead = RrfqEngine.ConvertLeadTime("");
        Assert.Null(lead.Days);
        Assert.Equal(string.Empty, lead.Total);
        Assert.Equal(string.Empty, lead.Warning);
    }

    [Theory]
    [InlineData("stock")]
    [InlineData("In stock")]
    [InlineData("на складе")]
    public void ConvertLeadTime_Stock_IsZeroDays(string leadTime)
    {
        var lead = RrfqEngine.ConvertLeadTime(leadTime);
        Assert.Equal(0.0, lead.Days);
        Assert.Equal("stock", lead.Total);
        Assert.Equal(string.Empty, lead.Warning);
    }

    [Fact]
    public void ConvertLeadTime_WeekRange_AddsLogisticBuffer()
    {
        // 8-12 недель -> дни по максимуму, итог с буфером +1/+2 недели.
        var lead = RrfqEngine.ConvertLeadTime("8-12 weeks");
        Assert.Equal(84.0, lead.Days);
        Assert.Equal("9-14 weeks", lead.Total);
        Assert.Equal(string.Empty, lead.Warning);
    }

    [Fact]
    public void ConvertLeadTime_ShortDays_RoundsToOneWeek()
    {
        var lead = RrfqEngine.ConvertLeadTime("3 days");
        Assert.Equal(3.0, lead.Days);
        Assert.Equal("1 week", lead.Total);
    }

    [Fact]
    public void ConvertLeadTime_Unrecognized_Warns()
    {
        var lead = RrfqEngine.ConvertLeadTime("скоро");
        Assert.Null(lead.Days);
        Assert.Contains("Не распознан срок", lead.Warning);
    }

    [Theory]
    [InlineData(1, 1, "1 week")]
    [InlineData(9, 9, "9 weeks")]
    [InlineData(9, 14, "9-14 weeks")]
    public void FormatWeeks_FormatsRanges(int min, int max, string expected)
        => Assert.Equal(expected, RrfqEngine.FormatWeeks(min, max));

    // ----- Особые поставщики (Compel/Ben) -----

    [Theory]
    [InlineData("Компэл", true)]
    [InlineData("COMPEL_2", true)]
    [InlineData("КОМПЕЛ", true)]
    [InlineData("Компэл 2", false)]          // пробел вместо суффикса _N
    [InlineData("CPR", false)]
    public void IsCompelSupplier_RecognizesNameWithOptionalSuffix(string supplier, bool expected)
        => Assert.Equal(expected, RrfqEngine.IsCompelSupplier(supplier));

    [Theory]
    [InlineData("Ben", true)]
    [InlineData("БЕН_1", true)]
    [InlineData("Бенадж", false)]
    public void IsBenSupplier_RecognizesNameWithOptionalSuffix(string supplier, bool expected)
        => Assert.Equal(expected, RrfqEngine.IsBenSupplier(supplier));

    [Fact]
    public void GetLeadTimeTotalForSupplier_Compel_DividesByWorkingWeek()
    {
        // 28 календарных дней -> 6 рабочих недель (28/5 вверх).
        var total = RrfqEngine.GetLeadTimeTotalForSupplier("Компэл", "4 weeks", RrfqEngine.ConvertLeadTime("4 weeks"));
        Assert.Equal("6 weeks", total);
    }

    [Fact]
    public void GetLeadTimeTotalForSupplier_Ben_KeepsRawText()
    {
        var total = RrfqEngine.GetLeadTimeTotalForSupplier("Ben", "4-6 weeks", RrfqEngine.ConvertLeadTime("4-6 weeks"));
        Assert.Equal("4-6 weeks", total);
    }

    [Fact]
    public void GetLeadTimeTotalForSupplier_Others_UsesConvertedTotal()
    {
        var total = RrfqEngine.GetLeadTimeTotalForSupplier("CPR", "8-12 weeks", RrfqEngine.ConvertLeadTime("8-12 weeks"));
        Assert.Equal("9-14 weeks", total);
    }

    // ----- ResolveQuoteTargets: сопоставление квот строкам RFQ -----

    private static RfqRow RfqRow(string id, string key, string keyNorm, string sheet = "RFQ", int row = 2)
        => new() { Id = id, Key = key, KeyNorm = keyNorm, SheetName = sheet, Row = row };

    private static Quote Quote(string keyNorm, string sheet = "RFQ", int row = 2, string format = "Standard")
        => new() { KeyNorm = keyNorm, SheetName = sheet, Row = row, Format = format };

    [Fact]
    public void ResolveQuoteTargets_StandardFormat_MatchesByExactRow()
    {
        var rfq = RfqRow("r1", "LM35", "LM35");
        var quotes = new List<Quote> { Quote("LM35") };

        RrfqEngine.ResolveQuoteTargets(new List<RfqRow> { rfq }, quotes);

        Assert.Equal("r1", quotes[0].TargetId);
        Assert.Equal("AutoRow", quotes[0].MatchStatus);
        Assert.Empty(quotes[0].Warning);
    }

    [Fact]
    public void ResolveQuoteTargets_NonStandard_MatchesByUniqueKey()
    {
        var rfq = RfqRow("r1", "LM35", "LM35");
        var quote = Quote("LM35", format: "Compel");

        RrfqEngine.ResolveQuoteTargets(new List<RfqRow> { rfq }, new List<Quote> { quote });

        Assert.Equal("r1", quote.TargetId);
        Assert.Equal("AutoKey", quote.MatchStatus);
    }

    [Fact]
    public void ResolveQuoteTargets_KeyOnTwoRows_MarksAmbiguous()
    {
        var rfq = new List<RfqRow> { RfqRow("r1", "LM35", "LM35", row: 2), RfqRow("r2", "LM35", "LM35", row: 3) };
        var quote = Quote("LM35", row: 5, format: "Compel");  // строка не совпадает — только по ключу

        RrfqEngine.ResolveQuoteTargets(rfq, new List<Quote> { quote });

        Assert.Null(quote.TargetId);
        Assert.Equal("Ambiguous", quote.MatchStatus);
        Assert.Equal(new[] { "r1", "r2" }, quote.CandidateIds);
    }

    [Fact]
    public void ResolveQuoteTargets_UnknownKey_IsUnmatched()
    {
        var quote = Quote("UNKNOWN");

        RrfqEngine.ResolveQuoteTargets(new List<RfqRow> { RfqRow("r1", "LM35", "LM35") }, new List<Quote> { quote });

        Assert.Null(quote.TargetId);
        Assert.Equal("Unmatched", quote.MatchStatus);
    }

    // ----- BuildDecisions: ядро выбора победителя -----

    private static Quote PricedQuote(string supplier, double price, double? leadDays, string targetId)
        => new() { Id = supplier + "-q", Supplier = supplier, UnitPrice = price, LeadDays = leadDays, TargetId = targetId };

    [Fact]
    public void BuildDecisions_PricePriority_WinsLowestPrice()
    {
        var row = RfqRow("r1", "LM35", "LM35");
        var quotes = new List<Quote>
        {
            PricedQuote("CPR", 10.0, 28, "r1"),
            PricedQuote("CRL", 8.5, 42, "r1"),
            PricedQuote("BMZ", 12.0, 14, "r1"),
        };

        var decisions = RrfqEngine.BuildDecisions(new List<RfqRow> { row }, quotes, "Price");

        var decision = Assert.Single(decisions);
        Assert.Equal("OK", decision.Status);
        Assert.Equal("CRL", decision.WinnerSupplier);        // минимальная цена
        Assert.Equal(8.5, decision.WinnerPrice);
    }

    [Fact]
    public void BuildDecisions_LeadTimePriority_WinsShortestLead()
    {
        var row = RfqRow("r1", "LM35", "LM35");
        var quotes = new List<Quote>
        {
            PricedQuote("CPR", 10.0, 28, "r1"),
            PricedQuote("BMZ", 12.0, 14, "r1"),
        };

        var decisions = RrfqEngine.BuildDecisions(new List<RfqRow> { row }, quotes, "LeadTime");

        Assert.Equal("BMZ", Assert.Single(decisions).WinnerSupplier);
    }

    [Fact]
    public void BuildDecisions_SamePrice_TieBreakByLeadDays()
    {
        var row = RfqRow("r1", "LM35", "LM35");
        var quotes = new List<Quote>
        {
            PricedQuote("CPR", 10.0, 28, "r1"),
            PricedQuote("BMZ", 10.0, 14, "r1"),
        };

        var decisions = RrfqEngine.BuildDecisions(new List<RfqRow> { row }, quotes, "Price");

        var decision = Assert.Single(decisions);
        Assert.Equal("BMZ", decision.WinnerSupplier);
        Assert.Contains("цена равна", decision.WinnerReason);
    }

    [Fact]
    public void BuildDecisions_NoSelectableQuotes_ReportsNoComparablePrice()
    {
        var row = RfqRow("r1", "LM35", "LM35");
        var quotes = new List<Quote>
        {
            new() { Id = "q1", Supplier = "CPR", UnitPrice = null, TargetId = "r1" },                       // нет цены
            new() { Id = "q2", Supplier = "BMZ", UnitPrice = 5.0, IsPriceComparable = false, TargetId = "r1" }, // несравнимая
        };

        var decisions = RrfqEngine.BuildDecisions(new List<RfqRow> { row }, quotes, "Price");

        var decision = Assert.Single(decisions);
        Assert.Equal("Нет сравнимой цены", decision.Status);
        Assert.Null(decision.Winner);
        Assert.Contains("CPR: нет цены", decision.Warning);
    }

    [Fact]
    public void BuildDecisions_QuotesWithoutTarget_AreIgnored()
    {
        var row = RfqRow("r1", "LM35", "LM35");
        var quotes = new List<Quote> { PricedQuote("CPR", 1.0, 28, "other-row") };

        var decisions = RrfqEngine.BuildDecisions(new List<RfqRow> { row }, quotes, "Price");

        Assert.Equal("Нет квот", Assert.Single(decisions).Status);
    }

    [Fact]
    public void BuildDecisions_ManualWinner_OverridesPrice()
    {
        var row = new RfqRow { Id = "r1", Key = "LM35", KeyNorm = "LM35", ManualWinnerId = "bmz-q" };
        var quotes = new List<Quote>
        {
            PricedQuote("CPR", 1.0, 28, "r1"),
            new Quote { Id = "bmz-q", Supplier = "BMZ", UnitPrice = 9.0, LeadDays = 10, TargetId = "r1" },
        };

        var decisions = RrfqEngine.BuildDecisions(new List<RfqRow> { row }, quotes, "Price");

        var decision = Assert.Single(decisions);
        Assert.Equal("BMZ", decision.WinnerSupplier);
        Assert.True(decision.IsManualWinner);
        Assert.Contains("Ручной выбор победителя: BMZ", decision.Warning);
    }

    [Fact]
    public void IsQuoteSelectable_RequiresPriceAndComparability()
    {
        Assert.True(RrfqEngine.IsQuoteSelectable(new Quote { UnitPrice = 1.0 }));
        Assert.False(RrfqEngine.IsQuoteSelectable(new Quote { UnitPrice = null }));
        Assert.False(RrfqEngine.IsQuoteSelectable(new Quote { UnitPrice = 1.0, IsPriceComparable = false }));
    }
}
