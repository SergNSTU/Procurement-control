using ProcurementControl.Models;
using ProcurementControl.Services;
using Xunit;

namespace ProcurementControl.Tests;

/// <summary>
/// Тесты чистой логики «Контроля закупки»: нормализация ключей, фильтры,
/// поиск, метрики и построение авто-пунктов напоминаний.
/// </summary>
public class PurchaseLogicTests
{
    // ----- NormalizeKey (перенос Normalize-Key из RrfqEngine.ps1) -----

    [Theory]
    [InlineData("  abc  ", "ABC")]
    [InlineData("ADS1278\u00A0IPZ", "ADS1278 IPZ")]          // NBSP -> пробел
    [InlineData("a\r\n\tb   c", "A B C")]                     // переводы строк и табы сжимаются
    [InlineData(null, "")]
    public void NormalizeKey_HandlesWhitespaceAndCase(string? value, string expected)
        => Assert.Equal(expected, PurchaseLogic.NormalizeKey(value));

    // ----- IsErpNotRequired (поставщики без ERP-документов) -----

    [Theory]
    [InlineData("Компэл", true)]
    [InlineData("  COMPEL_12  ", true)]
    [InlineData("ЧИД", true)]
    [InlineData("chid_3", true)]
    [InlineData("Promelec", true)]
    [InlineData("промэлектроника", true)]
    [InlineData("Ben", false)]
    [InlineData("Мосцентех", false)]
    [InlineData("", false)]
    public void IsErpNotRequired_MatchesKnownSuppliers(string supplier, bool expected)
        => Assert.Equal(expected, PurchaseLogic.IsErpNotRequired(supplier));

    // ----- FilterDeals (перенос WHERE Get-PurchaseDeals) -----

    private static PurchaseDealRow Deal(
        string number = "AE001",
        string status = "RFQ",
        string location = "Россия",
        string executor = "",
        int archived = 0,
        int suppliers = 0,
        int confirmed = 0)
        => new()
        {
            DealNumber = number,
            Status = status,
            AssemblyLocation = location,
            Executor = executor,
            Archived = archived,
            SupplierCount = suppliers,
            InvoiceConfirmedCount = confirmed,
        };

    [Fact]
    public void FilterDeals_Active_KeepsOperationalDealsWithoutExecutor()
    {
        var rows = new[]
        {
            Deal("AE001", "RFQ"),                    // остаётся
            Deal("AE002", "Закупка"),                // остаётся
            Deal("AE003", "Заказано", suppliers: 2, confirmed: 1), // Заказано с неполными подтверждениями — остаётся
            Deal("AE004", "Заказано", suppliers: 2, confirmed: 2), // всё подтверждено — не активная
            Deal("AE005", "RFQ", executor: "Иванов"), // есть исполнитель — не активная
        };

        var active = PurchaseLogic.FilterDeals(rows, DealFilter.Active);

        Assert.Equal(new[] { "AE001", "AE002", "AE003" }, active.Select(r => r.DealNumber));
    }

    [Fact]
    public void FilterDeals_NonChinaFilters_ExcludeChinaDeals()
    {
        var rows = new[] { Deal("AE001", location: "Китай"), Deal("AE002", location: "Россия") };

        var filtered = PurchaseLogic.FilterDeals(rows, DealFilter.Active);

        Assert.Equal(new[] { "AE002" }, filtered.Select(r => r.DealNumber));
    }

    [Fact]
    public void FilterDeals_China_KeepsOnlyChina()
    {
        var rows = new[] { Deal("AE001", location: "Китай"), Deal("AE002", location: "Россия") };

        var china = PurchaseLogic.FilterDeals(rows, DealFilter.China);

        Assert.Equal(new[] { "AE001" }, china.Select(r => r.DealNumber));
    }

    [Fact]
    public void FilterDeals_Done_RequiresAllInvoicesConfirmed()
    {
        var rows = new[]
        {
            Deal("AE001", "Заказано", suppliers: 2, confirmed: 2),
            Deal("AE002", "Заказано", suppliers: 2, confirmed: 1),
            Deal("AE003", "Заказано", suppliers: 0, confirmed: 0), // нет поставщиков — не «Готово»
        };

        var done = PurchaseLogic.FilterDeals(rows, DealFilter.Done);

        Assert.Equal(new[] { "AE001" }, done.Select(r => r.DealNumber));
    }

    [Fact]
    public void FilterDeals_Pause_KeepsRrfqAndPi()
    {
        var rows = new[] { Deal("AE001", "RRFQ"), Deal("AE002", "PI"), Deal("AE003", "RFQ") };

        var paused = PurchaseLogic.FilterDeals(rows, DealFilter.Pause);

        Assert.Equal(new[] { "AE001", "AE002" }, paused.Select(r => r.DealNumber));
    }

    [Fact]
    public void FilterDeals_Tracking_KeepsDealsWithExecutor_IncludingChina()
    {
        var rows = new[]
        {
            Deal("AE001", "RFQ", executor: "Иванов"),
            Deal("AE002", "RFQ", executor: "Иванов", location: "Китай"), // Tracking не отсекает Китай
            Deal("AE003", "RFQ"),
        };

        var tracking = PurchaseLogic.FilterDeals(rows, DealFilter.Tracking);

        Assert.Equal(new[] { "AE001", "AE002" }, tracking.Select(r => r.DealNumber));
    }

    [Fact]
    public void FilterDeals_Archived_KeepsOnlyArchived()
    {
        var rows = new[] { Deal("AE001", archived: 1), Deal("AE002") };

        var archived = PurchaseLogic.FilterDeals(rows, DealFilter.Archived);

        Assert.Equal(new[] { "AE001" }, archived.Select(r => r.DealNumber));
    }

    // ----- SearchDeals (аналог LIKE-поиска) -----

    [Fact]
    public void SearchDeals_FindsByNumberClientAndTitle_CaseInsensitive()
    {
        var rows = new[]
        {
            Deal("AE26072701TA-000334"),
            new PurchaseDealRow { DealNumber = "AE002", Client = "ООО Ромашка" },
            new PurchaseDealRow { DealNumber = "AE003", Title = "Плата управления" },
        };

        Assert.Single(PurchaseLogic.SearchDeals(rows, "000334"));
        Assert.Single(PurchaseLogic.SearchDeals(rows, "ромашка"));
        Assert.Single(PurchaseLogic.SearchDeals(rows, "плата УПРАВЛЕНИЯ"));
        Assert.Equal(3, PurchaseLogic.SearchDeals(rows, "").Count);
    }

    // ----- ComputeMetrics (дедупликация по номеру сделки) -----

    [Fact]
    public void ComputeMetrics_DeduplicatesDealsByNumber_AndCountsTasksAndAttention()
    {
        var rows = new[]
        {
            Deal("AE001", "RFQ"),
            Deal("AE001", "RFQ"),                       // дубль номера — считается один раз
            Deal("AE002", "Заказано", suppliers: 2, confirmed: 1), // attention
            new PurchaseDealRow { DealNumber = "AE003", Status = "Заказано", Archived = 1, SupplierCount = 1 }, // архив — мимо
        };
        var tasks = new List<CockpitTaskKey> { new() { DealNumber = "AE004", Id = 1 } };

        var (active, overdue, workAndPayment, attention) = PurchaseLogic.ComputeMetrics(rows, tasks);

        Assert.Equal(3, active);                    // AE001 (дубль), AE002, задача AE004
        Assert.Equal(1, workAndPayment);
        Assert.Equal(1, attention);
        Assert.Equal(0, overdue);
    }

    [Fact]
    public void ComputeMetrics_OverdueCountsReceiptsInThePast_ForNotDoneDeals()
    {
        var rows = new[]
        {
            new PurchaseDealRow { DealNumber = "AE001", Status = "RFQ", CompletionReceiptDate = DateTime.Today.AddDays(-3).ToString("dd.MM.yyyy") },
            new PurchaseDealRow { DealNumber = "AE002", Status = "Заказано", CompletionReceiptDate = DateTime.Today.AddDays(-1).ToString("dd.MM.yyyy") },
            new PurchaseDealRow { DealNumber = "AE003", Status = "Done", CompletionReceiptDate = DateTime.Today.AddDays(-1).ToString("dd.MM.yyyy") },
            new PurchaseDealRow { DealNumber = "AE004", Status = "RFQ", CompletionReceiptDate = DateTime.Today.ToString("dd.MM.yyyy") },
        };

        var (_, overdue, _, _) = PurchaseLogic.ComputeMetrics(rows, Array.Empty<CockpitTaskKey>());

        // Просрочка: AE001 и AE002 (статус «Заказано» не входит в Done-правило);
        // AE003 закрыт (Done), AE004 поступает сегодня.
        Assert.Equal(2, overdue);
    }

    // ----- BuildActionItems (перенос Get-PurchaseActionItems) -----

    private static SupplierActionRow Supplier(Action<SupplierActionRow>? setup = null)
    {
        var row = new SupplierActionRow
        {
            DealId = 10,
            DealNumber = "AE001",
            SupplierId = 20,
            Supplier = "Ben",
            PiAmountUsd = "100",
            InvoiceReceived = true,
            InvoiceConfirmedDate = "01.02.2026",
            // ERP-документы по умолчанию отправлены — тесты поступлений не должны
            // спотыкаться об ERP-пункты; отдельно они проверяются своим тестом.
            ErpSupplierSent = true,
            ErpRogerSent = true,
        };
        setup?.Invoke(row);
        return row;
    }

    [Fact]
    public void BuildActionItems_WarnsWhenPiAmountMissing_AndWhenNoInvoice()
    {
        var items = PurchaseLogic.BuildActionItems(
            new[] { Supplier(s => s.PiAmountUsd = ""), Supplier(s => s.InvoiceReceived = false) },
            Array.Empty<DealReminderRow>(),
            Array.Empty<ManualReminderRow>(),
            DateTime.Today);

        Assert.Contains(items, i => i.Title.Contains("не указана сумма PI"));
        Assert.Contains(items, i => i.Title.Contains("нет PI"));
        // После «нет PI» построение пунктов поставщика прекращается (continue).
        Assert.DoesNotContain(items, i => i.Title.Contains("нет даты подтверждения"));
    }

    [Fact]
    public void BuildActionItems_SkipsErpItemsForErpNotRequiredSuppliers()
    {
        var items = PurchaseLogic.BuildActionItems(
            new[] { Supplier(s => s.Supplier = "Компэл") },
            Array.Empty<DealReminderRow>(),
            Array.Empty<ManualReminderRow>(),
            DateTime.Today);

        Assert.DoesNotContain(items, i => i.Title.Contains("ERP"));
        Assert.Empty(items);
    }

    [Fact]
    public void BuildActionItems_FlagsErpDocumentsForRegularSuppliers()
    {
        var items = PurchaseLogic.BuildActionItems(
            new[] { Supplier(s =>
            {
                s.ErpSupplierSent = false;
                s.ErpRogerSent = false;
            }) },
            Array.Empty<DealReminderRow>(),
            Array.Empty<ManualReminderRow>(),
            DateTime.Today);

        Assert.Equal(2, items.Count);
        Assert.Contains(items, i => i.Title.Contains("Не отправлен ERP поставщику"));
        Assert.Contains(items, i => i.Title.Contains("Не отправлен ERP заказ"));
    }

    [Fact]
    public void BuildActionItems_ReceiptItems_UsePlannedDateAndWindow()
    {
        var today = new DateTime(2026, 9, 30);
        var items = PurchaseLogic.BuildActionItems(
            new[]
            {
                Supplier(s => s.ComponentsReceiptDate = "20.09.2026"),          // просрочено
                Supplier(s => s.ComponentsReceiptDate = "03.10.2026"),          // в окне 7 дней
                Supplier(s => s.ComponentsReceiptDate = "01.12.2026"),          // далеко — нет пункта
                Supplier(s =>                                                  // факт есть — нет пункта
                {
                    s.ComponentsReceiptDate = "01.12.2026";
                    s.ActualReceiptDate = "25.09.2026";
                }),
            },
            Array.Empty<DealReminderRow>(),
            Array.Empty<ManualReminderRow>(),
            today);

        var overdue = Assert.Single(items, i => i.Title.Contains("поступление просрочено"));
        Assert.Equal("Danger", overdue.Severity);
        var soon = Assert.Single(items, i => i.Title.Contains("скоро поступление"));
        Assert.Equal("Attention", soon.Severity);
        Assert.Equal(2, items.Count);
    }

    [Fact]
    public void BuildActionItems_SuppressedKeys_AreNotReturned()
    {
        var suppressed = new HashSet<string>
        {
            PurchaseLogic.AutomaticReminderKey(10, 20, "AE001 / Ben: нет PI"),
        };

        var items = PurchaseLogic.BuildActionItems(
            new[] { Supplier(s => s.InvoiceReceived = false) },
            Array.Empty<DealReminderRow>(),
            Array.Empty<ManualReminderRow>(),
            DateTime.Today,
            suppressed);

        Assert.DoesNotContain(items, i => i.Title.Contains("нет PI"));
        Assert.All(items, i => Assert.NotEqual(suppressed.Single(), i.SuppressionKey));
    }

    [Fact]
    public void BuildActionItems_DealAndManualItems_CarrySeverityAndSource()
    {
        var today = new DateTime(2026, 9, 30);
        var items = PurchaseLogic.BuildActionItems(
            Array.Empty<SupplierActionRow>(),
            new[] { new DealReminderRow { Id = 1, DealNumber = "AE001", ReminderDate = "25.09.2026" } },
            new[]
            {
                new ManualReminderRow { Id = 2, DealId = 10, SupplierId = 20, Title = "Позвонить", DueDate = "05.10.2026", Source = "manual" },
                new ManualReminderRow { Id = 3, DealId = 0, SupplierId = 0, Title = "Без даты", DueDate = "", Source = "quick_capture" },
            },
            today);

        var dealItem = Assert.Single(items, i => i.Source == "deal");
        Assert.Equal("Danger", dealItem.Severity);          // 25.09 < 30.09
        var manual = Assert.Single(items, i => i.Title == "Позвонить");
        Assert.Equal(("Attention", "manual", 10L, 20L, 2L), (manual.Severity, manual.Source, manual.DealId, manual.SupplierId, manual.ReminderId));
        var noDate = Assert.Single(items, i => i.Title == "Без даты");
        Assert.Equal("Warn", noDate.Severity);              // дата не парсится
    }

    [Fact]
    public void SearchItems_FiltersByTitleAndDealId()
    {
        var items = new List<ActionItemRow>
        {
            new() { Title = "AE001 / Ben: нет PI", DealId = 7 },
            new() { Title = "Другое дело", DealId = 42 },
        };

        Assert.Single(PurchaseLogic.SearchItems(items, "ben"));
        Assert.Single(PurchaseLogic.SearchItems(items, "42"));
        Assert.Equal(2, PurchaseLogic.SearchItems(items, "").Count);
    }
}
