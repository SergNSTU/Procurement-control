using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;
using ProcurementControl.Views;

namespace ProcurementControl.ViewModels;

public partial class SupplierAnalyticsViewModel : ObservableObject
{
    private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");
    private static readonly string[] MonthShorts =
        { "янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек" };
    private PurchaseSnapshot? _snapshot;
    private SupplierAnalyticsService? _service;
    private readonly DispatcherTimer _autoRefreshTimer;
    private DataTable _ordersMatrix = new();
    private bool _volumesLoaded;

    [ObservableProperty] private string _manufacturerSearch = string.Empty;
    [ObservableProperty] private string _supplierSearch = string.Empty;
    [ObservableProperty] private string _dealSearch = string.Empty;
    [ObservableProperty] private string _batchSearch = string.Empty;
    [ObservableProperty] private string _dateFromText = string.Empty;
    [ObservableProperty] private string _dateToText = string.Empty;
    [ObservableProperty] private int _selectedTabIndex;
    [ObservableProperty] private string _summaryText = "Введите производителя и нажмите «Рекомендовать».";
    [ObservableProperty] private string _errorMessage = string.Empty;
    [ObservableProperty] private AnalyticsManufacturerOption? _selectedManufacturer;
    [ObservableProperty] private ManufacturerAliasRow? _selectedAlias;
    [ObservableProperty] private AnalyticsManufacturerOption? _selectedTargetManufacturer;
    [ObservableProperty] private AnalyticsBatchRow? _selectedBatch;
    [ObservableProperty] private SupplierOverviewRow? _selectedOverview;
    [ObservableProperty] private string _responseAuditText = "Выберите поставщика, чтобы проверить расчёт ответов.";

    public ObservableCollection<AnalyticsManufacturerOption> Manufacturers { get; } = new();
    public ObservableCollection<SupplierRecommendationRow> Recommendations { get; } = new();
    public ObservableCollection<SupplierRecommendationRow> StatisticsRows { get; } = new();
    public ObservableCollection<SupplierOverviewRow> OverviewRows { get; } = new();
    public ObservableCollection<SupplierResponseAuditRow> ResponseAuditRows { get; } = new();
    public ObservableCollection<AnalyticsMonthlyRow> MonthlyTrend { get; } = new();
    public ObservableCollection<ManufacturerAliasRow> Aliases { get; } = new();
    public ObservableCollection<AnalyticsBatchRow> Batches { get; } = new();
    public ObservableCollection<AnalyticsBatchRow> FilteredBatches { get; } = new();
    public ObservableCollection<AnalyticsBatchQuoteRow> BatchQuotes { get; } = new();

    // Вкладка «Объёмы»: сумма заказов ($) по поставщикам и периодам
    [ObservableProperty] private SupplierVolumeModeOption? _selectedVolumeModeOption;
    [ObservableProperty] private SupplierVolumeMode _selectedVolumeMode = SupplierVolumeMode.Month;
    [ObservableProperty] private string _volumeSearch = string.Empty;
    [ObservableProperty] private string _volumeSummary = "Выберите группировку и нажмите «Пересчитать».";
    public IReadOnlyList<SupplierVolumeModeOption> VolumeGranularities { get; } = new[]
    {
        new SupplierVolumeModeOption(SupplierVolumeMode.Month, "По месяцам"),
        new SupplierVolumeModeOption(SupplierVolumeMode.Quarter, "По кварталам"),
        new SupplierVolumeModeOption(SupplierVolumeMode.Year, "По годам")
    };

    /// <summary>Сводка «поставщики × периоды» для DataGrid (AutoGenerateColumns).</summary>
    public DataView OrdersView => _ordersMatrix.DefaultView;

    public SupplierAnalyticsViewModel()
    {
        _autoRefreshTimer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(5) };
        _autoRefreshTimer.Tick += (_, _) => RefreshInBackground();
        _autoRefreshTimer.Start();
        SelectedVolumeModeOption = VolumeGranularities.First();
        Load();
    }

    partial void OnSelectedManufacturerChanged(AnalyticsManufacturerOption? value)
    {
        if (value is not null) ManufacturerSearch = value.Name;
    }

    partial void OnSelectedBatchChanged(AnalyticsBatchRow? value)
    {
        BatchQuotes.Clear();
        if (_service is null || value is null) return;
        foreach (var row in _service.GetBatchQuotes(value.Id)) BatchQuotes.Add(row);
    }

    partial void OnBatchSearchChanged(string value) => ApplyBatchFilter();
    partial void OnSelectedOverviewChanged(SupplierOverviewRow? value) => LoadResponseAudit(value);

    partial void OnSelectedTabIndexChanged(int value)
    {
        // Вкладка «Объёмы» — последняя; данные лениво, при первом открытии.
        if (value == 4 && !_volumesLoaded) ReloadVolumes();
    }

    partial void OnSelectedVolumeModeOptionChanged(SupplierVolumeModeOption? value)
    {
        if (value is not null) SelectedVolumeMode = value.Value;
    }

    partial void OnSelectedVolumeModeChanged(SupplierVolumeMode value)
    {
        if (_volumesLoaded) ReloadVolumes();
    }

    partial void OnVolumeSearchChanged(string value) => ApplyVolumeFilter();

    [RelayCommand]
    private void ReloadVolumes()
    {
        try
        {
            if (_service is null) throw new InvalidOperationException("Данные ещё не загружены.");
            var buckets = BuildVolumeBuckets();
            var sums = new Dictionary<string, decimal[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var raw in _service.GetSupplierOrderAmounts(buckets[0].Start, buckets[^1].End))
            {
                var orderedAt = ParseOrderDate(raw.OrderedAt);
                if (orderedAt is null) continue;
                var index = IndexOfBucket(buckets, orderedAt.Value);
                if (index < 0) continue;
                if (!sums.TryGetValue(raw.Supplier, out var perBucket))
                {
                    perBucket = new decimal[buckets.Count];
                    sums[raw.Supplier] = perBucket;
                }
                perBucket[index] += MoneyText.ParseUsd(raw.AmountUsd);
            }

            var table = new DataTable("supplierVolumes");
            table.Columns.Add("Поставщик", typeof(string));
            foreach (var bucket in buckets) table.Columns.Add(bucket.Label, typeof(string));
            table.Columns.Add("Итого, $", typeof(string));

            var totals = new decimal[buckets.Count];
            var grandTotal = 0m;
            foreach (var perBucket in sums.Values)
            {
                for (var i = 0; i < perBucket.Length; i++)
                {
                    totals[i] += perBucket[i];
                    grandTotal += perBucket[i];
                }
            }
            table.Rows.Add(TotalRow(table, buckets, totals, grandTotal));
            foreach (var pair in sums.OrderByDescending(item => item.Value.Sum()).ThenBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
            {
                table.Rows.Add(SupplierRow(table, pair.Key, pair.Value, buckets));
            }

            _ordersMatrix = table;
            OnPropertyChanged(nameof(OrdersView));
            ApplyVolumeFilter();
            _volumesLoaded = true;
            VolumeSummary = $"Заказы {buckets[0].Label} — {buckets[^1].Label}: {sums.Count} поставщиков, {FormatUsd(grandTotal)} всего. " +
                "Суммы — USD из PI сделок, заказанных в периоде; компоненты без поставщика в БД не учитываются.";
            ErrorMessage = string.Empty;
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
    }

    /// <summary>Последние 12 месяцев / 8 кварталов / 3 года, включая текущий период.</summary>
    private List<(DateTime Start, DateTime End, string Label)> BuildVolumeBuckets()
    {
        var today = DateTime.Today;
        var buckets = new List<(DateTime Start, DateTime End, string Label)>();
        switch (SelectedVolumeMode)
        {
            case SupplierVolumeMode.Quarter:
            {
                var start = new DateTime(today.Year, (today.Month - 1) / 3 * 3 + 1, 1);
                for (var i = 7; i >= 0; i--)
                {
                    var bucketStart = start.AddMonths(-3 * i);
                    var quarter = (bucketStart.Month - 1) / 3 + 1;
                    buckets.Add((bucketStart, bucketStart.AddMonths(3), $"{quarter} кв {bucketStart.Year % 100:00}"));
                }
                break;
            }
            case SupplierVolumeMode.Year:
            {
                for (var i = 2; i >= 0; i--)
                {
                    var bucketStart = new DateTime(today.Year - i, 1, 1);
                    buckets.Add((bucketStart, bucketStart.AddYears(1), bucketStart.Year.ToString(Ru)));
                }
                break;
            }
            default:
            {
                var start = new DateTime(today.Year, today.Month, 1);
                for (var i = 11; i >= 0; i--)
                {
                    var bucketStart = start.AddMonths(-i);
                    buckets.Add((bucketStart, bucketStart.AddMonths(1), $"{MonthShorts[bucketStart.Month - 1]} {bucketStart.Year % 100:00}"));
                }
                break;
            }
        }
        return buckets;
    }

    private static DataRow TotalRow(DataTable table, List<(DateTime Start, DateTime End, string Label)> buckets, decimal[] totals, decimal grandTotal)
    {
        var row = table.NewRow();
        row["Поставщик"] = "Итого";
        for (var i = 0; i < totals.Length; i++) row[buckets[i].Label] = FormatCell(totals[i]);
        row["Итого, $"] = FormatCell(grandTotal);
        return row;
    }

    private static DataRow SupplierRow(DataTable table, string supplier, decimal[] perBucket, List<(DateTime Start, DateTime End, string Label)> buckets)
    {
        var row = table.NewRow();
        row["Поставщик"] = supplier;
        var total = 0m;
        for (var i = 0; i < perBucket.Length; i++)
        {
            total += perBucket[i];
            row[buckets[i].Label] = FormatCell(perBucket[i]);
        }
        row["Итого, $"] = FormatCell(total);
        return row;
    }

    private static string FormatCell(decimal value) => value == 0m ? string.Empty : FormatUsd(value);

    private static string FormatUsd(decimal value) => value.ToString("N0", Ru);

    private static DateTime? ParseOrderDate(string value)
    {
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) return date;
        if (DateTime.TryParse(value, Ru, DateTimeStyles.None, out date)) return date;
        return null;
    }

    private static int IndexOfBucket(List<(DateTime Start, DateTime End, string Label)> buckets, DateTime moment)
    {
        for (var i = 0; i < buckets.Count; i++)
        {
            if (moment >= buckets[i].Start && moment < buckets[i].End) return i;
        }
        return -1;
    }

    private void ApplyVolumeFilter()
    {
        var needle = VolumeSearch.Trim().Replace("'", "''");
        _ordersMatrix.DefaultView.RowFilter = needle.Length == 0 ? string.Empty : $"[Поставщик] LIKE '%{needle}%'";
    }

    [RelayCommand]
    private void Refresh()
    {
        Load();
        Recommend();
    }

    [RelayCommand]
    private void Recommend()
    {
        try
        {
            if (_service is null) throw new InvalidOperationException("Данные ещё не загружены.");
            Recommendations.Clear();
            var rows = _service.GetRecommendations(ManufacturerSearch, DateForSort(DateFromText), DateForSort(DateToText));
            foreach (var row in rows) Recommendations.Add(row);
            SummaryText = rows.Count == 0
                ? "Недостаточно данных для рекомендации."
                : rows.All(row => row.IsFallback)
                    ? "По производителю данных нет. Показан общий рейтинг поставщиков."
                    : $"Рекомендованы поставщики для «{ManufacturerSearch.Trim()}». Учитываются только сравнительные позиции полных когорт; winner-only записи исключены.";
            ErrorMessage = string.Empty;
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
    }

    [RelayCommand]
    private void LoadStatistics()
    {
        try
        {
            // Импорт документов меняет реальную БД из другой страницы. Перед
            // расчётом открываем свежий снимок, а не используем данные момента
            // открытия вкладки аналитики.
            Load();
            if (_service is null) throw new InvalidOperationException("Данные ещё не загружены.");
            StatisticsRows.Clear();
            OverviewRows.Clear();
            ResponseAuditRows.Clear();
            ResponseAuditText = "Выберите поставщика, чтобы проверить расчёт ответов.";
            MonthlyTrend.Clear();
            var filter = CurrentAnalyticsFilter();
            foreach (var row in _service.GetSupplierOverview(filter)) OverviewRows.Add(row);
            foreach (var row in _service.GetMonthlyTrend(ManufacturerSearch, SupplierSearch, DateForSort(DateFromText), DateForSort(DateToText), DealSearch)) MonthlyTrend.Add(row);
            SummaryText = string.IsNullOrWhiteSpace(ManufacturerSearch)
                ? $"Прозрачный обзор по {OverviewRows.Count} поставщикам. Оценка 0–100 остаётся на вкладке «Рекомендация»."
                : $"Обзор производителя «{ManufacturerSearch.Trim()}»: {OverviewRows.Count} поставщиков.";
            ErrorMessage = string.Empty;
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
    }

    [RelayCommand]
    private void ExportOverview()
    {
        try
        {
            if (OverviewRows.Count == 0) LoadStatistics();
            if (OverviewRows.Count == 0) throw new InvalidOperationException("Нет строк для экспорта.");
            var path = SupplierOverviewExporter.ExportToExcel(OverviewRows.ToList());
            SummaryText = "Обзор экспортирован: " + path;
            ErrorMessage = string.Empty;
        }
        catch (Exception ex) { ErrorMessage = ex.Message; }
    }

    [RelayCommand]
    private void CreateManufacturer()
    {
        try
        {
            if (_service is null) return;
            var dialog = new InputDialogWindow("Производитель", "Каноническое название производителя");
            if (dialog.ShowDialog() != true || string.IsNullOrWhiteSpace(dialog.Value)) return;
            var unresolvedAliasId = SelectedAlias?.IsUnresolved == true ? SelectedAlias.Id : 0;
            _service.CreateManufacturer(dialog.Value);
            Load();
            if (unresolvedAliasId > 0 && _service is not null)
            {
                var target = Manufacturers.FirstOrDefault(item => string.Equals(item.Name, dialog.Value.Trim(), StringComparison.OrdinalIgnoreCase));
                if (target is not null)
                {
                    _service.AssignAlias(unresolvedAliasId, target.Id);
                    Load();
                }
            }
            ManufacturerSearch = dialog.Value.Trim();
            Recommend();
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Производители"); }
    }

    [RelayCommand]
    private void AssignAlias()
    {
        try
        {
            if (_service is null || SelectedAlias is null) throw new InvalidOperationException("Сначала выберите строку алиаса.");
            if (SelectedTargetManufacturer is null) throw new InvalidOperationException("Выберите канонического производителя в поле справа.");
            _service.AssignAlias(SelectedAlias.Id, SelectedTargetManufacturer.Id);
            Load();
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Алиасы производителей"); }
    }

    [RelayCommand]
    private void MergeManufacturer()
    {
        try
        {
            if (_service is null || SelectedAlias is null || SelectedAlias.ManufacturerId <= 0)
                throw new InvalidOperationException("Выберите алиас уже привязанного производителя.");
            var dialog = new InputDialogWindow("Объединить производителя", "Каноническое название целевого производителя");
            if (dialog.ShowDialog() != true || string.IsNullOrWhiteSpace(dialog.Value)) return;
            var target = Manufacturers.FirstOrDefault(item => string.Equals(item.Name, dialog.Value.Trim(), StringComparison.OrdinalIgnoreCase));
            if (target is null) throw new InvalidOperationException("Целевой производитель не найден.");
            _service.MergeManufacturers(SelectedAlias.ManufacturerId, target.Id);
            Load();
        }
        catch (Exception ex) { MessageBox.Show(ex.Message, "Объединение производителей"); }
    }

    private void Load()
    {
        try
        {
            _snapshot?.Dispose();
            _snapshot = PurchaseSnapshot.Create();
            _service = new SupplierAnalyticsService(_snapshot.Connection);
            Manufacturers.Clear();
            foreach (var option in _service.GetManufacturerOptions().GroupBy(item => item.Id).Select(group => group.First())) Manufacturers.Add(option);
            Aliases.Clear();
            foreach (var alias in _service.GetAliasRows()) Aliases.Add(alias);
            Batches.Clear();
            foreach (var batch in _service.GetBatches()) Batches.Add(batch);
            ApplyBatchFilter();
            ErrorMessage = string.Empty;
        }
        catch (Exception ex) { ErrorMessage = "Не удалось прочитать статистику: " + ex.Message; }
    }

    private void ApplyBatchFilter()
    {
        var selectedId = SelectedBatch?.Id;
        FilteredBatches.Clear();
        var search = BatchSearch.Trim();
        foreach (var batch in Batches.Where(batch => string.IsNullOrWhiteSpace(search)
                     || batch.DealNumber.Contains(search, StringComparison.OrdinalIgnoreCase)))
            FilteredBatches.Add(batch);
        SelectedBatch = FilteredBatches.FirstOrDefault(batch => batch.Id == selectedId);
    }

    private AnalyticsFilter CurrentAnalyticsFilter()
        => new(DateForSort(DateFromText), DateForSort(DateToText), ManufacturerSearch, SupplierSearch, DealSearch);

    private void LoadResponseAudit(SupplierOverviewRow? row)
    {
        ResponseAuditRows.Clear();
        if (_service is null || row is null) return;
        foreach (var item in _service.GetSupplierResponseAudit(row.Supplier, CurrentAnalyticsFilter())) ResponseAuditRows.Add(item);
        ResponseAuditText = ResponseAuditRows.Count == 0
            ? $"Для «{row.Supplier}» нет позиций, включённых в частоту ответа. Неполные PI и ручные импорты исключены."
            : $"Проверка ответов «{row.Supplier}»: {ResponseAuditRows.Count(item => item.HasQuote)} из {ResponseAuditRows.Count} позиций. Столбец «Основание» показывает тип знаменателя.";
    }

    private void RefreshInBackground()
    {
        try
        {
            Load();
            if (SelectedTabIndex == 1) LoadStatistics();
            else if (!string.IsNullOrWhiteSpace(ManufacturerSearch)) Recommend();
        }
        catch
        {
            // Фоновое обновление не должно прерывать работу пользователя.
        }
    }

    private static string DateForSort(string value)
        => DateTime.TryParse(value, out var date) ? date.ToString("yyyy-MM-dd") : string.Empty;
}
