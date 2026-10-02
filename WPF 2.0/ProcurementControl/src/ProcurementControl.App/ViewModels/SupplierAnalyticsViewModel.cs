using System.Collections.ObjectModel;
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
    private PurchaseSnapshot? _snapshot;
    private SupplierAnalyticsService? _service;
    private readonly DispatcherTimer _autoRefreshTimer;

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

    public SupplierAnalyticsViewModel()
    {
        _autoRefreshTimer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(5) };
        _autoRefreshTimer.Tick += (_, _) => RefreshInBackground();
        _autoRefreshTimer.Start();
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
