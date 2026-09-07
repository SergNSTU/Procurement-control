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
    [ObservableProperty] private string _dateFromText = string.Empty;
    [ObservableProperty] private string _dateToText = string.Empty;
    [ObservableProperty] private int _selectedTabIndex;
    [ObservableProperty] private string _summaryText = "Введите производителя и нажмите «Рекомендовать».";
    [ObservableProperty] private string _errorMessage = string.Empty;
    [ObservableProperty] private AnalyticsManufacturerOption? _selectedManufacturer;
    [ObservableProperty] private ManufacturerAliasRow? _selectedAlias;
    [ObservableProperty] private AnalyticsManufacturerOption? _selectedTargetManufacturer;
    [ObservableProperty] private AnalyticsBatchRow? _selectedBatch;

    public ObservableCollection<AnalyticsManufacturerOption> Manufacturers { get; } = new();
    public ObservableCollection<SupplierRecommendationRow> Recommendations { get; } = new();
    public ObservableCollection<SupplierRecommendationRow> StatisticsRows { get; } = new();
    public ObservableCollection<AnalyticsMonthlyRow> MonthlyTrend { get; } = new();
    public ObservableCollection<ManufacturerAliasRow> Aliases { get; } = new();
    public ObservableCollection<AnalyticsBatchRow> Batches { get; } = new();
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
            if (_service is null) throw new InvalidOperationException("Данные ещё не загружены.");
            StatisticsRows.Clear();
            MonthlyTrend.Clear();
            foreach (var row in _service.GetStatistics(ManufacturerSearch, DateForSort(DateFromText), DateForSort(DateToText))
                .Where(row => string.IsNullOrWhiteSpace(SupplierSearch) || row.Supplier.Contains(SupplierSearch.Trim(), StringComparison.OrdinalIgnoreCase))) StatisticsRows.Add(row);
            foreach (var row in _service.GetMonthlyTrend(ManufacturerSearch, SupplierSearch, DateForSort(DateFromText), DateForSort(DateToText))) MonthlyTrend.Add(row);
            SummaryText = string.IsNullOrWhiteSpace(ManufacturerSearch)
                ? $"Общий рейтинг по {StatisticsRows.Count} поставщикам."
                : $"Статистика производителя «{ManufacturerSearch.Trim()}»: {StatisticsRows.Count} поставщиков.";
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
            ErrorMessage = string.Empty;
        }
        catch (Exception ex) { ErrorMessage = "Не удалось прочитать статистику: " + ex.Message; }
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
