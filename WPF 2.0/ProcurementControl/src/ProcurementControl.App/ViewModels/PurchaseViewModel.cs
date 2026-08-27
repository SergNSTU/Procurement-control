using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>Пункт комбобокса фильтров.</summary>
public sealed record FilterOption(DealFilter Value, string Label)
{
    public override string ToString() => Label;
}

/// <summary>
/// Read-only страница «Контроль закупки». Снапшот БД живёт всё время страницы,
/// чтобы поставщики выбранной сделки читались без повторного копирования файла.
/// Исходная база не изменяется.
/// </summary>
public partial class PurchaseViewModel : ObservableObject
{
    private PurchaseSnapshot? _snapshot;
    private PurchaseRepository? _repository;
    private List<PurchaseDealRow> _allDeals = new();

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private FilterOption _selectedFilter;

    [ObservableProperty]
    private PurchaseDealRow? _selectedDeal;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    private string _countText = string.Empty;

    // Карточки метрик.
    [ObservableProperty] private int _activeCount;
    [ObservableProperty] private int _overdueCount;
    [ObservableProperty] private int _dueTodayCount;
    [ObservableProperty] private int _attentionCount;

    public string ReadOnlyNotice { get; } =
        "Режим чтения: изменение доступно в основной версии приложения";

    public IReadOnlyList<FilterOption> Filters { get; } =
        Enum.GetValues<DealFilter>().Select(f => new FilterOption(f, f.Label())).ToList();

    public ObservableCollection<PurchaseDealRow> Deals { get; } = new();

    public ObservableCollection<SupplierRow> Suppliers { get; } = new();

    public PurchaseViewModel()
    {
        _selectedFilter = Filters[0];
        Load();
    }

    partial void OnSearchTextChanged(string value) => ApplyView();

    partial void OnSelectedFilterChanged(FilterOption value) => ApplyView();

    partial void OnSelectedDealChanged(PurchaseDealRow? value) => LoadSuppliers();

    [RelayCommand]
    private void Reload() => Load();

    private void Load()
    {
        try
        {
            _snapshot?.Dispose();
            _snapshot = PurchaseSnapshot.Create();
            _repository = new PurchaseRepository(_snapshot.Connection);

            _allDeals = _repository.GetDeals();
            var taskKeys = _repository.GetCockpitTaskKeys();

            var metrics = PurchaseLogic.ComputeMetrics(_allDeals, taskKeys);
            ActiveCount = metrics.Active;
            OverdueCount = metrics.Overdue;
            DueTodayCount = metrics.DueToday;
            AttentionCount = metrics.Attention;

            ErrorMessage = string.Empty;
            ApplyView();
            LoadSuppliers();
        }
        catch (Exception ex)
        {
            _repository = null;
            ErrorMessage = "Не удалось прочитать данные: " + ex.Message;
        }
    }

    private void ApplyView()
    {
        var selectedId = SelectedDeal?.Id;

        var filtered = PurchaseLogic.FilterDeals(_allDeals, SelectedFilter.Value);
        var visible = PurchaseLogic.SearchDeals(filtered, SearchText);

        Deals.Clear();
        foreach (var row in visible)
        {
            Deals.Add(row);
        }

        CountText = "Всего: " + Deals.Count;
        SelectedDeal = Deals.FirstOrDefault(r => r.Id == selectedId);
    }

    private void LoadSuppliers()
    {
        Suppliers.Clear();
        var deal = SelectedDeal;
        if (deal is null || _repository is null)
        {
            return;
        }

        try
        {
            foreach (var supplier in _repository.GetSuppliers(deal.Id))
            {
                Suppliers.Add(supplier);
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось прочитать поставщиков: " + ex.Message;
        }
    }
}
