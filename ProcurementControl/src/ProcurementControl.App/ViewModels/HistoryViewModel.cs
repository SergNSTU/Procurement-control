using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «История» — перенос $historyPage (RRFQComparer.ps1,
/// строки 2692-2767, 4368-4381, 4460-4470, 5044-5063). Чтение идёт из
/// снапшота; очистка журнала — через <see cref="PurchaseWriteRepository"/>
/// с сохранением копии в корзине.
/// </summary>
public partial class HistoryViewModel : ObservableObject
{
    private const int LogLimit = 500;

    private readonly Func<long?> _selectedDealIdProvider;
    private PurchaseSnapshot? _snapshot;
    private HistoryRepository? _repository;
    private readonly DispatcherTimer _searchTimer;
    private long _filterDealId;

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    private string _statusText = string.Empty;

    [ObservableProperty]
    private string _filterNotice = string.Empty;

    public ObservableCollection<ActivityLogRow> Rows { get; } = new();

    public string Hint { get; } =
        "История фиксирует основные действия: сделки, поставщиков, документы, импорт RRFQ и изменения статусов.";

    /// <param name="selectedDealIdProvider">
    /// Аналог Get-SelectedDealId: возвращает сделку, выбранную в «Контроле
    /// закупки», или 0, если раздел ещё не открывали / выбор пуст.
    /// </param>
    public HistoryViewModel(Func<long?>? selectedDealIdProvider = null)
    {
        _selectedDealIdProvider = selectedDealIdProvider ?? (() => null);

        _searchTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(350),
        };
        _searchTimer.Tick += (_, _) =>
        {
            _searchTimer.Stop();
            Reload();
        };

        Load();
    }

    /// <summary>Аналог Add-DebouncedTextChanged $txtHistorySearch.</summary>
    partial void OnSearchTextChanged(string value)
    {
        _searchTimer.Stop();
        _searchTimer.Start();
    }

    /// <summary>Аналог $btnRefreshHistory.</summary>
    [RelayCommand]
    private void Refresh()
    {
        _filterDealId = 0;
        FilterNotice = string.Empty;
        Load();
    }

    /// <summary>Аналог $btnHistorySelectedDeal: фильтр по сделке из «Контроля закупки».</summary>
    [RelayCommand]
    private void FilterBySelectedDeal()
    {
        _filterDealId = _selectedDealIdProvider() ?? 0;
        FilterNotice = _filterDealId > 0
            ? "Фильтр: журнал выбранной сделки (сбрасывается кнопкой «Обновить»)."
            : "В разделе «Контроль закупки» не выбрана сделка — показан весь журнал.";
        Load();
    }

    /// <summary>Аналог $btnClearHistory + Clear-AllHistoryRows.</summary>
    [RelayCommand]
    private void ClearHistory()
    {
        var answer = MessageBox.Show(
            "Очистить всю историю действий? Перед этим будет создана запись в корзине, откуда данные можно восстановить.",
            "История",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);
        if (answer != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            PurchaseWriteRepository.ClearActivityLog();
            ErrorMessage = string.Empty;
            StatusText = "Журнал очищен.";
            ToastService.Show(StatusText, ToastKind.Success);
            Load();
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось очистить журнал: " + ex.Message;
        }
    }

    private void Load()
    {
        try
        {
            _snapshot?.Dispose();
            _snapshot = PurchaseSnapshot.Create();
            _repository = new HistoryRepository(_snapshot.Connection);

            ErrorMessage = string.Empty;
            Reload();
        }
        catch (Exception ex)
        {
            _repository = null;
            Rows.Clear();
            ErrorMessage = "Не удалось прочитать данные: " + ex.Message;
        }
    }

    /// <summary>Аналог Refresh-History: перечитать сетку с текущим поиском.</summary>
    private void Reload()
    {
        Rows.Clear();
        if (_repository is null)
        {
            return;
        }

        try
        {
            foreach (var row in _repository.GetActivityLog(_filterDealId, LogLimit, SearchText))
            {
                Rows.Add(row);
            }

            StatusText = "Записей: " + Rows.Count;
            ErrorMessage = string.Empty;
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось прочитать журнал: " + ex.Message;
        }
    }
}
