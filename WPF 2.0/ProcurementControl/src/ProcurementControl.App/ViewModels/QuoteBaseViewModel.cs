using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Read-only страница «База квот» — перенос $quoteBasePage (RRFQComparer.ps1,
/// строки 2554-2690, 4325-4458, 5002-5042). Кнопки изменения данных (удаление,
/// очистка, загрузка Globalist) в порт не входят. Снапшот БД живёт всё время
/// страницы; исходная база не изменяется.
/// </summary>
public partial class QuoteBaseViewModel : ObservableObject
{
    private const int QuoteLimit = 2000;
    private const int GlobalistLimit = 20000;
    private const int ExportLimit = 100000;

    private PurchaseSnapshot? _snapshot;
    private QuoteBaseRepository? _repository;
    private readonly DispatcherTimer _searchTimer;

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private string _dateFromText = string.Empty;

    [ObservableProperty]
    private string _dateToText = string.Empty;

    [ObservableProperty]
    private int _selectedTabIndex;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    public ObservableCollection<QuoteHistoryRow> QuoteRows { get; } = new();

    public ObservableCollection<GlobalistRow> GlobalistRows { get; } = new();

    public string ReadOnlyNotice { get; } =
        "Режим чтения: изменение доступно в основной версии приложения";

    public QuoteBaseViewModel()
    {
        _searchTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(350),
        };
        _searchTimer.Tick += (_, _) =>
        {
            _searchTimer.Stop();
            ReloadQuotes();
            if (SelectedTabIndex == 1)
            {
                ReloadGlobalist();
            }
        };

        Load();
    }

    /// <summary>Аналог Add-DebouncedTextChanged $txtQuoteSearch: перезапуск таймера.</summary>
    partial void OnSearchTextChanged(string value)
    {
        _searchTimer.Stop();
        _searchTimer.Start();
    }

    /// <summary>Аналог $quoteBaseTabs.SelectedIndexChanged: вкладка Globalist читается по открытию.</summary>
    partial void OnSelectedTabIndexChanged(int value)
    {
        if (value == 1)
        {
            ReloadGlobalist();
        }
    }

    /// <summary>Аналог $btnRefreshQuoteBase: перечитать снапшот и сетку «Наши квоты».</summary>
    [RelayCommand]
    private void Refresh() => Load();

    /// <summary>Аналог $btnExportQuoteBase: экспорт отфильтрованных квот в новый файл.</summary>
    [RelayCommand]
    private void Export()
    {
        try
        {
            if (_repository is null)
            {
                throw new InvalidOperationException("Данные ещё не загружены.");
            }

            var rows = _repository.GetQuoteHistory(SearchText, DateForSort(DateFromText), DateForSort(DateToText), ExportLimit);
            var path = QuoteBaseExporter.ExportToExcel(rows);
            MessageBox.Show("Экспорт готов:\r\n" + path, "База квот");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "База квот");
        }
    }

    private void Load()
    {
        try
        {
            _snapshot?.Dispose();
            _snapshot = PurchaseSnapshot.Create();
            _repository = new QuoteBaseRepository(_snapshot.Connection);

            ErrorMessage = string.Empty;
            ReloadQuotes();
            if (SelectedTabIndex == 1)
            {
                ReloadGlobalist();
            }
        }
        catch (Exception ex)
        {
            _repository = null;
            ErrorMessage = "Не удалось прочитать данные: " + ex.Message;
        }
    }

    /// <summary>Аналог Refresh-QuoteBase: запрос с текущим поиском и датами.</summary>
    public void ReloadQuotes()
    {
        QuoteRows.Clear();
        if (_repository is null)
        {
            return;
        }

        try
        {
            foreach (var row in _repository.GetQuoteHistory(SearchText, DateForSort(DateFromText), DateForSort(DateToText), QuoteLimit))
            {
                QuoteRows.Add(row);
            }

            ErrorMessage = string.Empty;
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось прочитать базу квот: " + ex.Message;
        }
    }

    /// <summary>Аналог Refresh-GlobalistQuotes.</summary>
    public void ReloadGlobalist()
    {
        GlobalistRows.Clear();
        if (_repository is null)
        {
            return;
        }

        try
        {
            foreach (var row in _repository.GetGlobalistQuotes(SearchText, GlobalistLimit))
            {
                GlobalistRows.Add(row);
            }

            ErrorMessage = string.Empty;
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось прочитать Globalist: " + ex.Message;
        }
    }

    /// <summary>Аналог Convert-PurchaseDateForSort: текст даты → «yyyy-MM-dd» или пусто.</summary>
    private static string DateForSort(string text)
    {
        var parsed = PurchaseFormatting.ParseDate(text);
        return parsed is null ? string.Empty : parsed.Value.ToString("yyyy-MM-dd");
    }
}
