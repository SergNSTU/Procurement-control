using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;
using ProcurementControl.Views;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «База квот» — перенос $quoteBasePage (RRFQComparer.ps1,
/// строки 2554-2690, 4325-4458, 5002-5042). Чтение идёт из снапшота БД;
/// удаление/очистка/импорт Globalist пишут в исходную базу через корзину.
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

    /// <summary>Выбранные строки сетки квот — синхронизируются из SelectionChanged представления.</summary>
    public ObservableCollection<QuoteHistoryRow> SelectedQuotes { get; } = new();

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

        SelectedQuotes.CollectionChanged += (_, _) => DeleteSelectedCommand.NotifyCanExecuteChanged();

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

    [RelayCommand]
    private void ImportQuotes()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Выберите RRFQ с квотами",
            Filter = "RRFQ/Excel (*.xlsx;*.xls;*.xlsm)|*.xlsx;*.xls;*.xlsm|Все файлы (*.*)|*.*",
            Multiselect = true,
        };
        if (dialog.ShowDialog() != true) return;

        try
        {
            var sources = dialog.FileNames.Select(path => (Path: path, Supplier: string.Empty)).ToList();
            var quotes = RrfqQuoteImportService.ReadQuotes(sources);
            if (quotes.Count == 0)
            {
                MessageBox.Show("В выбранных RRFQ не найдено квот.", "Импорт квот");
                return;
            }

            var confirmation = new QuoteImportConfirmationWindow(quotes, dialog.FileNames)
            {
                Owner = Application.Current.MainWindow,
            };
            if (confirmation.ShowDialog() != true) return;

            var saved = RrfqQuoteImportService.SaveQuotes(quotes, dialog.FileNames, "Импорт из базы квот");
            Load();
            MessageBox.Show($"Импортировано квот: {saved}\r\nФайлов: {dialog.FileNames.Length}", "Импорт квот");
        }
        catch (Exception ex)
        {
            MessageBox.Show("Не удалось импортировать RRFQ:\r\n" + ex.Message, "Импорт квот");
        }
    }

    /// <summary>
    /// Порт Delete-SelectedQuoteBaseRows (RRFQComparer.ps1, строки 4429-4446):
    /// выбранные квоты переносятся в корзину, затем сетка перечитывается.
    /// </summary>
    [RelayCommand(CanExecute = nameof(HasSelectedQuotes))]
    private void DeleteSelected()
    {
        try
        {
            var ids = SelectedQuotes.Select(row => row.Id).Where(id => id > 0).ToList();
            if (ids.Count == 0)
            {
                throw new InvalidOperationException("Не удалось определить выбранные квоты.");
            }

            var answer = MessageBox.Show(
                "Переместить выбранные квоты в корзину: " + ids.Count + "?",
                "База квот", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes)
            {
                return;
            }

            PurchaseWriteRepository.RemoveQuoteHistoryItems(ids);
            Load();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "База квот");
        }
    }

    /// <summary>
    /// Порт Clear-AllQuoteBaseRows (RRFQComparer.ps1, строки 4448-4458): вся база квот
    /// сохраняется в корзину, затем таблицы очищаются.
    /// </summary>
    [RelayCommand]
    private void ClearAll()
    {
        try
        {
            var answer = MessageBox.Show(
                "Очистить всю базу квот? Перед этим будет создана запись в корзине, откуда данные можно восстановить.",
                "База квот", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (answer != MessageBoxResult.Yes)
            {
                return;
            }

            PurchaseWriteRepository.ClearQuoteHistory();
            Load();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "База квот");
        }
    }

    /// <summary>
    /// Порт $btnImportGlobalist.Add_Click (RRFQComparer.ps1, строки 5005-5017):
    /// выбор файла, полный разбор и замена таблицы globalist_quotes.
    /// </summary>
    [RelayCommand]
    private void ImportGlobalist()
    {
        try
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Файлы Globalist (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*",
                Title = "Выберите файл Globalist",
            };
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            var count = GlobalistImporter.Import(dialog.FileName);
            SelectedTabIndex = 1;
            Load();
            ToastService.Show("Файлов Globalist загружено: " + count, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Ошибка Globalist");
        }
    }

    private bool HasSelectedQuotes() => SelectedQuotes.Count > 0;

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
