using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using ProcurementControl.Models;
using ProcurementControl.Services;
using ProcurementControl.Views;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Сравнение RRFQ» — перенос $rrfqPage (RRFQComparer.ps1,
/// строки 980-1380, 4572-4702, 5926-6062). Анализ выполняется в отдельном
/// STA-потоке, чтобы не замораживать интерфейс. База данных не используется:
/// Save-QuoteHistory и сохранение ширины колонок в порт не входят.
/// </summary>
public partial class RrfqViewModel : ObservableObject
{
    private RrfqAnalysis? _analysis;

    [ObservableProperty]
    private string _rfqPath = string.Empty;

    [ObservableProperty]
    private bool _isLeadTimePriority;

    [ObservableProperty]
    private int _selectedFilterIndex;

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private string _logText = "Выберите RFQ и добавьте файлы поставщиков.\r\nПосле preview здесь появится сводка.";

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private bool _canSaveResult;

    [ObservableProperty]
    private Decision? _selectedDecision;

    [ObservableProperty]
    private QuoteDetail? _selectedQuoteDetail;

    public ObservableCollection<SupplierFileEntry> Suppliers { get; } = new();

    public ObservableCollection<Decision> PreviewRows { get; } = new();

    public ObservableCollection<QuoteDetail> QuoteDetails { get; } = new();

    /// <summary>Пункты фильтра — как $cmbFilter (строки 1207-1211).</summary>
    public IReadOnlyList<string> FilterOptions { get; } = new[]
    {
        "Все строки",
        "Только победители",
        "Без квот",
        "Предупреждения",
        "Ручная проверка",
    };

    /// <summary>Радио «Цена»: обратная сторона IsLeadTimePriority.</summary>
    public bool IsPricePriority
    {
        get => !IsLeadTimePriority;
        set => IsLeadTimePriority = !value;
    }

    partial void OnIsLeadTimePriorityChanged(bool value)
    {
        OnPropertyChanged(nameof(IsPricePriority));
    }

    partial void OnSelectedFilterIndexChanged(int value) => RefreshPreview();

    partial void OnSearchTextChanged(string value) => RefreshPreview();

    partial void OnSelectedDecisionChanged(Decision? value) => RefreshQuoteDetails();

    partial void OnIsBusyChanged(bool value)
    {
        RunPreviewCommand.NotifyCanExecuteChanged();
        SaveResultCommand.NotifyCanExecuteChanged();
        OpenManualMatchCommand.NotifyCanExecuteChanged();
    }

    // ----- 1. Выбор файла -----

    [RelayCommand]
    private void ChooseRfq()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*",
            Multiselect = false,
        };
        if (dialog.ShowDialog() == true)
        {
            RfqPath = dialog.FileName;
        }
    }

    /// <summary>Аналог $btnAdd: имена поставщиков угадываются из имён файлов.</summary>
    [RelayCommand]
    private void AddSuppliers()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*",
            Multiselect = true,
        };
        if (dialog.ShowDialog() != true)
        {
            return;
        }

        foreach (var file in dialog.FileNames)
        {
            Suppliers.Add(new SupplierFileEntry
            {
                Path = file,
                Supplier = RrfqConfigService.GetSupplierFromFileName(file),
            });
        }
    }

    /// <summary>Аналог $btnRemove: выбранные строки списка передаёт представление.</summary>
    [RelayCommand]
    private void RemoveSuppliers(IEnumerable<SupplierFileEntry>? entries)
    {
        if (entries is null)
        {
            return;
        }

        foreach (var entry in entries.ToList())
        {
            Suppliers.Remove(entry);
        }
    }

    // ----- 2. Параметры и запуск -----

    /// <summary>Аналог $btnPreview: анализ на отдельном STA-потоке (там живёт Excel).</summary>
    [RelayCommand(CanExecute = nameof(CanRunPreview))]
    private void RunPreview()
    {
        try
        {
            var suppliers = Suppliers
                .Where(s => !string.IsNullOrWhiteSpace(s.Path))
                .Select(s => new SupplierFileEntry
                {
                    Path = s.Path,
                    Supplier = string.IsNullOrWhiteSpace(s.Supplier)
                        ? RrfqConfigService.GetSupplierFromFileName(s.Path)
                        : s.Supplier,
                })
                .ToList();

            if (string.IsNullOrWhiteSpace(RfqPath))
            {
                throw new InvalidOperationException("Выберите RFQ.");
            }

            if (suppliers.Count == 0)
            {
                throw new InvalidOperationException("Добавьте хотя бы один файл поставщика.");
            }

            var priorityMode = IsLeadTimePriority ? "LeadTime" : "Price";
            IsBusy = true;
            LogText = "Читаю Excel-файлы...";

            StartStaJob(
                () => RrfqEngine.Analyze(RfqPath, suppliers, priorityMode),
                (error, result) =>
                {
                    IsBusy = false;
                    if (error is not null)
                    {
                        MessageBox.Show(error.Message, "Ошибка");
                        return;
                    }

                    _analysis = (RrfqAnalysis)result!;
                    RefreshPreview();
                    CanSaveResult = true;
                    SaveResultCommand.NotifyCanExecuteChanged();
                    OpenManualMatchCommand.NotifyCanExecuteChanged();
                });
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Ошибка");
        }
    }

    private bool CanRunPreview() => !IsBusy;

    /// <summary>Аналог $btnSave: копия RFQ с победителями, тоже на STA-потоке.</summary>
    [RelayCommand(CanExecute = nameof(CanSave))]
    private void SaveResult()
    {
        try
        {
            if (_analysis is null)
            {
                throw new InvalidOperationException("Сначала сформируйте preview.");
            }

            var analysis = _analysis;
            IsBusy = true;
            LogText += "\r\n\r\nСоздаю результирующий RRFQ...";

            StartStaJob(
                () => RrfqResultWriter.WriteResultWorkbook(analysis),
                (error, result) =>
                {
                    IsBusy = false;
                    if (error is not null)
                    {
                        MessageBox.Show(error.Message, "Ошибка сохранения");
                        return;
                    }

                    var path = (string)result!;
                    LogText += "\r\nГотово: " + path;
                    MessageBox.Show("Готово.\r\n" + path, "RRFQ создан");
                });
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Ошибка сохранения");
        }
    }

    private bool CanSave() => CanSaveResult && !IsBusy;

    // ----- 3. Preview и ручная корректировка -----

    /// <summary>Аналог Refresh-CurrentPreview: фильтр, поиск и восстановление выбора.</summary>
    public void RefreshPreview()
    {
        if (_analysis is null)
        {
            return;
        }

        var selectedId = SelectedDecision?.Id ?? string.Empty;

        PreviewRows.Clear();
        foreach (var decision in GetFilteredDecisions())
        {
            PreviewRows.Add(decision);
        }

        Decision? restored = null;
        if (!string.IsNullOrWhiteSpace(selectedId))
        {
            restored = PreviewRows.FirstOrDefault(d => d.Id == selectedId);
        }

        SelectedDecision = restored;
        RefreshLogSummary();
    }

    /// <summary>Аналог Get-FilteredDecisions (RrfqUi.ps1, строки 52-100).</summary>
    private List<Decision> GetFilteredDecisions()
    {
        if (_analysis is null)
        {
            return new List<Decision>();
        }

        IEnumerable<Decision> items = _analysis.Decisions;
        switch (SelectedFilterIndex)
        {
            case 1: // Только победители
                items = items.Where(d => d.Winner is not null);
                break;
            case 2: // Без квот
                items = items.Where(d => d.Status == "Нет квот");
                break;
            case 3: // Предупреждения
                items = items.Where(d => !string.IsNullOrWhiteSpace(d.Warning) || d.Status == "Нет сравнимой цены");
                break;
            case 4: // Ручная проверка
                var manualIds = new HashSet<string>(StringComparer.Ordinal);
                foreach (var quote in RrfqEngine.GetUnresolvedQuotes(_analysis))
                {
                    foreach (var id in quote.CandidateIds)
                    {
                        manualIds.Add(id);
                    }

                    if (!string.IsNullOrWhiteSpace(quote.TargetId))
                    {
                        manualIds.Add(quote.TargetId!);
                    }
                }

                foreach (var decision in _analysis.Decisions)
                {
                    if (decision.IsManualWinner)
                    {
                        manualIds.Add(decision.Id);
                        continue;
                    }

                    if (decision.Quotes.Any(q => q.MatchStatus.StartsWith("Manual", StringComparison.Ordinal)))
                    {
                        manualIds.Add(decision.Id);
                    }
                }

                items = items.Where(d => manualIds.Contains(d.Id));
                break;
        }

        if (!string.IsNullOrWhiteSpace(SearchText))
        {
            var needle = SearchText.Trim().ToLowerInvariant();
            items = items.Where(d =>
                d.Value.ToLowerInvariant().Contains(needle)
                || d.WinnerPN.ToLowerInvariant().Contains(needle)
                || d.RussianRemark.ToLowerInvariant().Contains(needle)
                || d.ChinaRemark.ToLowerInvariant().Contains(needle)
                || d.WinnerSupplier.ToLowerInvariant().Contains(needle)
                || d.QuotesSummary.ToLowerInvariant().Contains(needle));
        }

        return items.ToList();
    }

    /// <summary>Сводка в лог — аналог конца Refresh-PreviewGrid (RrfqUi.ps1, строки 46-49).</summary>
    private void RefreshLogSummary()
    {
        if (_analysis is null)
        {
            return;
        }

        var unresolved = RrfqEngine.GetUnresolvedQuotes(_analysis);
        var okCount = _analysis.Decisions.Count(d => d.Winner is not null);
        var noQuoteCount = _analysis.Decisions.Count(d => d.Status == "Нет квот");
        LogText = $"RFQ позиций: {_analysis.RfqRows.Count}\r\n"
            + $"Квот прочитано: {_analysis.Quotes.Count}\r\n"
            + $"Победителей: {okCount}\r\n"
            + $"Без квот: {noQuoteCount}\r\n"
            + $"Требуют ручной проверки: {unresolved.Count}";
    }

    /// <summary>Аналог Refresh-DecisionDetails (RrfqUi.ps1, строки 169-202).</summary>
    private void RefreshQuoteDetails()
    {
        QuoteDetails.Clear();
        var decision = SelectedDecision;
        if (decision is null)
        {
            return;
        }

        foreach (var quote in decision.Quotes)
        {
            QuoteDetails.Add(new QuoteDetail
            {
                Quote = quote,
                IsWinner = decision.Winner is not null && quote.Id == decision.Winner.Id,
            });
        }
    }

    /// <summary>Аналог $btnIncludeSelected.</summary>
    [RelayCommand]
    private void IncludeSelected()
    {
        if (SelectedDecision is null)
        {
            return;
        }

        SelectedDecision.Include = true;
    }

    /// <summary>Аналог $btnExcludeSelected.</summary>
    [RelayCommand]
    private void ExcludeSelected()
    {
        if (SelectedDecision is null)
        {
            return;
        }

        SelectedDecision.Include = false;
    }

    /// <summary>Аналог Set-SelectedQuoteAsWinner + Set-RfqRowManualWinner.</summary>
    [RelayCommand]
    private void SetWinner()
    {
        try
        {
            if (_analysis is null || SelectedDecision is null)
            {
                throw new InvalidOperationException("Выберите строку preview.");
            }

            var decision = SelectedDecision;
            var quote = SelectedQuoteDetail?.Quote ?? throw new InvalidOperationException("Выберите квоту в таблице справа.");

            if (quote.TargetId != decision.Id)
            {
                throw new InvalidOperationException("Выбранная квота не относится к этой строке RFQ.");
            }

            if (!RrfqEngine.IsQuoteSelectable(quote))
            {
                throw new InvalidOperationException("У выбранной квоты нет сравнимой цены.");
            }

            decision.RfqRow.ManualWinnerId = quote.Id;
            quote.MatchStatus = "ManualWinner";
            RrfqEngine.RebuildDecisionsAfterManualMatch(_analysis);
            RefreshPreview();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Ручной выбор квоты");
        }
    }

    /// <summary>Аналог $btnManual: окно ручного сопоставления (строки 5989-5995).</summary>
    [RelayCommand(CanExecute = nameof(CanOpenManualMatch))]
    private void OpenManualMatch()
    {
        if (_analysis is null)
        {
            return;
        }

        var dialogViewModel = new ManualMatchViewModel(_analysis, RefreshPreview);
        var window = new ManualMatchWindow(dialogViewModel) { Owner = Application.Current?.MainWindow };
        window.ShowDialog();
        RefreshPreview();
    }

    private bool CanOpenManualMatch() => _analysis is not null && !IsBusy;

    /// <summary>Аналог $btnAddManualQuote: ручной ввод квоты для выбранной строки (строки 6010-6024).</summary>
    [RelayCommand]
    private void AddManualQuote()
    {
        try
        {
            if (_analysis is null || SelectedDecision is null)
            {
                throw new InvalidOperationException("Выберите строку preview.");
            }

            var dialogViewModel = new ManualQuoteViewModel(_analysis, SelectedDecision);
            var window = new ManualQuoteWindow(dialogViewModel) { Owner = Application.Current?.MainWindow };
            window.ShowDialog();
            if (dialogViewModel.Saved)
            {
                RefreshPreview();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Ручной ввод квоты");
        }
    }

    // ----- Фоновый STA-поток для Excel -----

    /// <summary>
    /// Выполняет работу в отдельном STA-потоке (нужен для Excel COM) и
    /// возвращает результат или ошибку в поток интерфейса.
    /// </summary>
    private void StartStaJob(Func<object?> work, Action<Exception?, object?> completed)
    {
        var dispatcher = Application.Current?.Dispatcher;
        var thread = new Thread(() =>
        {
            Exception? error = null;
            object? result = null;
            try
            {
                result = work();
            }
            catch (Exception ex)
            {
                error = ex;
            }
            finally
            {
                if (dispatcher is not null)
                {
                    dispatcher.Invoke(() => completed(error, result));
                }
                else
                {
                    completed(error, result);
                }
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
    }
}
