using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Диалог «Ручное сопоставление» — перенос Show-ManualMatchDialog
/// (RrfqUi.ps1, строки 394-538): слева несопоставленные квоты, справа
/// строки RFQ; выбранные пара связывается вручную.
/// </summary>
public partial class ManualMatchViewModel : ObservableObject
{
    private readonly RrfqAnalysis _analysis;
    private readonly Action _onChanged;

    public ManualMatchViewModel(RrfqAnalysis analysis, Action onChanged)
    {
        _analysis = analysis;
        _onChanged = onChanged;
        RefreshLists();
    }

    /// <summary>Поиск по несопоставленным квотам (аналог $txtQuoteSearch).</summary>
    [ObservableProperty]
    private string _quoteSearch = string.Empty;

    /// <summary>Поиск по строкам RFQ (аналог $txtRfqSearch).</summary>
    [ObservableProperty]
    private string _rfqSearch = string.Empty;

    partial void OnQuoteSearchChanged(string value) => RefreshLists();

    partial void OnRfqSearchChanged(string value) => RefreshLists();

    public ObservableCollection<Quote> UnmatchedQuotes { get; } = new();

    public ObservableCollection<RfqRow> RfqRows { get; } = new();

    [ObservableProperty]
    private Quote? _selectedQuote;

    [ObservableProperty]
    private RfqRow? _selectedRfqRow;

    /// <summary>Вывод ошибок: по умолчанию MessageBox; подменяется в проверках.</summary>
    public Action<string, string> ShowError { get; set; } = (message, title) => MessageBox.Show(message, title);

    /// <summary>Аналог $btnMatch: связывает выбранную квоту со строкой RFQ.</summary>
    [RelayCommand]
    private void Match()
    {
        if (SelectedQuote is null || SelectedRfqRow is null)
        {
            ShowError("Выберите квоту слева и строку RFQ справа.", "Ручное сопоставление");
            return;
        }

        SelectedQuote.TargetId = SelectedRfqRow.Id;
        SelectedQuote.MatchStatus = "Manual";
        RrfqEngine.RebuildDecisionsAfterManualMatch(_analysis);
        RefreshLists();
        _onChanged();
    }

    /// <summary>Сигнал окну закрыться (аналог $btnClose).</summary>
    [RelayCommand]
    private void Close() => CloseRequested?.Invoke(this, EventArgs.Empty);

    public event EventHandler? CloseRequested;

    /// <summary>Аналог Refresh-ManualLists с учётом поисковых строк.</summary>
    private void RefreshLists()
    {
        UnmatchedQuotes.Clear();
        var quoteNeedle = (QuoteSearch ?? string.Empty).Trim().ToLowerInvariant();
        foreach (var quote in RrfqEngine.GetUnresolvedQuotes(_analysis))
        {
            if (quoteNeedle.Length > 0)
            {
                var hay = $"{quote.Supplier} {quote.FileName} {quote.SheetName} {quote.Key} {quote.PN}".ToLowerInvariant();
                if (!hay.Contains(quoteNeedle))
                {
                    continue;
                }
            }

            UnmatchedQuotes.Add(quote);
        }

        RfqRows.Clear();
        var rfqNeedle = (RfqSearch ?? string.Empty).Trim().ToLowerInvariant();
        foreach (var rfqRow in _analysis.RfqRows)
        {
            if (rfqNeedle.Length > 0)
            {
                var hay = $"{rfqRow.SheetName} {rfqRow.Row} {rfqRow.Key} {rfqRow.PN} {rfqRow.RussianRemark}".ToLowerInvariant();
                if (!hay.Contains(rfqNeedle))
                {
                    continue;
                }
            }

            RfqRows.Add(rfqRow);
        }
    }
}
