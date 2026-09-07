using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Поиск цен» — перенос $priceSearchPage (RRFQComparer.ps1,
/// строки 3100-3137, 4825-4864). Источники ищутся асинхронно; локальные
/// читаются из снапшота БД, внешние — по HTTP.
/// </summary>
public partial class PriceSearchViewModel : ObservableObject
{
    private PurchaseSnapshot? _snapshot;

    [ObservableProperty]
    private string _pnText = string.Empty;

    [ObservableProperty]
    private string _qtyText = string.Empty;

    [ObservableProperty]
    private string _statusText = string.Empty;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SearchCommand))]
    private bool _isSearching;

    public ObservableCollection<PriceSearchRow> Rows { get; } = new();

    /// <summary>Аналог $btnPriceSearch: обход всех источников из Get-PriceSearchResults.</summary>
    [RelayCommand(CanExecute = nameof(CanSearch))]
    private async Task Search()
    {
        try
        {
            var pn = PnText.Trim();
            if (string.IsNullOrWhiteSpace(pn))
            {
                throw new InvalidOperationException("Укажите PN.");
            }

            var quantity = 0;
            if (!string.IsNullOrWhiteSpace(QtyText) && !int.TryParse(QtyText.Trim(), out quantity))
            {
                throw new InvalidOperationException("Количество должно быть целым числом.");
            }

            _snapshot?.Dispose();
            _snapshot = PurchaseSnapshot.Create();

            IsSearching = true;
            ErrorMessage = string.Empty;
            StatusText = "Поиск предложений...";

            var connection = _snapshot.Connection;
            var items = await Task.Run(() => PriceSearchService.GetResultsAsync(connection, pn, quantity));

            Rows.Clear();
            foreach (var row in items)
            {
                Rows.Add(row);
            }

            StatusText = "Найдено предложений: " + Rows.Count;
            if (Rows.Count == 0)
            {
                MessageBox.Show("По этому PN предложения не найдены.", "Поиск цены");
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            MessageBox.Show(ex.Message, "Поиск цены");
        }
        finally
        {
            IsSearching = false;
        }
    }

    private bool CanSearch() => !IsSearching;

    /// <summary>Аналог $btnPriceSearchGpt: запрос в буфер + открытие chatgpt.com.</summary>
    [RelayCommand]
    private void SearchInGpt()
    {
        try
        {
            var pn = PnText.Trim();
            if (string.IsNullOrWhiteSpace(pn))
            {
                throw new InvalidOperationException("Укажите PN.");
            }

            _ = int.TryParse(QtyText.Trim(), out var quantity);
            Clipboard.SetText(PriceSearchService.GetGptPrompt(pn, quantity));
            Process.Start(new ProcessStartInfo("https://chatgpt.com/") { UseShellExecute = true });
            MessageBox.Show("Запрос отправлен в чат GPT.", "Поиск в GPT");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Поиск в GPT");
        }
    }
}
