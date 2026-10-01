using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Поиск аналогов»: кросс-справочник LCSC (первый источник;
/// позже сюда добавляются ChipFind и Nexar). Запрос выполняется асинхронно,
/// строка-оригинал идёт первой, дальше — аналоги LCSC.
/// </summary>
public partial class CrossReferenceViewModel : ObservableObject
{
    [ObservableProperty]
    private string _mpnText = string.Empty;

    [ObservableProperty]
    private bool _inStockOnly;

    [ObservableProperty]
    private string _statusText = string.Empty;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SearchCommand))]
    private bool _isSearching;

    public ObservableCollection<CrossReferenceRow> Rows { get; } = new();

    [RelayCommand(CanExecute = nameof(CanSearch))]
    private async Task Search()
    {
        try
        {
            var mpn = MpnText.Trim();
            if (string.IsNullOrWhiteSpace(mpn))
            {
                throw new InvalidOperationException("Укажите MPN (полный артикул производителя).");
            }

            IsSearching = true;
            ErrorMessage = string.Empty;
            StatusText = "Поиск аналогов на LCSC...";

            var inStockOnly = InStockOnly;
            var result = await Task.Run(() => LcscCrossService.SearchAsync(mpn, inStockOnly));

            Rows.Clear();
            foreach (var row in result.Rows)
            {
                Rows.Add(row);
            }

            if (result.ResolvedCode.Length == 0)
            {
                StatusText = string.Empty;
                MessageBox.Show($"MPN «{mpn}» в LCSC не найден.", "Поиск аналогов");
                return;
            }

            var head = result.ExactMatch
                ? $"Исходный: {result.ResolvedModel} ({result.ResolvedCode})"
                : $"Точное совпадение не найдено — аналоги для {result.ResolvedModel} ({result.ResolvedCode})";
            StatusText = $"{head} · строк: {Rows.Count}";
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            MessageBox.Show(ex.Message, "Поиск аналогов");
        }
        finally
        {
            IsSearching = false;
        }
    }

    private bool CanSearch() => !IsSearching;
}
