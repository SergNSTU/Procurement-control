using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Read-only страница «Напоминания» (аналог $remindersPage оригинала).
/// Список строится по правилам Get-PurchaseActionItems: авто-пункты из статусов
/// закупки + напоминания сделок + ручные напоминания. Исходная база не изменяется.
/// </summary>
public partial class RemindersViewModel : ObservableObject
{
    private List<ActionItemRow> _allItems = new();

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    private string _countText = string.Empty;

    public string ReadOnlyNotice { get; } =
        "Режим чтения: изменение доступно в основной версии приложения";

    public ObservableCollection<ActionItemRow> Items { get; } = new();

    public RemindersViewModel()
    {
        Load();
    }

    partial void OnSearchTextChanged(string value)
    {
        ApplyFilter();
    }

    [RelayCommand]
    private void Reload()
    {
        Load();
    }

    private void Load()
    {
        try
        {
            // Ленивых обращений к данным нет — снапшот освобождается сразу после чтения.
            using var snapshot = PurchaseSnapshot.Create();
            var repo = new PurchaseRepository(snapshot.Connection);
            var suppliers = repo.GetSupplierActionRows();
            var dealReminders = repo.GetDealReminderRows();
            var manuals = repo.GetManualReminderRows();

            _allItems = PurchaseLogic.BuildActionItems(suppliers, dealReminders, manuals, DateTime.Today);

            ErrorMessage = string.Empty;
            ApplyFilter();
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось прочитать данные: " + ex.Message;
        }
    }

    /// <summary>Поиск оригинала по заголовку и сделке, выполняется в памяти.</summary>
    private void ApplyFilter()
    {
        var visible = PurchaseLogic.SearchItems(_allItems, SearchText);

        Items.Clear();
        foreach (var item in visible)
        {
            Items.Add(item);
        }

        var danger = visible.Count(i => i.Severity == "Danger");
        var attention = visible.Count(i => i.Severity == "Attention");
        var warn = visible.Count(i => i.Severity == "Warn");
        CountText = $"Всего: {visible.Count}  ·  Просрочено: {danger}  ·  Скоро: {attention}  ·  Внимание: {warn}";
    }
}
