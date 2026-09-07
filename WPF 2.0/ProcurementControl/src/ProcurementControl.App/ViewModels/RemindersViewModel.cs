using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;
using ProcurementControl.Views;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Напоминания» (аналог $remindersPage оригинала).
/// Список строится по правилам Get-PurchaseActionItems: авто-пункты из статусов закупки + напоминания сделок + ручные.
/// Действия — порты $btnDoneReminder, $btnReceiptArrived, $btnNewReminder и
/// $remindersDeleteMenuItem (RRFQComparer.ps1, строки 4949-5001, 4274-4289).
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

    [ObservableProperty]
    private ActionItemRow? _selectedItem;

    /// <summary>
    /// Чекбокс «Уведомления включены» из тулбара оригинала (строки 2484-2490,
    /// 5873-5879): значение сохраняется в настройку notifications.enabled.
    /// </summary>
    [ObservableProperty]
    private bool _notificationsEnabled;

    public ObservableCollection<ActionItemRow> Items { get; } = new();

    public RemindersViewModel()
    {
        NotificationsEnabled = NotificationCenter.Enabled;
        Load();
    }

    partial void OnNotificationsEnabledChanged(bool value)
    {
        NotificationCenter.Enabled = value;
        try
        {
            PurchaseWriteRepository.SetSetting("notifications.enabled", value ? "1" : "0");
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось сохранить настройку: " + ex.Message;
        }
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

    /// <summary>
    /// Аналог ветки reminder из Open-NotificationTarget (RRFQComparer.ps1,
    /// строки 734-746): обновить список и выделить пункт по ручному напоминанию.
    /// </summary>
    public void OpenReminder(long reminderId)
    {
        Load();
        SelectedItem = Items.FirstOrDefault(i => i.ReminderId == reminderId);
    }

    /// <summary>
    /// Порт $btnDoneReminder (RRFQComparer.ps1, строки 4965-4983): для напоминания сделки —
    /// очистка reminder_date, для ручного — статус Done.
    /// </summary>
    [RelayCommand]
    private void DoneReminder()
    {
        try
        {
            var item = SelectedItem ?? throw new InvalidOperationException("Выберите ручное напоминание.");
            if (item.Source == "deal")
            {
                PurchaseWriteRepository.ClearDealReminder(item.DealId);
            }
            else if (item.ReminderId > 0)
            {
                PurchaseWriteRepository.SetReminderDone(item.ReminderId, done: true);
            }
            else
            {
                throw new InvalidOperationException("Выберите ручное напоминание.");
            }

            Load();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Напоминания");
        }
    }

    /// <summary>
    /// Порт Confirm-SelectedReminderReceipt (RRFQComparer.ps1, строки 4274-4289):
    /// фактическая дата поступления для авто-напоминания по поступлению.
    /// </summary>
    [RelayCommand]
    private void ReceiptArrived()
    {
        try
        {
            var item = SelectedItem;
            if (item is null || item.SupplierId <= 0 || item.Source != "auto" ||
                !item.Title.ToLowerInvariant().Contains("поступление"))
            {
                throw new InvalidOperationException("Выберите автоматическое напоминание по поступлению.");
            }

            // Аналог Show-DatePickerDialog: ввод даты в формате дд.мм.гггг.
            var dialog = new InputDialogWindow("Поступление", "Фактическая дата поступления (дд.мм.гггг)", DateTime.Now.ToString("dd.MM.yyyy"));
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            var date = PurchaseFormatting.ParseDate(dialog.Value);
            if (date is null)
            {
                throw new InvalidOperationException("Фактическая дата поступления должна быть в формате дд.мм.гггг.");
            }

            PurchaseWriteRepository.SetSupplierActualReceiptDate(item.SupplierId, date.Value);
            Load();
            MessageBox.Show("Поступление подтверждено\r\nФакт: " + date.Value.ToString("dd.MM.yyyy"), "Напоминания");
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Напоминания");
        }
    }

    /// <summary>
    /// Порт $btnNewReminder (RRFQComparer.ps1, строки 4991-5001). В оригинале сделка и поставщик брались из сеток закупки;
    /// здесь подставляются из выбранного пункта, если он есть.
    /// </summary>
    [RelayCommand]
    private void NewReminder()
    {
        try
        {
            var titleDialog = new InputDialogWindow("Новое напоминание", "Что нужно сделать?");
            if (titleDialog.ShowDialog() != true || string.IsNullOrWhiteSpace(titleDialog.Value))
            {
                return;
            }

            var dueDialog = new InputDialogWindow("Новое напоминание", "Дата (можно пусто, формат дд.мм.гггг)");
            var due = dueDialog.ShowDialog() == true ? dueDialog.Value : string.Empty;

            PurchaseWriteRepository.SaveReminder(
                titleDialog.Value,
                due,
                SelectedItem?.DealId ?? 0,
                SelectedItem?.SupplierId ?? 0,
                "manual");
            Load();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Напоминания");
        }
    }

    /// <summary>
    /// Порт $remindersDeleteMenuItem (RRFQComparer.ps1, строки 4949-4962): ручное напоминание уходит в корзину.
    /// </summary>
    [RelayCommand]
    private void DeleteReminder()
    {
        try
        {
            var item = SelectedItem;
            if (item is null) throw new InvalidOperationException("Выберите напоминание.");

            var answer = MessageBox.Show(
                "Переместить выбранное напоминание в корзину?",
                "Корзина", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes)
            {
                return;
            }

            // Автоматические строки не имеют ReminderId. Удаляем их по ключу,
            // даже если источник строки не был заполнен старой версией данных.
            if (item.Source == "auto" || (item.ReminderId <= 0 && item.Source != "deal"))
                PurchaseWriteRepository.RemoveAutomaticReminder(item.DealId, item.SupplierId, item.Title);
            else if (item.Source == "deal")
                PurchaseWriteRepository.ClearDealReminder(item.DealId);
            else if (item.ReminderId > 0)
                PurchaseWriteRepository.RemoveReminder(item.ReminderId);
            else
                throw new InvalidOperationException("Выберите напоминание.");
            Load();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Напоминания");
        }
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
            var suppressedAutomaticKeys = repo.GetAutomaticReminderSuppressions();

            _allItems = PurchaseLogic.BuildActionItems(suppliers, dealReminders, manuals, DateTime.Today, suppressedAutomaticKeys);

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
