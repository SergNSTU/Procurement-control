using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Navigation;
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
    /// Двойной щелчок по пункту: перейти в «Контроль закупки» и выделить сделку
    /// пункта — как ветка deal из Open-NotificationTarget (строки 747-763).
    /// </summary>
    public void OpenDealFromReminder()
    {
        try
        {
            var item = SelectedItem ?? throw new InvalidOperationException("Выберите напоминание.");
            if (item.DealId <= 0)
            {
                throw new InvalidOperationException("У этого напоминания нет связанной сделки.");
            }

            if (Application.Current.MainWindow?.DataContext is MainViewModel main)
            {
                main.Navigate(AppPage.Purchase);
                (main.CurrentViewModel as PurchaseViewModel)?.OpenDeal(item.DealId);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Напоминания");
        }
    }

    /// <summary>
    /// Порт $btnDoneReminder (RRFQComparer.ps1, строки 4965-4983): для напоминания сделки —
    /// очистка reminder_date, для ручного — статус Done. Работает и с несколькими
    /// выбранными строками: один вопрос, закрытие по каждой.
    /// </summary>
    [RelayCommand]
    private void DoneReminder()
        => DoneSelected(SelectedItem is null ? Array.Empty<ActionItemRow>() : new[] { SelectedItem });

    public void DoneSelected(System.Collections.IList rows)
    {
        try
        {
            var items = rows.OfType<ActionItemRow>().ToList();
            if (items.Count == 0)
            {
                throw new InvalidOperationException("Выберите напоминание.");
            }

            // Авто-пункты не имеют ReminderId и сделочной даты — закрывать «Готово» их нельзя.
            var closable = items.Where(i => i.Source == "deal" || i.ReminderId > 0).ToList();
            if (closable.Count == 0)
            {
                throw new InvalidOperationException("Кнопка «Готово» закрывает напоминания сделок и ручные; автоматические убираются удалением.");
            }

            if (closable.Count > 1)
            {
                var answer = MessageBox.Show(
                    "Отметить выполненными: " + closable.Count + "?",
                    "Напоминания", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (answer != MessageBoxResult.Yes)
                {
                    return;
                }
            }

            foreach (var item in closable)
            {
                if (item.Source == "deal")
                {
                    PurchaseWriteRepository.ClearDealReminder(item.DealId);
                }
                else
                {
                    PurchaseWriteRepository.SetReminderDone(item.ReminderId, done: true);
                }
            }

            Load();
            ToastService.Show(
                closable.Count == 1 ? "Напоминание закрыто" : "Закрыто напоминаний: " + closable.Count,
                ToastKind.Success);
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

            // Аналог Show-DatePickerDialog: календарь вместо текстового ввода даты.
            var dialog = new DatePickerWindow("Фактическая дата поступления", DateTime.Now.ToString("dd.MM.yyyy"));
            if (dialog.ShowDialog() != true || string.IsNullOrWhiteSpace(dialog.SelectedDateText))
            {
                return;
            }

            var date = PurchaseFormatting.ParseDate(dialog.SelectedDateText);
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
    /// здесь подставляются из выбранного пункта, если он есть. Вместо двух текстовых
    /// диалогов — окно NewReminderWindow с календарём.
    /// </summary>
    [RelayCommand]
    private void NewReminder()
    {
        try
        {
            var context = SelectedItem is not null && (SelectedItem.DealId > 0 || SelectedItem.SupplierId > 0)
                ? "Новая запись будет привязана к выбранному пункту: " + SelectedItem.Title
                : string.Empty;
            var window = new NewReminderWindow(context) { Owner = Application.Current?.MainWindow };
            if (window.ShowDialog() != true)
            {
                return;
            }

            PurchaseWriteRepository.SaveReminder(
                window.ReminderTitle,
                window.DueDateText,
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
    /// Порт $remindersDeleteMenuItem (RRFQComparer.ps1, строки 4949-4962): напоминания
    /// уходят в корзину. Работает и с несколькими выбранными строками: один вопрос,
    /// удаление по каждой — авто-пункты по ключу, сделочные очисткой даты, ручные по id.
    /// </summary>
    [RelayCommand]
    private void DeleteReminder()
        => DeleteSelected(SelectedItem is null ? Array.Empty<ActionItemRow>() : new[] { SelectedItem });

    public void DeleteSelected(System.Collections.IList rows)
    {
        try
        {
            var items = rows.OfType<ActionItemRow>().ToList();
            if (items.Count == 0)
            {
                throw new InvalidOperationException("Выберите напоминание.");
            }

            var question = items.Count == 1
                ? "Переместить выбранное напоминание в корзину?"
                : "Переместить в корзину: " + items.Count + "?";
            var answer = MessageBox.Show(question, "Корзина", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes)
            {
                return;
            }

            foreach (var item in items)
            {
                DeleteReminderCore(item);
            }

            SelectedItem = null;
            Load();
            ToastService.Show("Перемещено в корзину: " + items.Count, ToastKind.Info);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Напоминания");
        }
    }

    private void DeleteReminderCore(ActionItemRow item)
    {
        // Автоматические строки не имеют ReminderId. Удаляем их по ключу,
        // даже если источник строки не был заполнен старой версией данных.
        if (item.Source == "auto" || (item.ReminderId <= 0 && item.Source != "deal"))
        {
            PurchaseWriteRepository.RemoveAutomaticReminder(item.DealId, item.SupplierId, item.Title);
        }
        else if (item.Source == "deal")
        {
            PurchaseWriteRepository.ClearDealReminder(item.DealId);
        }
        else if (item.ReminderId > 0)
        {
            PurchaseWriteRepository.RemoveReminder(item.ReminderId);
        }
        else
        {
            throw new InvalidOperationException("Выберите напоминание.");
        }
    }

    /// <summary>
    /// Перенос выбранного напоминания в календарь Outlook: встреча-напоминание
    /// на 9:00 даты пункта (строку запланированной интеграции см. в «Инструкции»).
    /// </summary>
    [RelayCommand]
    private void ReminderToOutlook()
    {
        try
        {
            var item = SelectedItem ?? throw new InvalidOperationException("Выберите напоминание.");
            if (!OutlookService.IsAvailable)
            {
                throw new InvalidOperationException("Outlook не найден на этом компьютере.");
            }

            var date = PurchaseFormatting.ParseDate(item.DueDate);
            if (date is null)
            {
                var dialog = new DatePickerWindow("Дата напоминания", DateTime.Now.ToString("dd.MM.yyyy"));
                if (dialog.ShowDialog() != true || string.IsNullOrWhiteSpace(dialog.SelectedDateText))
                {
                    return;
                }

                date = PurchaseFormatting.ParseDate(dialog.SelectedDateText)
                    ?? throw new InvalidOperationException("Дата должна быть в формате дд.мм.гггг.");
            }

            OutlookService.CreateCalendarReminder(item.Title, item.Title, date.Value);
            ToastService.Show("Напоминание создано в Outlook", ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Напоминание в Outlook");
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
