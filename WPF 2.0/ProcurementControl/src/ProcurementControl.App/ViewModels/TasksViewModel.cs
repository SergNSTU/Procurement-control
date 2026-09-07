using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;
using ProcurementControl.Views;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Страница «Задачи» (аналог вкладки «Компоненты» оригинала). Чтение — одним
/// снапшотом, запись — через PurchaseWriteRepository: «Новая задача»
/// ($btnNewComponent), «Завершить» (Show-TaskCompletionDialog + обновление),
/// «Удалить» с папкой через корзину ($btnDeleteComponent).
/// </summary>
public partial class TasksViewModel : ObservableObject
{
    public IReadOnlyList<string> TaskStatuses { get; } =
        new[] { "В работе", "Ожидание ответа", "На отслеживании", "Заказано", "Подано в оплату", "Выполнено", "Не актуально" };

    public IReadOnlyList<string> TaskStages { get; } =
        new[] { "", "Запросил поставщиков", "RRFQ отправлено", "PI отправлен", "Контроль оплаты", "Отправлено в РФ" };

    public IReadOnlyList<string> TaskPriorities { get; } =
        new[] { "", "Срочно", "1", "2", "3", "4", "5" };

    public IReadOnlyList<string> TaskPeriods { get; } =
        new[] { "", "I квартал", "II квартал", "III квартал", "IV квартал" };

    private List<TaskRow> _allRows = new();

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private TaskRow? _selectedTask;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    private string _statusText = string.Empty;

    [ObservableProperty]
    private string _countText = string.Empty;

    public ObservableCollection<TaskRow> Tasks { get; } = new();

    [ObservableProperty]
    private string _editableNotes = string.Empty;

    public TasksViewModel()
    {
        Load();
    }

    partial void OnSearchTextChanged(string value)
    {
        ApplyFilter();
    }

    partial void OnSelectedTaskChanged(TaskRow? value)
    {
        EditableNotes = value?.Notes ?? string.Empty;
    }

    [RelayCommand]
    private void Reload()
    {
        Load();
    }

    [RelayCommand]
    private void CopyDealNumber()
    {
        try
        {
            var task = SelectedTask ?? throw new InvalidOperationException("Выберите задачу.");
            Clipboard.SetText(task.DealNumber);
            StatusText = "Номер сделки скопирован";
            ToastService.Show(StatusText, ToastKind.Info);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Скопировать номер сделки");
        }
    }

    /// <summary>Аналог $btnNewComponent: создать задачу и выделить её.</summary>
    [RelayCommand]
    private void NewTask()
    {
        try
        {
            var newId = PurchaseWriteRepository.CreateComponentDeal();
            Load();
            SelectedTask = Tasks.FirstOrDefault(r => r.Id == newId);
            StatusText = "Задача создана";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Задачи");
        }
    }

    /// <summary>Сохраняет все колонки выбранной задачи и текст записей.</summary>
    [RelayCommand]
    private void SaveTask()
    {
        try
        {
            var task = SelectedTask ?? throw new InvalidOperationException("Выберите задачу.");
            task.Notes = EditableNotes;
            PurchaseWriteRepository.SaveComponentDeal(task);
            Load();
            SelectedTask = Tasks.FirstOrDefault(row => row.Id == task.Id);
            StatusText = "Задача сохранена";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Сохранить задачу");
        }
    }

    /// <summary>Сохраняет дату из календаря, не требуя ручного текстового ввода.</summary>
    public void UpdateTaskDate(TaskRow task, string field, string value)
    {
        try
        {
            if (field == nameof(TaskRow.ReminderDate)) task.ReminderDate = value;
            else if (field == nameof(TaskRow.DeadlineDate)) task.DeadlineDate = value;
            else return;
            PurchaseWriteRepository.SaveComponentDeal(task);
            Load();
            SelectedTask = Tasks.FirstOrDefault(row => row.Id == task.Id);
            StatusText = "Дата сохранена";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Выбор даты");
        }
    }

    /// <summary>Сохраняет ячейку, отредактированную прямо в таблице задач.</summary>
    public void UpdateTaskGridField(TaskRow task, string field)
    {
        try
        {
            var value = field switch
            {
                nameof(TaskRow.DealNumber) => task.DealNumber,
                nameof(TaskRow.Status) => task.Status,
                nameof(TaskRow.Stage) => task.Stage,
                nameof(TaskRow.Description) => task.Description,
                nameof(TaskRow.NextAction) => task.NextAction,
                nameof(TaskRow.Priority) => task.Priority,
                nameof(TaskRow.Period) => task.Period,
                nameof(TaskRow.OrderAmount) => task.OrderAmount,
                _ => throw new ArgumentException("Поле нельзя изменить из таблицы.", nameof(field)),
            };
            PurchaseWriteRepository.UpdateComponentDealGridField(task.Id, field, value);
            ReloadSelection(task.Id);
            StatusText = "Изменения сохранены";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            ReloadSelection(task.Id);
            MessageBox.Show(ex.Message, "Сохранить задачу");
        }
    }

    /// <summary>
    /// Завершение задачи: диалог Show-TaskCompletionDialog (строки 649-702
    /// оригинала) — статус/этап, срок очищается.
    /// </summary>
    [RelayCommand]
    private void CompleteTask()
    {
        try
        {
            var task = SelectedTask;
            if (task is null)
            {
                throw new InvalidOperationException("Выберите задачу.");
            }

            var dialog = new TaskCompletionWindow(task.Status, task.Stage)
            {
                Owner = Application.Current.MainWindow,
            };
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            PurchaseWriteRepository.UpdateComponentDealCompletion(
                task.Id, dialog.Status, dialog.Stage,
                clearReminderDate: true, clearDeadlineDate: true);
            Load();
            SelectedTask = Tasks.FirstOrDefault(r => r.Id == task.Id);
            StatusText = "Задача завершена";
            ToastService.Show(StatusText, ToastKind.Success);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Задачи");
        }
    }

    /// <summary>Аналог $btnDeleteComponent: задача и папка уходят в корзину.</summary>
    [RelayCommand]
    private void DeleteTask()
    {
        try
        {
            var task = SelectedTask;
            if (task is null)
            {
                throw new InvalidOperationException("Выберите строку.");
            }

            var answer = MessageBox.Show(
                "Переместить задачу в корзину?\n" + task.DealNumber,
                "Корзина",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes)
            {
                return;
            }

            PurchaseWriteRepository.DeleteComponentDeal(task.Id);
            Load();
            StatusText = "Задача перемещена в корзину";
            ToastService.Show(StatusText, ToastKind.Info);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Задачи");
        }
    }

    /// <summary>
    /// Аналог ветки component из Open-NotificationTarget (RRFQComparer.ps1,
    /// строки 714-733): обновить список и выделить задачу; если поиск мешает —
    /// очистить его и повторить.
    /// </summary>
    public void OpenTask(long id)
    {
        Load();
        var target = Tasks.FirstOrDefault(r => r.Id == id);
        if (target is null && !string.IsNullOrWhiteSpace(SearchText))
        {
            SearchText = string.Empty;
            target = Tasks.FirstOrDefault(r => r.Id == id);
        }

        SelectedTask = target;
    }

    private void Load()
    {
        try
        {
            using var snapshot = PurchaseSnapshot.Create();
            var repo = new PurchaseRepository(snapshot.Connection);
            _allRows = repo.GetTasks(string.Empty);

            ErrorMessage = string.Empty;
            ApplyFilter();
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось прочитать данные: " + ex.Message;
        }
    }

    /// <summary>Аналог LIKE '%...%' оригинала по сделке и описанию, без учёта регистра.</summary>
    private void ApplyFilter()
    {
        var needle = SearchText.Trim();

        var selectedId = SelectedTask?.Id;
        Tasks.Clear();
        foreach (var row in _allRows)
        {
            if (needle.Length == 0
                || row.DealNumber.Contains(needle, StringComparison.OrdinalIgnoreCase)
                || row.Description.Contains(needle, StringComparison.OrdinalIgnoreCase))
            {
                Tasks.Add(row);
            }
        }

        CountText = "Всего: " + Tasks.Count;
        SelectedTask = selectedId is null
            ? null
            : Tasks.FirstOrDefault(r => r.Id == selectedId);
        EditableNotes = SelectedTask?.Notes ?? string.Empty;
    }

    private void ReloadSelection(long taskId)
    {
        Load();
        SelectedTask = Tasks.FirstOrDefault(row => row.Id == taskId);
    }
}
