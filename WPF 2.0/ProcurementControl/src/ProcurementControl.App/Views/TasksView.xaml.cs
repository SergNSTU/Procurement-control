using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;
using ProcurementControl.ViewModels;

namespace ProcurementControl.Views;

public partial class TasksView : UserControl
{
    private bool _restoringTaskColumns;
    private bool _savingTaskCell;
    public TasksView()
    {
        InitializeComponent();
    }

    private void TasksGrid_Loaded(object sender, RoutedEventArgs e)
    {
        _restoringTaskColumns = true;
        GridColumnSettings.Restore(TasksGrid, "ui.components.column_widths");
        _restoringTaskColumns = false;

        var widthProperty = DependencyPropertyDescriptor.FromProperty(DataGridColumn.WidthProperty, typeof(DataGridColumn));
        foreach (var column in TasksGrid.Columns)
        {
            widthProperty?.AddValueChanged(column, TasksGrid_ColumnWidthChanged);
        }
    }

    private void TasksGrid_ColumnWidthChanged(object? sender, EventArgs e)
    {
        if (!_restoringTaskColumns) GridColumnSettings.Save(TasksGrid, "ui.components.column_widths");
    }

    /// <summary>Двойной щелчок по напоминанию или дедлайну открывает календарь.</summary>
    private void TasksGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid grid || grid.CurrentItem is not TaskRow task ||
            DataContext is not TasksViewModel viewModel)
        {
            return;
        }

        var field = grid.CurrentCell.Column?.SortMemberPath;
        if (field is not nameof(TaskRow.ReminderDate) and not nameof(TaskRow.DeadlineDate)) return;
        var current = field == nameof(TaskRow.ReminderDate) ? task.ReminderDate : task.DeadlineDate;
        var dialog = new DatePickerWindow(field == nameof(TaskRow.ReminderDate) ? "Напоминание" : "Дедлайн", current)
        {
            Owner = Window.GetWindow(this),
        };
        if (dialog.ShowDialog() == true) viewModel.UpdateTaskDate(task, field, dialog.SelectedDateText);
    }

    /// <summary>Выделяет строку под курсором, чтобы команда контекстного меню работала именно с ней.</summary>
    private void TasksGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid grid || e.OriginalSource is not DependencyObject source) return;
        if (ItemsControl.ContainerFromElement(grid, source) is not DataGridRow row) return;
        row.IsSelected = true;
        row.Focus();
    }

    /// <summary>Сохраняет текст и выбранные значения списков сразу после изменения ячейки.</summary>
    private void TasksGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (_savingTaskCell || e.EditAction != DataGridEditAction.Commit ||
            e.Row.Item is not TaskRow task || DataContext is not TasksViewModel viewModel)
        {
            return;
        }

        var field = e.Column.SortMemberPath;
        if (field is not nameof(TaskRow.DealNumber) and not nameof(TaskRow.Status) and
            not nameof(TaskRow.Stage) and not nameof(TaskRow.Description) and
            not nameof(TaskRow.NextAction) and not nameof(TaskRow.Priority) and
            not nameof(TaskRow.Period) and not nameof(TaskRow.OrderAmount))
        {
            return;
        }

        Dispatcher.BeginInvoke(() =>
        {
            if (_savingTaskCell) return;
            _savingTaskCell = true;
            try { viewModel.UpdateTaskGridField(task, field); }
            finally { _savingTaskCell = false; }
        });
    }
}
