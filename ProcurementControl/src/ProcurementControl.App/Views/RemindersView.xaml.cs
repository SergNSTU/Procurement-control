using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ProcurementControl.ViewModels;

namespace ProcurementControl.Views;

/// <summary>Страница «Напоминания» — только чтение.</summary>
public partial class RemindersView : UserControl
{
    public RemindersView()
    {
        InitializeComponent();
    }

    /// <summary>Двойной щелчок по строке открывает сделку пункта в «Контроле закупки».</summary>
    private void RemindersGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid grid ||
            DataContext is not RemindersViewModel viewModel ||
            e.OriginalSource is not DependencyObject source ||
            ItemsControl.ContainerFromElement(grid, source) is not DataGridRow)
        {
            return;
        }

        viewModel.OpenDealFromReminder();
    }

    /// <summary>Правый клик выделяет строку под курсором: пункты меню действуют на неё.</summary>
    private void RemindersGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid grid || e.OriginalSource is not DependencyObject source) return;
        if (ItemsControl.ContainerFromElement(grid, source) is not DataGridRow row) return;
        grid.SelectedItems.Clear();
        row.IsSelected = true;
        row.Focus();
    }

    /// <summary>Массовое «Готово»: один вопрос, закрытие всех выбранных напоминаний.</summary>
    private void DoneClick(object sender, RoutedEventArgs e)
        => (DataContext as RemindersViewModel)?.DoneSelected(RemindersGrid.SelectedItems);

    /// <summary>Массовое удаление: один вопрос, все выбранные строки уходят в корзину.</summary>
    private void DeleteClick(object sender, RoutedEventArgs e)
        => (DataContext as RemindersViewModel)?.DeleteSelected(RemindersGrid.SelectedItems);
}
