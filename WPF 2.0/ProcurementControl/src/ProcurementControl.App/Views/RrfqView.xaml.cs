using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ProcurementControl.Models;
using ProcurementControl.ViewModels;

namespace ProcurementControl.Views;

/// <summary>
/// Представление раздела «Сравнение RRFQ». Обработчики кода нужны только там,
/// где требуется доступ к множественному выбору сетки или к событию мыши —
/// остальное живёт в <see cref="RrfqViewModel"/>.
/// </summary>
public partial class RrfqView : UserControl
{
    public RrfqView()
    {
        InitializeComponent();
    }

    /// <summary>Удаление выбранных файлов поставщиков (мультивыбор сетки).</summary>
    private void RemoveSelectedSuppliersClick(object sender, RoutedEventArgs e)
    {
        if (DataContext is not RrfqViewModel vm)
        {
            return;
        }

        var selected = supplierGrid.SelectedItems.Cast<SupplierFileEntry>().ToList();
        if (selected.Count > 0)
        {
            vm.RemoveSuppliersCommand.Execute(selected);
        }
    }

    /// <summary>Двойной клик по квоте — сделать её победителем (аналог оригинала).</summary>
    private void QuoteGridMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is not RrfqViewModel vm)
        {
            return;
        }

        if (vm.SetWinnerCommand.CanExecute(null))
        {
            vm.SetWinnerCommand.Execute(null);
        }
    }
}
