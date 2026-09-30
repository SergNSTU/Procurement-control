using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ProcurementControl.Models;
using ProcurementControl.ViewModels;
using ProcurementControl.Services;

namespace ProcurementControl.Views;

public partial class PurchaseView : UserControl
{
    private bool _restoringDealColumns;
    private bool _savingDealCell;
    private bool _savingSupplierCell;
    private bool _allowDealCellEdit;
    public PurchaseView()
    {
        InitializeComponent();
    }

    private void DealsGrid_Loaded(object sender, RoutedEventArgs e)
    {
        _restoringDealColumns = true;
        GridColumnSettings.Restore(DealsGrid, "ui.purchase.deals.column_widths");
        _restoringDealColumns = false;

        var widthProperty = DependencyPropertyDescriptor.FromProperty(DataGridColumn.WidthProperty, typeof(DataGridColumn));
        foreach (var column in DealsGrid.Columns)
        {
            widthProperty?.AddValueChanged(column, DealsGrid_ColumnWidthChanged);
        }
    }

    /// <summary>Открывает меню дополнительных действий по обычному левому клику на «⋯».</summary>
    private void PurchaseActionsButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { ContextMenu: { } menu } button) return;
        menu.PlacementTarget = button;
        menu.IsOpen = true;
    }

    private void DealsGrid_ColumnWidthChanged(object? sender, EventArgs e)
    {
        if (!_restoringDealColumns) GridColumnSettings.Save(DealsGrid, "ui.purchase.deals.column_widths");
    }

    /// <summary>Открывает специальные редакторы дат по двойному щелчку.</summary>
    private void DealsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid grid || grid.CurrentItem is not PurchaseDealRow deal ||
            DataContext is not PurchaseViewModel viewModel)
        {
            return;
        }

        var field = grid.CurrentCell.Column?.SortMemberPath;
        if (field == nameof(PurchaseDealRow.CompletionReceiptDate))
        {
            if (viewModel.Suppliers.Count == 0)
            {
                MessageBox.Show("У сделки нет поставщиков, для которых можно указать поступление.", "Поступление комплектации");
                return;
            }
            var receiptDialog = new DealCompletionReceiptWindow(viewModel.Suppliers) { Owner = Window.GetWindow(this) };
            if (receiptDialog.ShowDialog() == true)
                viewModel.UpdateCompletionReceiptDates(deal, receiptDialog.Updates);
            return;
        }
        if (field == nameof(PurchaseDealRow.ReminderDateText))
        {
            var dateDialog = new DatePickerWindow("Напоминание", deal.ReminderDate) { Owner = Window.GetWindow(this) };
            if (dateDialog.ShowDialog() == true) viewModel.UpdateDealReminderDate(deal, dateDialog.SelectedDateText);
            return;
        }
    }

    /// <summary>Редактирование полей сделки начинается только по двойному щелчку мыши.</summary>
    private void DealsGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _allowDealCellEdit = e.ClickCount == 2;
    }

    private void DealsGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
    {
        if (!_allowDealCellEdit) e.Cancel = true;
        _allowDealCellEdit = false;
    }

    /// <summary>Выделяет строку под курсором перед открытием контекстного меню.</summary>
    private void DealsGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid grid || e.OriginalSource is not DependencyObject source) return;
        if (ItemsControl.ContainerFromElement(grid, source) is not DataGridRow row) return;
        row.IsSelected = true;
        row.Focus();
    }

    /// <summary>Сохраняет изменённые в таблице текстовые поля и значения списков.</summary>
    private void DealsGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (_savingDealCell || e.EditAction != DataGridEditAction.Commit ||
            e.Row.Item is not PurchaseDealRow deal || DataContext is not PurchaseViewModel viewModel)
        {
            return;
        }

        var field = e.Column.SortMemberPath;
        if (field is not nameof(PurchaseDealRow.DealNumber) and not nameof(PurchaseDealRow.BoardCount) and
            not nameof(PurchaseDealRow.Client) and not nameof(PurchaseDealRow.Priority) and
            not nameof(PurchaseDealRow.Status) and not nameof(PurchaseDealRow.Comment) and
            not nameof(PurchaseDealRow.Period) and not nameof(PurchaseDealRow.Executor) and
            not nameof(PurchaseDealRow.AssemblyLocation) and not nameof(PurchaseDealRow.TrackingStatus) and
            not nameof(PurchaseDealRow.MasksText))
        {
            return;
        }

        Dispatcher.BeginInvoke(() =>
        {
            if (_savingDealCell) return;
            _savingDealCell = true;
            try { viewModel.UpdateDealGridField(deal, field); }
            finally { _savingDealCell = false; }
        });
    }

    /// <summary>
    /// Флажок поставщика переключается первым щелчком, без двойного входа в правку:
    /// отображаемый CheckBox колонки недоступен для мыши, поэтому ищем ячейку,
    /// меняем поле по SortMemberPath и сразу сохраняем.
    /// </summary>
    private void SuppliersGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid grid || DataContext is not PurchaseViewModel viewModel ||
            e.OriginalSource is not DependencyObject source)
        {
            return;
        }

        var node = source;
        while (node is not DataGridCell and not DataGrid and not null)
        {
            node = node is Visual ? VisualTreeHelper.GetParent(node) : LogicalTreeHelper.GetParent(node);
        }

        if (node is not DataGridCell cell || cell.Column is not DataGridCheckBoxColumn ||
            cell.DataContext is not SupplierRow supplier)
        {
            return;
        }

        var field = cell.Column.SortMemberPath;
        if (field is not nameof(SupplierRow.PaymentSubmitted) and not nameof(SupplierRow.Paid) and
            not nameof(SupplierRow.InvoiceReceived) and not nameof(SupplierRow.InvoiceConfirmed) and
            not nameof(SupplierRow.ErpSupplierSent) and not nameof(SupplierRow.ErpRogerSent))
        {
            return;
        }

        switch (field)
        {
            case nameof(SupplierRow.PaymentSubmitted): supplier.PaymentSubmitted = !supplier.PaymentSubmitted; break;
            case nameof(SupplierRow.Paid): supplier.Paid = !supplier.Paid; break;
            case nameof(SupplierRow.InvoiceReceived): supplier.InvoiceReceived = !supplier.InvoiceReceived; break;
            case nameof(SupplierRow.InvoiceConfirmed): supplier.InvoiceConfirmed = !supplier.InvoiceConfirmed; break;
            case nameof(SupplierRow.ErpSupplierSent): supplier.ErpSupplierSent = !supplier.ErpSupplierSent; break;
            case nameof(SupplierRow.ErpRogerSent): supplier.ErpRogerSent = !supplier.ErpRogerSent; break;
        }

        e.Handled = true;
        if (ItemsControl.ContainerFromElement(grid, source) is DataGridRow row)
        {
            row.IsSelected = true;
            row.Focus();
        }

        viewModel.UpdateSupplierFlags(supplier);
    }

    /// <summary>Сохраняет флажок поставщика сразу после щелчка.</summary>
    private void SuppliersGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
    {
        if (_savingSupplierCell || e.EditAction != DataGridEditAction.Commit ||
            e.Row.Item is not SupplierRow supplier ||
            (e.Column is not DataGridCheckBoxColumn &&
             e.Column.SortMemberPath is not nameof(SupplierRow.DeliveryWeeks) and not nameof(SupplierRow.PiAmountUsd) and
                 not nameof(SupplierRow.PiAmountCny) and not nameof(SupplierRow.PiAmountRub) and not nameof(SupplierRow.PaidAmount) and
                 not nameof(SupplierRow.Comment)) ||
            DataContext is not PurchaseViewModel viewModel)
        {
            return;
        }

        Dispatcher.BeginInvoke(() =>
        {
            if (_savingSupplierCell) return;
            _savingSupplierCell = true;
            try { viewModel.UpdateSupplierFlags(supplier); }
            finally { _savingSupplierCell = false; }
        });
    }

    /// <summary>Открывает календарь для дат поставщика по двойному щелчку.</summary>
    private void SuppliersGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid grid || grid.CurrentItem is not SupplierRow supplier ||
            DataContext is not PurchaseViewModel viewModel)
        {
            return;
        }

        var field = grid.CurrentCell.Column?.SortMemberPath;
        if (field is not nameof(SupplierRow.InvoiceConfirmedDate) and not nameof(SupplierRow.ComponentsReceiptDate) and
            not nameof(SupplierRow.ActualReceiptDate))
        {
            return;
        }

        var currentValue = field switch
        {
            nameof(SupplierRow.InvoiceConfirmedDate) => supplier.InvoiceConfirmedDate,
            nameof(SupplierRow.ComponentsReceiptDate) => supplier.ComponentsReceiptDate,
            nameof(SupplierRow.ActualReceiptDate) => supplier.ActualReceiptDate,
            _ => null,
        };
        var title = field == nameof(SupplierRow.InvoiceConfirmedDate)
            ? "Дата подтверждения инвойса"
            : field == nameof(SupplierRow.ComponentsReceiptDate)
                ? "Плановое поступление"
                : "Фактическое поступление";
        var dialog = new DatePickerWindow(title, currentValue ?? string.Empty) { Owner = Window.GetWindow(this) };
        if (dialog.ShowDialog() == true)
            viewModel.UpdateSupplierDate(supplier, field!, dialog.SelectedDateText);
    }

    private void PickDealReminderDateClick(object sender, RoutedEventArgs e)
    {
        if (DataContext is not PurchaseViewModel viewModel) return;
        var dialog = new DatePickerWindow("Напоминание", viewModel.EditReminderDate) { Owner = Window.GetWindow(this) };
        if (dialog.ShowDialog() == true) viewModel.EditReminderDate = dialog.SelectedDateText;
    }

    /// <summary>
    /// Перенос $purchaseDocsDragEnter (RRFQComparer.ps1, строки 5647-5653):
    /// принимать только файлы из проводника.
    /// </summary>
    private void DocumentsDragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    /// <summary>Перенос $purchaseDocsDragDrop (строки 5659-5666).</summary>
    private void DocumentsDrop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] paths &&
            DataContext is PurchaseViewModel viewModel)
        {
            viewModel.AddDocumentsFromPaths(paths);
        }
    }

    /// <summary>Открывает файл выбранного документа двойным щелчком.</summary>
    private void DocumentsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is DataGrid { SelectedItem: PurchaseDocumentRow document } &&
            DataContext is PurchaseViewModel viewModel)
        {
            viewModel.OpenDocument(document);
        }
    }
}
