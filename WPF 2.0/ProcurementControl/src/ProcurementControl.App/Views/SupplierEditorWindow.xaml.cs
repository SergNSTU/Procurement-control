using System.Windows;
using System.Windows.Controls;
using ProcurementControl.Models;

namespace ProcurementControl.Views;

public partial class SupplierEditorWindow : Window
{
    public SupplierEditorWindow(SupplierRow supplier)
    {
        InitializeComponent();
        SupplierName.Text = supplier.Supplier;
        InvoiceReceived.IsChecked = supplier.InvoiceReceived; InvoiceConfirmed.IsChecked = supplier.InvoiceConfirmed;
        SupplierOrderCreated.IsChecked = supplier.SupplierOrderCreated; ErpSupplierSent.IsChecked = supplier.ErpSupplierSent;
        ErpRogerSent.IsChecked = supplier.ErpRogerSent; PiUsd.Text = supplier.PiAmountUsd; PiCny.Text = supplier.PiAmountCny;
        PiRub.Text = supplier.PiAmountRub; PaidAmount.Text = supplier.PaidAmount; DeliveryWeeks.Text = supplier.DeliveryWeeks;
        PaymentSubmitted.IsChecked = supplier.PaymentSubmitted; Paid.IsChecked = supplier.Paid;
        InvoiceConfirmedDate.Text = supplier.InvoiceConfirmedDate; ComponentsReceiptDate.Text = supplier.ComponentsReceiptDate;
        ActualReceiptDate.Text = supplier.ActualReceiptDate; Comment.Text = supplier.Comment ?? string.Empty;
    }

    public PurchaseSupplierUpdate Update => new()
    {
        InvoiceReceived = InvoiceReceived.IsChecked == true, InvoiceConfirmed = InvoiceConfirmed.IsChecked == true,
        SupplierOrderCreated = SupplierOrderCreated.IsChecked == true, ErpSupplierSent = ErpSupplierSent.IsChecked == true,
        ErpRogerSent = ErpRogerSent.IsChecked == true, PiAmountUsd = PiUsd.Text, PiAmountCny = PiCny.Text,
        PiAmountRub = PiRub.Text, PaidAmount = PaidAmount.Text, DeliveryWeeks = DeliveryWeeks.Text,
        PaymentSubmitted = PaymentSubmitted.IsChecked == true, Paid = Paid.IsChecked == true,
        InvoiceConfirmedDate = InvoiceConfirmedDate.Text, ComponentsReceiptDate = ComponentsReceiptDate.Text,
        ActualReceiptDate = ActualReceiptDate.Text, Comment = Comment.Text,
    };

    private void SaveClick(object sender, RoutedEventArgs e) => DialogResult = true;
    private void CancelClick(object sender, RoutedEventArgs e) => DialogResult = false;

    private void PickDateClick(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement { Tag: string name } || FindName(name) is not TextBox box) return;
        var dialog = new DatePickerWindow("Выберите дату", box.Text) { Owner = this };
        if (dialog.ShowDialog() == true) box.Text = dialog.SelectedDateText;
    }
}
