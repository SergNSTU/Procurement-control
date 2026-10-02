using System.Windows;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.Views;

public partial class DocumentBindingWindow : Window
{
    private sealed record SupplierChoice(long Id, string Supplier);

    public DocumentBindingWindow(PurchaseDocumentRow document, IEnumerable<SupplierRow> suppliers)
    {
        InitializeComponent();
        foreach (var type in PurchaseDocumentsStore.DocumentTypes) TypeCombo.Items.Add(type);
        TypeCombo.SelectedItem = document.DocumentType;
        SupplierCombo.Items.Add(new SupplierChoice(0, "Поставщик не указан"));
        foreach (var supplier in suppliers) SupplierCombo.Items.Add(new SupplierChoice(supplier.Id, supplier.Supplier));
        SupplierCombo.SelectedItem = SupplierCombo.Items.OfType<SupplierChoice>().FirstOrDefault(item => item.Id == document.SupplierId)
            ?? SupplierCombo.Items.OfType<SupplierChoice>().First();
        TypeCombo.SelectionChanged += (_, _) => UpdateSupplierState();
        UpdateSupplierState();
    }

    public string SelectedType => TypeCombo.SelectedItem as string ?? string.Empty;
    public long SelectedSupplierId => (SupplierCombo.SelectedItem as SupplierChoice)?.Id ?? 0;

    private void UpdateSupplierState()
    {
        var required = string.Equals(SelectedType, "RRFQ", StringComparison.OrdinalIgnoreCase);
        SupplierCombo.IsEnabled = required;
        if (!required) SupplierCombo.SelectedIndex = 0;
    }

    private void AcceptClick(object sender, RoutedEventArgs e)
    {
        if (string.Equals(SelectedType, "RRFQ", StringComparison.OrdinalIgnoreCase) && SelectedSupplierId <= 0)
        {
            MessageBox.Show("Выберите поставщика для RRFQ.", "Документ");
            return;
        }
        DialogResult = true;
    }

    private void CancelClick(object sender, RoutedEventArgs e) => DialogResult = false;
}
