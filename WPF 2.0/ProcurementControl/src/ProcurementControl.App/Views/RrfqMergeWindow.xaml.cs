using System.Windows;
using ProcurementControl.Models;

namespace ProcurementControl.Views;

public partial class RrfqMergeWindow : Window
{
    public RrfqMergeWindow(IReadOnlyList<PurchaseDocumentRow> rfqs, IReadOnlyList<PurchaseDocumentRow> responses)
    {
        InitializeComponent();
        RfqCombo.ItemsSource = rfqs;
        RfqCombo.SelectedIndex = 0;
        ResponsesGrid.DataContext = responses;
        foreach (var response in responses) ResponsesGrid.SelectedItems.Add(response);
    }

    public PurchaseDocumentRow? SelectedRfq => RfqCombo.SelectedItem as PurchaseDocumentRow;
    public IReadOnlyList<PurchaseDocumentRow> SelectedResponses => ResponsesGrid.SelectedItems.OfType<PurchaseDocumentRow>().ToList();
    private void AcceptClick(object sender, RoutedEventArgs e) => DialogResult = true;
    private void CancelClick(object sender, RoutedEventArgs e) => DialogResult = false;
}
