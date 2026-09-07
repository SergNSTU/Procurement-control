using System.Windows;
using ProcurementControl.Services;

namespace ProcurementControl.Views;

/// <summary>
/// Аналог Show-DocumentTypeDialog из оригинала (RRFQComparer.ps1,
/// строки 1474-1517): выбор типа для загружаемых файлов.
/// </summary>
public partial class DocumentTypeWindow : Window
{
    public DocumentTypeWindow()
    {
        InitializeComponent();
        foreach (var type in PurchaseDocumentsStore.DocumentTypes)
        {
            TypeCombo.Items.Add(type);
        }
        TypeCombo.SelectedIndex = 0;
        Loaded += (_, _) => TypeCombo.Focus();
    }

    /// <summary>Выбранный тип документа (актуально после DialogResult == true).</summary>
    public string SelectedType => TypeCombo.SelectedItem as string ?? string.Empty;

    private void AcceptClick(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void CancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
