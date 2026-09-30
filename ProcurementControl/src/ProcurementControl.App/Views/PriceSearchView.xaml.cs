using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Input;
using ProcurementControl.Models;

namespace ProcurementControl.Views;

public partial class PriceSearchView : UserControl
{
    public PriceSearchView()
    {
        InitializeComponent();
    }

    /// <summary>Аналог $priceSearchGrid.Add_CellDoubleClick: открыть ссылку выбранной строки.</summary>
    private void ResultGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (ResultGrid.SelectedItem is not PriceSearchRow row || !row.HasLink)
        {
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo(row.Link) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(ex.Message, "Открытие ссылки");
        }
    }
}
