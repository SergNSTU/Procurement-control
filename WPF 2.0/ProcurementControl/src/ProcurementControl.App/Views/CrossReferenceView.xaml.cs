using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Input;
using ProcurementControl.Models;

namespace ProcurementControl.Views;

public partial class CrossReferenceView : UserControl
{
    public CrossReferenceView()
    {
        InitializeComponent();
    }

    /// <summary>Двойной клик по строке открывает карточку компонента на lcsc.com.</summary>
    private void ResultGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (ResultGrid.SelectedItem is not CrossReferenceRow row || !row.HasLink)
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
