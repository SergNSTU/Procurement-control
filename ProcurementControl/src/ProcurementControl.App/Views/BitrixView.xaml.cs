using System.Windows.Controls;
using System.Windows.Input;

namespace ProcurementControl.Views;

public partial class BitrixView : UserControl
{
    public BitrixView()
    {
        InitializeComponent();
    }

    /// <summary>Двойной клик по строке открывает задачу в Bitrix.</summary>
    private void RowsGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is ViewModels.BitrixViewModel vm && vm.OpenLinkCommand.CanExecute(null))
        {
            vm.OpenLinkCommand.Execute(null);
        }
    }
}
