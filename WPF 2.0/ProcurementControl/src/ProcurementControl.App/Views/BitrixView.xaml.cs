using System.Windows.Controls;
using System.Windows.Input;

namespace ProcurementControl.Views;

public partial class BitrixView : UserControl
{
    public BitrixView()
    {
        InitializeComponent();
    }

    /// <summary>
    /// Аналог $bitrixGrid.Add_CellDoubleClick: двойной клик по колонке «Ссылка»
    /// открывает задачу в браузере (в остальных колонках оригинал молчит).
    /// </summary>
    private void TasksGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not DataGrid grid || grid.CurrentCell.Column is null)
        {
            return;
        }

        if (grid.CurrentCell.Column.Header as string == "Ссылка" && DataContext is ViewModels.BitrixViewModel vm)
        {
            vm.OpenLinkCommand.Execute(null);
        }
    }
}
