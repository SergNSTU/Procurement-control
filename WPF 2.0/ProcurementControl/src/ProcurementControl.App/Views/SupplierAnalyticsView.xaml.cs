using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ProcurementControl.Views;

public partial class SupplierAnalyticsView : UserControl
{
    public SupplierAnalyticsView() => InitializeComponent();

    /// <summary>Колонки сводки «Объёмы» строятся из DataTable: поставщик слева, суммы справа.</summary>
    private void VolumesGrid_AutoGeneratingColumn(object? sender, DataGridAutoGeneratingColumnEventArgs e)
    {
        var column = (DataGridTextColumn)e.Column;
        if (e.PropertyName == "Поставщик")
        {
            column.Width = new DataGridLength(1, DataGridLengthUnitType.Star, 140, 220);
            column.ElementStyle = (Style)FindResource("VolumeSupplierCell");
            return;
        }
        column.ElementStyle = (Style)FindResource("VolumeAmountCell");
        column.Width = e.PropertyName == "Итого, $" ? 105 : 82;
        if (e.PropertyName == "Итого, $")
        {
            column.ElementStyle = new Style(typeof(TextBlock), (Style)FindResource("VolumeAmountCell"))
            {
                Setters = { new Setter(TextBlock.FontWeightProperty, FontWeights.Bold) }
            };
        }
    }
}
