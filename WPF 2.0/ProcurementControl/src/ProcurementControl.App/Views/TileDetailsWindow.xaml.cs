using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ProcurementControl.Models;

namespace ProcurementControl.Views;

/// <summary>
/// Список сделок/пунктов за плиткой метрик «Контроля закупки»: открывается
/// по клику на плитку; двойной клик по строке переходит к сделке.
/// </summary>
public partial class TileDetailsWindow : Window
{
    private readonly Action<long, string> _openDeal;

    public TileDetailsWindow(string title, IReadOnlyList<TileLineRow> lines, Action<long, string> openDeal)
    {
        InitializeComponent();
        Title = title;
        TitleText.Text = title + ": " + lines.Count;
        LinesList.ItemsSource = lines;
        _openDeal = openDeal;
    }

    /// <summary>Двойной щелчок по строке — выделить сделку в таблице и закрыть список.</summary>
    private void LinesList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not ListBox list ||
            e.OriginalSource is not DependencyObject source ||
            ItemsControl.ContainerFromElement(list, source) is not ListBoxItem ||
            list.SelectedItem is not TileLineRow line)
        {
            return;
        }

        if (line.DealId > 0 || !string.IsNullOrWhiteSpace(line.DealNumber))
        {
            _openDeal(line.DealId, line.DealNumber);
            Close();
        }
    }

    private void CloseClick(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
