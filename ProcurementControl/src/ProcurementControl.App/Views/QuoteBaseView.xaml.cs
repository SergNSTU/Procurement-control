using System.Windows.Controls;
using ProcurementControl.Models;
using ProcurementControl.ViewModels;

namespace ProcurementControl.Views;

/// <summary>Страница «База квот»; вся логика в <c>QuoteBaseViewModel</c>.</summary>
public partial class QuoteBaseView : UserControl
{
    public QuoteBaseView()
    {
        InitializeComponent();
    }

    /// <summary>Мультиселект сетки квот — зеркалим в SelectedQuotes модели (аналог $quoteBaseGrid.SelectedRows).</summary>
    private void QuotesGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (DataContext is not QuoteBaseViewModel model)
        {
            return;
        }

        foreach (var item in e.RemovedItems)
        {
            if (item is QuoteHistoryRow row)
            {
                model.SelectedQuotes.Remove(row);
            }
        }

        foreach (var item in e.AddedItems)
        {
            if (item is QuoteHistoryRow row && !model.SelectedQuotes.Contains(row))
            {
                model.SelectedQuotes.Add(row);
            }
        }
    }
}
