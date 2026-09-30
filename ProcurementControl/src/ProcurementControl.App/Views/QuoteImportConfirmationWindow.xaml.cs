using System.Windows;
using ProcurementControl.Models;

namespace ProcurementControl.Views;

/// <summary>Предпросмотр найденных квот перед записью в общую базу.</summary>
public partial class QuoteImportConfirmationWindow : Window
{
    public QuoteImportConfirmationWindow(IReadOnlyList<Quote> quotes, IEnumerable<string> paths)
    {
        InitializeComponent();
        Quotes = quotes;
        DataContext = this;
        SummaryText.Text = $"Найдено квот: {quotes.Count}. Проверьте список перед импортом.\n" +
            string.Join("\n", paths.Select(path => "• " + System.IO.Path.GetFileName(path)));
    }

    public IReadOnlyList<Quote> Quotes { get; }
    private void AcceptClick(object sender, RoutedEventArgs e) => DialogResult = true;
    private void CancelClick(object sender, RoutedEventArgs e) => DialogResult = false;
}
