using System.Windows;
using System.Windows.Input;
using ProcurementControl.ViewModels;

namespace ProcurementControl.Views;

/// <summary>
/// Большое окно preview «Сравнения RRFQ»: те же объекты <see cref="RrfqViewModel"/>,
/// что и на вкладке (фильтр, строки, квоты), поэтому выбор победителя из окна
/// сразу виден в обеих сетках. Окно немодальное — вкладку можно не бросать.
/// </summary>
public partial class RrfqPreviewWindow : Window
{
    private readonly RrfqViewModel _viewModel;

    public RrfqPreviewWindow(RrfqViewModel viewModel)
    {
        InitializeComponent();
        _viewModel = viewModel;
        DataContext = viewModel;
        Title = string.IsNullOrWhiteSpace(viewModel.DealContextText)
            ? "Preview RRFQ"
            : "Preview RRFQ — " + viewModel.DealContextText;
    }

    /// <summary>Двойной клик по квоте — сделать её победителем (как на вкладке).</summary>
    private void QuoteGridMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (_viewModel.SetWinnerCommand.CanExecute(null))
        {
            _viewModel.SetWinnerCommand.Execute(null);
        }
    }

    private void CloseClick(object sender, RoutedEventArgs e) => Close();
}
