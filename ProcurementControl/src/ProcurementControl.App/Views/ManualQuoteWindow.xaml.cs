using System.Windows;
using ProcurementControl.ViewModels;

namespace ProcurementControl.Views;

/// <summary>Окно «Ручной ввод квоты»; вся логика в <see cref="ManualQuoteViewModel"/>.</summary>
public partial class ManualQuoteWindow : Window
{
    public ManualQuoteWindow(ManualQuoteViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
        viewModel.CloseRequested += (_, _) => Close();
    }
}
