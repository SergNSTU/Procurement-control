using System.Windows;
using ProcurementControl.ViewModels;

namespace ProcurementControl.Views;

/// <summary>Окно «Ручное сопоставление»; вся логика в <see cref="ManualMatchViewModel"/>.</summary>
public partial class ManualMatchWindow : Window
{
    public ManualMatchWindow(ManualMatchViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
        viewModel.CloseRequested += (_, _) => Close();
    }
}
