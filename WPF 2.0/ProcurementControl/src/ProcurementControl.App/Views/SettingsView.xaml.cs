using System.Windows;
using System.Windows.Controls;
using ProcurementControl.ViewModels;

namespace ProcurementControl.Views;

/// <summary>Страница «Настройки»; вся логика в <c>SettingsViewModel</c>.</summary>
public partial class SettingsView : UserControl
{
    public SettingsView()
    {
        InitializeComponent();
        // Страницы кэшируются как ViewModel, поэтому список копий иначе остался бы
        // таким, каким был при первом открытии: автокопия создаётся в фоне при
        // старте и в списке не появилась бы. Обновляем при каждом показе страницы.
        IsVisibleChanged += (_, args) =>
        {
            if (args.NewValue is true && DataContext is SettingsViewModel viewModel)
            {
                viewModel.RefreshRecovery();
            }
        };
    }
}
