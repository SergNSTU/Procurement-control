using System.Windows;
using ProcurementControl.ViewModels;

namespace ProcurementControl;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainViewModel();
    }
}
