using CommunityToolkit.Mvvm.ComponentModel;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Заглушка для разделов, которые ещё не перенесены из PowerShell-версии.
/// </summary>
public partial class PlaceholderViewModel : ObservableObject
{
    [ObservableProperty]
    private string _title = string.Empty;

    public PlaceholderViewModel(string title)
    {
        _title = title;
    }
}
