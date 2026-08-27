using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ProcurementControl.Models;
using ProcurementControl.Services;

namespace ProcurementControl.ViewModels;

/// <summary>
/// Read-only страница «Дашборд». Переносит отображение метрик из оригинала,
/// данные читаются через безопасный снапшот (исходная БД не изменяется).
/// </summary>
public partial class DashboardViewModel : ObservableObject
{
    private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");
    private readonly DashboardService _service = new();

    [ObservableProperty]
    private string _monthLabel = string.Empty;

    [ObservableProperty]
    private string _trackingStartedText = string.Empty;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    // Сборки (deals)
    [ObservableProperty] private int _assembliesAdded;
    [ObservableProperty] private int _assembliesOrdered;
    [ObservableProperty] private string _assembliesCurrentConversion = "—";
    [ObservableProperty] private string _assembliesMatureConversion = "—";

    // Компоненты (component_deals)
    [ObservableProperty] private int _componentsAdded;
    [ObservableProperty] private int _componentsOrdered;
    [ObservableProperty] private string _componentsCurrentConversion = "—";
    [ObservableProperty] private string _componentsMatureConversion = "—";

    public ObservableCollection<MonthHistoryRow> History { get; } = new();

    public DashboardViewModel()
    {
        Reload();
    }

    [RelayCommand]
    private void Reload()
    {
        try
        {
            var data = _service.GetDashboard();
            ErrorMessage = string.Empty;

            MonthLabel = data.MonthLabel;
            TrackingStartedText = data.TrackingStartedAt is null
                ? "—"
                : data.TrackingStartedAt.Value.ToString("dd.MM.yyyy", Ru);

            AssembliesAdded = data.Assemblies.AddedThisMonth;
            AssembliesOrdered = data.Assemblies.OrderedThisMonth;
            AssembliesCurrentConversion = FormatPercent(data.Assemblies.CurrentConversion);
            AssembliesMatureConversion = FormatPercent(data.Assemblies.MatureConversion);

            ComponentsAdded = data.Components.AddedThisMonth;
            ComponentsOrdered = data.Components.OrderedThisMonth;
            ComponentsCurrentConversion = FormatPercent(data.Components.CurrentConversion);
            ComponentsMatureConversion = FormatPercent(data.Components.MatureConversion);

            History.Clear();
            foreach (var row in data.History)
            {
                History.Add(row);
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = "Не удалось прочитать данные: " + ex.Message;
        }
    }

    private static string FormatPercent(double? value)
        => value is null ? "—" : value.Value.ToString("0.0", Ru) + " %";
}
