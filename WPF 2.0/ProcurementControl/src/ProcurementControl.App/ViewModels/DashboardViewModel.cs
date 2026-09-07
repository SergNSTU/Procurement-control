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
    private string _periodLabel = string.Empty;

    [ObservableProperty]
    private string _periodUnitLabel = "месяц";

    [ObservableProperty]
    private string _trackingStartedText = string.Empty;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    // Сборки (deals)
    [ObservableProperty] private int _assembliesAdded;
    [ObservableProperty] private int _assembliesOrdered;
    [ObservableProperty] private string _assembliesOrderAmount = "0 $";
    [ObservableProperty] private string _assembliesCurrentConversion = "—";
    [ObservableProperty] private string _assembliesMatureConversion = "—";

    // Компоненты (component_deals)
    [ObservableProperty] private int _componentsAdded;
    [ObservableProperty] private int _componentsOrdered;
    [ObservableProperty] private string _componentsOrderAmount = "0 $";
    [ObservableProperty] private string _componentsCurrentConversion = "—";
    [ObservableProperty] private string _componentsMatureConversion = "—";

    public ObservableCollection<MonthHistoryRow> History { get; } = new();
    public ObservableCollection<int> AvailableYears { get; } = new();
    public IReadOnlyList<DashboardPeriodModeOption> PeriodModes { get; } = new[]
    {
        new DashboardPeriodModeOption(DashboardPeriodMode.Month, "Месяц"),
        new DashboardPeriodModeOption(DashboardPeriodMode.Quarter, "Квартал")
    };
    public ObservableCollection<DashboardPeriodOption> PeriodOptions { get; } = new();

    [ObservableProperty] private int _selectedYear;
    [ObservableProperty] private DashboardPeriodMode _selectedPeriodMode = DashboardPeriodMode.Month;
    [ObservableProperty] private DashboardPeriodModeOption? _selectedPeriodModeOption;
    [ObservableProperty] private DashboardPeriodOption? _selectedPeriod;
    private bool _isInitialized;

    public DashboardViewModel()
    {
        foreach (var year in _service.GetAvailableYears()) AvailableYears.Add(year);
        SelectedYear = DateTime.Today.Year;
        SelectedPeriodModeOption = PeriodModes.First();
        RebuildPeriodOptions();
        SelectedPeriod = PeriodOptions.First(option => option.Value == DateTime.Today.Month);
        _isInitialized = true;
        Reload();
    }

    partial void OnSelectedYearChanged(int value) => ReloadForSelectionChange();

    partial void OnSelectedPeriodModeChanged(DashboardPeriodMode value)
    {
        RebuildPeriodOptions();
        ReloadForSelectionChange();
    }

    partial void OnSelectedPeriodModeOptionChanged(DashboardPeriodModeOption? value)
    {
        if (value is not null) SelectedPeriodMode = value.Value;
    }

    partial void OnSelectedPeriodChanged(DashboardPeriodOption? value) => ReloadForSelectionChange();

    [RelayCommand]
    private void Reload()
    {
        try
        {
            if (SelectedPeriod is null) return;
            var data = _service.GetDashboard(new DashboardPeriod(SelectedYear, SelectedPeriodMode, SelectedPeriod.Value));
            ErrorMessage = string.Empty;

            PeriodLabel = data.PeriodLabel;
            PeriodUnitLabel = data.PeriodUnitLabel;
            TrackingStartedText = data.TrackingStartedAt is null
                ? "—"
                : data.TrackingStartedAt.Value.ToString("dd.MM.yyyy", Ru);

            AssembliesAdded = data.Assemblies.AddedThisMonth;
            AssembliesOrdered = data.Assemblies.OrderedThisMonth;
            AssembliesOrderAmount = FormatAmount(data.Assemblies.OrderAmountUsd);
            AssembliesCurrentConversion = FormatPercent(data.Assemblies.CurrentConversion);
            AssembliesMatureConversion = FormatPercent(data.Assemblies.MatureConversion);

            ComponentsAdded = data.Components.AddedThisMonth;
            ComponentsOrdered = data.Components.OrderedThisMonth;
            ComponentsOrderAmount = FormatAmount(data.Components.OrderAmountUsd);
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

    private static string FormatAmount(decimal value)
        => value.ToString("N2", Ru) + " $";

    private void ReloadForSelectionChange()
    {
        if (_isInitialized) Reload();
    }

    private void RebuildPeriodOptions()
    {
        var previousValue = SelectedPeriod?.Value;
        PeriodOptions.Clear();
        if (SelectedPeriodMode == DashboardPeriodMode.Month)
        {
            for (var month = 1; month <= 12; month++)
                PeriodOptions.Add(new DashboardPeriodOption(month, new DateTime(SelectedYear, month, 1).ToString("MMMM", Ru)));
        }
        else
        {
            for (var quarter = 1; quarter <= 4; quarter++)
                PeriodOptions.Add(new DashboardPeriodOption(quarter, quarter + " квартал"));
        }
        SelectedPeriod = PeriodOptions.FirstOrDefault(option => option.Value == previousValue) ?? PeriodOptions.First();
    }
}
