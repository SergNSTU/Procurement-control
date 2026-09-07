using System.Globalization;
using System.Windows;

namespace ProcurementControl.Views;

/// <summary>Единое окно выбора даты для сделок, задач и карточек поставщиков.</summary>
public partial class DatePickerWindow : Window
{
    private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");

    public DatePickerWindow(string prompt, string initialValue = "")
    {
        InitializeComponent();
        PromptText.Text = prompt;
        if (DateTime.TryParse(initialValue, Ru, DateTimeStyles.None, out var value)) DateBox.SelectedDate = value;
        Loaded += (_, _) => DateBox.Focus();
    }

    public string SelectedDateText => DateBox.SelectedDate?.ToString("dd.MM.yyyy", Ru) ?? string.Empty;

    private void AcceptClick(object sender, RoutedEventArgs e) => DialogResult = true;
    private void ClearClick(object sender, RoutedEventArgs e) { DateBox.SelectedDate = null; DialogResult = true; }
    private void CancelClick(object sender, RoutedEventArgs e) => DialogResult = false;
}
