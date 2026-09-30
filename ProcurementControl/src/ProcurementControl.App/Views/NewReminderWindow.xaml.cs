using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace ProcurementControl.Views;

/// <summary>
/// Единое окно создания напоминания: текст и дата календарём — замена двум
/// текстовым диалогам подряд («Что нужно сделать?», «Дата дд.мм.гггг»).
/// </summary>
public partial class NewReminderWindow : Window
{
    private static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");

    /// <summary>Подсказка о привязке к выбранному пункту; пустая строка скрывает её.</summary>
    public NewReminderWindow(string contextCaption = "")
    {
        InitializeComponent();
        ContextText.Text = contextCaption;
        ContextText.Visibility = string.IsNullOrEmpty(contextCaption)
            ? Visibility.Collapsed
            : Visibility.Visible;
        Loaded += (_, _) => TitleBox.Focus();
    }

    public string ReminderTitle => TitleBox.Text.Trim();

    /// <summary>Дата в формате дд.мм.гггг или пустая строка, если дата не выбрана.</summary>
    public string DueDateText => DateBox.SelectedDate?.ToString("dd.MM.yyyy", Ru) ?? string.Empty;

    private void AcceptClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(TitleBox.Text))
        {
            MessageBox.Show("Введите текст напоминания.", "Новое напоминание");
            TitleBox.Focus();
            return;
        }

        DialogResult = true;
    }

    private void ClearDateClick(object sender, RoutedEventArgs e) => DateBox.SelectedDate = null;

    private void CancelClick(object sender, RoutedEventArgs e) => DialogResult = false;
}
