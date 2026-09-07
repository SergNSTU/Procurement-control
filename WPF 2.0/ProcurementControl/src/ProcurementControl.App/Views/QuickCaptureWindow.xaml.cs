using System.Windows;
using System.Windows.Media;
using ProcurementControl.Services;

namespace ProcurementControl.Views;

/// <summary>
/// Окно быстрого захвата — порт диалога агента (аналог
/// Show-QuickCaptureDialog из QuickCaptureAgent.ps1, строки 335-446),
/// перерисованный в стиле дизайн-системы приложения.
/// </summary>
public partial class QuickCaptureWindow : Window
{
    private readonly string _kind;
    private readonly string _sourceText;

    public QuickCaptureWindow(string kind, string text)
    {
        InitializeComponent();
        _kind = kind;
        _sourceText = Normalize(text);

        ApplyKindTheme();
        FillInitialValues();
        Loaded += (_, _) =>
        {
            TitleBox.Focus();
            TitleBox.CaretIndex = TitleBox.Text.Length;
        };
    }

    /// <summary>Акценты и состав полей различаются для задач и напоминаний.</summary>
    private void ApplyKindTheme()
    {
        var isReminder = _kind == "reminder";
        if (isReminder)
        {
            var amber = (Brush)FindResource("AmberBrush");
            var amberLight = (Brush)FindResource("AmberLightBrush");
            AccentBar.Background = amber;
            IconBox.Background = amberLight;
            IconGlyph.Foreground = amber;
            IconGlyph.Text = "\uE823";
            HeadingText.Text = "Новое напоминание";
            Title = "Создать напоминание";
            DealBlock.Visibility = Visibility.Collapsed;
            DateCaption.Text = "Напомнить";
            HasDueCheck.Visibility = Visibility.Collapsed;
            DateHint.Visibility = Visibility.Visible;
            SaveButton.Content = "Создать напоминание";
            SaveButton.Style = (Style)FindResource("AmberButtonStyle");
        }
        else
        {
            Title = "Создать задачу";
            DateHint.Visibility = Visibility.Collapsed;
            DueDatePicker.IsEnabled = false;
            DueTimeBox.IsEnabled = false;
        }
    }

    private void FillInitialValues()
    {
        var title = _sourceText;
        if (title.Length > 240)
        {
            title = title[..237] + "...";
        }

        TitleBox.Text = title;
        var defaultDue = DateTime.Now.AddHours(1);
        DueDatePicker.SelectedDate = defaultDue.Date;
        DueTimeBox.Text = defaultDue.ToString("HH:mm");
    }

    private void HasDueChanged(object sender, RoutedEventArgs e)
    {
        var enabled = HasDueCheck.IsChecked == true;
        DueDatePicker.IsEnabled = enabled;
        DueTimeBox.IsEnabled = enabled;
    }

    private void CancelClicked(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    /// <summary>Аналог обработчика кнопки ОК (строки 434-441 оригинала).</summary>
    private void SaveClicked(object sender, RoutedEventArgs e)
    {
        var isReminder = _kind == "reminder";
        var title = TitleBox.Text.Replace("\0", string.Empty).Trim();
        if (title.Length == 0)
        {
            ToastService.Show("Введите заголовок.", ToastKind.Danger);
            TitleBox.Focus();
            return;
        }

        var needDue = isReminder || HasDueCheck.IsChecked == true;
        var dueDate = string.Empty;
        if (needDue)
        {
            var parsed = TryBuildDueDate();
            if (parsed is null)
            {
                ToastService.Show(
                    isReminder ? "Для напоминания укажите дату и время." : "Не удалось разобрать дату и время.",
                    ToastKind.Danger);
                return;
            }

            dueDate = parsed;
        }

        try
        {
            if (isReminder)
            {
                PurchaseWriteRepository.SaveReminder(title, dueDate, source: "quick_capture");
            }
            else
            {
                PurchaseWriteRepository.CreateQuickCaptureTask(title, _sourceText, dueDate, DealBox.Text);
            }
        }
        catch (Exception ex)
        {
            ToastService.Show(ex.Message, ToastKind.Danger);
            return;
        }

        DialogResult = true;
    }

    /// <summary>Дата из календаря + время «чч:мм» → формат базы «гггг-мм-дд чч:мм:сс».</summary>
    private string? TryBuildDueDate()
    {
        if (DueDatePicker.SelectedDate is not DateTime date)
        {
            return null;
        }

        if (!TimeSpan.TryParse(DueTimeBox.Text.Trim(), out var time))
        {
            return null;
        }

        return date.Date.Add(time).ToString("yyyy-MM-dd HH:mm:ss");
    }

    private static string Normalize(string value)
        => (value ?? string.Empty).Replace("\0", string.Empty).Trim();
}
