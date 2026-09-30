using System.Windows;
using System.Windows.Input;

namespace ProcurementControl.Views;

/// <summary>
/// Постоянное окно уведомления о загруженных из почты RRFQ: висит поверх окон,
/// пока пользователь не нажмёт ОК или крестик. Новые файлы дописываются в
/// открытое окно. Создаётся через MailImportNotifier.
/// </summary>
public partial class MailImportWindow : Window
{
    private readonly List<string> _lines = new();

    public MailImportWindow()
    {
        InitializeComponent();
    }

    /// <summary>Дописывает строки и обновляет текст (окно можно переиспользовать).</summary>
    public void AppendLines(IEnumerable<string> lines)
    {
        _lines.AddRange(lines);
        LinesText.Text = string.Join(Environment.NewLine, _lines);
    }

    private void OkClicked(object sender, RoutedEventArgs e) => Close();

    private void CloseClicked(object sender, RoutedEventArgs e) => Close();

    private void BorderDragMove(object sender, MouseButtonEventArgs e)
    {
        if (e.ButtonState == MouseButtonState.Pressed)
        {
            try
            {
                DragMove();
            }
            catch (InvalidOperationException)
            {
                // DragMove бросает исключение, если кнопка уже отпущена — не критично.
            }
        }
    }
}
