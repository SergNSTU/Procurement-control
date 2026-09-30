using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace ProcurementControl.Views;

/// <summary>
/// Аналог Show-Toast из UiFoundation.ps1 (строки 179-218): тёмное всплывающее
/// окно без рамки в правом нижнем углу экрана, цветная полоса по виду,
/// автозакрытие через 3.5 секунды.
/// </summary>
public partial class ToastWindow : Window
{
    private const int LifetimeMs = 3500;

    public ToastWindow(string text, string kind = "Info")
    {
        InitializeComponent();
        ToastText.Text = text;
        KindBar.Fill = new SolidColorBrush(KindColor(kind));

        // Как $toast.Add_Shown: позиция от рабочей области экрана.
        Loaded += (_, _) =>
        {
            var area = SystemParameters.WorkArea;
            Left = area.Right - ActualWidth - 24;
            Top = area.Bottom - ActualHeight - 56;

            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(LifetimeMs) };
            timer.Tick += (_, _) =>
            {
                timer.Stop();
                Close();
            };
            timer.Start();
        };
    }

    /// <summary>Цвета полосы — палитра оригинального Show-Toast.</summary>
    private static Color KindColor(string kind) => kind switch
    {
        "Success" => Color.FromRgb(74, 222, 128),
        "Warn" => Color.FromRgb(250, 204, 21),
        "Danger" => Color.FromRgb(248, 113, 113),
        _ => Color.FromRgb(96, 165, 250),
    };
}
