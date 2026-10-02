using System.Windows;
using ProcurementControl.Views;

namespace ProcurementControl.Services;

/// <summary>
/// Постоянное уведомление о пришедших RRFQ файлах: окно висит до нажатия ОК
/// или закрытия (в отличие от тостов, сами не исчезают). Если окно уже открыто,
/// новые строки дописываются в него. Гарантирует создание окна в UI-потоке.
/// </summary>
public static class MailImportNotifier
{
    private static MailImportWindow? _window;

    public static void Show(IReadOnlyList<string> lines)
    {
        if (lines.Count == 0)
        {
            return;
        }

        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is null)
        {
            return;
        }

        if (!dispatcher.CheckAccess())
        {
            dispatcher.BeginInvoke(() => Show(lines));
            return;
        }

        ShowCore(lines);
    }

    private static void ShowCore(IReadOnlyList<string> lines)
    {
        if (_window is null)
        {
            _window = new MailImportWindow();
            _window.Closed += (_, _) => _window = null;
            _window.AppendLines(lines);
            _window.Show();
            _window.Activate();
            return;
        }

        _window.AppendLines(lines);
        _window.Activate();
    }

    /// <summary>Закрыть уведомление при выходе из приложения.</summary>
    public static void Shutdown()
    {
        var window = _window;
        _window = null;
        window?.Close();
    }
}
