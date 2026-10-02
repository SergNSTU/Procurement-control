using System.Windows;
using ProcurementControl.Views;

namespace ProcurementControl.Services;

/// <summary>
/// Окно отчёта по приходу со склада: каждый новый отчёт заменяет содержимое —
/// окно всегда показывает актуальное состояние (вопросы и сделанные записи).
/// Гарантирует создание окна в UI-потоке.
/// </summary>
public static class WarehouseArrivalNotifier
{
    private static WarehouseArrivalWindow? _window;

    public static void Show(WarehouseArrivalReport report)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is null)
        {
            return;
        }

        if (!dispatcher.CheckAccess())
        {
            dispatcher.BeginInvoke(() => Show(report));
            return;
        }

        ShowCore(report);
    }

    private static void ShowCore(WarehouseArrivalReport report)
    {
        if (_window is null)
        {
            _window = new WarehouseArrivalWindow();
            _window.Closed += (_, _) => _window = null;
        }

        _window.ResetReport();
        _window.AppendReport(report);
        _window.Show();
        _window.Activate();
    }

    /// <summary>Закрыть отчёт при выходе из приложения.</summary>
    public static void Shutdown()
    {
        var window = _window;
        _window = null;
        window?.Close();
    }
}
