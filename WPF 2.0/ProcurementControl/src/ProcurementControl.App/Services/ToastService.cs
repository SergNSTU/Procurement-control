using System.Windows;
using ProcurementControl.Views;

namespace ProcurementControl.Services;

/// <summary>Вид тоста — значения параметра $Kind из Show-Toast.</summary>
public enum ToastKind
{
    Info,
    Success,
    Warn,
    Danger,
}

/// <summary>
/// Глобальный вызов тостов — аналог Show-Toast из UiFoundation.ps1.
/// Гарантирует создание окна в UI-потоке, чтобы метод можно было звать
/// из фоновых задач (например, после асинхронного поиска цен).
/// </summary>
public static class ToastService
{
    public static void Show(string text, ToastKind kind = ToastKind.Info)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is null || dispatcher.CheckAccess())
        {
            ShowCore(text, kind);
        }
        else
        {
            dispatcher.BeginInvoke(() => ShowCore(text, kind));
        }
    }

    private static void ShowCore(string text, ToastKind kind)
    {
        new ToastWindow(text, kind.ToString()).Show();
    }
}
