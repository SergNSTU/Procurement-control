using System.Windows;

namespace ProcurementControl.Services;

/// <summary>Состояние индикатора долгой операции для статусной строки.</summary>
public sealed class ProgressEventArgs : EventArgs
{
    /// <summary>Показывать полосу (есть хоть одна активная операция).</summary>
    public bool Active { get; init; }

    public string Text { get; init; } = string.Empty;

    /// <summary>Процент неизвестен — полоса в режиме «занят, но без процентов».</summary>
    public bool IsIndeterminate { get; init; } = true;

    public double Percent { get; init; }
}

/// <summary>
/// Глобальный индикатор долгих операций (Excel COM, импорт почты, синхронизация
/// Bitrix) в статусной строке главного окна — по образцу Show-Toast/ToastService.
/// Методы можно звать из любых потоков: событие уходит в UI-поток. Begin/End
/// считаются стеком, чтобы одновременные операции (почта + Bitrix) не гасили
/// друг друга: полоса прячется, только когда закончилась последняя.
/// </summary>
public static class ProgressService
{
    private static readonly object Gate = new();
    private static int _active;
    private static string _text = string.Empty;
    private static double? _percent;

    public static event EventHandler<ProgressEventArgs>? Changed;

    /// <summary>Операция началась: полоса появляется, по умолчанию без процентов.</summary>
    public static void Begin(string text)
    {
        lock (Gate)
        {
            _active++;
            _text = text;
            _percent = null;
        }

        Raise();
    }

    /// <summary>Уточняет текст текущей операции.</summary>
    public static void Report(string text)
    {
        lock (Gate)
        {
            _text = text;
        }

        Raise();
    }

    /// <summary>Текст и процент выполнения (0–100) — для операций с известным объёмом.</summary>
    public static void Report(string text, double percent)
    {
        lock (Gate)
        {
            _text = text;
            _percent = Math.Clamp(percent, 0, 100);
        }

        Raise();
    }

    /// <summary>Операция закончилась; полоса прячется, если это была последняя.</summary>
    public static void End()
    {
        lock (Gate)
        {
            _active = Math.Max(0, _active - 1);
        }

        Raise();
    }

    private static void Raise()
    {
        ProgressEventArgs args;
        lock (Gate)
        {
            args = new ProgressEventArgs
            {
                Active = _active > 0,
                Text = _text,
                IsIndeterminate = _percent is null,
                Percent = _percent ?? 0,
            };
        }

        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is null || dispatcher.CheckAccess())
        {
            Changed?.Invoke(null, args);
        }
        else
        {
            dispatcher.BeginInvoke(() => Changed?.Invoke(null, args));
        }
    }
}
