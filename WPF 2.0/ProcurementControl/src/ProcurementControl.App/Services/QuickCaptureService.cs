using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using ProcurementControl.Views;

namespace ProcurementControl.Services;

/// <summary>
/// Быстрый захват прямо в WPF-приложении — порт агента
/// app\QuickCaptureAgent.ps1 (хоткеи, захват выделения, диалог):
/// глобальные сочетания Ctrl+Shift+T / Ctrl+Shift+R (с запасным
/// вариантом Ctrl+Alt, как RegisterPair в оригинале, строки 72-81),
/// копирование выделения через Ctrl+C с восстановлением буфера
/// (Get-SelectedText / Restore-QuickCaptureClipboard, строки 448-477),
/// собственное окно вместо WinForms-диалога агента и HTTP-слушатель на
/// 127.0.0.1:8765 для браузерного расширения (строки 498-531).
/// </summary>
public static class QuickCaptureService
{
    private const int WmHotkey = 0x0312;
    private const int TaskHotkeyId = 1;
    private const int ReminderHotkeyId = 2;
    private const uint ModControl = 0x0002;
    private const uint ModShift = 0x0004;
    private const uint ModAlt = 0x0001;
    private const uint ModNoRepeat = 0x4000;

    private static HwndSource? _hwndSource;
    private static bool _captureInProgress;
    private static TcpListener? _httpListener;

    /// <summary>Порт браузерного расширения — $Port из параметров агента (строка 3).</summary>
    private const int HttpPort = 8765;

    /// <summary>Журнал — аналог Write-QuickCaptureLog агента.</summary>
    private static void Log(string message)
    {
        try
        {
            var logPath = Path.Combine(AppPaths.DataRoot, "logs", "quick_capture_wpf.log");
            Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
            File.AppendAllText(logPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}");
        }
        catch
        {
            // Журнал не должен мешать захвату.
        }
    }

    /// <summary>Данные добавлены — главный экран обновляет кэшированные страницы.</summary>
    public static event Action? DataChanged;

    /// <summary>Зарегистрированное сочетание: «Ctrl+Shift» или «Ctrl+Alt».</summary>
    public static string HotkeyMode { get; private set; } = "Ctrl+Shift";

    /// <summary>
    /// Регистрация глобальных хоткеев на окне. Вызывается из
    /// MainWindow.SourceInitialized (аналог $hotkeyWindow.Register()).
    /// </summary>
    public static void Register(Window window)
    {
        // HTTP для браузерного расширения работает независимо от хоткеев,
        // как в оригинале: $hotkeyWindow.Register() и затем Start-QuickCaptureHttp.
        StartHttp();
        window.Closed += (_, _) => Shutdown();

        try
        {
            var handle = new WindowInteropHelper(window).EnsureHandle();
            _hwndSource = HwndSource.FromHwnd(handle);
            if (_hwndSource is null)
            {
                return;
            }

            if (TryRegisterPair(handle, ModControl | ModShift | ModNoRepeat))
            {
                HotkeyMode = "Ctrl+Shift";
            }
            else if (TryRegisterPair(handle, ModControl | ModAlt | ModNoRepeat))
            {
                HotkeyMode = "Ctrl+Alt";
            }
            else
            {
                // Сочетания заняты другим приложением — захват молча недоступен.
                Log("Не удалось зарегистрировать сочетания быстрого захвата.");
                return;
            }

            _hwndSource.AddHook(WndProc);
            Log($"Хоткеи зарегистрированы: {HotkeyMode}+T, {HotkeyMode}+R.");
        }
        catch
        {
            // Хоткеи — необязательная надстройка: приложение работает и без них.
        }
    }

    private static void Shutdown()
    {
        try
        {
            StopHttp();
            if (_hwndSource is not null)
            {
                _hwndSource.RemoveHook(WndProc);
                UnregisterHotKey(_hwndSource.Handle, TaskHotkeyId);
                UnregisterHotKey(_hwndSource.Handle, ReminderHotkeyId);
                _hwndSource = null;
            }
        }
        catch
        {
            // Окно уже закрывается — ошибки снятия хоткеев не важны.
        }
    }

    /// <summary>
    /// Порт Start-QuickCaptureHttp (строки 498-501): TcpListener на loopback.
    /// Порт занят (например, запущен старый агент) — расширение просто не достучится.
    /// </summary>
    private static void StartHttp()
    {
        try
        {
            var listener = new TcpListener(IPAddress.Loopback, HttpPort);
            listener.Start();
            _httpListener = listener;
            Log($"HTTP-слушатель запущен: 127.0.0.1:{HttpPort}.");
            _ = Task.Run(AcceptLoop);
        }
        catch (Exception exception)
        {
            _httpListener = null;
            Log($"Не удалось запустить HTTP-слушатель: {exception.Message}");
        }
    }

    private static void StopHttp()
    {
        var listener = _httpListener;
        _httpListener = null;
        try
        {
            listener?.Stop();
        }
        catch
        {
            // Слушатель уже остановлен.
        }
    }

    private static async Task AcceptLoop()
    {
        while (true)
        {
            var listener = _httpListener;
            if (listener is null)
            {
                break;
            }

            TcpClient client;
            try
            {
                client = await listener.AcceptTcpClientAsync();
            }
            catch
            {
                // Listener остановлен — выходим из цикла приёма.
                break;
            }

            _ = Task.Run(() => HandleHttpClient(client));
        }
    }

    /// <summary>
    /// Порт Process-QuickCaptureHttp (строки 503-531): разбор запроса,
    /// валидация {kind, text}, открытие формы и ответ расширению.
    /// Ответ уходит после закрытия формы — как в оригинале (строка 527-528).
    /// </summary>
    private static void HandleHttpClient(TcpClient client)
    {
        using (client)
        {
            NetworkStream? stream = null;
            try
            {
                stream = client.GetStream();
                stream.ReadTimeout = 5000;
                var buffer = new byte[65536];
                var count = stream.Read(buffer, 0, buffer.Length);
                var raw = Encoding.UTF8.GetString(buffer, 0, count);
                var parts = raw.Split(new[] { "\r\n\r\n" }, 2, StringSplitOptions.None);
                var requestLine = parts[0].Split(new[] { "\r\n" }, 2, StringSplitOptions.None)[0];
                if (requestLine.StartsWith("OPTIONS ", StringComparison.OrdinalIgnoreCase))
                {
                    SendHttpResponse(stream, 204, null);
                    return;
                }

                if (!requestLine.StartsWith("POST /quick-capture/", StringComparison.OrdinalIgnoreCase))
                {
                    SendHttpResponse(stream, 404, new { ok = false, error = "Not found" });
                    return;
                }

                var contentLength = 0;
                var match = Regex.Match(parts[0], @"^Content-Length:\s*(\d+)", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                if (match.Success)
                {
                    contentLength = int.Parse(match.Groups[1].Value);
                }

                var body = parts.Length > 1 ? parts[1] : string.Empty;
                while (Encoding.UTF8.GetByteCount(body) < contentLength)
                {
                    var read = stream.Read(buffer, 0, buffer.Length);
                    if (read <= 0)
                    {
                        break;
                    }
                    body += Encoding.UTF8.GetString(buffer, 0, read);
                }

                string kind, text;
                try
                {
                    using var document = JsonDocument.Parse(body);
                    kind = document.RootElement.TryGetProperty("kind", out var kindElement)
                        ? kindElement.GetString() ?? string.Empty
                        : string.Empty;
                    text = NormalizeCaptureText(document.RootElement.TryGetProperty("text", out var textElement)
                        ? textElement.GetString() ?? string.Empty
                        : string.Empty);
                }
                catch (JsonException)
                {
                    throw new InvalidOperationException("Тело запроса не является JSON.");
                }

                if (kind != "task" && kind != "reminder")
                {
                    throw new InvalidOperationException("Неизвестный тип быстрого захвата.");
                }
                if (string.IsNullOrWhiteSpace(text))
                {
                    throw new InvalidOperationException("Выделенный текст пуст.");
                }

                ShowCaptureFromRequest(kind, text);
                SendHttpResponse(stream, 202, new { ok = true, accepted = true });
            }
            catch (Exception exception)
            {
                try
                {
                    if (stream is not null)
                    {
                        SendHttpResponse(stream, 400, new { ok = false, error = exception.Message });
                    }
                }
                catch
                {
                    // Соединение уже закрыто — отвечать некому.
                }
            }
        }
    }

    /// <summary>
    /// Открывает окно захвата в UI-потоке и ждёт его закрытия на потоке сокета,
    /// чтобы ответ 202 ушёл как в оригинале — после работы пользователя с формой.
    /// </summary>
    private static void ShowCaptureFromRequest(string kind, string text)
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher is null || dispatcher.HasShutdownStarted)
        {
            throw new InvalidOperationException("Приложение недоступно.");
        }

        var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        dispatcher.BeginInvoke(() =>
        {
            if (_captureInProgress)
            {
                completion.TrySetException(new InvalidOperationException("Форма быстрого захвата уже открыта."));
                return;
            }

            _captureInProgress = true;
            Log($"Запрос браузерного расширения получен: {kind}.");
            try
            {
                ShowCaptureWindow(kind, text);
                completion.TrySetResult(true);
            }
            catch (Exception exception)
            {
                completion.TrySetException(exception);
            }
            finally
            {
                _captureInProgress = false;
            }
        });
        completion.Task.Wait();
    }

    /// <summary>
    /// Порт Send-QuickCaptureResponse (строки 297-306): статусы и CORS-заголовки,
    /// чтобы браузерное расширение могло читать ответ.
    /// </summary>
    private static void SendHttpResponse(NetworkStream stream, int statusCode, object? body)
    {
        var bytes = body is null ? Array.Empty<byte>() : Encoding.UTF8.GetBytes(JsonSerializer.Serialize(body));
        var status = statusCode switch
        {
            202 => "202 Accepted",
            204 => "204 No Content",
            404 => "404 Not Found",
            _ => "400 Bad Request",
        };
        var header = $"HTTP/1.1 {status}\r\n" +
                     "Content-Type: application/json; charset=utf-8\r\n" +
                     "Access-Control-Allow-Origin: *\r\n" +
                     "Access-Control-Allow-Methods: POST, OPTIONS\r\n" +
                     "Access-Control-Allow-Headers: Content-Type\r\n" +
                     $"Content-Length: {bytes.Length}\r\n" +
                     "Connection: close\r\n\r\n";
        var headerBytes = Encoding.ASCII.GetBytes(header);
        stream.Write(headerBytes, 0, headerBytes.Length);
        if (bytes.Length > 0)
        {
            stream.Write(bytes, 0, bytes.Length);
        }
        stream.Flush();
    }

    /// <summary>Порт Normalize-QuickCaptureText (строки 245-250): убирает NUL и пробелы по краям.</summary>
    private static string NormalizeCaptureText(string value)
        => value.Replace("\0", string.Empty).Trim();

    private static bool TryRegisterPair(IntPtr handle, uint modifiers)
    {
        if (!RegisterHotKey(handle, TaskHotkeyId, modifiers, (uint)'T'))
        {
            return false;
        }

        if (!RegisterHotKey(handle, ReminderHotkeyId, modifiers, (uint)'R'))
        {
            UnregisterHotKey(handle, TaskHotkeyId);
            return false;
        }

        return true;
    }

    private static IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WmHotkey)
        {
            var id = wParam.ToInt32();
            if (id == TaskHotkeyId)
            {
                StartCapture("task");
                handled = true;
            }
            else if (id == ReminderHotkeyId)
            {
                StartCapture("reminder");
                handled = true;
            }
        }

        return IntPtr.Zero;
    }

    /// <summary>
    /// Аналог Invoke-QuickCaptureFromHotkey (строки 479-496): захват текста
    /// идёт в отдельном STA-потоке, чтобы не блокировать интерфейс ожиданием
    /// буфера обмена; окно открывается уже в UI-потоке.
    /// </summary>
    private static void StartCapture(string kind)
    {
        if (_captureInProgress)
        {
            return;
        }

        Log($"Хоткей получен: {kind}.");
        _captureInProgress = true;
        var thread = new Thread(() =>
        {
            var text = string.Empty;
            try
            {
                text = CaptureSelectedText();
            }
            catch
            {
                // Пустой текст — окно откроется с пустым заголовком, как в оригинале.
            }

            Log($"Захват завершён: {text.Length} символов.");

            Application.Current?.Dispatcher.BeginInvoke(() =>
            {
                try
                {
                    ShowCaptureWindow(kind, text);
                }
                finally
                {
                    _captureInProgress = false;
                }
            });
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
    }

    /// <summary>
    /// Порт Get-SelectedText (строки 448-467): снимок буфера, ожидание
    /// отпускания модификаторов, эмуляция Ctrl+C и ожидание нового текста.
    /// Буфер обмена возвращается владельцу в любом случае.
    /// </summary>
    private static string CaptureSelectedText()
    {
        IDataObject? snapshot = null;
        try
        {
            snapshot = Clipboard.GetDataObject();
        }
        catch
        {
            // Буфер занят другим приложением — снимаем без снимка.
        }

        try
        {
            var beforeSequence = GetClipboardSequenceNumber();
            for (var attempt = 0; attempt < 20; attempt++)
            {
                if (!IsKeyDown(VkControl) && !IsKeyDown(VkShift))
                {
                    break;
                }

                Thread.Sleep(25);
            }

            SendCtrlC();
            for (var attempt = 0; attempt < 15; attempt++)
            {
                Thread.Sleep(100);
                try
                {
                    var value = Clipboard.GetText();
                    if (GetClipboardSequenceNumber() != beforeSequence && !string.IsNullOrWhiteSpace(value))
                    {
                        return value;
                    }
                }
                catch
                {
                    // Буфер мог перехватить другой процесс — пробуем дальше.
                }
            }

            return string.Empty;
        }
        finally
        {
            RestoreClipboard(snapshot);
        }
    }

    /// <summary>Порт Restore-QuickCaptureClipboard (строки 469-477).</summary>
    private static void RestoreClipboard(IDataObject? snapshot)
    {
        if (snapshot is null)
        {
            return;
        }

        for (var attempt = 0; attempt < 8; attempt++)
        {
            try
            {
                Clipboard.SetDataObject(snapshot, true);
                return;
            }
            catch
            {
                Thread.Sleep(100);
            }
        }
    }

    private static void ShowCaptureWindow(string kind, string text)
    {
        var window = new QuickCaptureWindow(kind, text)
        {
            Owner = Application.Current.MainWindow,
            Topmost = true
        };

        var saved = window.ShowDialog();
        if (saved != true)
        {
            return;
        }

        ToastService.Show(kind == "reminder" ? "Напоминание создано." : "Задача создана.", ToastKind.Success);
        DataChanged?.Invoke();
    }

    private static bool IsKeyDown(int virtualKey)
        => (GetAsyncKeyState(virtualKey) & 0x8000) != 0;

    private static void SendCtrlC()
    {
        // Как в оригинальном агенте: нажатие и отпускание Ctrl и C.
        keybd_event(VkControl, 0, 0, UIntPtr.Zero);
        keybd_event(VkC, 0, 0, UIntPtr.Zero);
        keybd_event(VkC, 0, KeyUp, UIntPtr.Zero);
        keybd_event(VkControl, 0, KeyUp, UIntPtr.Zero);
    }

    private const byte VkControl = 0x11;
    private const byte VkShift = 0x10;
    private const byte VkC = 0x43;
    private const uint KeyUp = 0x0002;

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int virtualKey);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, UIntPtr extraInfo);

    [DllImport("user32.dll")]
    private static extern uint GetClipboardSequenceNumber();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint modifiers, uint key);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
}
