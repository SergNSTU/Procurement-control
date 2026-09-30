using System.IO;

namespace ProcurementControl.Services;

/// <summary>
/// Общий журнал ошибок: {DataRoot}/logs/errors.log. Тот же формат, который
/// App раньше писал приватным методом; теперь тот же канал используют
/// точечные catch, которые раньше молча проглатывали сбой.
/// </summary>
public static class ErrorLog
{
    public static void Append(string kind, Exception? exception)
        => Append(kind, exception?.ToString() ?? "?");

    public static void Append(string kind, string message)
    {
        try
        {
            var directory = Path.Combine(AppPaths.DataRoot, "logs");
            Directory.CreateDirectory(directory);
            var line = string.Format(
                "{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}\r\n\r\n",
                DateTime.Now, kind, message);
            File.AppendAllText(Path.Combine(directory, "errors.log"), line);
        }
        catch
        {
            // Лог не должен становиться причиной второго сбоя.
        }
    }
}
