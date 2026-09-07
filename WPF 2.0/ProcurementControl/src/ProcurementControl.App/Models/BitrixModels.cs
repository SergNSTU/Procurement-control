namespace ProcurementControl.Models;

/// <summary>
/// Настройки интеграции с Bitrix — порт нужной части Get-BitrixConfig
/// (app/BitrixIntegration.ps1, строки 87-113). Значения читаются из таблицы
/// settings; дефолты соответствуют оригиналу.
/// </summary>
public sealed record BitrixConfig(string BaseUrl, int WebhookUserId, string WebhookToken, int ProjectId)
{
    public bool IsConfigured => !string.IsNullOrWhiteSpace(WebhookToken) && WebhookUserId > 0;
}

/// <summary>
/// Строка таблицы задач на странице «Bitrix API Test» — аналог строк
/// $bitrixGrid (RRFQComparer.ps1, строки 361-374).
/// </summary>
public sealed class BitrixTaskRow
{
    public int TaskId { get; init; }
    public string Title { get; init; } = string.Empty;
    public string StatusText { get; init; } = string.Empty;
    public string RfqStatus { get; init; } = string.Empty;
    public string Deadline { get; init; } = string.Empty;

    /// <summary>Текст колонки «Ссылка» — в оригинале всегда «Открыть».</summary>
    public string LinkText => "Открыть";
}
