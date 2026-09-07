using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.RegularExpressions;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Интеграция с Bitrix24 — порт Invoke-BitrixMethod, New-BitrixMethodUri,
/// Normalize-BitrixBaseUrl (app/BitrixIntegration.ps1) и обработчика
/// $btnBitrixRefresh.Add_Click (RRFQComparer.ps1, строки 421-490).
/// </summary>
public static class BitrixService
{
    private const int TimeoutSeconds = 20;

    /// <summary>
    /// Заголовок пункта RFQ-чеклиста — аналог Get-BitrixRfqChecklistTitle
    /// (BitrixIntegration.ps1, строки 608-610).
    /// </summary>
    private const string RfqChecklistTitle =
        "Заполнение сводного RFQ и выделение компонентов на расчет (Сенько Е., Сидоров М.)";

    /// <summary>Префикс для поиска пункта: часть заголовка до '('.</summary>
    private static readonly string RfqChecklistPrefix =
        RfqChecklistTitle[..RfqChecklistTitle.IndexOf('(')].Trim();

    /// <summary>Статусы задач, попадающие в таблицу (аналог $allowedStatuses).</summary>
    private static readonly HashSet<int> AllowedStatuses = new() { 1, 2, 3, 6 };

    /// <summary>Аналог Normalize-BitrixBaseUrl.</summary>
    public static string NormalizeBaseUrl(string baseUrl)
    {
        var text = (baseUrl ?? string.Empty).Trim();
        if (text.Length == 0)
        {
            return string.Empty;
        }

        if (!Regex.IsMatch(text, "^[a-z][a-z0-9+\\.-]*://", RegexOptions.IgnoreCase))
        {
            text = "https://" + text;
        }

        if (!Uri.TryCreate(text, UriKind.Absolute, out var uri))
        {
            throw new InvalidOperationException("Bitrix BaseUrl must be a valid absolute URL.");
        }

        if (uri.Scheme != Uri.UriSchemeHttps)
        {
            var builder = new UriBuilder(uri) { Scheme = Uri.UriSchemeHttps, Port = -1 };
            uri = builder.Uri;
        }

        return uri.AbsoluteUri.TrimEnd('/');
    }

    /// <summary>Аналог New-BitrixMethodUri.</summary>
    public static string BuildMethodUri(BitrixConfig config, string method)
    {
        var baseUrl = NormalizeBaseUrl(config.BaseUrl);
        if (string.IsNullOrWhiteSpace(baseUrl))
        {
            throw new InvalidOperationException("Bitrix BaseUrl is required.");
        }
        if (config.WebhookUserId <= 0)
        {
            throw new InvalidOperationException("Bitrix WebhookUserId is required.");
        }
        if (string.IsNullOrWhiteSpace(config.WebhookToken))
        {
            throw new InvalidOperationException("Bitrix WebhookToken is required.");
        }

        var methodName = method.Trim().TrimStart('/');
        if (string.IsNullOrWhiteSpace(methodName))
        {
            throw new InvalidOperationException("Bitrix Method is required.");
        }

        return $"{baseUrl}/rest/{config.WebhookUserId}/{config.WebhookToken.Trim()}/{methodName}";
    }

    /// <summary>Аналог Get-BitrixTaskLink: ссылка на просмотр задачи в портала.</summary>
    public static string GetTaskViewUrl(BitrixConfig config, int taskId)
        => $"{config.BaseUrl}/company/personal/user/{config.WebhookUserId}/tasks/task/view/{taskId}/";

    /// <summary>
    /// Аналог Invoke-BitrixMethod: POST JSON-вызов метода портала. Возвращает
    /// узел result; при наличии error бросает исключение с текстом оригинала.
    /// </summary>
    public static async Task<JsonElement> InvokeMethodAsync(BitrixConfig config, string method, object parameters)
    {
        var uri = BuildMethodUri(config, method);
        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(TimeoutSeconds) };
        using var response = await client.PostAsJsonAsync(uri, parameters);
        response.EnsureSuccessStatusCode();
        using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        var root = document.RootElement;

        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("error", out var error)
            && error.ValueKind != JsonValueKind.Null)
        {
            var code = error.ToString();
            var description = root.TryGetProperty("error_description", out var desc) ? desc.GetString() : null;
            if (string.IsNullOrWhiteSpace(description))
            {
                description = "Bitrix returned an API error.";
            }
            throw new InvalidOperationException($"Bitrix API error {code}: {description}");
        }

        return root.TryGetProperty("result", out var result) ? result.Clone() : default;
    }

    /// <summary>
    /// Полный цикл обновления страницы — порт тела $btnBitrixRefresh.Add_Click
    /// (RRFQComparer.ps1, строки 431-483): список задач проекта, фильтрация по
    /// блокировкам и статусам, проверка RFQ-чеклиста каждой задачи.
    /// </summary>
    public static async Task<List<BitrixTaskRow>> GetTasksAsync(BitrixConfig config, IReadOnlySet<int> blockedTaskIds)
    {
        var taskParams = new Dictionary<string, object>
        {
            ["filter"] = new Dictionary<string, object>
            {
                ["GROUP_ID"] = config.ProjectId,
                [">=REAL_STATUS"] = 1,
                ["<=REAL_STATUS"] = 6,
            },
            ["select"] = new[] { "id", "title", "deadline", "status", "groupId" },
            ["order"] = new Dictionary<string, object> { ["ID"] = "DESC" },
        };

        var result = await InvokeMethodAsync(config, "tasks.task.list", taskParams);
        var rows = new List<BitrixTaskRow>();
        if (result.ValueKind != JsonValueKind.Object || !result.TryGetProperty("tasks", out var tasksElement)
            || tasksElement.ValueKind != JsonValueKind.Array)
        {
            return rows;
        }

        foreach (var task in tasksElement.EnumerateArray())
        {
            var id = ToInt(task, "ID", ToInt(task, "id", 0));
            if (id <= 0 || blockedTaskIds.Contains(id))
            {
                continue;
            }

            var status = ToInt(task, "STATUS", ToInt(task, "status", 0));
            if (!AllowedStatuses.Contains(status))
            {
                continue;
            }

            var deadline = ToText(task, "DEADLINE", ToText(task, "deadline"));
            if (deadline.Length > 10)
            {
                deadline = deadline[..10];
            }

            var (rfqStatusText, rfqCompleted) = await GetRfqStatusAsync(config, id);
            if (rfqCompleted)
            {
                continue;
            }

            rows.Add(new BitrixTaskRow
            {
                TaskId = id,
                Title = ToText(task, "TITLE", ToText(task, "title")),
                StatusText = StatusText(status),
                RfqStatus = rfqStatusText,
                Deadline = deadline,
            });
        }

        return rows;
    }

    /// <summary>
    /// Проверка пункта RFQ-чеклиста задачи — аналог блока с вызовом
    /// task.checklistitem.getlist (строки 455-471). Ошибка запроса даёт
    /// «Ошибка проверки», как в оригинале.
    /// </summary>
    private static async Task<(string StatusText, bool Completed)> GetRfqStatusAsync(BitrixConfig config, int taskId)
    {
        try
        {
            var result = await InvokeMethodAsync(
                config, "task.checklistitem.getlist", new Dictionary<string, object> { ["TASKID"] = taskId });
            if (result.ValueKind != JsonValueKind.Array)
            {
                return ("Пункт не найден", false);
            }

            foreach (var item in result.EnumerateArray())
            {
                var itemTitle = NormalizeChecklistTitle(
                    ToText(item, "TITLE", ToText(item, "title")));
                if (!itemTitle.StartsWith(RfqChecklistPrefix, StringComparison.Ordinal))
                {
                    continue;
                }

                var completed = ToText(item, "IS_COMPLETE", ToText(item, "isComplete"));
                if (completed is "Y" or "true" or "True")
                {
                    return ("Выполнен", true);
                }

                return ("Не выполнен", false);
            }

            return ("Пункт не найден", false);
        }
        catch
        {
            return ("Ошибка проверки", false);
        }
    }

    /// <summary>Аналог switch по статусу задачи (строки 472-478).</summary>
    private static string StatusText(int status) => status switch
    {
        1 => "Новая",
        2 => "Ожидает выполнения",
        3 => "В работе",
        6 => "Отложена",
        _ => status.ToString(),
    };

    /// <summary>Аналог Normalize-BitrixChecklistTitle.</summary>
    private static string NormalizeChecklistTitle(string title)
        => Regex.Replace(title ?? string.Empty, "\\s+", " ").Trim();

    private static string ToText(JsonElement element, string name, string fallback = "")
        => element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var value)
            && value.ValueKind is JsonValueKind.String or JsonValueKind.Number
                ? value.ToString()
                : fallback;

    private static int ToInt(JsonElement element, string name, int fallback)
        => element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var value)
            && value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var parsed)
                ? parsed
                : fallback;
}
