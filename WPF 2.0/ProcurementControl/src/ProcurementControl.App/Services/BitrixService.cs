using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.RegularExpressions;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Интеграция с Bitrix24 — порт Invoke-BitrixMethod, New-BitrixMethodUri,
/// Normalize-BitrixBaseUrl (app/BitrixIntegration.ps1) и, с поправкой на
/// лояльность к формату, Convert-BitrixTaskTitleToProcurement (там же,
/// строки 232-287). Проект «Заявка на расчет КС в РФ» — это группа задач,
/// каждая задача = будущая сделка «Контроля закупки».
/// </summary>
public static class BitrixService
{
    private const int TimeoutSeconds = 20;

    /// <summary>Сколько страниц по 50 задач выбираем за один обход (защита от зацикливания).</summary>
    private const int MaxPages = 40;

    /// <summary>Хвост названия с количеством: « - 5 шт.», «-3шт» и т.п.</summary>
    private static readonly Regex QuantityRegex =
        new("^(?<description>.*?)(?:\\s*-\\s*(?<quantity>\\d+)\\s*(?:шт\\.?)?\\s*)$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

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

    /// <summary>Ссылка на просмотр задачи в портале (аналог Get-BitrixTaskLink).</summary>
    public static string GetTaskViewUrl(BitrixConfig config, int taskId)
        => $"{NormalizeBaseUrl(config.BaseUrl)}/company/personal/user/{config.WebhookUserId}/tasks/task/view/{taskId}/";

    /// <summary>Ссылка на редактирование задачи: форма заполнена, можно поправить и сохранить.</summary>
    public static string GetTaskEditUrl(BitrixConfig config, int taskId)
        => $"{NormalizeBaseUrl(config.BaseUrl)}/company/personal/user/{config.WebhookUserId}/tasks/task/edit/{taskId}/";

    /// <summary>
    /// Форма создания задачи с применённым шаблоном (?TEMPLATE=id — проверено:
    /// новая форма предзаполняет название, проект, исполнителя, наблюдателей,
    /// чек-лист из шаблона).
    /// </summary>
    public static string GetTemplateFormUrl(BitrixConfig config, int templateId)
        => $"{NormalizeBaseUrl(config.BaseUrl)}/company/personal/user/{config.WebhookUserId}/tasks/task/edit/0/?TEMPLATE={templateId}";

    /// <summary>Обновляет поля шаблона задачи. Параметр строго templateId
    /// («id» портал не признаёт: Could not find value for parameter).</summary>
    public static async Task UpdateTemplateAsync(BitrixConfig config, int templateId, Dictionary<string, object> fields)
    {
        await InvokeRootAsync(config, "tasks.template.update", new Dictionary<string, object>
        {
            ["templateId"] = templateId,
            ["fields"] = fields,
        });
    }

    /// <summary>
    /// Аналог Invoke-BitrixMethod: POST JSON-вызов метода портала. Возвращает
    /// корень ответа: у списковых методов next/total лежат рядом с result,
    /// а не внутри него. При наличии error бросает исключение с текстом.
    /// </summary>
    public static async Task<JsonElement> InvokeRootAsync(BitrixConfig config, string method, object parameters)
    {
        var uri = BuildMethodUri(config, method);
        // UseProxy=false: портал доступен напрямую, а системный прокси
        // (Hiddify на 127.0.0.1:10809) включён не всегда и ронял все запросы.
        using var client = new HttpClient(new HttpClientHandler { UseProxy = false })
        {
            Timeout = TimeSpan.FromSeconds(TimeoutSeconds),
        };
        using var response = await client.PostAsJsonAsync(uri, parameters);
        var content = await response.Content.ReadAsStringAsync();

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(content);
        }
        catch (JsonException)
        {
            // Не-JSON тело бывает у ошибок уровня сервера — показываем хотя бы код.
            response.EnsureSuccessStatusCode();
            throw new InvalidOperationException("Bitrix вернул не-JSON ответ для " + method + ".");
        }

        using var _ = document;
        var root = document.RootElement;

        // Битрикс ошибки отдаёт и не-200 кодом, и JSON-телом — берём текст из тела.
        if (!response.IsSuccessStatusCode)
        {
            if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("error", out var httpError)
                && httpError.ValueKind != JsonValueKind.Null)
            {
                var httpDescription = root.TryGetProperty("error_description", out var httpDesc) ? httpDesc.GetString() : null;
                throw new InvalidOperationException($"Bitrix API error {httpError}: {httpDescription}");
            }

            response.EnsureSuccessStatusCode();
        }

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

        return root.Clone();
    }

    /// <summary>Вызов метода с разбором узла result (для несписковых методов).</summary>
    public static async Task<JsonElement> InvokeMethodAsync(BitrixConfig config, string method, object parameters)
    {
        var root = await InvokeRootAsync(config, method, parameters);
        return root.TryGetProperty("result", out var result) ? result.Clone() : default;
    }

    /// <summary>
    /// Ищет задачу по номеру заявки в названии («Номер/Клиент/…») по всему
    /// порталу (без привязки к проекту — задача может лежать в другой группе).
    /// Точное совпадение не нашлось — ищется «похожая»: тот же хвост после
    /// дефиса и совпадающий префикс номера (опечатки вида VNN/VNV). Возвращает
    /// самую свежую подходящую задачу; Scanned — сколько задач просмотрено.
    /// </summary>
    public static async Task<BitrixTaskSearchResult> FindTaskByRequestNumberAsync(BitrixConfig config, string requestNumber)
    {
        var needle = requestNumber.Trim();
        if (needle.Length == 0)
        {
            return new BitrixTaskSearchResult();
        }

        var scanned = 0;
        List<BitrixTaskInfo> filtered;
        try
        {
            var filteredRoot = await InvokeRootAsync(config, "tasks.task.list", new Dictionary<string, object>
            {
                ["filter"] = new Dictionary<string, object>
                {
                    ["TITLE"] = "%" + needle + "%",
                },
                ["select"] = new[] { "id", "title", "createdDate" },
                ["order"] = new Dictionary<string, object> { ["ID"] = "DESC" },
            });
            filtered = ReadTaskList(filteredRoot);
            scanned = filtered.Count;
        }
        catch (Exception ex)
        {
            // Портал мог не принять %-фильтр — идём полным обходом.
            ErrorLog.Append("bitrix", "поиск по TITLE задачи-заявки: " + ex.Message);
            filtered = new List<BitrixTaskInfo>();
        }

        var exact = PickTaskByRequestNumber(filtered, needle);
        if (exact is not null)
        {
            return new BitrixTaskSearchResult { Task = exact };
        }

        // Похожая задача: тот же хвост номера (после дефиса) — ловит опечатки
        // в буквах префикса (VNN/VNV), но не подставляет чужой номер молча.
        var tail = needle.Contains('-') ? needle[(needle.LastIndexOf('-') + 1)..] : string.Empty;
        List<BitrixTaskInfo> similar = new();
        if (tail.Length >= 4)
        {
            try
            {
                var similarRoot = await InvokeRootAsync(config, "tasks.task.list", new Dictionary<string, object>
                {
                    ["filter"] = new Dictionary<string, object>
                    {
                        ["TITLE"] = "%" + tail + "%",
                    },
                    ["select"] = new[] { "id", "title", "createdDate" },
                    ["order"] = new Dictionary<string, object> { ["ID"] = "DESC" },
                });
                var tailTasks = ReadTaskList(similarRoot);
                scanned = Math.Max(scanned, tailTasks.Count);
                similar = tailTasks
                    .Select(task => (task, parts: ParseTaskTitle(task.Title)))
                    .Where(item => item.parts is not null
                        && SameRequestTail(item.parts.RequestNumber, needle)
                        && CommonPrefixLength(item.parts.RequestNumber, needle) >= 10)
                    .Select(item => item.task)
                    .OrderByDescending(task => task.TaskId)
                    .ToList();
            }
            catch (Exception ex)
            {
                ErrorLog.Append("bitrix", "поиск по хвосту номера задачи-заявки: " + ex.Message);
            }
        }

        var similarMatch = similar.FirstOrDefault();
        if (similarMatch is not null)
        {
            return new BitrixTaskSearchResult { Task = similarMatch, IsSimilarMatch = true, Scanned = scanned };
        }

        // Последний шанс: полный обход доступных задач (LIKE мог не сработать).
        for (var page = 0; page < MaxPages; page++)
        {
            var parameters = new Dictionary<string, object>
            {
                ["select"] = new[] { "id", "title", "createdDate" },
                ["order"] = new Dictionary<string, object> { ["ID"] = "ASC" },
            };
            if (page > 0)
            {
                parameters["start"] = page * 50;
            }

            var root = await InvokeRootAsync(config, "tasks.task.list", parameters);
            var batch = ReadTaskList(root);
            scanned += batch.Count;
            var found = PickTaskByRequestNumber(batch, needle);
            if (found is not null)
            {
                return new BitrixTaskSearchResult { Task = found };
            }

            if (batch.Count == 0
                || !(root.ValueKind == JsonValueKind.Object && root.TryGetProperty("next", out var next)
                    && next.ValueKind == JsonValueKind.Number && next.TryGetInt32(out var nextStart) && nextStart > 0))
            {
                break;
            }
        }

        ErrorLog.Append("bitrix", "задача-заявка «" + needle + "» не найдена: доступно вебхуку " + scanned +
            " задач. Проверьте, что вебхук-пользователь видит задачу и номер в названии совпадает.");
        return new BitrixTaskSearchResult { Scanned = scanned };
    }

    /// <summary>Общий хвост номера после дефиса (без учёта регистра).</summary>
    private static bool SameRequestTail(string candidate, string request)
    {
        var candidateTail = candidate.Contains('-') ? candidate[(candidate.LastIndexOf('-') + 1)..] : string.Empty;
        var requestTail = request.Contains('-') ? request[(request.LastIndexOf('-') + 1)..] : string.Empty;
        return candidateTail.Length > 0
            && string.Equals(candidateTail, requestTail, StringComparison.OrdinalIgnoreCase);
    }

    private static int CommonPrefixLength(string first, string second)
    {
        var length = Math.Min(first.Length, second.Length);
        var common = 0;
        while (common < length && char.ToUpperInvariant(first[common]) == char.ToUpperInvariant(second[common]))
        {
            common++;
        }

        return common;
    }

    private static BitrixTaskInfo? PickTaskByRequestNumber(List<BitrixTaskInfo> tasks, string requestNumber)
        => tasks.Where(task => string.Equals(
                    ParseTaskTitle(task.Title)?.RequestNumber,
                    requestNumber,
                    StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(task => task.TaskId)
            .FirstOrDefault();

    /// <summary>Вложение задачи Bitrix (диск-объект, привязанный к задаче, чек-листу или комментарию).</summary>
    public sealed class BitrixTaskFile
    {
        public long AttachedObjectId { get; init; }

        public string Name { get; set; } = string.Empty;

        public long? Size { get; set; }

        /// <summary>Где найдено: «задача», «чек-лист» или «комментарий».</summary>
        public string Source { get; init; } = string.Empty;
    }

    /// <summary>Ключи, внутри которых ищутся id вложений диска.</summary>
    private static readonly Regex FileKeyRegex = new("file|attach|webdav", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Все вложения задачи: файлы задачи (UF_TASK_WEBDAV_FILES и подобные),
    /// чек-лист и комментарии. Нестандартные поля портала собираются рекурсивным
    /// обходом JSON — где именно лежат файлы, зависит от версии портала.
    /// После сбора у каждого вложения запрашивается имя/размер
    /// (disk.attachedObject.get): вызывающий отсеивает RRFQ-ответы по имени до
    /// скачивания. Вложения, на которые нет прав, выбрасываются.
    /// </summary>
    public static async Task<List<BitrixTaskFile>> GetTaskFilesAsync(BitrixConfig config, int taskId)
    {
        var files = new List<BitrixTaskFile>();
        var seen = new HashSet<long>();

        var taskResult = await InvokeMethodAsync(config, "tasks.task.get", new Dictionary<string, object>
        {
            ["taskId"] = taskId,
            ["select"] = new[] { "*", "UF_*" },
        });
        var element = taskResult.ValueKind == JsonValueKind.Object && taskResult.TryGetProperty("task", out var task)
            ? task
            : taskResult;
        CollectFileIds(element, "задача", files, seen);

        try
        {
            var checklist = await InvokeMethodAsync(config, "tasks.task.checklist.get", new Dictionary<string, object>
            {
                ["taskId"] = taskId,
            });
            CollectFileIds(checklist, "чек-лист", files, seen);
        }
        catch (Exception ex)
        {
            ErrorLog.Append("bitrix", "чек-лист задачи " + taskId + ": " + ex.Message);
        }

        try
        {
            var comments = await InvokeMethodAsync(config, "task.commentitem.getlist", new Dictionary<string, object>
            {
                ["TASKID"] = taskId,
            });
            CollectFileIds(comments, "комментарий", files, seen);
        }
        catch (Exception ex)
        {
            ErrorLog.Append("bitrix", "комментарии задачи " + taskId + ": " + ex.Message);
        }

        // Имена нужны до скачивания: исходный RFQ отличается от RRFQ-ответа
        // именно по имени файла.
        for (var index = files.Count - 1; index >= 0; index--)
        {
            try
            {
                var attached = await InvokeMethodAsync(config, "disk.attachedObject.get", new Dictionary<string, object>
                {
                    ["id"] = files[index].AttachedObjectId,
                });
                files[index].Name = ToText(attached, "NAME", ToText(attached, "name"));
                var size = ToInt(attached, "SIZE", ToInt(attached, "size", 0));
                if (size > 0)
                {
                    files[index].Size = size;
                }
            }
            catch (Exception ex)
            {
                // Не диск-вложение (или нет прав) — не предлагаем.
                ErrorLog.Append("bitrix", "вложение " + files[index].AttachedObjectId + ": " + ex.Message);
                files.RemoveAt(index);
            }
        }

        return files;
    }

    /// <summary>
    /// Рекурсивно собирает id вложений диска: числа и числовые строки внутри
    /// полей, чьё имя содержит file/attach/webdav (FILES, UF_TASK_WEBDAV_FILES,
    /// attachedObjectId и т.п.).
    /// </summary>
    private static void CollectFileIds(JsonElement element, string source, List<BitrixTaskFile> files, HashSet<long> seen, string? contextKey = null, int depth = 0)
    {
        if (depth > 8 || files.Count >= 50)
        {
            return;
        }

        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                // Файловые коллекции часто приходят объектами вида
                // {"2213": {"ID": ..., "NAME": ...}} — берём ID прямо из объекта.
                if (contextKey is not null && FileKeyRegex.IsMatch(contextKey))
                {
                    foreach (var idKeyName in new[] { "ID", "ATTACHED_OBJECT_ID", "ATTACHEDOBJECTID" })
                    {
                        if (element.TryGetProperty(idKeyName, out var idValue)
                            && (idValue.ValueKind == JsonValueKind.Number || idValue.ValueKind == JsonValueKind.String)
                            && long.TryParse(idValue.ToString(), out var objectId)
                            && objectId > 0
                            && seen.Add(objectId))
                        {
                            files.Add(new BitrixTaskFile { AttachedObjectId = objectId, Source = source });
                        }
                    }
                }

                foreach (var property in element.EnumerateObject())
                {
                    CollectFileIds(property.Value, source, files, seen, property.Name, depth + 1);
                }

                break;
            case JsonValueKind.Array:
                foreach (var item in element.EnumerateArray())
                {
                    CollectFileIds(item, source, files, seen, contextKey, depth + 1);
                }

                break;
            case JsonValueKind.Number or JsonValueKind.String:
                if (contextKey is null || !FileKeyRegex.IsMatch(contextKey))
                {
                    break;
                }

                if (long.TryParse(element.ToString(), out var id) && id > 0 && seen.Add(id))
                {
                    files.Add(new BitrixTaskFile { AttachedObjectId = id, Source = source });
                }

                break;
        }
    }

    /// <summary>
    /// Скачивает вложение в destinationPath (через disk.attachedObject.get →
    /// DOWNLOAD_URL; без URL — второй запрос disk.file.get по FILE_ID).
    /// Заполняет имя и размер файла из портала.
    /// </summary>
    public static async Task DownloadTaskFileAsync(BitrixConfig config, BitrixTaskFile file, string destinationPath)
    {
        var attached = await InvokeMethodAsync(config, "disk.attachedObject.get", new Dictionary<string, object>
        {
            ["id"] = file.AttachedObjectId,
        });

        if (file.Name.Length == 0)
        {
            file.Name = ToText(attached, "NAME", ToText(attached, "name"));
        }

        if (file.Size is null)
        {
            var size = ToInt(attached, "SIZE", ToInt(attached, "size", 0));
            if (size > 0)
            {
                file.Size = size;
            }
        }

        var url = ToText(attached, "DOWNLOAD_URL", ToText(attached, "downloadUrl"));
        if (string.IsNullOrWhiteSpace(url))
        {
            var fileId = ToInt(attached, "FILE_ID", ToInt(attached, "fileId", 0));
            if (fileId <= 0)
            {
                throw new InvalidOperationException("Bitrix не вернул ссылку на скачивание для вложения " + file.AttachedObjectId + ".");
            }

            var diskFile = await InvokeMethodAsync(config, "disk.file.get", new Dictionary<string, object>
            {
                ["id"] = fileId,
            });
            url = ToText(diskFile, "DOWNLOAD_URL", ToText(diskFile, "downloadUrl"));
        }

        if (string.IsNullOrWhiteSpace(url))
        {
            throw new InvalidOperationException("Bitrix не вернул ссылку на скачивание для вложения " + file.AttachedObjectId + ".");
        }

        // UseProxy=false: портал доступен напрямую, а системный прокси
        // (Hiddify на 127.0.0.1:10809) включён не всегда и ронял все запросы.
        using var client = new HttpClient(new HttpClientHandler { UseProxy = false })
        {
            Timeout = TimeSpan.FromMinutes(2),
        };
        using var response = await client.GetAsync(url);
        response.EnsureSuccessStatusCode();
        await File.WriteAllBytesAsync(destinationPath, await response.Content.ReadAsByteArrayAsync());
    }

    /// <summary>
    /// Разбор названия задачи «Номер/Клиент/Описание - N шт.». В отличие от
    /// оригинала не бросает исключений: null означает, что номер заявки
    /// выделить нельзя (нет '/' до него или он пуст).
    /// </summary>
    public static BitrixTaskTitleParts? ParseTaskTitle(string? title)
    {
        var original = title?.Trim() ?? string.Empty;
        if (original.Length == 0)
        {
            return null;
        }

        var firstSlash = original.IndexOf('/');
        if (firstSlash < 1)
        {
            return null;
        }

        var requestNumber = original[..firstSlash].Trim();
        if (requestNumber.Length == 0)
        {
            return null;
        }

        var lastSlash = original.LastIndexOf('/');
        string client;
        string itemText;
        if (lastSlash > firstSlash)
        {
            client = original.Substring(firstSlash + 1, lastSlash - firstSlash - 1).Trim();
            itemText = original[(lastSlash + 1)..].Trim();
        }
        else
        {
            // Клиента в названии нет — всё после номера считаем описанием.
            client = string.Empty;
            itemText = original[(firstSlash + 1)..].Trim();
        }

        var quantityMatch = QuantityRegex.Match(itemText);
        int? quantity = null;
        if (quantityMatch.Success)
        {
            itemText = quantityMatch.Groups["description"].Value.Trim();
            quantity = int.Parse(quantityMatch.Groups["quantity"].Value, System.Globalization.CultureInfo.InvariantCulture);
        }

        return new BitrixTaskTitleParts
        {
            RequestNumber = requestNumber,
            Client = client,
            ItemDescription = itemText,
            Quantity = quantity,
            OriginalTitle = original,
        };
    }

    /// <summary>
    /// Задачи проекта, созданные строго позже момента sinceRaw. Момент
    /// передаётся в исходном формате портала (например
    /// «2026-09-25T18:11:54+07:00»): Bitrix сравнивает в своём часовом поясе,
    /// конвертация на нашей стороне только портит окно выборки. Выборка
    /// постраничная, от старых к новым. Статусы задач сознательно не
    /// проверяются — критерий новизны только дата создания.
    /// </summary>
    public static async Task<List<BitrixTaskInfo>> GetTasksCreatedAfterAsync(BitrixConfig config, string sinceRaw)
    {
        var filter = new Dictionary<string, object>
        {
            ["GROUP_ID"] = config.ProjectId,
            [">CREATED_DATE"] = sinceRaw,
        };

        var tasks = new List<BitrixTaskInfo>();
        var start = 0;
        for (var page = 0; page < MaxPages; page++)
        {
            var parameters = new Dictionary<string, object>
            {
                ["filter"] = filter,
                ["select"] = new[] { "id", "title", "createdDate" },
                ["order"] = new Dictionary<string, object> { ["ID"] = "ASC" },
            };
            if (start > 0)
            {
                parameters["start"] = start;
            }

            var root = await InvokeRootAsync(config, "tasks.task.list", parameters);
            var batch = ReadTaskList(root);
            tasks.AddRange(batch);

            // next/total лежат в корне ответа, а не внутри result.
            if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("next", out var next))
            {
                var nextStart = ToInt(root, "next", -1);
                if (nextStart > 0)
                {
                    start = nextStart;
                }
                else
                {
                    break;
                }
            }
            else
            {
                break;
            }
        }

        return tasks;
    }

    /// <summary>
    /// Момент создания самой свежей задачи проекта в исходном формате портала —
    /// точка отсчёта первого запуска: всё, что уже есть в проекте, считается
    /// старым.
    /// </summary>
    public static async Task<string?> GetLatestCreatedDateRawAsync(BitrixConfig config)
    {
        var root = await InvokeRootAsync(config, "tasks.task.list", new Dictionary<string, object>
        {
            ["filter"] = new Dictionary<string, object> { ["GROUP_ID"] = config.ProjectId },
            ["select"] = new[] { "id", "createdDate" },
            ["order"] = new Dictionary<string, object> { ["ID"] = "DESC" },
        });

        var tasks = ReadTaskList(root);
        return tasks.Count == 0 ? null : tasks[0].CreatedDate;
    }

    /// <summary>Задачи по списку ID (повторная проверка нераспознанных названий).</summary>
    public static async Task<List<BitrixTaskInfo>> GetTasksByIdsAsync(BitrixConfig config, IEnumerable<int> taskIds)
    {
        var tasks = new List<BitrixTaskInfo>();
        foreach (var taskId in taskIds)
        {
            if (taskId <= 0)
            {
                continue;
            }

            var result = await InvokeMethodAsync(config, "tasks.task.get", new Dictionary<string, object>
            {
                ["taskId"] = taskId,
                ["select"] = new[] { "id", "title", "createdDate" },
            });

            var element = result.ValueKind == JsonValueKind.Object && result.TryGetProperty("task", out var task)
                ? task
                : result;
            if (element.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var id = ToInt(element, "ID", ToInt(element, "id", 0));
            if (id == 0)
            {
                continue;
            }

            tasks.Add(new BitrixTaskInfo
            {
                TaskId = id,
                Title = ToText(element, "TITLE", ToText(element, "title")),
                CreatedDate = ToText(element, "CREATED_DATE", ToText(element, "createdDate")),
            });
        }

        return tasks;
    }

    /// <summary>
    /// Разбор ответа tasks.task.list. Принимает корень ответа: задачи лежат в
    /// result.tasks, при этом сам result бывает и массивом (зависит от портала).
    /// </summary>
    private static List<BitrixTaskInfo> ReadTaskList(JsonElement root)
    {
        var tasks = new List<BitrixTaskInfo>();
        if (root.ValueKind != JsonValueKind.Object || !root.TryGetProperty("result", out var resultNode))
        {
            return tasks;
        }

        JsonElement tasksElement;
        if (resultNode.ValueKind == JsonValueKind.Array)
        {
            tasksElement = resultNode;
        }
        else if (resultNode.ValueKind == JsonValueKind.Object
            && resultNode.TryGetProperty("tasks", out var inner)
            && inner.ValueKind == JsonValueKind.Array)
        {
            tasksElement = inner;
        }
        else
        {
            return tasks;
        }

        foreach (var task in tasksElement.EnumerateArray())
        {
            var id = ToInt(task, "ID", ToInt(task, "id", 0));
            if (id <= 0)
            {
                continue;
            }

            tasks.Add(new BitrixTaskInfo
            {
                TaskId = id,
                Title = ToText(task, "TITLE", ToText(task, "title")),
                CreatedDate = ToText(task, "CREATED_DATE", ToText(task, "createdDate")),
            });
        }

        return tasks;
    }

    // ------------------------------------------------------------------
    // Заказ переработчику: создание задачи в группе «Внесение документов
    // в ЕРП» с чек-листом и вложениями сделки.
    // ------------------------------------------------------------------

    /// <summary>Создаёт задачу (tasks.task.add), возвращает её id.</summary>
    public static async Task<int> AddTaskAsync(BitrixConfig config, Dictionary<string, object> fields)
    {
        var result = await InvokeMethodAsync(config, "tasks.task.add", new Dictionary<string, object>
        {
            ["fields"] = fields,
        });

        var element = result.ValueKind == JsonValueKind.Object && result.TryGetProperty("task", out var task)
            ? task
            : result;
        var id = ToInt(element, "ID", ToInt(element, "id", 0));
        if (id <= 0)
        {
            throw new InvalidOperationException("Bitrix не вернул номер созданной задачи.");
        }

        return id;
    }

    /// <summary>Название задачи по id (tasks.task.get), пустая строка — не найдена.</summary>
    public static async Task<string> GetTaskTitleAsync(BitrixConfig config, int taskId)
    {
        var result = await InvokeMethodAsync(config, "tasks.task.get", new Dictionary<string, object>
        {
            ["taskId"] = taskId,
            ["select"] = new[] { "id", "title" },
        });

        var element = result.ValueKind == JsonValueKind.Object && result.TryGetProperty("task", out var task)
            ? task
            : result;
        return ToText(element, "TITLE", ToText(element, "title"));
    }

    /// <summary>
    /// Задачи группы (постранично). LIKE-фильтр по TITLE на портале ненадёжен
    /// (поиск «%TRPR-…%» не находит существующую задачу), поэтому для проверки
    /// дублей группу выбираем целиком и сопоставляем названия на нашей стороне.
    /// </summary>
    public static async Task<List<BitrixTaskInfo>> GetGroupTasksAsync(BitrixConfig config, int groupId)
    {
        var tasks = new List<BitrixTaskInfo>();
        var start = 0;
        for (var page = 0; page < MaxPages; page++)
        {
            var parameters = new Dictionary<string, object>
            {
                ["filter"] = new Dictionary<string, object> { ["GROUP_ID"] = groupId },
                ["select"] = new[] { "id", "title", "createdDate" },
                ["order"] = new Dictionary<string, object> { ["ID"] = "DESC" },
            };
            if (start > 0)
            {
                parameters["start"] = start;
            }

            var root = await InvokeRootAsync(config, "tasks.task.list", parameters);
            var batch = ReadTaskList(root);
            tasks.AddRange(batch);

            if (!(root.ValueKind == JsonValueKind.Object && root.TryGetProperty("next", out var next)
                && next.ValueKind == JsonValueKind.Number && next.TryGetInt32(out var nextStart) && nextStart > 0))
            {
                break;
            }

            start = nextStart;
        }

        return tasks;
    }

    /// <summary>Пункт чек-листа задачи. Старый метод task.checklistitem.add —
    /// поля строго массивом в FIELDS (TITLE отдельным параметром даёт
    /// WRONG_ARGUMENTS).</summary>
    public static async Task AddChecklistItemAsync(BitrixConfig config, int taskId, string title)
    {
        await InvokeRootAsync(config, "task.checklistitem.add", new Dictionary<string, object>
        {
            ["TASKID"] = taskId.ToString(),
            ["FIELDS"] = new Dictionary<string, object> { ["TITLE"] = title },
        });
    }

    /// <summary>
    /// Корневая папка диска группы (disk.storage.getlist → ROOT_OBJECT_ID);
    /// null — хранилище группы недоступно вебхуку.
    /// </summary>
    public static async Task<string?> FindGroupStorageRootAsync(BitrixConfig config, int groupId)
    {
        var result = await InvokeMethodAsync(config, "disk.storage.getlist", new Dictionary<string, object>());
        if (result.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        foreach (var storage in result.EnumerateArray())
        {
            if (ToText(storage, "ENTITY_TYPE") != "group" || ToInt(storage, "ENTITY_ID", 0) != groupId)
            {
                continue;
            }

            var root = ToText(storage, "ROOT_OBJECT_ID", ToText(storage, "rootObjectId"));
            if (root.Length > 0)
            {
                return root;
            }
        }

        return null;
    }

    /// <summary>
    /// Папка name внутри parentFolderId: находит существующую или создаёт
    /// (disk.folder.addsubfolder). Возвращает id папки.
    /// </summary>
    public static async Task<string> EnsureFolderAsync(BitrixConfig config, string parentFolderId, string name)
    {
        var children = await InvokeMethodAsync(config, "disk.folder.getchildren", new Dictionary<string, object>
        {
            ["id"] = parentFolderId,
        });

        if (children.ValueKind == JsonValueKind.Array)
        {
            foreach (var child in children.EnumerateArray())
            {
                if (string.Equals(ToText(child, "NAME", ToText(child, "name")), name, StringComparison.OrdinalIgnoreCase)
                    && ToText(child, "TYPE", ToText(child, "type")) == "folder")
                {
                    return ToText(child, "ID", ToText(child, "id"));
                }
            }
        }

        var created = await InvokeMethodAsync(config, "disk.folder.addsubfolder", new Dictionary<string, object>
        {
            // Ровно «id»: параметр parentId портал отвергает с ERROR_ARGUMENT.
            ["id"] = parentFolderId,
            ["data"] = new Dictionary<string, object> { ["NAME"] = name },
        });
        var id = ToText(created, "ID", ToText(created, "id"));
        if (id.Length == 0)
        {
            throw new InvalidOperationException("Bitrix не вернул номер созданной папки «" + name + "».");
        }

        return id;
    }

    /// <summary>
    /// Загружает локальный файл в папку диска. Проверенный контракт портала:
    /// disk.folder.uploadfile с fileContent (base64) — одиночный вызов (вариант
    /// с multipart портал переводит в двухэтапный handshake с uploadUrl).
    /// Файл с тем же именем: тот же размер — переиспользуем id (повторный запуск
    /// после сбоя); другой размер — disk.file.uploadversion (новая версия).
    /// Возвращает id файла на диске — он же идёт в UF_TASK_WEBDAV_FILES.
    /// </summary>
    public static async Task<long> UploadFileAsync(BitrixConfig config, string folderId, string localPath, string fileName)
    {
        var children = await InvokeMethodAsync(config, "disk.folder.getchildren", new Dictionary<string, object>
        {
            ["id"] = folderId,
        });

        long existingId = 0;
        long existingSize = -1;
        if (children.ValueKind == JsonValueKind.Array)
        {
            foreach (var child in children.EnumerateArray())
            {
                if (!string.Equals(ToText(child, "NAME", ToText(child, "name")), fileName, StringComparison.OrdinalIgnoreCase)
                    || ToText(child, "TYPE", ToText(child, "type")) != "file")
                {
                    continue;
                }

                existingId = ToLong(child, "ID", ToLong(child, "id", 0));
                existingSize = ToLong(child, "SIZE", ToLong(child, "size", -1));
                break;
            }
        }

        var fileContent = Convert.ToBase64String(await File.ReadAllBytesAsync(localPath));

        if (existingId > 0)
        {
            if (existingSize == new FileInfo(localPath).Length)
            {
                return existingId;
            }

            var updated = await InvokeMethodAsync(config, "disk.file.uploadversion", new Dictionary<string, object>
            {
                ["id"] = existingId,
                ["data"] = new Dictionary<string, object> { ["NAME"] = fileName },
                ["fileContent"] = fileContent,
            });
            var updatedId = ToLong(updated, "ID", ToLong(updated, "id", 0));
            return updatedId > 0 ? updatedId : existingId;
        }

        var created = await InvokeMethodAsync(config, "disk.folder.uploadfile", new Dictionary<string, object>
        {
            ["id"] = folderId,
            ["data"] = new Dictionary<string, object> { ["NAME"] = fileName },
            ["fileContent"] = fileContent,
        });

        var id = ToLong(created, "ID", ToLong(created, "id", 0));
        if (id <= 0)
        {
            throw new InvalidOperationException("Bitrix не вернул номер файла после загрузки " + fileName + ".");
        }

        return id;
    }

    /// <summary>
    /// Привязывает файлы диска к задаче (UF_TASK_WEBDAV_FILES). Проверено на
    /// портале: принимается только строковый формат «n{id}» (id файла диска,
    /// портал сам создаёт attached object); голое число даёт ERROR_CORE
    /// «Не удалось найти файл». Поэтому «n{id}» — первым, число — запасным.
    /// </summary>
    public static async Task AttachFilesToTaskAsync(BitrixConfig config, int taskId, IReadOnlyList<long> diskFileIds)
    {
        if (diskFileIds.Count == 0)
        {
            return;
        }

        foreach (var attempt in new Func<long, object>[] { id => "n" + id, id => id })
        {
            var payload = diskFileIds.Select(attempt).ToList();
            try
            {
                await InvokeRootAsync(config, "tasks.task.update", new Dictionary<string, object>
                {
                    ["taskId"] = taskId,
                    ["fields"] = new Dictionary<string, object> { ["UF_TASK_WEBDAV_FILES"] = payload },
                });
            }
            catch (Exception ex)
            {
                ErrorLog.Append("bitrix", "привязка файлов к задаче " + taskId + " (формат " +
                    (payload[0] is string ? "nXXX" : "число") + "): " + ex.Message);
                continue;
            }

            var attached = await GetAttachedFileCountAsync(config, taskId);
            if (attached >= diskFileIds.Count)
            {
                return;
            }
        }

        throw new InvalidOperationException(
            "Задача " + taskId + " создана, но файлы не прикрепились (ids на диске: " +
            string.Join(", ", diskFileIds) + "). Прикрепите их вручную.");
    }

    /// <summary>Сколько файлов сейчас висит на задаче (UF_TASK_WEBDAV_FILES).</summary>
    private static async Task<int> GetAttachedFileCountAsync(BitrixConfig config, int taskId)
    {
        try
        {
            var result = await InvokeMethodAsync(config, "tasks.task.get", new Dictionary<string, object>
            {
                ["taskId"] = taskId,
                // «*» + «UF_*»: выборка только по имени UF-поля на портале терялась.
                ["select"] = new[] { "*", "UF_*" },
            });

            var element = result.ValueKind == JsonValueKind.Object && result.TryGetProperty("task", out var task)
                ? task
                : result;
            if (element.ValueKind == JsonValueKind.Object
                && element.TryGetProperty("ufTaskWebdavFiles", out var files)
                && files.ValueKind == JsonValueKind.Array)
            {
                return files.GetArrayLength();
            }
        }
        catch (Exception ex)
        {
            ErrorLog.Append("bitrix", "проверка вложений задачи " + taskId + ": " + ex.Message);
        }

        return 0;
    }

    /// <summary>CREATED_DATE портала («yyyy-MM-dd HH:mm:ss») в DateTime.</summary>
    public static DateTime? ParsePortalDate(string? value)
        => DateTime.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None, out var parsed)
            ? parsed
            : null;

    private static string ToText(JsonElement element, string name, string fallback = "")
        => element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var value)
            && value.ValueKind is JsonValueKind.String or JsonValueKind.Number
                ? value.ToString()
                : fallback;

    /// <summary>
    /// Bitrix возвращает числовые поля задач строками («id»: «5643»), поэтому
    /// принимаем и JSON-число, и числовую строку.
    /// </summary>
    private static int ToInt(JsonElement element, string name, int fallback)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var value))
        {
            return fallback;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var parsed))
        {
            return parsed;
        }

        return value.ValueKind == JsonValueKind.String && int.TryParse(value.GetString(), out var fromText)
            ? fromText
            : fallback;
    }

    /// <summary>Как ToInt, но для id диска (могут не влезать в int).</summary>
    private static long ToLong(JsonElement element, string name, long fallback)
    {
        if (element.ValueKind != JsonValueKind.Object || !element.TryGetProperty(name, out var value))
        {
            return fallback;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out var parsed))
        {
            return parsed;
        }

        return value.ValueKind == JsonValueKind.String && long.TryParse(value.GetString(), out var fromText)
            ? fromText
            : fallback;
    }
}
