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

/// <summary>Задача проекта «Заявка на расчет КС в РФ», прочитанная из API.</summary>
public sealed class BitrixTaskInfo
{
    public int TaskId { get; init; }
    public string Title { get; init; } = string.Empty;

    /// <summary>CREATED_DATE как вернул портал (текст в часовом поясе портала).</summary>
    public string CreatedDate { get; init; } = string.Empty;
}

/// <summary>
/// Результат поиска задачи по номеру сделки: точный или похожий (в номерах
/// бывают опечатки — например VNN/VNV); Scanned — сколько задач просмотрено.
/// </summary>
public sealed class BitrixTaskSearchResult
{
    public BitrixTaskInfo? Task { get; init; }

    /// <summary>Номер задачи отличается от номера сделки — требует подтверждения.</summary>
    public bool IsSimilarMatch { get; init; }

    public int Scanned { get; init; }
}

/// <summary>
/// Разобранное название задачи формата «Номер/Клиент/Описание - N шт.» —
/// порт Convert-BitrixTaskTitleToProcurement (BitrixIntegration.ps1,
/// строки 232-287) в облегчённом варианте: клиент и количество не обязательны.
/// </summary>
public sealed class BitrixTaskTitleParts
{
    /// <summary>Номер заявки — текст до первого '/' (будущий номер сделки).</summary>
    public string RequestNumber { get; init; } = string.Empty;

    /// <summary>Клиент — текст между первым и последним '/' (может быть пустым).</summary>
    public string Client { get; init; } = string.Empty;

    /// <summary>Описание изделия после последнего '/' без хвоста «- N шт.».</summary>
    public string ItemDescription { get; init; } = string.Empty;

    /// <summary>Количество плат из хвоста «- N шт.»; null, если не указано.</summary>
    public int? Quantity { get; init; }

    public string OriginalTitle { get; init; } = string.Empty;
}

/// <summary>
/// Параметры задачи «Заказ переработчику». Задача создаёт сам пользователь
/// в Битриксе через предзаполненную форму (шаблон id 67 «Заказ
/// переработчику…»: группа «Внесение документов в ЕРП», исполнитель
/// Дарья Пономаренко (249), наблюдатель Карина Рахметуллина (25), чек-лист
/// «Заказы поставщикам»); приложение лишь обновляет шаблон под текущий
/// заказ и открывает форму. Ключи settings: bitrix.recycler_*.
/// </summary>
public sealed record RecyclerOrderOptions(
    int GroupId,
    int ResponsibleId,
    IReadOnlyList<int> AuditorIds,
    int DeadlineDays,
    IReadOnlyList<string> ChecklistItems,
    string DiskFolderName,
    int TemplateId);

/// <summary>
/// Строка таблицы импортированных сделок на странице «Bitrix API Test»:
/// связь bitrix_task_links + данные сделки из deals.
/// </summary>
public sealed class BitrixImportRow
{
    public long DealId { get; init; }
    public int TaskId { get; init; }
    public string DealNumber { get; init; } = string.Empty;
    public string Client { get; init; } = string.Empty;
    public string BoardCount { get; init; } = string.Empty;
    public string Status { get; init; } = string.Empty;
    public string CreatedAt { get; init; } = string.Empty;

    public string CreatedAtText => PurchaseFormatting.DateOrRaw(CreatedAt);

    /// <summary>Текст колонки «Ссылка».</summary>
    public string LinkText => "Открыть";
}

/// <summary>Итог одной проверки проекта Bitrix.</summary>
public sealed class BitrixImportReport
{
    /// <summary>Текст ошибки (нет вебхука, сеть, API).</summary>
    public string? Error { get; set; }

    /// <summary>Первый запуск: точка отсчёта установлена, старые задачи не импортируются.</summary>
    public bool Initialized { get; set; }

    /// <summary>Сколько новых задач увидел портал.</summary>
    public int Fetched { get; set; }

    /// <summary>Созданные сделки: строки «Номер — клиент».</summary>
    public List<string> Imported { get; } = new();

    /// <summary>Сколько задач уже были связаны или существовали по номеру.</summary>
    public int AlreadyPresent { get; set; }

    /// <summary>Задачи с нераспознанным названием: строки «ID 123 — название».</summary>
    public List<string> Skipped { get; } = new();

    /// <summary>Накопленный список нераспознанных задач (ждут исправления названия).</summary>
    public int PendingSkippedCount { get; set; }
}
