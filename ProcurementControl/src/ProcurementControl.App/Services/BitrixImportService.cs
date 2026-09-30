using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Автоматическое подтягивание новых сделок из проекта Bitrix «Заявка на
/// расчет КС в РФ» (группа задач bitrix_project_id) в «Контроль закупки».
///
/// Новизна считается по дате создания задачи, а не по статусам: при первом
/// запуске новой версии приложения запоминается CREATED_DATE самой свежей
/// задачи проекта («точка отсчёта») — всё, что было в проекте раньше,
/// считается старым. Дальше импортируются только задачи, созданные позже
/// точки отсчёта; после каждого успешного обхода точка сдвигается.
///
/// Название задачи разбирается по формату «Номер/Клиент/Описание - N шт.»:
/// номер → «Сделка», клиент → «Клиент», N → «Кол-во плат», этап — RFQ.
/// Задачи с нераспознанным названием не пропадают: их ID ждут в настройках,
/// и при каждом обходе название проверяется заново (его могли исправить).
/// Связь сделка↔задача хранится в bitrix_task_links и защищает от дублей.
/// </summary>
public static class BitrixImportService
{
    private const string EnabledKey = "bitrix.import.enabled";
    private const string AnchorKey = "bitrix.import.anchor";
    private const string LastCheckKey = "bitrix.import.last_check";
    private const string SkippedKey = "bitrix.import.skipped";

    /// <summary>Раз в сколько минут автопроверка (ручная кнопка — без ожидания).</summary>
    private static readonly TimeSpan AutoCheckInterval = TimeSpan.FromMinutes(10);

    /// <summary>Лимит накопленных нераспознанных задач.</summary>
    private const int MaxSkipped = 300;

    private static readonly SemaphoreSlim Gate = new(1, 1);

    /// <summary>Чекбокс автоимпорта; хранится в настройках, включён по умолчанию.</summary>
    public static bool IsEnabled
    {
        get
        {
            try
            {
                using var snapshot = PurchaseSnapshot.Create();
                return new SettingsRepository(snapshot.Connection).GetBooleanSetting(EnabledKey, true);
            }
            catch
            {
                return true;
            }
        }
        set => PurchaseWriteRepository.SetSetting(EnabledKey, value ? "1" : "0");
    }

    public static DateTime? LastCheck => PurchaseFormatting.ParseDate(ReadSetting(LastCheckKey));

    /// <summary>Пора ли автопроверке бежать (ручная — всегда).</summary>
    public static bool DueForAutoCheck
        => IsEnabled && (LastCheck is null || DateTime.Now - LastCheck.Value >= AutoCheckInterval);

    /// <summary>
    /// Один обход проекта: новые задачи → сделки «Контроля закупки».
    /// Вызывается автопроверкой раз в 10 минут и кнопкой «Проверить сейчас».
    /// Прогресс виден в статусной строке главного окна (ProgressService).
    /// </summary>
    public static async Task<BitrixImportReport> SyncAsync(bool manual = false)
    {
        var report = new BitrixImportReport();
        await Gate.WaitAsync();
        // Прогресс после захвата семафора: полоса отражает реальную работу,
        // а не ожидание чужого обхода.
        ProgressService.Begin("Синхронизация Bitrix: обход проекта «Заявка на расчет КС в РФ»…");
        try
        {
            await SyncCore(report);
        }
        catch (Exception ex)
        {
            report.Error = ex.Message;
        }
        finally
        {
            ProgressService.End();
            Gate.Release();
        }
        return report;
    }

    /// <summary>Строки для таблицы вкладки: импортированные сделки с датами.</summary>
    public static List<BitrixImportRow> GetImportedRows(int limit = 200)
    {
        using var snapshot = PurchaseSnapshot.Create();
        using var command = snapshot.Connection.CreateCommand();
        command.CommandText = @"
SELECT d.id, d.deal_number, IFNULL(d.client, ''), IFNULL(d.board_count, ''),
       IFNULL(d.status, ''), IFNULL(d.created_at, ''), l.bitrix_task_id
FROM bitrix_task_links l
JOIN deals d ON d.id = l.deal_id
ORDER BY d.created_at DESC, d.id DESC
LIMIT @limit;";
        PurchaseWriteRepository.Add(command, "@limit", limit);
        using var reader = command.ExecuteReader();
        var rows = new List<BitrixImportRow>();
        while (reader.Read())
        {
            rows.Add(new BitrixImportRow
            {
                DealId = Convert.ToInt64(reader.GetValue(0)),
                DealNumber = Convert.ToString(reader.GetValue(1)) ?? string.Empty,
                Client = Convert.ToString(reader.GetValue(2)) ?? string.Empty,
                BoardCount = Convert.ToString(reader.GetValue(3)) ?? string.Empty,
                Status = Convert.ToString(reader.GetValue(4)) ?? string.Empty,
                CreatedAt = Convert.ToString(reader.GetValue(5)) ?? string.Empty,
                TaskId = Convert.ToInt32(reader.GetValue(6)),
            });
        }
        return rows;
    }

    private static async Task SyncCore(BitrixImportReport report)
    {
        var config = BitrixRepository.GetConfig();
        if (!config.IsConfigured)
        {
            report.Error = "Bitrix не настроен: нет вебхука (bitrix_webhook_token).";
            return;
        }

        // Якорь хранится в исходном формате портала (с его часовым поясом):
        // и фильтр выборки, и сравнение дат идут в одном домене, конвертация
        // на нашей стороне смещала бы окно на разницу поясов.
        var anchorRaw = ReadSetting(AnchorKey);

        // Первый запуск версии: точкой отсчёта становится самая свежая задача
        // проекта — существующие сделки в «Контроль закупки» не переносятся.
        if (string.IsNullOrWhiteSpace(anchorRaw))
        {
            ProgressService.Report("Bitrix: первый запуск — определяю точку отсчёта…");
            var latestRaw = await BitrixService.GetLatestCreatedDateRawAsync(config);
            if (string.IsNullOrWhiteSpace(latestRaw))
            {
                report.Error = "В проекте нет ни одной задачи — не от чего отмерять новизну.";
                return;
            }

            PurchaseWriteRepository.SetSetting(AnchorKey, latestRaw);
            PurchaseWriteRepository.SetSetting(LastCheckKey, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            report.Initialized = true;
            return;
        }

        var anchorAbsolute = BitrixService.ParsePortalDate(anchorRaw);
        var linked = BitrixRepository.GetLinkedTaskIds();

        ProgressService.Report("Bitrix: загрузка новых задач…");
        var tasks = await BitrixService.GetTasksCreatedAfterAsync(config, anchorRaw);
        report.Fetched = tasks.Count;

        // Нераспознанные раньше задачи проверяем снова — название могли исправить.
        var skippedIds = ReadSkippedIds();
        if (skippedIds.Count > 0)
        {
            var retryIds = skippedIds.Where(id => !linked.Contains(id)).ToList();
            if (retryIds.Count > 0)
            {
                var known = tasks.Select(t => t.TaskId).ToHashSet();
                foreach (var retriedTask in await BitrixService.GetTasksByIdsAsync(config, retryIds))
                {
                    if (!known.Contains(retriedTask.TaskId))
                    {
                        tasks.Add(retriedTask);
                    }
                }
            }
        }

        if (tasks.Count > 0)
        {
            ProgressService.Report("Bitrix: обработка задач — " + tasks.Count + " шт.…");
            var existingDeals = GetExistingDeals();
            var stillSkipped = new List<int>();

            // От старых задач к новым, чтобы точка отсчёта двигалась корректно.
            foreach (var task in tasks.OrderBy(t => BitrixService.ParsePortalDate(t.CreatedDate) ?? DateTime.MinValue))
            {
                var created = BitrixService.ParsePortalDate(task.CreatedDate);
                if (created is not null && (anchorAbsolute is null || created.Value > anchorAbsolute.Value))
                {
                    anchorAbsolute = created;
                    anchorRaw = task.CreatedDate;
                }

                if (linked.Contains(task.TaskId))
                {
                    report.AlreadyPresent++;
                    continue;
                }

                var parts = BitrixService.ParseTaskTitle(task.Title);
                if (parts is null)
                {
                    // Название не по формату — но задача не потеряется: ID
                    // остаётся в списке, название перепроверится при следующих
                    // обходах (его можно исправить в Bitrix).
                    if (!stillSkipped.Contains(task.TaskId))
                    {
                        stillSkipped.Add(task.TaskId);
                        report.Skipped.Add("ID " + task.TaskId + " — «" + task.Title + "»");
                    }
                    continue;
                }

                var number = parts.RequestNumber;
                var existing = existingDeals.FirstOrDefault(d =>
                    string.Equals(d.Number, number, StringComparison.OrdinalIgnoreCase));
                if (existing.Id > 0)
                {
                    // Сделка с таким номером уже есть (создана вручную) — не
                    // трогаем её статус, только запоминаем связь с задачей.
                    BitrixRepository.LinkTask(existing.Id, task.TaskId);
                    linked.Add(task.TaskId);
                    report.AlreadyPresent++;
                    continue;
                }

                var dealId = PurchaseWriteRepository.CreateOrUpdateDeal(new PurchaseDealDraft
                {
                    DealNumber = number,
                    Client = parts.Client,
                    Title = parts.OriginalTitle,
                    Status = "RFQ",
                    BoardCount = parts.Quantity?.ToString(System.Globalization.CultureInfo.InvariantCulture),
                });
                BitrixRepository.LinkTask(dealId, task.TaskId);
                linked.Add(task.TaskId);
                existingDeals.Add((dealId, number));
                report.Imported.Add(number + (parts.Client.Length > 0 ? " — " + parts.Client : string.Empty));
            }

            // Импортированные сейчас задачи из списка ожидания убираются.
            stillSkipped.AddRange(skippedIds.Where(id => !linked.Contains(id) && !stillSkipped.Contains(id)));
            PurchaseWriteRepository.SetSetting(SkippedKey,
                string.Join("|", stillSkipped.Distinct().Take(MaxSkipped)));
            report.PendingSkippedCount = stillSkipped.Distinct().Count();
        }

        if (anchorRaw.Length > 0)
        {
            PurchaseWriteRepository.SetSetting(AnchorKey, anchorRaw);
        }
        PurchaseWriteRepository.SetSetting(LastCheckKey, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
    }

    /// <summary>Существующие сделки по номерам — защита от дублей при ручном создании.</summary>
    private static List<(long Id, string Number)> GetExistingDeals()
    {
        using var snapshot = PurchaseSnapshot.Create();
        var repo = new PurchaseRepository(snapshot.Connection);
        return repo.GetDeals()
            .Select(deal => (deal.Id, deal.DealNumber.Trim()))
            .Where(deal => deal.Item2.Length > 0)
            .ToList();
    }

    private static List<int> ReadSkippedIds()
        => (ReadSetting(SkippedKey) ?? string.Empty)
            .Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(text => int.TryParse(text, out var id) ? id : 0)
            .Where(id => id > 0)
            .Distinct()
            .ToList();

    private static string? ReadSetting(string key)
    {
        try
        {
            using var snapshot = PurchaseSnapshot.Create();
            return new SettingsRepository(snapshot.Connection).GetSetting(key);
        }
        catch
        {
            return null;
        }
    }
}
