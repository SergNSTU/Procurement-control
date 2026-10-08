using System.IO;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>Результат одной проверки почты поставщиков.</summary>
public sealed class MailImportReport
{
    /// <summary>Функция выключена настройкой — проверка не выполнялась.</summary>
    public bool Disabled { get; set; }

    /// <summary>Текст ошибки (Outlook не найден, ящик не найден и т.п.).</summary>
    public string? Error { get; set; }

    /// <summary>Сколько писем просмотрено с момента прошлой проверки.</summary>
    public int Scanned { get; set; }

    /// <summary>Строки вида «сделка ← файл.xlsx (отправитель)» по успешным загрузкам.</summary>
    public List<string> Imported { get; } = new();
}

/// <summary>
/// Тестовая фича: слежение за ящиком comp@npkexpert.ru. Письма поставщиков
/// с номером АКТИВНОЙ сделки (фильтр «Активные» из «Контроля закупки») в теме
/// и Excel-вложением загружаются в документы сделки как RRFQ. Отправители из
/// таблицы SupplierByEmail привязываются к сделке как поставщики, документ
/// кладётся в их папку; остальные — без привязки, папка = адрес отправителя.
/// Никаких отправок и удалений — только чтение почты. Обработанные письма
/// запоминаются по EntryID, чтобы не грузить файлы дважды.
/// </summary>
public static class OutlookMailImportService
{
    /// <summary>Имя хранилища в профиле Outlook, которое смотрит сервис.</summary>
    public const string MailboxName = "comp@npkexpert.ru";

    /// <summary>
    /// Соответствие почты отправителя и имени поставщика сделки. Ключи в нижнем
    /// регистре. Отправитель вне таблицы импортирует RRFQ без привязки
    /// (папка = адрес), как раньше.
    /// </summary>
    private static readonly string[] ExcelExtensions = { ".xlsx", ".xlsm", ".xls" };

    /// <summary>Свои домены: письма коллег и ERP поставщиками не считаются.</summary>
    private static readonly string[] OwnDomains = { "@unirec.ru", "@npkexpert.ru" };

    private const string EnabledKey = "mail.import.enabled";
    private const string LastCheckKey = "mail.import.last_check";
    private const string ProcessedKey = "mail.import.processed.v1";
    private const int MaxProcessedEntries = 400;

    /// <summary>Окно догоняющего скана: сколько назад смотреть письма на каждой
    /// проверке. Повторная загрузка предотвращается дедупликацией по EntryID,
    /// поэтому окно сознательно широкое — оно накрывает простой приложения и
    /// зависшие проверки (2026-10-06: прежние «5 минут от прошлой проверки»
    /// молча теряли письма, пришедшие между проверками).</summary>
    private static readonly TimeSpan CatchUpWindow = TimeSpan.FromDays(3);

    private static readonly object Gate = new();
    private static HashSet<string> _processedEntryIds = new(StringComparer.Ordinal);
    private static bool _processedLoaded;

    /// <summary>Чекбокс «Импорт RRFQ из почты» (настройка mail.import.enabled).</summary>
    public static bool IsEnabled
    {
        get
        {
            try
            {
                using var snapshot = PurchaseSnapshot.Create();
                var settings = new SettingsRepository(snapshot.Connection);
                return settings.GetBooleanSetting(EnabledKey, false);
            }
            catch
            {
                return false;
            }
        }
        set => PurchaseWriteRepository.SetSetting(EnabledKey, value ? "1" : "0");
    }

    /// <summary>
    /// Одна проверка ящика. Вызывается таймером раз в минуту и командой
    /// «Проверить почту поставщиков». Read-only по отношению к Outlook.
    /// </summary>
    public static MailImportReport Check(bool manual = false)
    {
        var report = new MailImportReport();
        // Прогресс виден в статусной строке главного окна: холодный запуск
        // Outlook и обход ящика занимают заметное время.
        ProgressService.Begin("Импорт почты: проверка ящик " + MailboxName + "…");
        try
        {
            // Прошлая проверка ещё выполняется (обычно — зависший COM-вызов
            // Outlook): не копим очередь потоков, а пропускаем такт. Обычная
            // проверка занимает секунды, поэтому авто-такт подождёт 5 с —
            // пропуск означает реальное зависание. Письма при этом не теряются:
            // окно скана догоняет их задним числом.
            if (!Monitor.TryEnter(Gate, manual ? TimeSpan.FromSeconds(30) : TimeSpan.FromSeconds(5)))
            {
                ReportStallStart("Проверка почты не выполняется: предыдущая проверка ещё работает (" +
                    (manual ? "ждали 30 с" : "такт пропущен") + ").");
                report.Error = "предыдущая проверка почты ещё выполняется";
                return report;
            }

            try
            {
                RepairMailImportFileNames();
                CheckCore(report, manual);
                _lastLoggedError = null;
                ReportStallEnd();
            }
            catch (Exception ex)
            {
                // Проверка выполняется каждую минуту: в лог пишем только смену
                // ошибки, чтобы постоянный сбой (например, закрытый Outlook)
                // не создавал запись в errors.log на каждый тик.
                if (ex.Message != _lastLoggedError)
                {
                    _lastLoggedError = ex.Message;
                    ErrorLog.Append("mail", ex);
                }

                report.Error = ex.Message;
            }
            finally
            {
                Monitor.Exit(Gate);
            }
        }
        finally
        {
            ProgressService.End();
        }
        return report;
    }

    private static string? _lastLoggedError;

    private static DateTime? _stallSince;

    /// <summary>Зависшая проверка: в лог пишем только начало и конец простоя,
    /// а не каждый пропущенный такт.</summary>
    private static void ReportStallStart(string message)
    {
        if (_stallSince is not null)
        {
            return;
        }
        _stallSince = DateTime.Now;
        ErrorLog.Append("mail", new InvalidOperationException(message));
    }

    private static void ReportStallEnd()
    {
        if (_stallSince is not { } since)
        {
            return;
        }
        _stallSince = null;
        ErrorLog.Append("mail", new InvalidOperationException(
            "Предыдущая проверка почты завершилась после простоя " +
            (DateTime.Now - since).ToString(@"hh\:mm\:ss") +
            " — письма за время простоя будут догнаны следующим сканом."));
    }

    private static bool _repairDone;

    /// <summary>
    /// Разовая починка документов (один раз за сессию перед первой проверкой почты):
    /// 1) документы дефектной сборки от 2026-09-29 со служебным именем временного
    /// файла "mailimport_&lt;guid&gt;_&lt;настоящее имя&gt;" — имя восстанавливается из суффикса;
    /// 2) документы типа RRFQ с именем "RFQ_..."/"RFQ ..." — переименовываются
    /// в "RRFQ_..." по правилу от 2026-09-29. Файл переименовывается на диске,
    /// ссылка в базе обновляется.
    /// </summary>
    private static void RepairMailImportFileNames()
    {
        if (_repairDone)
        {
            return;
        }
        _repairDone = true;

        const int prefixLength = 11;   // "mailimport_"
        const int guidLength = 32;     // Guid.NewGuid().ToString("N")

        try
        {
            PurchaseWriteSession.Execute(connection =>
            {
                var rows = new List<(long Id, string Name, string StoredPath)>();
                using (var select = connection.CreateCommand())
                {
                    select.CommandText = @"
SELECT id, original_name, stored_path FROM documents
WHERE substr(original_name, 1, @prefix) = 'mailimport_'
   OR (document_type = 'RRFQ'
       AND (original_name LIKE 'RFQ\_%' ESCAPE '\' OR original_name LIKE 'RFQ %'));";
                    PurchaseWriteRepository.Add(select, "@prefix", prefixLength);
                    using var reader = select.ExecuteReader();
                    while (reader.Read())
                    {
                        rows.Add((reader.GetInt64(0), reader.GetString(1), reader.GetString(2)));
                    }
                }

                foreach (var row in rows)
                {
                    var name = row.Name;
                    string? realName;
                    if (name.Length > prefixLength + guidLength + 1
                        && name[prefixLength + guidLength] == '_')
                    {
                        // Формат дефектной сборки: mailimport_<guid>_<настоящее имя>.
                        realName = name[(prefixLength + guidLength + 1)..].Trim();
                    }
                    else if (IsRfqNamedFile(name))
                    {
                        // Правило 2026-09-29: RFQ_.../RFQ ... → RRFQ_... (для документов RRFQ).
                        realName = NormalizeFileName(name);
                    }
                    else
                    {
                        continue; // старый формат без суффикса — настоящее имя восстановить нельзя
                    }

                    if (realName.Length == 0 || string.Equals(realName, name, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var currentPath = PurchaseWriteRepository.ResolvePortable(row.StoredPath);
                    var directory = Path.GetDirectoryName(currentPath);
                    if (string.IsNullOrWhiteSpace(directory) || !File.Exists(currentPath))
                    {
                        continue;
                    }

                    var stem = Path.GetFileNameWithoutExtension(realName);
                    var extension = Path.GetExtension(realName);
                    var candidate = Path.Combine(directory, realName);
                    var index = 2;
                    while (File.Exists(candidate))
                    {
                        candidate = Path.Combine(directory, $"{stem} ({index}){extension}");
                        index++;
                    }
                    File.Move(currentPath, candidate);

                    using var update = connection.CreateCommand();
                    update.CommandText = "UPDATE documents SET original_name = @name, stored_path = @path WHERE id = @id;";
                    PurchaseWriteRepository.Add(update, "@name", realName);
                    PurchaseWriteRepository.Add(update, "@path", PurchaseDocumentsStore.ToPortablePath(candidate));
                    PurchaseWriteRepository.Add(update, "@id", row.Id);
                    update.ExecuteNonQuery();
                }

                return 0;
            });
        }
        catch
        {
            // Починка не критична: строки останутся со старыми именами до следующего запуска.
        }
    }

    private static void CheckCore(MailImportReport report, bool manual)
    {
        if (!IsEnabled)
        {
            report.Disabled = true;
            return;
        }

        var activeDeals = GetActiveDeals();
        if (activeDeals.Count == 0)
        {
            return;
        }

        if (!OutlookService.IsAvailable)
        {
            report.Error = "Outlook не найден на этом компьютере";
            return;
        }

        dynamic application = Activator.CreateInstance(
            Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("Outlook.Application не зарегистрирован"))!;
        dynamic ns = application.GetNamespace("MAPI");

        dynamic inbox = FindInbox(ns);
        if (inbox is null)
        {
            report.Error = "В профиле Outlook нет ящика " + MailboxName;
            return;
        }

        var ownAddresses = CollectOwnAddresses(ns);

        // Время начала скана: оно же попадёт в last_check. Если проверка
        // зависла в COM-вызове Outlook на часы, окно следующей проверки всё
        // равно накроет письма, пришедшие за время зависания, — раньше они
        // терялись, потому что last_check двигало временем завершения.
        var scanStartedAt = DateTime.Now;

        var lastCheckText = ReadSetting(LastCheckKey);
        var windowStart = (PurchaseFormatting.ParseDate(lastCheckText) ?? scanStartedAt).Subtract(CatchUpWindow);

        EnsureProcessedLoaded();
        var processed = new List<string>(_processedEntryIds);

        // Restrict отбирает свежие письма силами Outlook — без сортировки
        // и перебора всего ящика через COM (заметно быстрее на больших ящиках).
        // DASL с UTC-значением, а не Jet «[ReceivedTime] >= 'MM/dd/yyyy …'»:
        // Jet-литерал парсится по локали, и на русской машине «MM/dd» молча
        // читается как «dd/MM» — фильтр пропускал месяцы писем вместо окна.
        var items = inbox.Items.Restrict("@SQL=\"urn:schemas:httpmail:datereceived\" >= '" +
            windowStart.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture) + "'");
        items.Sort("[ReceivedTime]", true);
        var total = items.Count;

        for (var index = 1; index <= total; index++)
        {
            dynamic item = items.Item(index);
            DateTime received;
            try
            {
                received = (DateTime)item.ReceivedTime;
            }
            catch
            {
                continue;
            }

            // Список отсортирован от новых к старым — дальше окна можно не смотреть.
            if (received < windowStart)
            {
                break;
            }

            report.Scanned++;

            string entryId;
            string subject;
            try
            {
                if ((int)item.Class != 43) // 43 = olMail
                {
                    continue;
                }
                entryId = (string)item.EntryID;
                subject = (string)(item.Subject ?? string.Empty);
            }
            catch
            {
                continue;
            }

            if (subject.Length == 0 || _processedEntryIds.Contains(entryId))
            {
                continue;
            }

            var matchedDeals = activeDeals
                .Where(deal => subject.IndexOf(deal.Number, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
            if (matchedDeals.Count == 0)
            {
                continue;
            }

            string sender = GetSenderSmtp(item);
            if (IsOwnSender(sender, ownAddresses))
            {
                continue;
            }

            // Письмо разобрано по текущим правилам — больше не возвращаемся к нему.
            processed.Add(entryId);

            foreach (dynamic attachment in EnumerateAttachments(item))
            {
                var fileName = string.Empty;
                try
                {
                    fileName = (string)attachment.FileName;
                }
                catch
                {
                    continue;
                }

                if (!IsExcelFile(fileName))
                {
                    continue;
                }

                // Имя файла как пришло от поставщика; "RFQ_..." переименовывается в "RRFQ_...".
                var finalName = NormalizeFileName(fileName);
                // Временный файл со служебным именем: настоящее имя передаётся
                // в AddDocument через displayName, чтобы в хранилище лёгло имя поставщика.
                var tempPath = Path.Combine(Path.GetTempPath(),
                    "mailimport_" + Guid.NewGuid().ToString("N") + Path.GetExtension(finalName));

                foreach (var deal in matchedDeals)
                {
                    long supplierId = 0;
                    string? supplierName = null;
                    var mapped = SupplierDirectory.FindSupplierByEmail(sender);
                    if (mapped is not null)
                    {
                        // Поставщик создаётся или «освежается» (ON CONFLICT по паре сделка+имя).
                        PurchaseWriteRepository.AddOrTouchSupplier(deal.Id, mapped);
                        supplierId = GetSupplierId(deal.Id, mapped);
                        supplierName = mapped;
                    }

                    try
                    {
                        attachment.SaveAsFile(tempPath);
                        // Привязанному поставщику папку задаёт AddDocument сам (имя поставщика);
                        // без привязки папка = адрес отправителя.
                        PurchaseDocumentsStore.AddDocument(
                            deal.Id, supplierId, "RRFQ", tempPath,
                            folderName: supplierId > 0 ? string.Empty : sender,
                            displayName: finalName);
                        report.Imported.Add(supplierName is null
                            ? deal.Number + " ← " + finalName + " (" + sender + ")"
                            : deal.Number + " ← " + finalName + " → " + supplierName);
                    }
                    finally
                    {
                        try { File.Delete(tempPath); } catch { /* временный файл не критичен */ }
                    }
                }
            }
        }

        _processedEntryIds = new HashSet<string>(processed, StringComparer.Ordinal);
        PurchaseWriteRepository.SetSetting(ProcessedKey, string.Join("|", processed.TakeLast(MaxProcessedEntries)));
        PurchaseWriteRepository.SetSetting(LastCheckKey, scanStartedAt.ToString("yyyy-MM-dd HH:mm:ss"));
    }

    /// <summary>Активные сделки тем же правилом, что и фильтр «Активные» на странице.</summary>
    private static List<(long Id, string Number)> GetActiveDeals()
    {
        using var snapshot = PurchaseSnapshot.Create();
        var repo = new PurchaseRepository(snapshot.Connection);
        return PurchaseLogic.FilterDeals(repo.GetDeals(), DealFilter.Active)
            .Select(deal => (deal.Id, deal.DealNumber.Trim()))
            .Where(deal => deal.Item2.Length > 0)
            .ToList();
    }

    /// <summary>Входящие ящика comp@npkexpert.ru; запасной вариант — дефолтный ящик.</summary>
    private static dynamic? FindInbox(dynamic ns)
    {
        foreach (dynamic store in ns.Stores)
        {
            if (string.Equals((string)store.DisplayName, MailboxName, StringComparison.OrdinalIgnoreCase))
            {
                return store.GetDefaultFolder(6); // olFolderInbox
            }
        }
        return null;
    }

    private static HashSet<string> CollectOwnAddresses(dynamic ns)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            foreach (dynamic account in ns.Accounts)
            {
                result.Add(((string)account.SmtpAddress).Trim().ToLowerInvariant());
            }
        }
        catch
        {
            // Список аккаунтов недоступен — останутся доменные правила.
        }
        return result;
    }

    /// <summary>SMTP отправителя: для внутренних адресов Outlook возвращает X.500 — разворачиваем через PR_SMTP_ADDRESS.</summary>
    private static string GetSenderSmtp(dynamic mail)
    {
        try
        {
            if (string.Equals((string)mail.SenderEmailType, "EX", StringComparison.OrdinalIgnoreCase))
            {
                return ((string)mail.PropertyAccessor
                    .GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")).Trim();
            }
        }
        catch
        {
            // Нет доступа к SMTP-свойству — берём адрес как есть.
        }

        try
        {
            return ((string)mail.SenderEmailAddress).Trim();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static bool IsOwnSender(string sender, HashSet<string> ownAddresses)
    {
        if (sender.Length == 0)
        {
            return true;
        }
        var lower = sender.ToLowerInvariant();
        if (ownAddresses.Contains(lower))
        {
            return true;
        }
        foreach (var domain in OwnDomains)
        {
            if (lower.EndsWith(domain, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        return false;
    }

    private static IEnumerable<dynamic> EnumerateAttachments(dynamic mail)
    {
        dynamic attachments;
        try
        {
            attachments = mail.Attachments;
        }
        catch
        {
            yield break;
        }

        var count = 0;
        try
        {
            count = (int)attachments.Count;
        }
        catch
        {
            yield break;
        }

        for (var index = 1; index <= count; index++)
        {
            yield return attachments.Item(index);
        }
    }

    private static bool IsExcelFile(string fileName)
    {
        var extension = Path.GetExtension(fileName);
        foreach (var excel in ExcelExtensions)
        {
            if (string.Equals(extension, excel, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Имя как пришло от поставщика, но "RFQ_..." и "RFQ ..." (в любом регистре)
    /// переименовываются в "RRFQ_...".
    /// </summary>
    internal static string NormalizeFileName(string fileName)
    {
        var name = fileName.TrimStart();
        if (name.StartsWith("RFQ_", StringComparison.OrdinalIgnoreCase))
        {
            return "RRFQ_" + name[4..];
        }
        if (name.StartsWith("RFQ ", StringComparison.OrdinalIgnoreCase))
        {
            return "RRFQ_" + name[5..];
        }
        return name;
    }

    /// <summary>True, если имя начинается с "RFQ_" или "RFQ " — кандидаты на RRFQ_.</summary>
    private static bool IsRfqNamedFile(string fileName)
        => fileName.StartsWith("RFQ_", StringComparison.OrdinalIgnoreCase)
           || fileName.StartsWith("RFQ ", StringComparison.OrdinalIgnoreCase);

    /// <summary>Id поставщика сделки по имени (после AddOrTouchSupplier он точно есть).</summary>
    private static long GetSupplierId(long dealId, string supplier)
    {
        using var snapshot = PurchaseSnapshot.Create();
        using var command = snapshot.Connection.CreateCommand();
        command.CommandText = @"
SELECT id FROM deal_suppliers
WHERE deal_id = @deal_id AND supplier = @supplier
ORDER BY id DESC LIMIT 1;";
        PurchaseWriteRepository.Add(command, "@deal_id", dealId);
        PurchaseWriteRepository.Add(command, "@supplier", supplier);
        var value = command.ExecuteScalar();
        return value is null or DBNull ? 0 : Convert.ToInt64(value);
    }

    private static void EnsureProcessedLoaded()
    {
        if (_processedLoaded)
        {
            return;
        }

        var raw = ReadSetting(ProcessedKey) ?? string.Empty;
        _processedEntryIds = new HashSet<string>(
            raw.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
            StringComparer.Ordinal);
        _processedLoaded = true;
    }

    private static string? ReadSetting(string key)
    {
        using var snapshot = PurchaseSnapshot.Create();
        return new SettingsRepository(snapshot.Connection).GetSetting(key);
    }
}
