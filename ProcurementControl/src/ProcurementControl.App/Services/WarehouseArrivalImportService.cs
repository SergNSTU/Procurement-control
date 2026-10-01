using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>Результат одной проверки файла приходов со склада.</summary>
public sealed class WarehouseArrivalReport
{
    /// <summary>Функция выключена настройкой — проверка не выполнялась.</summary>
    public bool Disabled { get; set; }

    /// <summary>Текст ошибки (Outlook не найден, папки нет, Excel не читается…).</summary>
    public string? Error { get; set; }

    /// <summary>Тема обработанного письма (для заголовка отчёта).</summary>
    public string? FileSubject { get; set; }

    /// <summary>Куда сохранён Excel из письма (архив _arrivals).</summary>
    public string? SavedPath { get; set; }

    /// <summary>Сколько строкам поставщиков записано «Факт поступл.».</summary>
    public int FilledFacts { get; set; }

    /// <summary>Информационные строки: обработанные заказы «пришло целиком».</summary>
    public List<string> Lines { get; } = new();

    /// <summary>Замечания: частичные приходы и заказы, которых нет в файле.</summary>
    public List<WarehouseArrivalIssue> Issues { get; } = new();
}

/// <summary>
/// Замечание по одной паре «сделка + поставщик»: пришло не всё или заказ ещё
/// не появился в файле прихода. Служит строкой отчёта и базой для черновика
/// письма поставщику («пушнуть»).
/// </summary>
public sealed class WarehouseArrivalIssue
{
    public long DealId { get; init; }
    public string DealNumber { get; init; } = string.Empty;
    public long SupplierId { get; init; }
    public string Supplier { get; init; } = string.Empty;

    /// <summary>Номер заказа с суффиксом из комментария поставщика.</summary>
    public string PoNumber { get; init; } = string.Empty;

    /// <summary>Сколько позиций должно прийти (строки ERP-файла поставщика).</summary>
    public int Expected { get; init; }

    /// <summary>Сколько позиций пришло (строки с заполненной датой прибытия).</summary>
    public int Arrived { get; init; }

    /// <summary>Последняя дата прибытия по заказу в файле склада.</summary>
    public DateTime? ArrivalDate { get; init; }

    /// <summary>Плановая дата поступления («План поступл.» в карточке сделки).</summary>
    public DateTime? Deadline { get; init; }

    /// <summary>«partial» — пришло не всё; «missing» — заказа нет в файле прихода.</summary>
    public string Status { get; init; } = "missing";

    /// <summary>Пояснение (ERP-файл не найден, в файле есть без даты прибытия…).</summary>
    public string Note { get; init; } = string.Empty;

    /// <summary>Имя поставщика без служебного суффикса («Компэл_2» → «Компэл»).</summary>
    public string BaseSupplier => Regex.Replace(Supplier, @"_\d+$", string.Empty);

    public bool Overdue => Deadline is not null && Deadline.Value.Date < DateTime.Today;

    public string DeadlineText => Deadline?.ToString("dd.MM.yyyy") ?? "план не указан";

    /// <summary>Строка отчёта по замечанию.</summary>
    public string DisplayText
    {
        get
        {
            var prefix = Status == "partial" ? "△" : "✗";
            var head = prefix + " " + PoNumber + " (" + DealNumber + " · " + Supplier + "): ";
            var body = Status == "partial"
                ? "частичный приход: " + Arrived + " из " + Expected + " позиций"
                : "нет в файле прихода";
            if (Expected == 0 && Status != "missing")
            {
                body = "пришло " + Arrived + " позиций, ERP-файл не найден — сверка количества невозможна";
            }
            if (Note.Length > 0)
            {
                body += " (" + Note + ")";
            }

            return head + body + "; " + DeadlineText +
                   (Overdue ? " — просрочен" : string.Empty);
        }
    }
}

/// <summary>
/// Импорт ежедневного файла прихода компонентов на склад: письма «The CE CT WE
/// file daily update …» из папки «CE CT WE» ящика comp@npkexpert.ru несут
/// Excel, где по номерам заказов (номер сделки + суффикс поставщика, напр.
/// «AE26060501EB-002968-C1») перечислены позиции с датами прибытия.
///
/// Связка «строка поставщика → номер заказа» ведётся вручную в комментарии
/// поставщика («AE…-C4 / Маски»). Ожидаемое число позиций берётся из ERP-файла
/// поставщика (документы «ERP», колонка PI Number). Пришедшие целиком заказы
/// заполняют «Факт поступл.» (если пусто), по остальным — отчёт и черновик
/// письма поставщику. Статусы и этапы сделок не меняются, почта только читается.
/// </summary>
public static class WarehouseArrivalImportService
{
    /// <summary>Имя хранилища и папки с письмами прихода (правило Outlook уже складывает их туда).</summary>
    public const string MailboxName = OutlookMailImportService.MailboxName;
    public const string FolderName = "CE CT WE";

    /// <summary>Устойчивая часть темы письма склада.</summary>
    private const string SubjectMarker = "CE CT WE file daily update";

    private const string EnabledKey = "warehouse.import.enabled";
    private const string LastCheckKey = "warehouse.import.last_check";
    private const string ProcessedKey = "warehouse.import.processed.v1";
    private const int MaxProcessedEntries = 400;

    /// <summary>Период авто-проверки (решение от 2026-10-01: раз в два часа).</summary>
    private static readonly TimeSpan AutoCheckInterval = TimeSpan.FromHours(2);

    /// <summary>Окно первой проверки: файлы прихода ежедневные, недели хватает.</summary>
    private static readonly TimeSpan FirstRunWindow = TimeSpan.FromDays(7);

    /// <summary>За сколько дней до «Плана поступл.» предупреждать о непришедшем заказе.</summary>
    public const int DeadlineWarningDays = 5;

    // Колонки листа «Orders' materials onlly (…)», 1-базные (шапка 4 строки).
    private const int OrderColumn = 2;
    private const int ArrivalColumn = 14;

    private static readonly object Gate = new();
    private static HashSet<string> _processedEntryIds = new(StringComparer.Ordinal);
    private static bool _processedLoaded;
    private static string? _lastLoggedError;

    /// <summary>Чекбокс «Импорт приходов со склада» (настройка warehouse.import.enabled).</summary>
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

    /// <summary>Пора ли авто-проверке выполняться (прошло 2 часа с прошлой).</summary>
    public static bool DueForAutoCheck
    {
        get
        {
            var last = PurchaseFormatting.ParseDate(ReadSetting(LastCheckKey));
            return last is null || DateTime.Now - last.Value >= AutoCheckInterval;
        }
    }

    /// <summary>
    /// Одна проверка папки «CE CT WE». Вызывается авто-таймером раз в 2 часа
    /// и командой ручной проверки. Read-only по отношению к Outlook.
    /// </summary>
    public static WarehouseArrivalReport Check(bool manual = false)
    {
        var report = new WarehouseArrivalReport();
        ProgressService.Begin("Приход со склада: проверка папки " + FolderName + "…");
        lock (Gate)
        {
            try
            {
                CheckCore(report);
                _lastLoggedError = null;
            }
            catch (Exception ex)
            {
                // Авто-проверка редкая, но при постоянном сбое не засоряем лог.
                if (ex.Message != _lastLoggedError)
                {
                    _lastLoggedError = ex.Message;
                    ErrorLog.Append("warehouse", ex);
                }

                report.Error = ex.Message;
            }
            finally
            {
                ProgressService.End();
            }
        }
        return report;
    }

    private static void CheckCore(WarehouseArrivalReport report)
    {
        if (!IsEnabled)
        {
            report.Disabled = true;
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

        dynamic? folder = FindArrivalFolder(ns);
        if (folder is null)
        {
            report.Error = "В ящике " + MailboxName + " нет папки «" + FolderName + "»";
            return;
        }

        var lastCheckText = ReadSetting(LastCheckKey);
        var firstRun = string.IsNullOrWhiteSpace(lastCheckText);
        var windowStart = firstRun
            ? DateTime.Now.Subtract(FirstRunWindow)
            : (PurchaseFormatting.ParseDate(lastCheckText) ?? DateTime.Now).Subtract(TimeSpan.FromMinutes(30));

        EnsureProcessedLoaded();

        // Новых писем обычно 0-1, файл накопительный — достаточно самого свежего.
        dynamic? newestMail = null;
        var seenIds = new List<string>();
        var items = folder.Items.Restrict("[ReceivedTime] >= '" +
            windowStart.ToString("MM/dd/yyyy HH:mm", CultureInfo.InvariantCulture) + "'");
        items.Sort("[ReceivedTime]", true);
        var total = (int)items.Count;
        for (var index = 1; index <= total; index++)
        {
            dynamic item = items.Item(index);
            try
            {
                if ((int)item.Class != 43 || _processedEntryIds.Contains((string)item.EntryID))
                {
                    continue;
                }

                var subject = (string)(item.Subject ?? string.Empty);
                if (subject.IndexOf(SubjectMarker, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                seenIds.Add((string)item.EntryID);
                newestMail ??= item;
            }
            catch
            {
                // Битый элемент папки пропускаем.
            }
        }

        if (newestMail is null)
        {
            // Нового файла нет — отметку времени всё равно двигаем, чтобы не
            // пересканировать окно на каждом тике.
            PurchaseWriteRepository.SetSetting(LastCheckKey, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            return;
        }

        dynamic mail = newestMail;
        report.FileSubject = (string)(mail.Subject ?? string.Empty);

        // 1. Сохранить вложение в архив _arrivals.
        var savedPath = SaveAttachment(mail);
        if (savedPath is null)
        {
            report.Error = "В письме «" + report.FileSubject + "» нет Excel-вложения";
            return;
        }
        report.SavedPath = savedPath;

        // 2. Разобрать файл прихода: номер заказа → даты прибытия его строк.
        var arrivals = ParseArrivalWorkbook(savedPath);

        // 3. Кандидаты: неархивные сделки, у поставщиков которых в комментарии номер с суффиксом.
        var candidates = LoadCandidates();
        var matchCache = new Dictionary<string, List<DateTime?>?>(StringComparer.Ordinal);
        var processedLines = 0;

        foreach (var candidate in candidates)
        {
            if (!matchCache.TryGetValue(candidate.Po, out var dates))
            {
                dates = FindGroupDates(arrivals, candidate.Po);
                matchCache[candidate.Po] = dates;
            }

            var arrived = dates?.Count(date => date is not null) ?? 0;
            var arrivalDate = dates?.Where(date => date is not null).DefaultIfEmpty(null).Max();
            var expected = GetExpectedPositions(candidate.Deal.Id, candidate.Po);

            if (expected > 0 && arrived >= expected)
            {
                // Заказ пришёл целиком: заполняем «Факт поступл.», если он не был отмечен.
                processedLines++;
                if (candidate.Supplier.ActualReceiptDate.Trim().Length == 0 && arrivalDate is not null)
                {
                    PurchaseWriteRepository.SetSupplierActualReceiptDate(candidate.Supplier.Id, arrivalDate);
                    report.FilledFacts++;
                    report.Lines.Add("✓ " + candidate.Po + " (" + candidate.Supplier.Supplier + "): пришло " +
                                     arrived + "/" + expected + ", приход " + arrivalDate.Value.ToString("dd.MM.yyyy") +
                                     " → записан «Факт поступл.»");
                }
                else
                {
                    report.Lines.Add("✓ " + candidate.Po + " (" + candidate.Supplier.Supplier + "): пришло " +
                                     arrived + "/" + expected + ", приход " +
                                     (arrivalDate?.ToString("dd.MM.yyyy") ?? "—") + " (факт уже был отмечен)");
                }
                continue;
            }

            if (arrived == 0 && expected == 0)
            {
                // Ни в файле (или без дат), ни ERP: следить за количеством не по чем.
                report.Issues.Add(new WarehouseArrivalIssue
                {
                    DealId = candidate.Deal.Id, DealNumber = candidate.Deal.DealNumber,
                    SupplierId = candidate.Supplier.Id, Supplier = candidate.Supplier.Supplier,
                    PoNumber = candidate.Po, Expected = 0, Arrived = 0,
                    Deadline = ParseDeadline(candidate.Supplier.ComponentsReceiptDate),
                    Status = "missing",
                    Note = dates is not null ? "в файле есть строки без даты прибытия (подготовлено)" : string.Empty
                });
                continue;
            }

            if (arrived == 0)
            {
                processedLines++;
                report.Issues.Add(new WarehouseArrivalIssue
                {
                    DealId = candidate.Deal.Id, DealNumber = candidate.Deal.DealNumber,
                    SupplierId = candidate.Supplier.Id, Supplier = candidate.Supplier.Supplier,
                    PoNumber = candidate.Po, Expected = expected, Arrived = 0,
                    Deadline = ParseDeadline(candidate.Supplier.ComponentsReceiptDate),
                    Status = "missing",
                    Note = dates is not null ? "в файле без даты прибытия (подготовлено)" : string.Empty
                });
                continue;
            }

            // Частичный приход (или сверка невозможна без ERP).
            processedLines++;
            report.Issues.Add(new WarehouseArrivalIssue
            {
                DealId = candidate.Deal.Id, DealNumber = candidate.Deal.DealNumber,
                SupplierId = candidate.Supplier.Id, Supplier = candidate.Supplier.Supplier,
                PoNumber = candidate.Po, Expected = expected, Arrived = arrived,
                ArrivalDate = arrivalDate,
                Deadline = ParseDeadline(candidate.Supplier.ComponentsReceiptDate),
                Status = "partial"
            });
        }

        if (processedLines > 0 || report.FilledFacts > 0)
        {
            report.Lines.Insert(0, "Заказов в работе: " + candidates.Count +
                                    ", с изменениями: " + processedLines +
                                    ", «Факт поступл.» записан: " + report.FilledFacts);
        }

        // 4. Письмо (все новые письма папки) больше не возвращается к разбору.
        var processed = new List<string>(_processedEntryIds);
        processed.AddRange(seenIds);
        _processedEntryIds = new HashSet<string>(processed, StringComparer.Ordinal);
        PurchaseWriteRepository.SetSetting(ProcessedKey,
            string.Join("|", processed.TakeLast(MaxProcessedEntries)));
        PurchaseWriteRepository.SetSetting(LastCheckKey, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
    }

    /// <summary>Папка «CE CT WE» в корне хранилища comp@npkexpert.ru.</summary>
    private static dynamic? FindArrivalFolder(dynamic ns)
    {
        foreach (dynamic store in ns.Stores)
        {
            if (!string.Equals((string)store.DisplayName, MailboxName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            foreach (dynamic folder in store.GetRootFolder().Folders)
            {
                if (string.Equals((string)folder.Name, FolderName, StringComparison.OrdinalIgnoreCase))
                {
                    return folder;
                }
            }
        }
        return null;
    }

    /// <summary>Excel-вложение письма → архив data\…\files\_arrivals; путь или null.</summary>
    private static string? SaveAttachment(dynamic mail)
    {
        foreach (dynamic attachment in EnumerateAttachments(mail))
        {
            string fileName;
            try
            {
                fileName = (string)attachment.FileName;
            }
            catch
            {
                continue;
            }

            var extension = Path.GetExtension(fileName);
            if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                && !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase)
                && !extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var directory = Path.Combine(PurchaseDocumentsStore.GetDocumentsRoot(), "_arrivals");
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, fileName);
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                attachment.SaveAsFile(path);
                return path;
            }
            catch
            {
                return null;
            }
        }
        return null;
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

    /// <summary>
    /// Разбор книги прихода: листы «Orders' materials onlly (…)». Номер заказа
    /// стоит только в первой строке группы (объединённые ячейки), поэтому он
    /// «протягивается» вниз; для каждой группы собираем даты прибытия строк.
    /// Ключ словаря — текст ячейки с номером как есть (матчинг — подстрокой).
    /// </summary>
    internal static Dictionary<string, List<DateTime?>> ParseArrivalWorkbook(string path)
    {
        var result = new Dictionary<string, List<DateTime?>>(StringComparer.Ordinal);
        dynamic? excel = null;
        try
        {
            excel = RrfqExcel.OpenExcel();
            dynamic workbook = RrfqExcel.OpenWorkbookReadOnly(excel, path);
            try
            {
                var sheetCount = (int)workbook.Worksheets.Count;
                for (var index = 1; index <= sheetCount; index++)
                {
                    dynamic ws = workbook.Worksheets.Item(index);
                    string name;
                    try
                    {
                        name = (string)ws.Name;
                    }
                    catch
                    {
                        continue;
                    }

                    // Четыре листа-«выгрузки» (CE/CT/AE/AT); остальные листы
                    // (CE&WE, CT, PE, Direct payment) — другая раскладка.
                    if (name.IndexOf("materials onlly", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    ReadArrivalSheet(ws, result);
                }
            }
            finally
            {
                workbook.Close(false);
                RrfqExcel.Release(workbook);
            }
        }
        finally
        {
            RrfqExcel.CloseExcel(excel);
        }
        return result;
    }

    private static void ReadArrivalSheet(dynamic ws, Dictionary<string, List<DateTime?>> result)
    {
        // (object) разрывает динамическую диспетчеризацию: иначе весь метод
        // компилируется как dynamic и out var не выводится.
        var values = ToMatrix((object)ws.UsedRange.Value2);
        if (values is null || values.GetLength(1) < OrderColumn)
        {
            return;
        }

        var rows = values.GetLength(0);
        var hasDates = values.GetLength(1) >= ArrivalColumn;
        string? current = null;
        for (var row = 0; row < rows; row++)
        {
            var cell = ToText(values[row, OrderColumn - 1]);
            if (cell.Length > 0)
            {
                current = cell;
                if (!result.TryGetValue(cell, out var group))
                {
                    group = new List<DateTime?>();
                    result[cell] = group;
                }
            }

            if (current is null)
            {
                continue;
            }

            result[current].Add(hasDates ? ToDate(values[row, ArrivalColumn - 1]) : null);
        }
    }

    /// <summary>COM-срез Value2 (1-базовый) → обычный 0-базовый массив.</summary>
    private static object?[,]? ToMatrix(object? raw)
    {
        if (raw is object[,] com)
        {
            var rows = com.GetLength(0);
            var cols = com.GetLength(1);
            var rowBase = com.GetLowerBound(0);
            var colBase = com.GetLowerBound(1);
            var result = new object?[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    result[r, c] = com.GetValue(rowBase + r, colBase + c);
                }
            }
            return result;
        }
        return null;
    }

    private static string ToText(object? value)
        => Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;

    /// <summary>Дата в файле: числовой serial Excel или текст (yyyy-MM-dd / yyyy.MM.dd / dd.MM.yyyy).</summary>
    private static DateTime? ToDate(object? value)
    {
        if (value is double serial && serial > 20000 && serial < 80000)
        {
            try
            {
                return DateTime.FromOADate(serial).Date;
            }
            catch
            {
                return null;
            }
        }

        if (value is DateTime direct)
        {
            return direct.Date;
        }

        var text = ToText(value);
        if (text.Length < 8)
        {
            return null;
        }

        string[] formats = { "yyyy-MM-dd", "yyyy.MM.dd", "yyyy/MM/dd", "dd.MM.yyyy", "MM/dd/yyyy" };
        foreach (var culture in new[] { CultureInfo.InvariantCulture, CultureInfo.GetCultureInfo("ru-RU") })
        {
            if (DateTime.TryParseExact(text, formats, culture, DateTimeStyles.None, out var parsed))
            {
                return parsed.Date;
            }
        }
        return null;
    }

    /// <summary>Кандидаты: (сделка, поставщик, номер заказа с суффиксом из комментария).</summary>
    private sealed record Candidate(PurchaseDealRow Deal, SupplierRow Supplier, string Po);

    private static List<Candidate> LoadCandidates()
    {
        var candidates = new List<Candidate>();
        using var snapshot = PurchaseSnapshot.Create();
        var repo = new PurchaseRepository(snapshot.Connection);
        foreach (var deal in repo.GetDeals())
        {
            if (deal.Archived == 1)
            {
                continue;
            }

            var number = deal.DealNumber.Trim();
            if (number.Length == 0)
            {
                continue;
            }

            foreach (var supplier in repo.GetSuppliers(deal.Id))
            {
                var po = ExtractPoNumber(number, supplier.Comment);
                if (po is not null)
                {
                    candidates.Add(new Candidate(deal, supplier, po));
                }
            }
        }
        return candidates;
    }

    /// <summary>
    /// Номер заказа из комментария поставщика: «&lt;номер сделки&gt;-&lt;суффикс&gt; / текст»,
    /// короткая форма «-2», или просто номер сделки (без суффикса). Суффиксы:
    /// «1…5», «C1…C8».
    /// </summary>
    internal static string? ExtractPoNumber(string dealNumber, string? comment)
    {
        if (string.IsNullOrWhiteSpace(comment) || dealNumber.Length == 0)
        {
            return null;
        }

        var text = comment.Trim();
        var suffix = @"(\d{1,3}|[A-Za-z]\d{1,3})(?![\dA-Za-z])";

        var full = Regex.Match(text, Regex.Escape(dealNumber) + @"\s*-" + suffix,
            RegexOptions.IgnoreCase);
        if (full.Success)
        {
            return dealNumber + "-" + full.Groups[1].Value;
        }

        var shortForm = Regex.Match(text, @"^-\s*" + suffix, RegexOptions.IgnoreCase);
        if (shortForm.Success)
        {
            return dealNumber + "-" + shortForm.Groups[1].Value;
        }

        if (text.StartsWith(dealNumber, StringComparison.OrdinalIgnoreCase))
        {
            return dealNumber;
        }

        return null;
    }

    /// <summary>
    /// Поиск группы заказа: ячейка файла содержит номер целиком. Границы:
    /// перед номером не буква/цифра (префиксы «2-», «C-» разрешены), после —
    /// не буква/цифра и не «-» (чтобы «AE…-002968» не совпал с «AE…-002968-5»).
    /// </summary>
    private static List<DateTime?>? FindGroupDates(Dictionary<string, List<DateTime?>> arrivals, string po)
    {
        if (arrivals.TryGetValue(po, out var exact))
        {
            return exact;
        }

        foreach (var (cell, dates) in arrivals)
        {
            if (ContainsToken(cell, po))
            {
                return dates;
            }
        }
        return null;
    }

    private static bool ContainsToken(string cell, string token)
    {
        var scan = 0;
        int start;
        while ((start = cell.IndexOf(token, scan, StringComparison.OrdinalIgnoreCase)) >= 0)
        {
            if (start > 0 && char.IsLetterOrDigit(cell[start - 1]))
            {
                scan = start + 1;
                continue;
            }

            var end = start + token.Length;
            if (end < cell.Length)
            {
                var next = cell[end];
                if (char.IsLetterOrDigit(next) || next == '-')
                {
                    scan = start + 1;
                    continue;
                }
            }
            return true;
        }
        return false;
    }

    /// <summary>Ожидаемое число позиций: строки ERP-файлов сделки с этим PI Number.</summary>
    private static int GetExpectedPositions(long dealId, string po)
    {
        var expected = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var document in PurchaseDocumentsStore.GetDocuments(dealId))
        {
            if (!string.Equals(document.DocumentType, "ERP", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var path = PurchaseWriteRepository.ResolvePortable(document.StoredPath);
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                continue;
            }

            try
            {
                ReadErpFile(path, expected);
            }
            catch
            {
                // Один нечитаемый ERP-файл не должен ронять всю проверку.
            }
        }
        return expected.GetValueOrDefault(po);
    }

    /// <summary>Чтение ERP-файла: PI Number (кол. 5, с протягиванием) → число строк позиций (кол. 3).</summary>
    private static void ReadErpFile(string path, Dictionary<string, int> expected)
    {
        dynamic? excel = null;
        try
        {
            excel = RrfqExcel.OpenExcel();
            dynamic workbook = RrfqExcel.OpenWorkbookReadOnly(excel, path);
            try
            {
                var sheetCount = (int)workbook.Worksheets.Count;
                for (var index = 1; index <= sheetCount; index++)
                {
                    var values = ToMatrix((object)workbook.Worksheets.Item(index).UsedRange.Value2);
                    if (values is null || values.GetLength(1) < 5)
                    {
                        continue;
                    }

                    string? currentPi = null;
                    for (var row = 0; row < values.GetLength(0); row++)
                    {
                        var pi = ToText(values[row, 4]);
                        if (pi.Length > 0)
                        {
                            currentPi = pi;
                        }

                        var erpCode = ToText(values[row, 2]);
                        if (currentPi is not null && erpCode.Length > 0
                            && !erpCode.Equals("ERP-Code", StringComparison.OrdinalIgnoreCase))
                        {
                            expected[currentPi] = expected.GetValueOrDefault(currentPi) + 1;
                        }
                    }
                }
            }
            finally
            {
                workbook.Close(false);
                RrfqExcel.Release(workbook);
            }
        }
        finally
        {
            RrfqExcel.CloseExcel(excel);
        }
    }

    private static DateTime? ParseDeadline(string value) => PurchaseFormatting.ParseDate(value);

    /// <summary>Черновик письма поставщику по замечанию («пушнуть»). Только Display().</summary>
    public static void PushSupplier(WarehouseArrivalIssue issue)
    {
        var email = SupplierDirectory.FindEmailBySupplier(issue.BaseSupplier) ?? string.Empty;
        var overdue = issue.Overdue ? " — дата уже прошла" : string.Empty;
        var subject = issue.PoNumber + ": сроки поставки? (сделка " + issue.DealNumber + ")";
        var body = "Добрый день!\n\n"
            + "Просьба перепроверить поставку " + issue.PoNumber
            + " (сделка " + issue.DealNumber + ", позиций: " + issue.Expected + ").\n"
            + (issue.Arrived > 0
                ? "По файлу прихода на складе пока " + issue.Arrived + " из " + issue.Expected
                  + " позиций (последний приход " + (issue.ArrivalDate?.ToString("dd.MM.yyyy") ?? "—") + ").\n"
                : "Заказа пока нет в ежедневном файле прихода на склад.\n")
            + "Плановая дата поступления: " + issue.DeadlineText + overdue + ".\n\n"
            + "Просьба уточнить статус и подтвердить срок.\n\nСпасибо!";

        OutlookService.CreateMailDraft(email, subject, body);
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
