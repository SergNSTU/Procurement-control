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

    /// <summary>«partial» — пришло не всё; «missing» — заказа нет в файле прихода;
    /// «nodate» — строки совпадают с ERP, но даты прибытия в файле нет.</summary>
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
            var prefix = Status == "missing" ? "✗" : "△";
            var head = prefix + " " + PoNumber + " (" + DealNumber + " · " + Supplier + "): ";
            string body;
            if (Status == "nodate")
            {
                body = "строки совпадают (" + Arrived + "/" + Expected + "), но даты прибытия в файле нет" +
                       " — проставьте «Факт поступл.» вручную";
            }
            else if (Status == "partial")
            {
                body = Expected == 0
                    ? "пришло " + Arrived + " позиций, ERP-файл не найден — сверка количества невозможна"
                    : "частичный приход: " + Arrived + " из " + Expected + " позиций";
            }
            else
            {
                body = "нет в файле прихода";
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
/// поставщика (документы «ERP», колонка PI Number). Пришедшие целиком по ERP
/// заказы синхронизируют «Факт поступл.» с файлом: пустое поле заполняется,
/// расходящееся с датой прихода — исправляется; без ERP-файла дата не пишется
/// (неизвестно, всё ли должно прийти). По остальным — отчёт и черновик письма
/// поставщику. Статусы и этапы сделок не меняются, почта только читается.
///
/// Семантика отчёта (2026-10-01): строки «записан/исправлен» показываются
/// один раз — в прогоне, где произошла запись; уже совпадающие заказы не
/// выводятся. Вопросы (нет в файле, частичный приход, ERP не найден, факт
/// без заказа) показываются при каждой проверке, пока не разрешатся.
///
/// Каждая проверка заново разбирает САМОЕ СВЕЖЕЕ письмо папки, даже если оно
/// уже разбиралось: файл накопительный, а пользователь догружает ERP-файлы и
/// повторяет проверку, чтобы увидеть изменения (решение 2026-10-01).
/// Поэтому дедупликация писем не нужна — только отметка времени для графика
/// авто-проверки (раз в 2 часа).
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

    /// <summary>Период авто-проверки (решение от 2026-10-01: раз в два часа).</summary>
    private static readonly TimeSpan AutoCheckInterval = TimeSpan.FromHours(2);

    /// <summary>За сколько дней до «Плана поступл.» предупреждать о непришедшем заказе.</summary>
    public const int DeadlineWarningDays = 5;

    // Колонки листа «Orders' materials onlly (…)», 1-базные (шапка 4 строки).
    private const int OrderColumn = 2;
    private const int ArrivalColumn = 14;

    private static readonly object Gate = new();
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

        // Разбираем самое свежее письмо папки при каждой проверке — даже если
        // оно уже разбиралось (см. описание класса).
        dynamic? newestMail = null;
        dynamic items = folder.Items;
        items.Sort("[ReceivedTime]", true);
        var total = (int)items.Count;
        for (var index = 1; index <= total; index++)
        {
            dynamic item = items.Item(index);
            try
            {
                if ((int)item.Class != 43)
                {
                    continue;
                }

                var subject = (string)(item.Subject ?? string.Empty);
                if (subject.IndexOf(SubjectMarker, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    newestMail = item; // список от новых к старым — первое совпадение свежее
                    break;
                }
            }
            catch
            {
                // Битый элемент папки пропускаем.
            }
        }

        if (newestMail is null)
        {
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

        // 2–3. Разбор файла прихода и сверка с ERP — в ОДНОМ экземпляре Excel.
        // Прошлый вариант открывал свой Excel на каждый файл (до ~30 за прогон),
        // и недогашенные экземпляры копились как скрытые EXCEL.EXE с «открытыми
        // файлами» — Windows ругалась на них при выключении (случай 2026-10-02).
        dynamic? excel = null;
        try
        {
            excel = RrfqExcel.OpenExcel();

            var arrivals = ParseArrivalWorkbook((object)excel, savedPath);
            var candidates = LoadCandidates();
            var matchCache = new Dictionary<string, List<DateTime?>?>(StringComparer.Ordinal);
            var erpCache = new Dictionary<long, Dictionary<string, int>>();

            foreach (var candidate in candidates)
            {
                if (!matchCache.TryGetValue(candidate.Po, out var dates))
                {
                    dates = FindGroupDates(arrivals, candidate.Po);
                    matchCache[candidate.Po] = dates;
                }

                var arrived = dates?.Count(date => date is not null) ?? 0;
                var arrivalDate = dates?.Where(date => date is not null).DefaultIfEmpty(null).Max();
                if (!erpCache.TryGetValue(candidate.Deal.Id, out var erp))
                {
                    erp = LoadErpExpectations(excel, candidate.Deal.Id);
                    erpCache[candidate.Deal.Id] = erp;
                }
                var expected = erp.GetValueOrDefault(candidate.Po);

            // «Факт поступл.» синхронизируем с файлом только когда заказ пришёл
            // целиком по ERP. Без ERP-файла неизвестно, всё ли должно прийти, —
            // дату не пишем и не исправляем (решение 2026-10-01).
            // Строки успеха показываются один раз: «записан/исправлен» возникают
            // только в прогоне с записью, дальше факт совпадает и не выводится.
            if (expected > 0 && arrived >= expected)
            {
                var label = candidate.Po + " (" + candidate.Supplier.Supplier + ")";
                var factText = candidate.Supplier.ActualReceiptDate.Trim();
                var factDate = PurchaseFormatting.ParseDate(factText);
                if (factText.Length == 0)
                {
                    PurchaseWriteRepository.SetSupplierActualReceiptDate(candidate.Supplier.Id, arrivalDate);
                    report.FilledFacts++;
                    report.Lines.Add("✓ " + label + ": пришло " + arrived + "/" + expected +
                                     ", приход " + arrivalDate.Value.ToString("dd.MM.yyyy") +
                                     " → записан «Факт поступл.»");
                }
                else if (factDate is null || factDate.Value.Date != arrivalDate.Value.Date)
                {
                    PurchaseWriteRepository.SetSupplierActualReceiptDate(candidate.Supplier.Id, arrivalDate);
                    report.FilledFacts++;
                    report.Lines.Add("✓ " + label + ": приход по файлу " +
                                     arrivalDate.Value.ToString("dd.MM.yyyy") +
                                     " → исправлен «Факт поступл.» (было: " +
                                     (factDate is not null ? factDate.Value.ToString("dd.MM.yyyy") : factText) + ")");
                }
                continue;
            }

            // Строк в файле столько, сколько позиций в ERP, но даты прибытия
            // склад не проставил/стёр (бывает на старых заказах): сверка по
            // строкам пройдена, дате программе взяться неоткуда — вопрос
            // закрывается, когда пользователь сам проставит «Факт поступл.»
            // (решение 2026-10-01).
            var totalRows = dates?.Count ?? 0;
            if (expected > 0 && totalRows >= expected && arrived < expected)
            {
                if (candidate.Supplier.ActualReceiptDate.Trim().Length == 0)
                {
                    report.Issues.Add(new WarehouseArrivalIssue
                    {
                        DealId = candidate.Deal.Id, DealNumber = candidate.Deal.DealNumber,
                        SupplierId = candidate.Supplier.Id, Supplier = candidate.Supplier.Supplier,
                        PoNumber = candidate.Po, Expected = expected, Arrived = totalRows,
                        ArrivalDate = arrivalDate,
                        Deadline = ParseDeadline(candidate.Supplier.ComponentsReceiptDate),
                        Status = "nodate",
                        Note = arrivalDate is not null
                            ? "частичная дата в файле есть: " + arrivalDate.Value.ToString("dd.MM.yyyy")
                            : string.Empty
                    });
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
                    Note = dates is not null ? "в файле есть строки, но без даты прибытия" : string.Empty
                });
                continue;
            }

            if (arrived == 0)
            {
                // Факт отмечен, но файла прихода заказ не касается — не ошибка
                // (номер мог не попасть в файл), но человеку показать стоит.
                var staleFact = PurchaseFormatting.ParseDate(candidate.Supplier.ActualReceiptDate);
                if (staleFact is not null)
                {
                    report.Lines.Add("⚠ " + candidate.Po + " (" + candidate.Supplier.Supplier +
                                     "): «Факт поступл.» " + staleFact.Value.ToString("dd.MM.yyyy") +
                                     ", но заказа нет в файле прихода — перепроверьте вручную");
                }

                report.Issues.Add(new WarehouseArrivalIssue
                {
                    DealId = candidate.Deal.Id, DealNumber = candidate.Deal.DealNumber,
                    SupplierId = candidate.Supplier.Id, Supplier = candidate.Supplier.Supplier,
                    PoNumber = candidate.Po, Expected = expected, Arrived = 0,
                    Deadline = ParseDeadline(candidate.Supplier.ComponentsReceiptDate),
                    Status = "missing",
                    Note = dates is not null ? "в файле есть строки, но без даты прибытия" : string.Empty
                });
                continue;
            }

            // Частичный приход (или сверка невозможна без ERP) — вопрос,
            // показывается при каждой проверке, пока не разрешится.
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

            report.Lines.Insert(0, "Заказов в работе: " + candidates.Count +
                                    ", с замечаниями: " + report.Issues.Count +
                                    ", «Факт поступл.» записан/исправлен: " + report.FilledFacts);
        }
        finally
        {
            CloseExcelQuietly(excel);
        }

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
    /// Разбор книги прихода. Две раскладки:
    /// 1) листы «Orders' materials onlly (…)» (CE/CT/AE/AT): номер заказа в
    ///    кол. 2 только в первой строке группы (объединённые ячейки, протягиваем),
    ///    дата прибытия в кол. 14;
    /// 2) сводные журналы CE&amp;WE, CT, PE, Direct payment: номер в кол. «PO.»,
    ///    дата — «Arrival date of goods», номер в КАЖДОЙ строке. Колонки ищем по
    ///    английской шапке — они сдвигаются между листами. Читаем ограниченно
    ///    по колонкам (A:R): UsedRange сводных листов раздут форматированием
    ///    до 16 тыс. колонок, полный Value2 раздувает память.
    /// Ключ словаря — текст ячейки с номером как есть (матчинг — подстрокой).
    /// </summary>
    internal static Dictionary<string, List<DateTime?>> ParseArrivalWorkbook(dynamic excel, string path)
    {
        var result = new Dictionary<string, List<DateTime?>>(StringComparer.Ordinal);
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

                // 1) Четыре листа-«выгрузки» (CE/CT/AE/AT) с группами по номеру.
                // 2) Сводные журналы (CE&WE и т.п.) — там заказы, которых
                //    нет на листах-выгрузках (находка 2026-10-01).
                if (name.IndexOf("materials onlly", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ReadArrivalSheet(ws, result);
                }
                else
                {
                    ReadFlatArrivalSheet(ws, result);
                }
            }
        }
        finally
        {
            workbook.Close(false);
            RrfqExcel.Release(workbook);
            GC.Collect();
            GC.WaitForPendingFinalizers();
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
            var cell = NormalizePo(ToText(values[row, OrderColumn - 1]));
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

    /// <summary>
    /// Сводные журналы (CE&amp;WE, CT, PE, Direct payment): одна строка = одна
    /// позиция, номер заказа в колонке «PO.» в каждой строке. Колонки номера и
    /// даты прибытия ищем по английской шапке («PO.», «Arrival date of goods»,
    /// но не «VAT arrival date») — между листами они сдвигаются. Протягивание
    /// номера не делаем: без своего номера строка не считается.
    /// </summary>
    private static void ReadFlatArrivalSheet(dynamic ws, Dictionary<string, List<DateTime?>> result)
    {
        int lastRow;
        try
        {
            lastRow = (int)ws.UsedRange.Rows.Count;
        }
        catch
        {
            return;
        }

        if (lastRow < 10)
        {
            return;
        }

        // (object) разрывает динамическую диспетчеризацию — как в ReadArrivalSheet.
        // Ограничение A:R: нужные колонки в пределах 11-й, а UsedRange сводных
        // листов раздут форматированием до 16 тыс. колонок.
        var values = ToMatrix((object)ws.Range("A1:R" + lastRow).Value2);
        if (values is null)
        {
            return;
        }

        var rows = values.GetLength(0);
        var cols = values.GetLength(1);

        // Шапка: строка с «PO.» и «Arrival date of goods» в первых 15 строках.
        var headerRow = -1;
        var poCol = -1;
        var arrivalCol = -1;
        for (var row = 0; row < Math.Min(rows, 15) && headerRow < 0; row++)
        {
            for (var col = 0; col < cols; col++)
            {
                if (!ToText(values[row, col]).Equals("PO.", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                for (var dateCol = 0; dateCol < cols; dateCol++)
                {
                    var header = ToText(values[row, dateCol]);
                    if (header.IndexOf("arrival date", StringComparison.OrdinalIgnoreCase) >= 0
                        && header.IndexOf("vat", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        headerRow = row;
                        poCol = col;
                        arrivalCol = dateCol;
                        break;
                    }
                }
                if (headerRow >= 0)
                {
                    break;
                }
            }
        }

        if (headerRow < 0 || arrivalCol == poCol)
        {
            return; // не наш лист — молча пропускаем
        }

        for (var row = headerRow + 1; row < rows; row++)
        {
            var po = NormalizePo(ToText(values[row, poCol]));
            if (po.Length == 0)
            {
                continue;
            }

            if (!result.TryGetValue(po, out var group))
            {
                group = new List<DateTime?>();
                result[po] = group;
            }
            group.Add(ToDate(values[row, arrivalCol]));
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
    /// Приведение похожих кириллических букв к латинским: номера заказов
    /// печатаются то латиницей, то кириллицей («-С4» с кириллической «С»),
    /// и без нормализации матчинг их не сводит (находка 2026-10-01).
    /// Применяется и к номерам из комментариев, и к ячейкам файла прихода.
    /// </summary>
    private static readonly Dictionary<char, char> CyrillicToLatin = new()
    {
        ['А'] = 'A', ['а'] = 'A',
        ['В'] = 'B', ['в'] = 'B',
        ['С'] = 'C', ['с'] = 'C',
        ['Е'] = 'E', ['е'] = 'E',
        ['Н'] = 'H', ['н'] = 'H',
        ['К'] = 'K', ['к'] = 'K',
        ['М'] = 'M', ['м'] = 'M',
        ['О'] = 'O', ['о'] = 'O',
        ['Р'] = 'P', ['р'] = 'P',
        ['Т'] = 'T', ['т'] = 'T',
        ['Х'] = 'X', ['х'] = 'X',
        ['У'] = 'Y', ['у'] = 'Y',
    };

    internal static string NormalizePo(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        var builder = new System.Text.StringBuilder(text.Length);
        foreach (var ch in text)
        {
            builder.Append(CyrillicToLatin.TryGetValue(ch, out var latin) ? latin : ch);
        }
        return builder.ToString();
    }

    /// <summary>
    /// Номер заказа из комментария поставщика: «&lt;номер сделки&gt;-&lt;суффикс&gt; / текст»,
    /// короткая форма «-2», или просто номер сделки (без суффикса). Суффиксы:
    /// «1…5», «C1…C8». Комментарий нормализуется (кириллица → латиница).
    /// </summary>
    internal static string? ExtractPoNumber(string dealNumber, string? comment)
    {
        if (string.IsNullOrWhiteSpace(comment) || dealNumber.Length == 0)
        {
            return null;
        }

        var number = NormalizePo(dealNumber.Trim());
        var text = NormalizePo(comment.Trim());
        var suffix = @"(\d{1,3}|[A-Za-z]\d{1,3})(?![\dA-Za-z])";

        var full = Regex.Match(text, Regex.Escape(number) + @"\s*-" + suffix,
            RegexOptions.IgnoreCase);
        if (full.Success)
        {
            return number + "-" + full.Groups[1].Value;
        }

        var shortForm = Regex.Match(text, @"^-\s*" + suffix, RegexOptions.IgnoreCase);
        if (shortForm.Success)
        {
            return number + "-" + shortForm.Groups[1].Value;
        }

        if (text.StartsWith(number, StringComparison.OrdinalIgnoreCase))
        {
            return number;
        }

        return null;
    }

    /// <summary>
    /// Даты строк заказа: собираем ВСЕ группы, чей номер содержит искомый —
    /// у склада встречаются варианты с префиксами («2-AE…», «C-AE…»), это те
    /// же поставки, и их строки не должны теряться. Суффиксные группы
    /// («AE…-002968-5») отсекаются граничным правилом ContainsToken.
    /// </summary>
    private static List<DateTime?>? FindGroupDates(Dictionary<string, List<DateTime?>> arrivals, string po)
    {
        List<DateTime?>? merged = arrivals.TryGetValue(po, out var exact)
            ? new List<DateTime?>(exact)
            : null;
        foreach (var (cell, dates) in arrivals)
        {
            if (cell.Equals(po, StringComparison.Ordinal) || !ContainsToken(cell, po))
            {
                continue;
            }

            merged ??= new List<DateTime?>();
            merged.AddRange(dates);
        }
        return merged;
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

    /// <summary>
    /// Ожидаемые позиции сделки: ERP-файлы (тип «ERP») → PI Number → число строк
    /// позиций. Файлы с одним номером — перезагрузки/переэкспорты одного заказа,
    /// поэтому берём максимум по файлам, а не сумму (иначе «BMZ… (2)» удваивал
    /// заказ: 6+6=12 — находка 2026-10-01).
    /// </summary>
    private static Dictionary<string, int> LoadErpExpectations(dynamic excel, long dealId)
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
                // Файл могли переименовать на диске уже после загрузки
                // («…-С4.xlsx» кириллицей → «…-C4.xlsx»): пробуем нормализованный
                // вариант пути (находка 2026-10-02).
                var normalized = PurchaseWriteRepository.ResolvePortable(NormalizePo(document.StoredPath));
                if (!string.IsNullOrWhiteSpace(normalized) && File.Exists(normalized))
                {
                    path = normalized;
                }
            }

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                continue;
            }

            try
            {
                // (object) разрывает динамическую диспетчеризацию: иначе pair
                // становится dynamic и расширения словаря не компилируются.
                foreach (var pair in ReadErpFile((object)excel, path))
                {
                    if (pair.Value > expected.GetValueOrDefault(pair.Key))
                    {
                        expected[pair.Key] = pair.Value;
                    }
                }
            }
            catch
            {
                // Один нечитаемый ERP-файл не должен ронять всю проверку.
            }
        }
        return expected;
    }

    /// <summary>
    /// Чтение одного ERP-файла: PI Number → число строк позиций. Колонки
    /// «ERP-Code» и «PI Number» ищем по шапке — форматы две: старый
    /// («CPR-…_ERPCodes.xls», шапка в 3-й строке, PI в кол. 5) и новый
    /// («ERP_Codes_….xlsx», шапка в 1-й строке, PI в кол. 4). PI Number
    /// протягивается вниз на случай объединённых ячеек.
    /// </summary>
    private static Dictionary<string, int> ReadErpFile(dynamic excel, string path)
    {
        var expected = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        dynamic workbook = RrfqExcel.OpenWorkbookReadOnly(excel, path);
        try
        {
            var sheetCount = (int)workbook.Worksheets.Count;
            for (var index = 1; index <= sheetCount; index++)
            {
                var values = ToMatrix((object)workbook.Worksheets.Item(index).UsedRange.Value2);
                if (values is null || values.GetLength(1) < 2)
                {
                    continue;
                }

                var rows = values.GetLength(0);
                var cols = values.GetLength(1);

                // Шапка: строка, где рядом есть «ERP-Code» и «PI Number».
                var headerRow = -1;
                var codeCol = -1;
                var piCol = -1;
                for (var row = 0; row < Math.Min(rows, 15) && headerRow < 0; row++)
                {
                    var code = -1;
                    var pi = -1;
                    for (var col = 0; col < cols; col++)
                    {
                        var header = NormalizePo(ToText(values[row, col]));
                        if (header.Equals("ERP-Code", StringComparison.OrdinalIgnoreCase))
                        {
                            code = col;
                        }
                        else if (header.Equals("PI Number", StringComparison.OrdinalIgnoreCase))
                        {
                            pi = col;
                        }
                    }

                    if (code >= 0 && pi >= 0)
                    {
                        headerRow = row;
                        codeCol = code;
                        piCol = pi;
                    }
                }

                if (headerRow < 0)
                {
                    continue; // не ERP-раскладка — лист пропускаем
                }

                string? currentPi = null;
                for (var row = headerRow + 1; row < rows; row++)
                {
                    var pi = NormalizePo(ToText(values[row, piCol]));
                    if (pi.Length > 0)
                    {
                        currentPi = pi;
                    }

                    var erpCode = ToText(values[row, codeCol]);
                    if (currentPi is not null && erpCode.Length > 0)
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
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        return expected;
    }

    /// <summary>
    /// Закрытие Excel с гарантированным гашением процесса: до Quit собираем
    /// мусор, чтобы «хвостовые» COM-обёртки (Worksheets, Range из dynamic)
    /// не держали процесс живым; после Quit — контрольные сборки. Иначе
    /// скрытые EXCEL.EXE с открытыми книгами копятся, и Windows при
    /// выключении жалуется на «незакрытые файлы» (случай 2026-10-02).
    /// </summary>
    private static void CloseExcelQuietly(dynamic? excel)
    {
        if (excel is null)
        {
            return;
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
        try
        {
            excel.Quit();
        }
        catch
        {
            // Quit может не удаться на занятом процессе — не критично.
        }
        RrfqExcel.Release(excel);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
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
                ? "По файлу прихода: " + (issue.Status == "nodate"
                    ? "позиции числятся прибывшими (" + issue.Arrived + " из " + issue.Expected
                      + " строк), но даты прибытия не проставлены"
                    : "на складе пока " + issue.Arrived + " из " + issue.Expected + " позиций")
                  + (issue.ArrivalDate is not null
                    ? " (последний приход " + issue.ArrivalDate.Value.ToString("dd.MM.yyyy") + ")"
                    : string.Empty) + ".\n"
                : "Заказа пока нет в ежедневном файле прихода на склад.\n")
            + "Плановая дата поступления: " + issue.DeadlineText + overdue + ".\n\n"
            + "Просьба уточнить статус и подтвердить срок.\n\nСпасибо!";

        OutlookService.CreateMailDraft(email, subject, body);
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
