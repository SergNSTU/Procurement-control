using System.Globalization;
using System.IO;

namespace ProcurementControl.Services;

/// <summary>
/// Одна разобранная строка книги Globalist — аналог записи, которую собирает
/// Import-GlobalistWorkbook (app/modules/QuoteHistory.ps1, строки 135-143).
/// </summary>
public sealed class GlobalistImportRecord
{
    public string ImportedAt { get; init; } = string.Empty;
    public string SourceFile { get; init; } = string.Empty;
    public string SheetName { get; init; } = string.Empty;
    public int RowNumber { get; init; }
    public string Factory { get; init; } = string.Empty;
    public string Pn { get; init; } = string.Empty;
    public string Comment { get; init; } = string.Empty;
    public string PiNumber { get; init; } = string.Empty;
    public string Replacement { get; init; } = string.Empty;
    public string ChineseRemark { get; init; } = string.Empty;
    public string Package { get; init; } = string.Empty;
    public string Brand { get; init; } = string.Empty;
    public string Datacode { get; init; } = string.Empty;
    public string Moq { get; init; } = string.Empty;
    public string Qty { get; init; } = string.Empty;
    public string Stock { get; init; } = string.Empty;
    public string NeedSpq { get; init; } = string.Empty;
    public string Spq { get; init; } = string.Empty;
    public double? UnitPrice { get; init; }
    public double? TotalAmount { get; init; }
    public string LeadTime { get; init; } = string.Empty;
    public string Weight { get; init; } = string.Empty;
    public string Target { get; init; } = string.Empty;
    public string SupplierQuoteId { get; init; } = string.Empty;
}

/// <summary>
/// Порт Import-GlobalistWorkbook (app/modules/QuoteHistory.ps1, строки 96-162):
/// чтение книги через Excel-автоматизацию, поиск строки заголовков с колонками
/// PN и PRICE в первых 25 строках, разбор строк и полная замена таблицы
/// globalist_quotes. Запись идёт одной сессией (см. PurchaseWriteRepository.ReplaceGlobalistRows).
/// </summary>
public static class GlobalistImporter
{
    /// <summary>Импортирует книгу и возвращает число записанных строк.</summary>
    public static int Import(string path)
    {
        if (!File.Exists(path))
        {
            throw new InvalidOperationException("Файл Globalist не найден: " + path);
        }

        var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var records = new List<GlobalistImportRecord>();

        dynamic excel = RrfqExcel.OpenExcel();
        dynamic? workbook = null;
        try
        {
            workbook = RrfqExcel.OpenWorkbookReadOnly(excel, path);
            foreach (dynamic ws in workbook.Worksheets)
            {
                try
                {
                    ParseSheet(ws, path, now, records);
                }
                finally
                {
                    RrfqExcel.Release(ws);
                }
            }
        }
        finally
        {
            if (workbook is not null)
            {
                try
                {
                    workbook.Close(false);
                }
                catch
                {
                    // Книга закроется вместе с Excel.
                }
                RrfqExcel.Release(workbook);
            }

            RrfqExcel.CloseExcel(excel);
        }

        PurchaseWriteRepository.ReplaceGlobalistRows(records);
        return records.Count;
    }

    private static void ParseSheet(dynamic ws, string sourceFile, string importedAt, List<GlobalistImportRecord> records)
    {
        var sheet = RrfqExcel.ReadSheetData(ws);

        // Строка заголовков: ищем среди первых 25 строк строку с колонками PN и PRICE.
        var headerRow = 0;
        var columns = new Dictionary<string, int>();
        var scanEnd = Math.Min(sheet.LastRow, sheet.FirstRow + 24);
        for (var row = sheet.FirstRow; row <= scanEnd; row++)
        {
            var candidate = new Dictionary<string, int>();
            for (var col = sheet.FirstCol; col <= sheet.LastCol; col++)
            {
                var header = RrfqEngine.NormalizeKey(GetCell(sheet, row, col));
                if (!string.IsNullOrWhiteSpace(header))
                {
                    candidate[header] = col;
                }
            }

            if (candidate.ContainsKey("PN") && candidate.ContainsKey("PRICE"))
            {
                headerRow = row;
                columns = candidate;
                break;
            }
        }

        if (headerRow == 0)
        {
            return;
        }

        for (var row = headerRow + 1; row <= sheet.LastRow; row++)
        {
            var pn = Get(sheet, row, columns, "PN").Trim();
            if (string.IsNullOrWhiteSpace(pn))
            {
                continue;
            }

            records.Add(new GlobalistImportRecord
            {
                ImportedAt = importedAt,
                SourceFile = sourceFile,
                SheetName = sheet.Name,
                RowNumber = row,
                Factory = Get(sheet, row, columns, "FACTORY"),
                Pn = pn,
                Comment = Get(sheet, row, columns, "COMMENT"),
                PiNumber = Get(sheet, row, columns, "PI NUMBER"),
                Replacement = Get(sheet, row, columns, "REPLACEMENT"),
                ChineseRemark = Get(sheet, row, columns, "CHINESE REMARK"),
                Package = Get(sheet, row, columns, "PACKAGE"),
                Brand = Get(sheet, row, columns, "BRAND"),
                Datacode = Get(sheet, row, columns, "DATACODE (DC)"),
                Moq = Get(sheet, row, columns, "MINIMUM ORDER QUANTITY (MOQ)"),
                Qty = Get(sheet, row, columns, "QNTY"),
                Stock = Get(sheet, row, columns, "QTY ON STOCK"),
                NeedSpq = Get(sheet, row, columns, "NEED SPQ"),
                Spq = Get(sheet, row, columns, "SPQ"),
                UnitPrice = ConvertNumber(Get(sheet, row, columns, "PRICE")),
                TotalAmount = ConvertNumber(Get(sheet, row, columns, "TOTAL AMOUNT")),
                LeadTime = Get(sheet, row, columns, "LEAD TIME (LT), WKS"),
                Weight = Get(sheet, row, columns, "WEIGHT, G"),
                Target = Get(sheet, row, columns, "TARGET"),
                SupplierQuoteId = Get(sheet, row, columns, "SUPPLIERS QUOTE ID"),
            });
        }
    }

    private static string Get(
        ProcurementControl.Models.SheetData sheet,
        int row,
        Dictionary<string, int> columns,
        string name)
        => columns.TryGetValue(name, out var col) ? GetCell(sheet, row, col).Trim() : string.Empty;

    private static string GetCell(ProcurementControl.Models.SheetData sheet, int row, int col)
    {
        if (row < sheet.FirstRow || row > sheet.LastRow || col < sheet.FirstCol || col > sheet.LastCol)
        {
            return string.Empty;
        }

        if (sheet.Values is null)
        {
            return sheet.ScalarValue is null ? string.Empty : Convert.ToString(sheet.ScalarValue, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        var value = sheet.Values[row - sheet.FirstRow, col - sheet.FirstCol];
        return value is null ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    }

    /// <summary>Порт Convert-GlobalistNumber (QuoteHistory.ps1, строки 85-94).</summary>
    private static double? ConvertNumber(string value)
    {
        var text = value.Replace('\u00A0', ' ').Replace(" ", string.Empty).Trim();
        if (text.Length == 0)
        {
            return null;
        }

        text = text.Replace(',', '.');
        return double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var number) ? number : null;
    }
}
