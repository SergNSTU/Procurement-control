using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using CommunityToolkit.Mvvm.ComponentModel;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Строка исходного RFQ в окне подготовки. Included — «уходит поставщикам»
/// (предзаполняется детекцией зелёной заливки), остальные строки считаются
/// «для Компэла». Поля редактируются прямо в таблице подготовки.
/// </summary>
public partial class RfqSourceRow : ObservableObject
{
    public int SourceRow { get; init; }

    public string SourceSheet { get; init; } = string.Empty;

    /// <summary>Заливка строки определена как зелёная (нет — решение за пользователем).</summary>
    public bool DetectedGreen { get; init; }

    [ObservableProperty]
    private bool _included;

    [ObservableProperty]
    private string _value = string.Empty;

    [ObservableProperty]
    private string _russianRemark = string.Empty;

    [ObservableProperty]
    private string _chinaRemark = string.Empty;

    // PN и DC — явные свойства: генератор ObservableProperty превратил бы
    // _pn/_dc в Pn/Dc, а привязки и движок используют именно PN/DC.
    private string _pn = string.Empty;
    public string PN { get => _pn; set => SetProperty(ref _pn, value); }

    private string _dc = string.Empty;
    public string DC { get => _dc; set => SetProperty(ref _dc, value); }

    [ObservableProperty]
    private string _mfgRussia = string.Empty;

    [ObservableProperty]
    private string _mfgChina = string.Empty;

    [ObservableProperty]
    private string _qtyPacking = string.Empty;

    [ObservableProperty]
    private string _qtyToBuy = string.Empty;
}

/// <summary>Результат чтения исходника: строки + какие колонки распознаны.</summary>
public sealed class RfqSourceReadResult
{
    public List<RfqSourceRow> Rows { get; init; } = new();

    public int GreenCount => Rows.Count(r => r.DetectedGreen);

    public bool HasFillInfo { get; init; }
}

/// <summary>
/// Подготовка «поставщицкого» RFQ из исходной заявки: чтение строк с детекцией
/// зелёной заливки (позиции, уходящие поставщикам; остальные считаются в
/// Компэле) и сборка двух файлов в формате, который читает RrfqEngine:
/// RFQ для поставщиков и список «Для Компэла» из исключённых строк.
/// </summary>
public static class RfqPrepareService
{
    /// <summary>Колонки поставщицкого RFQ (заголовки — формат Standard движка RRFQ).</summary>
    public static readonly (string Header, string Field)[] SupplierColumns =
    {
        ("Value", "Value"),
        ("Russian remark", "RussianRemark"),
        ("China remark", "ChinaRemark"),
        ("PN", "PN"),
        ("D/C", "DC"),
        ("Mfg from Russia", "MfgRussia"),
        ("Mfg from China", "MfgChina"),
        ("Q-ty in packing", "QtyPacking"),
        ("Q-ty to Buy/pcs", "QtyToBuy"),
    };

    private static readonly Regex TotalRowRegex = new("^(TOTAL|ИТОГО)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Читает исходник: строки с непустым Value, зелёная заливка колонки Value
    /// помечает позиции поставщикам. Выполнять в STA-потоке (Excel COM).
    /// </summary>
    public static RfqSourceReadResult ReadSource(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("Файл исходного RFQ не найден: " + path, path);
        }

        dynamic? excel = null;
        dynamic? workbook = null;
        try
        {
            excel = RrfqExcel.OpenExcel();
            workbook = RrfqExcel.OpenWorkbookReadOnly(excel, path);
            var sheetCount = (int)workbook.Worksheets.Count;
            for (var i = 1; i <= sheetCount; i++)
            {
                dynamic ws = workbook.Worksheets.Item(i);
                try
                {
                    var sheet = RrfqExcel.ReadSheetData(ws);
                    var map = RrfqEngine.FindTableMap(sheet, "Standard", "Delivery");
                    if (map is null)
                    {
                        continue;
                    }

                    var startRow = map.HeaderRow + 1;
                    if (map.Type == "Standard" && startRow < 49)
                    {
                        startRow = 49;
                    }

                    var rows = ReadRowsFromSheet(sheet, map, i);
                    if (rows.Count == 0)
                    {
                        continue;
                    }

                    // Заливка читается отдельным проходом по диапазону колонки
                    // Value — одним запросом к COM на весь столбец.
                    var hasFill = TryReadGreenFlags(ws, map, startRow, sheet.LastRow, rows);

                    return new RfqSourceReadResult { Rows = rows, HasFillInfo = hasFill };
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
                try { workbook.Close(false); } catch (Exception ex) { ErrorLog.Append("excel", ex); }
                RrfqExcel.Release(workbook);
            }

            RrfqExcel.CloseExcel(excel);
        }

        throw new InvalidOperationException(
            "Не удалось распознать таблицу позиций в файле. Нужен формат RFQ с шапкой (Value / Russian remark / ...).");
    }

    private static List<RfqSourceRow> ReadRowsFromSheet(SheetData sheet, HeaderMap map, int sheetIndex)
    {
        var rows = new List<RfqSourceRow>();
        var startRow = map.HeaderRow + 1;
        if (map.Type == "Standard" && startRow < 49)
        {
            startRow = 49;
        }

        for (var row = startRow; row <= sheet.LastRow; row++)
        {
            if (sheet.IsRowHidden(row))
            {
                continue;
            }

            var value = RrfqEngine.GetFieldText(sheet, map, "Value", row);
            if (string.IsNullOrWhiteSpace(value) || TotalRowRegex.IsMatch(RrfqEngine.NormalizeKey(value)))
            {
                continue;
            }

            rows.Add(new RfqSourceRow
            {
                SourceRow = row,
                SourceSheet = sheet.Name,
                Value = value,
                RussianRemark = RrfqEngine.GetFieldText(sheet, map, "RussianRemark", row),
                ChinaRemark = RrfqEngine.GetFieldText(sheet, map, "ChinaRemark", row),
                PN = RrfqEngine.GetFieldText(sheet, map, "PN", row),
                DC = RrfqEngine.GetFieldText(sheet, map, "DC", row),
                MfgRussia = RrfqEngine.GetFieldText(sheet, map, "MfgRussia", row),
                MfgChina = RrfqEngine.GetFieldText(sheet, map, "MfgChina", row),
                QtyPacking = RrfqEngine.GetFieldText(sheet, map, "QtyPacking", row),
                QtyToBuy = RrfqEngine.GetFieldText(sheet, map, "QtyToBuy", row),
            });
        }

        return rows;
    }

    /// <summary>
    /// Читает заливку колонки Value одним диапазоном и помечает зелёные строки.
    /// Возвращает false, если ни одной зелёной не нашлось (разметки нет или
    /// условное форматирование, которое Interior.Color не видит).
    /// </summary>
    private static bool TryReadGreenFlags(dynamic ws, HeaderMap map, int startRow, int lastRow, List<RfqSourceRow> rows)
    {
        if (!map.Fields.TryGetValue("Value", out var valueColumn) || rows.Count == 0 || lastRow < startRow)
        {
            return false;
        }

        object? raw;
        try
        {
            var range = ws.Range(ws.Cells(startRow, valueColumn), ws.Cells(lastRow, valueColumn));
            raw = range.Interior.Color;
            RrfqExcel.Release((object)range);
        }
        catch (Exception ex)
        {
            ErrorLog.Append("excel", ex);
            return false;
        }

        Dictionary<int, double> colors;
        if (raw is double single)
        {
            colors = new Dictionary<int, double> { [startRow] = single };
        }
        else if (raw is object[,] array)
        {
            colors = new Dictionary<int, double>();
            for (var i = array.GetLowerBound(1); i <= array.GetUpperBound(1); i++)
            {
                if (array[array.GetLowerBound(0), i] is double color)
                {
                    colors[startRow + i - 1] = color;
                }
            }
        }
        else
        {
            return false;
        }

        var anyGreen = false;
        foreach (var rowItem in rows)
        {
            if (colors.TryGetValue(rowItem.SourceRow, out var color) && IsGreenish(color))
            {
                rowItem.Included = true;
                anyGreen = true;
            }
        }

        return anyGreen;
    }

    /// <summary>Excel-цвет (BGR): зелёный доминирует над красным и синим.</summary>
    private static bool IsGreenish(double oleColor)
    {
        var value = (long)oleColor & 0xFFFFFF;
        if (value == 0xFFFFFF || value == 0)
        {
            return false;
        }

        var red = value & 0xFF;
        var green = (value >> 8) & 0xFF;
        var blue = (value >> 16) & 0xFF;
        return green > red + 20 && green > blue + 20;
    }

    /// <summary>Имена файлов результата для сделки.</summary>
    public static string BuildSupplierFilePath(string directory, string dealNumber)
        => Path.Combine(directory, SanitizeName("RFQ_" + (string.IsNullOrWhiteSpace(dealNumber) ? "без-номера" : dealNumber)) + ".xlsx");

    public static string BuildCompelFilePath(string directory, string dealNumber)
        => Path.Combine(directory, SanitizeName("Для_Компэла_" + (string.IsNullOrWhiteSpace(dealNumber) ? "без-номера" : dealNumber)) + ".xlsx");

    public static string SanitizeName(string name)
        => Regex.Replace(name, @"[^\wА-Яа-я\-]+", "_").Trim('_');

    /// <summary>Собирает RFQ для поставщиков из включённых строк (формат Standard).</summary>
    public static string WriteSupplierFile(IReadOnlyList<RfqSourceRow> rows, string outputPath)
        => WriteWorkbook(rows.Where(r => r.Included).ToList(), outputPath, "RFQ");

    /// <summary>Собирает список «Для Компэла» из исключённых строк (тот же формат).</summary>
    public static string WriteCompelFile(IReadOnlyList<RfqSourceRow> rows, string outputPath)
        => WriteWorkbook(rows.Where(r => !r.Included).ToList(), outputPath, "Для Компэла");

    private static string WriteWorkbook(IReadOnlyList<RfqSourceRow> rows, string path, string sheetName)
    {
        var fullPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);

        dynamic? excel = null;
        dynamic? workbook = null;
        dynamic? sheet = null;
        try
        {
            excel = RrfqExcel.OpenExcel();
            workbook = excel.Workbooks.Add();
            sheet = workbook.Worksheets.Item(1);
            sheet.Name = SanitizeName(sheetName);

            for (var col = 0; col < SupplierColumns.Length; col++)
            {
                sheet.Cells[1, col + 1].Value2 = SupplierColumns[col].Header;
            }

            for (var index = 0; index < rows.Count; index++)
            {
                var rowItem = rows[index];
                for (var col = 0; col < SupplierColumns.Length; col++)
                {
                    var text = SupplierColumns[col].Field switch
                    {
                        "Value" => rowItem.Value,
                        "RussianRemark" => rowItem.RussianRemark,
                        "ChinaRemark" => rowItem.ChinaRemark,
                        "PN" => rowItem.PN,
                        "DC" => rowItem.DC,
                        "MfgRussia" => rowItem.MfgRussia,
                        "MfgChina" => rowItem.MfgChina,
                        "QtyPacking" => rowItem.QtyPacking,
                        "QtyToBuy" => rowItem.QtyToBuy,
                        _ => string.Empty,
                    };
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        sheet.Cells[index + 2, col + 1].Value2 = text;
                    }
                }
            }

            var lastColumnLetter = ((char)('A' + SupplierColumns.Length - 1)).ToString();
            sheet.Range("A1:" + lastColumnLetter + "1").Font.Bold = true;
            sheet.Range("A1:" + lastColumnLetter + "1").Interior.ColorIndex = 15;
            sheet.Range("A1:" + lastColumnLetter + Math.Max(1, rows.Count + 1)).AutoFilter();
            sheet.Columns.AutoFit();
            workbook.SaveAs(fullPath, 51);
            return fullPath;
        }
        finally
        {
            if (workbook is not null)
            {
                try { workbook.Close(false); } catch (Exception ex) { ErrorLog.Append("excel", ex); }
                RrfqExcel.Release(workbook);
            }

            RrfqExcel.Release(sheet);
            RrfqExcel.CloseExcel(excel);
        }
    }
}
