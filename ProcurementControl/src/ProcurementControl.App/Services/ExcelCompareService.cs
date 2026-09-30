using System.IO;

namespace ProcurementControl.Services;

/// <summary>
/// Порт Compare-ExcelWorkbooks (RRFQComparer.ps1, строки 4704-4768):
/// сравнивает два файла ячейка за ячейкой и красит несовпадения красным
/// в копии первого файла. Работает через Excel COM, как оригинал
/// (помощники из <see cref="RrfqExcel"/>).
/// </summary>
public static class ExcelCompareService
{
    private const int DiffFillColor = 255; // Красный, как Interior.Color оригинала.
    private const int XlsxFileFormat = 51; // xlOpenXMLWorkbook.

    /// <summary>Аналог пути результата: {данные}/excel_compare_{дата}.xlsx.</summary>
    public static string NewOutputPath()
        => Path.Combine(
            AppPaths.PurchaseDataDirectory,
            "excel_compare_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

    /// <summary>Возвращает число несовпадений; результат сохраняется в outputPath.</summary>
    public static int CompareWorkbooks(string firstPath, string secondPath, string outputPath)
    {
        if (!File.Exists(firstPath))
        {
            throw new InvalidOperationException("Первый Excel-файл не найден.");
        }
        if (!File.Exists(secondPath))
        {
            throw new InvalidOperationException("Второй Excel-файл не найден.");
        }

        var excel = RrfqExcel.OpenExcel();
        dynamic? first = null;
        dynamic? second = null;
        dynamic? result = null;
        var diffCount = 0;
        try
        {
            first = RrfqExcel.OpenWorkbookReadOnly(excel, firstPath);
            second = RrfqExcel.OpenWorkbookReadOnly(excel, secondPath);

            Directory.CreateDirectory(AppPaths.PurchaseDataDirectory);
            File.Copy(firstPath, outputPath, overwrite: true);
            result = excel.Workbooks.Open(outputPath, 0, false);

            foreach (dynamic sheet in result.Worksheets)
            {
                dynamic? other = null;
                try
                {
                    other = second.Worksheets.Item((string)sheet.Name);
                }
                catch
                {
                    other = null;
                }

                var used = sheet.UsedRange;
                try
                {
                    var firstRow = (int)used.Row;
                    var firstCol = (int)used.Column;
                    var firstLastRow = firstRow + (int)used.Rows.Count - 1;
                    var firstLastCol = firstCol + (int)used.Columns.Count - 1;

                    if (other is null)
                    {
                        // Листа нет во втором файле — красим весь диапазон.
                        sheet.Range(
                            sheet.Cells.Item(firstRow, firstCol),
                            sheet.Cells.Item(firstLastRow, firstLastCol)).Interior.Color = DiffFillColor;
                        diffCount += (int)used.Rows.Count * (int)used.Columns.Count;
                        continue;
                    }

                    diffCount += CompareSheets(sheet, used, other, firstRow, firstCol, firstLastRow, firstLastCol);
                }
                finally
                {
                    RrfqExcel.Release(used);
                    RrfqExcel.Release(other);
                }
            }

            result.SaveAs(outputPath, XlsxFileFormat);
            return diffCount;
        }
        finally
        {
            CloseWorkbook(result);
            CloseWorkbook(second);
            CloseWorkbook(first);
            if (excel is not null)
            {
                excel.ScreenUpdating = true;
            }
            RrfqExcel.CloseExcel(excel);
        }
    }

    /// <summary>Ячейка за ячейкой, как циклы оригинала (строки 4745-4759).</summary>
    private static int CompareSheets(
        dynamic sheet,
        dynamic used,
        dynamic other,
        int firstRow,
        int firstCol,
        int firstLastRow,
        int firstLastCol)
    {
        var otherUsed = other.UsedRange;
        try
        {
            var secondRow = (int)otherUsed.Row;
            var secondCol = (int)otherUsed.Column;
            var secondLastRow = secondRow + (int)otherUsed.Rows.Count - 1;
            var secondLastCol = secondCol + (int)otherUsed.Columns.Count - 1;
            var maxRow = Math.Max(firstLastRow, secondLastRow);
            var maxCol = Math.Max(firstLastCol, secondLastCol);

            object? firstValues = used.Value2;
            object? secondValues = otherUsed.Value2;

            var diffCount = 0;
            for (var r = Math.Min(firstRow, secondRow); r <= maxRow; r++)
            {
                for (var c = Math.Min(firstCol, secondCol); c <= maxCol; c++)
                {
                    var aText = CellText(firstValues, r, c, firstRow, firstCol, firstLastRow, firstLastCol);
                    var bText = CellText(secondValues, r, c, secondRow, secondCol, secondLastRow, secondLastCol);
                    if (aText != bText)
                    {
                        sheet.Cells.Item(r, c).Interior.Color = DiffFillColor;
                        diffCount++;
                    }
                }
            }
            return diffCount;
        }
        finally
        {
            RrfqExcel.Release(otherUsed);
        }
    }

    /// <summary>Аналог выбора $a/$b из Value2: массив, скаляр или пусто; текст с обрезкой.</summary>
    private static string CellText(
        object? values,
        int row,
        int col,
        int rangeFirstRow,
        int rangeFirstCol,
        int rangeLastRow,
        int rangeLastCol)
    {
        object? value = null;
        if (row >= rangeFirstRow && row <= rangeLastRow && col >= rangeFirstCol && col <= rangeLastCol)
        {
            if (values is Array array)
            {
                value = array.GetValue(row - rangeFirstRow + 1, col - rangeFirstCol + 1);
            }
            else if (row == rangeFirstRow && col == rangeFirstCol)
            {
                value = values;
            }
        }

        return value is null ? string.Empty : Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)!.Trim();
    }

    private static void CloseWorkbook(dynamic? workbook)
    {
        if (workbook is null)
        {
            return;
        }

        try
        {
            workbook.Close(false);
        }
        catch
        {
            // Книга закроется вместе с Excel.
        }
        finally
        {
            RrfqExcel.Release(workbook);
        }
    }
}
