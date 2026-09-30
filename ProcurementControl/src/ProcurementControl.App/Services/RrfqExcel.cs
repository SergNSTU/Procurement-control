using System.Runtime.InteropServices;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Общие помощники Excel-автоматизации раздела Сравнение RRFQ — перенос
/// Open-ExcelApp / Close-ExcelApp / New-SheetData и чтения ячеек из
/// app/modules/RrfqEngine.ps1. Работает через dynamic/COM, как оригинал.
/// Каждый вызов открывает собственный экземпляр Excel и гарантированно
/// закрывает его, чтобы не оставлять процессов.
/// </summary>
public static class RrfqExcel
{
    /// <summary>Аналог Open-ExcelApp: скрытый Excel без макросов и алертов.</summary>
    public static dynamic OpenExcel()
    {
        var type = Type.GetTypeFromProgID("Excel.Application");
        if (type is null)
        {
            throw new InvalidOperationException("Excel не установлен на этом компьютере — сравнение недоступно.");
        }

        var instance = Activator.CreateInstance(type);
        if (instance is null)
        {
            throw new InvalidOperationException("Не удалось запустить Excel для сравнения.");
        }

        dynamic excel = instance;
        excel.Visible = false;
        excel.DisplayAlerts = false;
        excel.EnableEvents = false;
        excel.ScreenUpdating = false;
        excel.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
        return excel;
    }

    /// <summary>Аналог Close-ExcelApp: Quit + Release + сборка мусора.</summary>
    public static void CloseExcel(dynamic? excel)
    {
        if (excel is null)
        {
            return;
        }

        try
        {
            excel.Quit();
        }
        catch
        {
            // Excel закроется сам при выходе.
        }
        finally
        {
            Release(excel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>Аналог Release-ComObject.</summary>
    public static void Release(object? comObject)
    {
        if (comObject is null)
        {
            return;
        }

        try
        {
            if (Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch
        {
            // Освобождение не критично для результата.
        }
    }

    /// <summary>Открывает книгу только для чтения (аналог $excel.Workbooks.Open($Path, 0, $true)).</summary>
    public static dynamic OpenWorkbookReadOnly(dynamic excel, string path)
        => excel.Workbooks.Open(path, 0, true);

    /// <summary>
    /// Аналог New-SheetData: слепок UsedRange.Value2 плюс скрытые строки и столбцы.
    /// </summary>
    public static SheetData ReadSheetData(dynamic ws)
    {
        var used = ws.UsedRange;
        try
        {
            var firstRow = (int)used.Row;
            var firstCol = (int)used.Column;
            var rowCount = (int)used.Rows.Count;
            var colCount = (int)used.Columns.Count;
            var lastRow = firstRow + rowCount - 1;
            var lastCol = firstCol + colCount - 1;

            var hiddenRows = new HashSet<int>();
            for (var row = firstRow; row <= lastRow; row++)
            {
                var rowRange = ws.Rows.Item(row);
                try
                {
                    if ((bool)rowRange.Hidden)
                    {
                        hiddenRows.Add(row);
                    }
                }
                finally
                {
                    Release(rowRange);
                }
            }

            var hiddenCols = new HashSet<int>();
            for (var col = firstCol; col <= lastCol; col++)
            {
                var colRange = ws.Columns.Item(col);
                try
                {
                    if ((bool)colRange.Hidden)
                    {
                        hiddenCols.Add(col);
                    }
                }
                finally
                {
                    Release(colRange);
                }
            }

            var values = used.Value2;
            object?[,]? array = null;
            object? scalar = null;
            if (values is not null)
            {
                if (values is object?[,] direct)
                {
                    array = NormalizeArray(direct);
                }
                else if (values is object[] singleRow)
                {
                    // Редкий случай одномерного среза — приводим к матрице.
                    array = new object?[1, singleRow.Length];
                    for (var i = 0; i < singleRow.Length; i++)
                    {
                        array[0, i] = singleRow[i];
                    }
                }
                else
                {
                    scalar = values;
                }
            }

            return new SheetData
            {
                Name = (string)ws.Name,
                Index = (int)ws.Index,
                FirstRow = firstRow,
                FirstCol = firstCol,
                LastRow = lastRow,
                LastCol = lastCol,
                Values = array,
                ScalarValue = scalar,
                HiddenRows = hiddenRows,
                HiddenCols = hiddenCols,
            };
        }
        finally
        {
            Release(used);
        }
    }

    /// <summary>Переводит 1-базовый COM-массив в 0-базовый с тем же содержимым.</summary>
    private static object?[,] NormalizeArray(object?[,] source)
    {
        var rows = source.GetLength(0);
        var cols = source.GetLength(1);
        var rowBase = source.GetLowerBound(0);
        var colBase = source.GetLowerBound(1);
        var result = new object?[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                result[r, c] = source.GetValue(rowBase + r, colBase + c);
            }
        }

        return result;
    }
}
