using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Выгрузка разбора Компэла в Excel — перенос Write-CompelWorkbook
/// (CompelParser.ps1, строки 361-509). Работает через Excel-автоматизацию
/// (dynamic/COM), как оригинал; новые NuGet-пакеты не требуются.
/// </summary>
public static class CompelExcelExporter
{
    /// <summary>Создаёт книгу «Сводка» + «Позиции» и возвращает полный путь.</summary>
    public static string WriteWorkbook(string path, IReadOnlyList<CompelRow> rows, CompelSummary summary)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Не задан файл выгрузки.", nameof(path));
        }
        if (summary is null)
        {
            throw new InvalidOperationException("Сначала разберите HTML-файл Компэла.");
        }

        var fullPath = Path.GetFullPath(path);
        var dir = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        dynamic? excel = null;
        dynamic? wb = null;
        dynamic? summaryWs = null;
        dynamic? detailWs = null;
        try
        {
            excel = OpenExcel();
            wb = excel.Workbooks.Add();
            while (wb.Worksheets.Count > 1)
            {
                var sheet = wb.Worksheets.Item(wb.Worksheets.Count);
                sheet.Delete();
                Marshal.ReleaseComObject(sheet);
            }

            summaryWs = wb.Worksheets.Item(1);
            summaryWs.Name = "Сводка";
            detailWs = wb.Worksheets.Add(Type.Missing, summaryWs);
            detailWs.Name = "Позиции";

            WriteSummarySheet(summaryWs, summary);
            WriteDetailSheet(detailWs, rows, summary);

            // SaveAs формат 51 = xlsx.
            wb.SaveAs(fullPath, 51);
            return fullPath;
        }
        finally
        {
            if (wb is not null)
            {
                try { wb.Close(false); } catch { /* Книга уже могла быть закрыта. */ }
            }
            Release(detailWs);
            Release(summaryWs);
            Release(wb);
            if (excel is not null)
            {
                try { excel.Quit(); } catch { /* Excel закроется сам при выходе. */ }
                Release(excel);
            }
        }
    }

    private static void WriteSummarySheet(dynamic ws, CompelSummary summary)
    {
        var titleColor = ToOle(31, 78, 121);
        var subColor = ToOle(217, 234, 247);
        var whiteColor = ToOle(255, 255, 255);

        SetCell(ws, 1, 1, "Расчет SDS Compel");
        var titleRange = ws.Range("A1:D1");
        titleRange.Merge();
        titleRange.Interior.Color = titleColor;
        titleRange.Font.Bold = true;
        titleRange.Font.Color = whiteColor;
        titleRange.Font.Size = 14;
        Release(titleRange);

        SetCell(ws, 3, 1, "Показатель");
        SetCell(ws, 3, 2, "USD");
        SetCell(ws, 3, 3, "Комментарий");
        var headRange = ws.Range("A3:D3");
        headRange.Interior.Color = subColor;
        headRange.Font.Bold = true;
        Release(headRange);

        var missingText = string.Join(", ", summary.MissingRows);
        SetCell(ws, 4, 1, "Итого строк с подобранной ценой");
        SetCell(ws, 4, 2, summary.SelectedTotalUsd);
        SetCell(ws, 4, 3, "Сумма по листу Позиции");
        SetCell(ws, 5, 1, "Итог сайта: оптимизировано по цене");
        SetCell(ws, 5, 2, summary.CheapTotalUsd);
        SetCell(ws, 5, 3, summary.CheapLeadTime);
        SetCell(ws, 6, 1, "Итог сайта: оптимизировано по сроку");
        SetCell(ws, 6, 2, summary.OptimalTotalUsd);
        SetCell(ws, 6, 3, summary.OptimalLeadTime);
        SetCell(ws, 7, 1, "Всего строк");
        SetCell(ws, 7, 2, summary.TotalRows);
        SetCell(ws, 8, 1, "Строк с ценой");
        SetCell(ws, 8, 2, summary.PricedRows);
        SetCell(ws, 9, 1, "Строк без подобранного предложения");
        SetCell(ws, 9, 2, summary.MissingRows.Count);
        SetCell(ws, 9, 3, missingText);
        SetCell(ws, 10, 1, "Самый долгий срок по позициям");
        SetCell(ws, 10, 2, summary.LongestLeadDays);
        SetCell(ws, 10, 3, summary.LongestLeadTime);

        for (var row = 4; row <= 10; row++)
        {
            var labelCell = ws.Cells.Item(row, 1);
            labelCell.Font.Bold = true;
            Release(labelCell);
            var valueCell = ws.Cells.Item(row, 2);
            SetNumberFormat(valueCell, "#,##0.00", "# ##0,00");
            Release(valueCell);
        }

        SetColumnWidth(ws, 1, 34);
        SetColumnWidth(ws, 2, 16);
        SetColumnWidth(ws, 3, 38);
        SetColumnWidth(ws, 4, 8);
        var bodyRange = ws.Range("A1:D10");
        bodyRange.WrapText = true;
        bodyRange.VerticalAlignment = -4160; // xlTop
        Release(bodyRange);
    }

    private static void WriteDetailSheet(dynamic ws, IReadOnlyList<CompelRow> rows, CompelSummary summary)
    {
        var titleColor = ToOle(31, 78, 121);
        var subColor = ToOle(217, 234, 247);
        var whiteColor = ToOle(255, 255, 255);

        var headers = new[] { "N", "Исходное наименование", "Подобранный товар", "Производитель", "Срок поставки", "Количество", "Цена за шт", "Валюта", "Сумма", "Валюта суммы" };
        for (var i = 0; i < headers.Length; i++)
        {
            SetCell(ws, 1, i + 1, headers[i]);
        }

        var rowIndex = 2;
        foreach (var row in rows)
        {
            SetCell(ws, rowIndex, 1, row.Number);
            SetCell(ws, rowIndex, 2, row.SourceName);
            SetCell(ws, rowIndex, 3, row.MatchedPart);
            SetCell(ws, rowIndex, 4, row.Manufacturer);
            SetCell(ws, rowIndex, 5, row.LeadTime);
            SetCell(ws, rowIndex, 6, row.Quantity);
            SetCell(ws, rowIndex, 7, row.UnitPrice);
            SetCell(ws, rowIndex, 8, row.Currency);
            SetCell(ws, rowIndex, 9, row.Total);
            SetCell(ws, rowIndex, 10, row.TotalCurrency);
            rowIndex++;
        }

        var totalRow = rows.Count + 2;
        SetCell(ws, totalRow, 2, "Итого");
        SetCell(ws, totalRow, 5, "Самый долгий срок");
        SetCell(ws, totalRow, 6, CompelParserService.GetMaxLeadTime(rows));
        SetCell(ws, totalRow, 9, summary.SelectedTotalUsd);
        SetCell(ws, totalRow, 10, "USD");

        var headRange = ws.Range("A1:J1");
        headRange.Interior.Color = titleColor;
        headRange.Font.Bold = true;
        headRange.Font.Color = whiteColor;
        Release(headRange);

        var totalRange = ws.Range($"A{totalRow}:J{totalRow}");
        totalRange.Interior.Color = subColor;
        totalRange.Font.Bold = true;
        Release(totalRange);

        SetColumnWidth(ws, 1, 6);
        SetColumnWidth(ws, 2, 28);
        SetColumnWidth(ws, 3, 32);
        SetColumnWidth(ws, 4, 18);
        SetColumnWidth(ws, 5, 13);
        SetColumnWidth(ws, 6, 13);
        SetColumnWidth(ws, 7, 13);
        SetColumnWidth(ws, 8, 10);
        SetColumnWidth(ws, 9, 14);
        SetColumnWidth(ws, 10, 13);
        SetColumnFormat(ws, 1, "0", "0");
        SetColumnFormat(ws, 6, "0", "0");
        SetColumnFormat(ws, 7, "0.00000", "0,00000");
        SetColumnFormat(ws, 9, "#,##0.00", "# ##0,00");

        var bodyRange = ws.Range($"A1:J{totalRow}");
        bodyRange.WrapText = true;
        bodyRange.VerticalAlignment = -4160; // xlTop
        Release(bodyRange);

        var filterLastRow = Math.Max(1, rows.Count + 1);
        var filterRange = ws.Range($"A1:J{filterLastRow}");
        filterRange.AutoFilter();
        Release(filterRange);

        ws.Activate();
        FreezeFirstRow(ws.Application);

        // Активируем сводку, как оригинал (строка 497).
        ws.Parent.Worksheets.Item("Сводка").Activate();
    }

    /// <summary>Заморозка первой строки активного листа через окно приложения.</summary>
    private static void FreezeFirstRow(dynamic excel)
    {
        try
        {
            excel.ActiveWindow.SplitRow = 1;
            excel.ActiveWindow.FreezePanes = true;
        }
        catch
        {
            // Заморозка не критична для выгрузки.
        }
    }

    /// <summary>Ширина колонки одним обращением к COM-объекту.</summary>
    private static void SetColumnWidth(dynamic ws, int index, double width)
    {
        var column = ws.Columns.Item(index);
        try
        {
            column.ColumnWidth = width;
        }
        finally
        {
            Release(column);
        }
    }

    /// <summary>Формат колонки одним обращением к COM-объекту.</summary>
    private static void SetColumnFormat(dynamic ws, int index, string format, string localFormat)
    {
        var column = ws.Columns.Item(index);
        try
        {
            SetNumberFormat(column, format, localFormat);
        }
        finally
        {
            Release(column);
        }
    }

    private static dynamic OpenExcel()
    {
        var type = Type.GetTypeFromProgID("Excel.Application");
        if (type is null)
        {
            throw new InvalidOperationException("Excel не установлен на этом компьютере — выгрузка недоступна.");
        }
        var instance = Activator.CreateInstance(type);
        if (instance is null)
        {
            throw new InvalidOperationException("Не удалось запустить Excel для выгрузки.");
        }
        return instance;
    }

    /// <summary>Аналог Set-CompelCellValue: числа пишутся числами, остальное строкой, пустые пропускаются.</summary>
    private static void SetCell(dynamic ws, int row, int column, object? value)
    {
        if (value is null)
        {
            return;
        }
        var cell = ws.Cells.Item(row, column);
        try
        {
            cell.Value2 = ToCellValue(value);
        }
        finally
        {
            Release(cell);
        }
    }

    /// <summary>Числа — числами, остальное — строкой в invariant-культуре.</summary>
    private static object ToCellValue(object value)
    {
        return value switch
        {
            int i => (double)i,
            long l => (double)l,
            double d => d,
            float f => (double)f,
            _ => Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
        };
    }

    /// <summary>Аналог Set-CompelNumberFormat с запасным локальным форматом.</summary>
    private static void SetNumberFormat(dynamic range, string format, string localFormat)
    {
        if (range is null || string.IsNullOrWhiteSpace(format))
        {
            return;
        }
        try
        {
            range.NumberFormat = format;
        }
        catch
        {
            if (!string.IsNullOrWhiteSpace(localFormat))
            {
                try { range.NumberFormatLocal = localFormat; } catch { /* Формат не критичен. */ }
            }
        }
    }

    /// <summary>Аналог ColorTranslator.ToOle: B*65536 + G*256 + R.</summary>
    private static int ToOle(int r, int g, int b) => b * 65536 + g * 256 + r;

    private static void Release(object? comObject)
    {
        if (comObject is not null)
        {
            try { Marshal.ReleaseComObject(comObject); } catch { }
        }
    }

    /// <summary>Открывает проводник с выделенным файлом (аналог кнопки «Открыть папку»).</summary>
    public static void RevealInExplorer(string filePath)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = "/select,\"" + filePath + "\"",
            UseShellExecute = false,
        });
    }
}
