using System.IO;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Экспорт базы квот в Excel — перенос Export-QuoteBaseToExcel
/// (RRFQComparer.ps1, строки 4383-4428). Пишет только НОВЫЙ файл
/// в каталог данных, база не изменяется.
/// </summary>
public static class QuoteBaseExporter
{
    private static readonly string[] Headers =
    {
        "Дата квоты", "Value", "PN", "Поставщик", "Цена", "Срок",
        "Lead time total", "MFG", "Победитель", "Почему выбран", "Предупреждение", "RFQ",
    };

    /// <summary>
    /// Создаёт книгу с отфильтрованными квотами.
    /// <paramref name="targetDirectory"/> — каталог результата (по умолчанию <see cref="AppPaths.DataRoot"/>).
    /// </summary>
    public static string ExportToExcel(IReadOnlyList<QuoteHistoryRow> rows, string? targetDirectory = null)
    {
        if (rows.Count == 0)
        {
            throw new InvalidOperationException("Нет строк для экспорта.");
        }

        var directory = string.IsNullOrWhiteSpace(targetDirectory) ? AppPaths.DataRoot : targetDirectory!;
        var path = Path.Combine(directory, $"quote_history_export_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        dynamic? excel = null;
        dynamic? wb = null;
        dynamic? ws = null;
        try
        {
            excel = RrfqExcel.OpenExcel();
            wb = excel.Workbooks.Add();
            ws = wb.Worksheets.Item(1);

            for (var c = 0; c < Headers.Length; c++)
            {
                ws.Cells.Item(1, c + 1).Value2 = Headers[c];
            }

            var rowNumber = 2;
            foreach (var row in rows)
            {
                ws.Cells.Item(rowNumber, 1).Value2 = row.QuoteDate;
                ws.Cells.Item(rowNumber, 2).Value2 = row.RfqValue;
                ws.Cells.Item(rowNumber, 3).Value2 = row.PN;
                ws.Cells.Item(rowNumber, 4).Value2 = row.Supplier;
                // Цена — числом, как в оригинале (пусто при отсутствии).
                if (row.UnitPrice is not null)
                {
                    ws.Cells.Item(rowNumber, 5).Value2 = row.UnitPrice.Value;
                }

                ws.Cells.Item(rowNumber, 6).Value2 = row.LeadTime;
                ws.Cells.Item(rowNumber, 7).Value2 = row.LeadTimeTotal;
                ws.Cells.Item(rowNumber, 8).Value2 = row.Mfg;
                ws.Cells.Item(rowNumber, 9).Value2 = row.WinnerMark;
                ws.Cells.Item(rowNumber, 10).Value2 = row.WinnerReason;
                ws.Cells.Item(rowNumber, 11).Value2 = row.Warning;
                ws.Cells.Item(rowNumber, 12).Value2 = row.RfqPath;
                rowNumber++;
            }

            var columns = ws.Columns;
            try
            {
                columns.AutoFit();
            }
            finally
            {
                RrfqExcel.Release(columns);
            }

            wb.SaveAs(path, 51); // xlOpenXMLWorkbook
        }
        finally
        {
            RrfqExcel.Release(ws);
            if (wb is not null)
            {
                try { wb.Close(false); } catch { /* Книга уже могла быть закрыта. */ }
                RrfqExcel.Release(wb);
            }

            RrfqExcel.CloseExcel(excel);
        }

        return path;
    }
}
