using System.IO;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>Экспортирует прозрачный обзор с числителями и знаменателями.</summary>
public static class SupplierOverviewExporter
{
    private static readonly string[] Headers =
    {
        "Поставщик", "Когорты", "Позиции", "Уникальные PN", "Ответы: числитель", "Ответы: знаменатель", "Частота ответа",
        "Победы", "Решения", "Доля побед", "Лучшие цены", "Сравнимые цены", "Доля лучших цен",
        "Среднее отклонение цены", "Медиана отклонения цены", "Средний срок, дн.", "Медиана срока, дн.",
        "Лучшие сроки", "Сравнимые сроки", "Доля лучших сроков"
    };

    public static string ExportToExcel(IReadOnlyList<SupplierOverviewRow> rows, string? targetDirectory = null)
    {
        if (rows.Count == 0) throw new InvalidOperationException("Нет строк для экспорта.");
        var directory = string.IsNullOrWhiteSpace(targetDirectory) ? AppPaths.DataRoot : targetDirectory!;
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, $"supplier_overview_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        dynamic? excel = null; dynamic? workbook = null; dynamic? sheet = null;
        try
        {
            excel = RrfqExcel.OpenExcel(); workbook = excel.Workbooks.Add(); sheet = workbook.Worksheets.Item(1); sheet.Name = "Обзор";
            for (var column = 0; column < Headers.Length; column++) sheet.Cells.Item(1, column + 1).Value2 = Headers[column];
            var rowNumber = 2;
            foreach (var row in rows)
            {
                object?[] values =
                {
                    row.Supplier, row.Cohorts, row.Positions, row.UniquePn, row.ResponseNumerator, row.ResponseDenominator,
                    row.ResponseDenominator == 0 ? null : row.ResponseNumerator / (double)row.ResponseDenominator,
                    row.Wins, row.WinDenominator, row.WinDenominator == 0 ? null : row.Wins / (double)row.WinDenominator,
                    row.BestPriceCount, row.PriceDenominator, row.PriceDenominator == 0 ? null : row.BestPriceCount / (double)row.PriceDenominator,
                    row.MeanPriceDeviation, row.MedianPriceDeviation, row.MeanLeadDays, row.MedianLeadDays,
                    row.BestLeadCount, row.LeadDenominator, row.LeadDenominator == 0 ? null : row.BestLeadCount / (double)row.LeadDenominator,
                };
                for (var column = 0; column < values.Length; column++) if (values[column] is not null) sheet.Cells.Item(rowNumber, column + 1).Value2 = values[column];
                rowNumber++;
            }
            var columns = sheet.Columns;
            try { columns.AutoFit(); } finally { RrfqExcel.Release(columns); }
            workbook.SaveAs(path, 51);
        }
        finally
        {
            RrfqExcel.Release(sheet);
            if (workbook is not null) { try { workbook.Close(false); } catch (Exception ex) { ErrorLog.Append("excel", ex); } RrfqExcel.Release(workbook); }
            RrfqExcel.CloseExcel(excel);
        }
        return path;
    }
}
