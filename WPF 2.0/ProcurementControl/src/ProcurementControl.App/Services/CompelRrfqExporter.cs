using System.IO;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>Файл Компэла в формате, который напрямую читает движок RRFQ.</summary>
public static class CompelRrfqExporter
{
    public static string WriteWorkbook(string path, IReadOnlyList<CompelRow> rows)
    {
        if (rows.Count == 0) throw new InvalidOperationException("Сначала разберите HTML-файл Компэла.");
        var fullPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        dynamic? excel = null; dynamic? workbook = null; dynamic? sheet = null;
        try
        {
            excel = RrfqExcel.OpenExcel(); workbook = excel.Workbooks.Add(); sheet = workbook.Worksheets.Item(1); sheet.Name = "RRFQ";
            var headers = new[] { "Value", "Russian remark", "China remark", "PN", "D/C", "Mfg from Russia", "Mfg from China", "Q-ty in packing", "Q-ty to Buy/pcs", "Unit price", "Currency", "Lead time", "Supplier" };
            for (var col = 0; col < headers.Length; col++) sheet.Cells[1, col + 1].Value2 = headers[col];
            for (var index = 0; index < rows.Count; index++)
            {
                var item = rows[index]; var pn = string.Equals(item.SourceName.Trim(), item.MatchedPart.Trim(), StringComparison.OrdinalIgnoreCase) ? string.Empty : item.MatchedPart;
                object?[] values = { item.SourceName, "", "", pn, "", "", item.Manufacturer, item.Package, item.Quantity, item.UnitPrice, item.Currency, item.LeadTime, "Компэл" };
                for (var col = 0; col < values.Length; col++) sheet.Cells[index + 2, col + 1].Value2 = values[col];
            }
            sheet.Range("A1:M1").Font.Bold = true; sheet.Range("A1:M1").Interior.ColorIndex = 15;
            sheet.Range("A1:M" + Math.Max(1, rows.Count + 1)).AutoFilter(); sheet.Columns.AutoFit();
            workbook.SaveAs(fullPath, 51); return fullPath;
        }
        finally
        {
            if (workbook is not null) { try { workbook.Close(false); } catch { } }
            RrfqExcel.Release(sheet); RrfqExcel.Release(workbook); RrfqExcel.CloseExcel(excel);
        }
    }
}
