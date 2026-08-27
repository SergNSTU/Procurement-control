using System.IO;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Создание результирующего RRFQ — перенос Get-OutputPath и
/// Write-ResultWorkbook с помощниками из app/modules/RrfqEngine.ps1
/// (строки 1477-1677). Пишет только НОВЫЙ файл (копию RFQ), исходник не трогает.
/// </summary>
public static class RrfqResultWriter
{
    private static readonly string[] ResultFields =
    {
        "ChinaRemark", "PN", "DC", "MfgChina", "QtyPacking",
        "UnitPrice", "LeadTime", "LeadTimeTotal", "Supplier", "UnitPriceVat",
    };

    /// <summary>Аналог Get-OutputPath: рядом с RFQ, при занятом имени — с меткой времени.</summary>
    public static string GetOutputPath(string rfqPath)
    {
        var dir = Path.GetDirectoryName(rfqPath) ?? string.Empty;
        var stem = Path.GetFileNameWithoutExtension(rfqPath);
        var candidate = Path.Combine(dir, stem + "_RRFQ_result.xlsx");
        if (!File.Exists(candidate))
        {
            return candidate;
        }

        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        return Path.Combine(dir, $"{stem}_RRFQ_result_{stamp}.xlsx");
    }

    /// <summary>Аналог Write-ResultWorkbook.</summary>
    public static string WriteResultWorkbook(RrfqAnalysis analysis)
    {
        var outputPath = GetOutputPath(analysis.RfqPath);
        File.Copy(analysis.RfqPath, outputPath);

        var vatDivisor = RrfqConfigService.Current.VatDivisor.ToString(System.Globalization.CultureInfo.InvariantCulture);
        dynamic? excel = null;
        dynamic? wb = null;
        bool calculationSet = false;
        try
        {
            excel = RrfqExcel.OpenExcel();
            try
            {
                excel.Calculation = -4135; // xlCalculationManual
                calculationSet = true;
            }
            catch
            {
                calculationSet = false;
            }

            wb = excel.Workbooks.Open(outputPath, 0, false);
            var sheetCount = (int)wb.Worksheets.Count;
            for (var i = 1; i <= sheetCount; i++)
            {
                dynamic ws = wb.Worksheets.Item(i);
                try
                {
                    var sheetName = (string)ws.Name;
                    var sheetDecisionsAll = analysis.Decisions.Where(d => d.SheetName == sheetName).ToList();
                    if (sheetDecisionsAll.Count == 0)
                    {
                        continue;
                    }

                    var sheetDecisionsWithWork = sheetDecisionsAll.Where(d => d.Winner is not null || d.Quotes.Count > 0).ToList();
                    if (sheetDecisionsWithWork.Count == 0)
                    {
                        continue;
                    }

                    var sheetData = RrfqExcel.ReadSheetData(ws);
                    var map = RrfqEngine.FindTableMap(sheetData, "Standard", "Delivery");
                    if (map is null)
                    {
                        continue;
                    }

                    if (map.Fields.ContainsKey("LeadTime"))
                    {
                        map = EnsureHeaderColumn(ws, map, "LeadTimeTotal", "Lead time (total)", insertBeforeColumn: 0, insertAfterColumn: map.Fields["LeadTime"]);
                    }

                    var hasCompel = sheetDecisionsAll.Any(d => d.Include && d.Winner is not null && d.Winner.Supplier == "Компэл");
                    if (hasCompel && map.Fields.ContainsKey("UnitPrice"))
                    {
                        map = EnsureHeaderColumn(ws, map, "UnitPriceVat", "Unit price (с НДС)", insertBeforeColumn: map.Fields["UnitPrice"], insertAfterColumn: 0);
                    }

                    if (!map.Fields.ContainsKey("Supplier"))
                    {
                        map = EnsureHeaderColumn(ws, map, "Supplier", "Supplier", insertBeforeColumn: 0, insertAfterColumn: 0);
                    }

                    foreach (var decision in sheetDecisionsWithWork)
                    {
                        ClearResultRow(ws, map, decision.Row);
                    }

                    foreach (var decision in sheetDecisionsWithWork.Where(d => d.Include && d.Winner is not null))
                    {
                        WriteDecisionToWorksheet(ws, map, decision, vatDivisor);
                    }
                }
                finally
                {
                    RrfqExcel.Release(ws);
                }
            }

            wb.Save();
        }
        finally
        {
            if (calculationSet && excel is not null)
            {
                try { excel.Calculation = -4105; } catch { /* xlCalculationAutomatic, не критично. */ }
            }

            if (wb is not null)
            {
                try { wb.Close(true); } catch { /* Книга уже могла быть закрыта. */ }
                RrfqExcel.Release(wb);
            }

            RrfqExcel.CloseExcel(excel);
        }

        return outputPath;
    }

    /// <summary>Аналог Ensure-HeaderColumn: вставка недостающей колонки с корректировкой карты.</summary>
    private static HeaderMap EnsureHeaderColumn(
        dynamic ws,
        HeaderMap map,
        string field,
        string headerText,
        int insertBeforeColumn,
        int insertAfterColumn)
    {
        if (map.Fields.ContainsKey(field))
        {
            return map;
        }

        int newCol;
        if (insertBeforeColumn > 0)
        {
            var columnRange = ws.Columns.Item(insertBeforeColumn);
            try
            {
                columnRange.Insert();
            }
            finally
            {
                RrfqExcel.Release(columnRange);
            }

            newCol = insertBeforeColumn;
        }
        else if (insertAfterColumn > 0)
        {
            var columnRange = ws.Columns.Item(insertAfterColumn + 1);
            try
            {
                columnRange.Insert();
            }
            finally
            {
                RrfqExcel.Release(columnRange);
            }

            newCol = insertAfterColumn + 1;
        }
        else
        {
            var used = ws.UsedRange;
            try
            {
                newCol = (int)used.Column + (int)used.Columns.Count;
            }
            finally
            {
                RrfqExcel.Release(used);
            }
        }

        var headerCell = ws.Cells.Item(map.HeaderRow, newCol);
        try
        {
            headerCell.Value2 = headerText;
        }
        finally
        {
            RrfqExcel.Release(headerCell);
        }

        var newColumnRange = ws.Columns.Item(newCol);
        try
        {
            newColumnRange.Hidden = false;
        }
        finally
        {
            RrfqExcel.Release(newColumnRange);
        }

        var updatedFields = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var (key, col) in map.Fields)
        {
            updatedFields[key] = col >= newCol ? col + 1 : col;
        }

        updatedFields[field] = newCol;
        return new HeaderMap
        {
            Type = map.Type,
            HeaderRow = map.HeaderRow,
            Fields = updatedFields,
            Score = map.Score,
        };
    }

    /// <summary>Аналог Clear-ResultRow.</summary>
    private static void ClearResultRow(dynamic ws, HeaderMap map, int row)
    {
        foreach (var field in ResultFields)
        {
            if (!map.Fields.TryGetValue(field, out var col))
            {
                continue;
            }

            var cell = ws.Cells.Item(row, col);
            try
            {
                cell.ClearContents();
            }
            finally
            {
                RrfqExcel.Release(cell);
            }
        }
    }

    /// <summary>Аналог Set-CellIfPresent.</summary>
    private static void SetCellIfPresent(dynamic ws, HeaderMap map, string field, int row, string? value)
    {
        if (!map.Fields.TryGetValue(field, out var col))
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        var cell = ws.Cells.Item(row, col);
        try
        {
            cell.Value2 = value;
        }
        finally
        {
            RrfqExcel.Release(cell);
        }
    }

    /// <summary>Аналог Write-DecisionToWorksheet.</summary>
    private static void WriteDecisionToWorksheet(dynamic ws, HeaderMap map, Decision decision, string vatDivisor)
    {
        var row = decision.Row;
        var quote = decision.Winner;
        if (quote is null)
        {
            return;
        }

        SetCellIfPresent(ws, map, "ChinaRemark", row, quote.ChinaRemark);
        SetCellIfPresent(ws, map, "PN", row, quote.PN);
        SetCellIfPresent(ws, map, "DC", row, quote.DC);

        if (!string.IsNullOrWhiteSpace(quote.MfgChina))
        {
            SetCellIfPresent(ws, map, "MfgChina", row, quote.MfgChina);
        }
        else if (!string.IsNullOrWhiteSpace(quote.MfgRussia) && string.IsNullOrWhiteSpace(decision.RfqRow.MfgRussia))
        {
            SetCellIfPresent(ws, map, "MfgChina", row, quote.MfgRussia);
        }

        var qtyPacking = quote.QtyPacking;
        if (RrfqEngine.IsCompelSupplier(quote.Supplier))
        {
            // Ячейка QtyToBuy в RFQ формульная — её трогать нельзя. Если Компэлу
            // нужно больше штук, чем в RFQ, показываем его количество в QtyPacking.
            var rfqQtyToBuy = RrfqEngine.ConvertToNumber(decision.RfqRow.QtyToBuy);
            var compelQtyToBuy = RrfqEngine.ConvertToNumber(quote.QtyToBuy);
            if (rfqQtyToBuy is not null && compelQtyToBuy is not null && compelQtyToBuy.Value > rfqQtyToBuy.Value)
            {
                qtyPacking = quote.QtyToBuy;
            }
        }

        SetCellIfPresent(ws, map, "QtyPacking", row, qtyPacking);

        if (quote.Supplier == "Компэл" && map.Fields.ContainsKey("UnitPriceVat") && quote.UnitPriceVat is not null)
        {
            var vatCol = map.Fields["UnitPriceVat"];
            var unitCol = map.Fields["UnitPrice"];
            dynamic vatCell = ws.Cells.Item(row, vatCol);
            dynamic unitCell = ws.Cells.Item(row, unitCol);
            try
            {
                try
                {
                    vatCell.NumberFormat = unitCell.NumberFormat;
                }
                catch
                {
                    vatCell.NumberFormat = "0.000000";
                }

                vatCell.Value2 = quote.UnitPriceVat.Value;
                var vatAddress = (string)vatCell.Address(false, false);
                unitCell.Formula = $"={vatAddress}/{vatDivisor}";
            }
            finally
            {
                RrfqExcel.Release(vatCell);
                RrfqExcel.Release(unitCell);
            }
        }
        else if (map.Fields.ContainsKey("UnitPrice") && quote.UnitPrice is not null)
        {
            var cell = ws.Cells.Item(row, map.Fields["UnitPrice"]);
            try
            {
                cell.Value2 = quote.UnitPrice.Value;
            }
            finally
            {
                RrfqExcel.Release(cell);
            }
        }

        SetCellIfPresent(ws, map, "LeadTime", row, quote.LeadTime);
        SetCellIfPresent(ws, map, "LeadTimeTotal", row, quote.LeadTimeTotal);
        SetCellIfPresent(ws, map, "Supplier", row, quote.Supplier);
    }
}
