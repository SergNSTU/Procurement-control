using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Движок сравнения RRFQ — перенос app/modules/RrfqEngine.ps1
/// (шапки, квоты, сопоставление, победители, решения). Аналог Invoke-Analysis
/// без записи в базу: Save-QuoteHistory в порт не входит.
/// </summary>
public static class RrfqEngine
{
    private static readonly Regex WhitespaceRuns = new(@"\s+", RegexOptions.Compiled);
    private static readonly Regex NewlineRuns = new(@"[\r\n\t]+", RegexOptions.Compiled);
    private static readonly Regex CurrencyChars = new(@"[\$€?]", RegexOptions.Compiled);
    private static readonly Regex CurrencyWords = new("usd|eur|rub|rur", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex NumberInText = new(@"\d+(?:[\.,]\d+)?", RegexOptions.Compiled);

    /// <summary>Читает квоты из одного файла для импорта в «Базу квот».</summary>
    public static List<Quote> ReadQuotesFromWorkbook(string path, string supplier)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            throw new FileNotFoundException("Файл RRFQ не найден: " + path, path);
        dynamic? excel = null;
        try
        {
            excel = RrfqExcel.OpenExcel();
            return GetQuotesFromWorkbook(excel, path, supplier, RrfqConfigService.Current.VatDivisor);
        }
        finally
        {
            RrfqExcel.CloseExcel(excel);
        }
    }

    // ----- Нормализация (строки 39-53) -----

    /// <summary>Аналог Normalize-Header.</summary>
    public static string NormalizeHeader(object? value)
    {
        var text = Convert.ToString(value ?? string.Empty, CultureInfo.InvariantCulture) ?? string.Empty;
        text = text.Replace('\u00A0', ' ').Trim().ToLowerInvariant();
        text = NewlineRuns.Replace(text, " ");
        text = WhitespaceRuns.Replace(text, " ");
        return text;
    }

    /// <summary>Аналог Normalize-Key.</summary>
    public static string NormalizeKey(object? value)
    {
        var text = Convert.ToString(value ?? string.Empty, CultureInfo.InvariantCulture) ?? string.Empty;
        text = text.Replace('\u00A0', ' ').Trim().ToUpperInvariant();
        text = NewlineRuns.Replace(text, " ");
        text = WhitespaceRuns.Replace(text, " ");
        return text;
    }

    /// <summary>Аналог Convert-ToNumber: валюта, пробелы, запятая как разделитель.</summary>
    public static double? ConvertToNumber(object? value)
    {
        if (value is null)
        {
            return null;
        }

        switch (value)
        {
            case double d:
                return d;
            case int i:
                return i;
            case decimal dec:
                return (double)dec;
            case long l:
                return l;
        }

        var text = Convert.ToString(value, CultureInfo.InvariantCulture)!.Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        text = CurrencyChars.Replace(text, string.Empty);
        text = CurrencyWords.Replace(text, string.Empty);
        text = text.Replace('\u00A0', ' ');
        text = WhitespaceRuns.Replace(text, string.Empty);
        if (text.Contains(',') && !text.Contains('.'))
        {
            text = text.Replace(',', '.');
        }
        else if (text.Contains(',') && text.Contains('.'))
        {
            text = text.Replace(",", string.Empty);
        }

        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var number))
        {
            return number;
        }

        return null;
    }

    // ----- Шапки таблиц (строки 565-735) -----

    private sealed record HeaderText(int Column, string Text, bool Visible);

    /// <summary>Аналог Get-HeaderTexts.</summary>
    private static List<HeaderText> GetHeaderTexts(SheetData sheet, int row)
    {
        var items = new List<HeaderText>();
        for (var col = sheet.FirstCol; col <= sheet.LastCol; col++)
        {
            var text = NormalizeHeader(sheet.GetCellText(row, col));
            if (!string.IsNullOrWhiteSpace(text))
            {
                items.Add(new HeaderText(col, text, !sheet.IsColumnHidden(col)));
            }
        }

        return items;
    }

    private static bool HasHeader(List<HeaderText> headers, string pattern)
        => headers.Any(h => Regex.IsMatch(h.Text, pattern));

    /// <summary>Аналог Build-HeaderMapForRow: тип шапки и карта полей с очками.</summary>
    public static HeaderMap? BuildHeaderMapForRow(SheetData sheet, int row)
    {
        var headers = GetHeaderTexts(sheet, row);
        if (headers.Count == 0)
        {
            return null;
        }

        string type;
        if (HasHeader(headers, "^исходное наименование$") && HasHeader(headers, "^цена за шт"))
        {
            type = "Compel";
        }
        else if (HasHeader(headers, "^original inquiry$") && (HasHeader(headers, "^buy unit") || HasHeader(headers, "^l/t$")))
        {
            type = "Delivery";
        }
        else if (HasHeader(headers, "^value$") && HasHeader(headers, "^pn$") && HasHeader(headers, "^unit price$"))
        {
            type = "Standard";
        }
        else
        {
            return null;
        }

        var rules = GetHeaderRules(type);
        var candidates = new Dictionary<string, List<(int Column, bool Visible)>>(StringComparer.Ordinal);
        foreach (var header in headers)
        {
            foreach (var (pattern, field) in rules)
            {
                if (Regex.IsMatch(header.Text, pattern))
                {
                    if (!candidates.TryGetValue(field, out var list))
                    {
                        list = new List<(int, bool)>();
                        candidates[field] = list;
                    }

                    list.Add((header.Column, header.Visible));
                    break;
                }
            }
        }

        var fields = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var (field, list) in candidates)
        {
            // Аналог Select-HeaderColumn: первый видимый, иначе первый.
            var visible = list.FirstOrDefault(c => c.Visible, list[0]);
            fields[field] = visible.Column;
        }

        var score = fields.Count;
        if (fields.ContainsKey("Value"))
        {
            score += 5;
        }

        if (fields.ContainsKey("UnitPrice") || fields.ContainsKey("UnitPriceVat"))
        {
            score += 4;
        }

        if (fields.ContainsKey("LeadTime"))
        {
            score += 3;
        }

        return new HeaderMap { Type = type, HeaderRow = row, Fields = fields, Score = score };
    }

    /// <summary>Правила сопоставления заголовков по типу формата (строки 647-689).</summary>
    private static (string Pattern, string Field)[] GetHeaderRules(string type) => type switch
    {
        "Standard" => new[]
        {
            ("^value$", "Value"),
            ("^russian remark$", "RussianRemark"),
            ("^(china|chinese) remark$", "ChinaRemark"),
            ("^pn$", "PN"),
            ("^d/c$", "DC"),
            ("^mfg from russia$", "MfgRussia"),
            ("^mfg from china$", "MfgChina"),
            ("^q-?ty in packing$|^qty in packing$", "QtyPacking"),
            ("^qty to buy/pcs$", "QtyToBuy"),
            ("^unit price \\(с ндс\\)$", "UnitPriceVat"),
            ("^unit price$", "UnitPrice"),
            ("^lead time \\(total\\)$", "LeadTimeTotal"),
            ("^lead time", "LeadTime"),
            ("^supplier$|^поставщик$|^source$", "Supplier"),
        },
        "Delivery" => new[]
        {
            ("^original inquiry$", "Value"),
            ("^russian remark$", "RussianRemark"),
            ("^(china|chinese) remark$", "ChinaRemark"),
            ("^replacement$", "PN"),
            ("^mfg$", "MfgChina"),
            ("^d/c$", "DC"),
            ("^qty in packing$|^q-?ty in packing$", "QtyPacking"),
            ("^qty to buy/pcs$", "QtyToBuy"),
            ("^buy unit", "UnitPrice"),
            ("^lead time \\(total\\)$", "LeadTimeTotal"),
            ("^l/t$|^lead time", "LeadTime"),
            ("^supplier$|^поставщик$|^source$", "Supplier"),
        },
        _ => new[]
        {
            ("^исходное наименование$", "Value"),
            ("^подобранный товар$", "PN"),
            ("^производитель$", "MfgChina"),
            ("^срок поставки$", "LeadTime"),
            ("^количество$", "QtyToBuy"),
            ("^цена за шт", "UnitPriceVat"),
            ("^валюта$", "Currency"),
        },
    };

    /// <summary>Аналог Find-TableMap: лучшая шапка в первых 180 строках.</summary>
    public static HeaderMap? FindTableMap(SheetData sheet, params string[] allowedTypes)
    {
        HeaderMap? best = null;
        var scanLastRow = Math.Min(sheet.LastRow, sheet.FirstRow + 180);
        for (var row = sheet.FirstRow; row <= scanLastRow; row++)
        {
            var map = BuildHeaderMapForRow(sheet, row);
            if (map is null || !allowedTypes.Contains(map.Type))
            {
                continue;
            }

            if (best is null || map.Score > best.Score)
            {
                best = map;
            }
        }

        return best;
    }

    /// <summary>Аналог Get-FieldText.</summary>
    public static string GetFieldText(SheetData sheet, HeaderMap map, string field, int row)
        => map.Fields.TryGetValue(field, out var col) ? sheet.GetCellText(row, col) : string.Empty;

    /// <summary>Аналог Get-FieldNumber.</summary>
    public static double? GetFieldNumber(SheetData sheet, HeaderMap map, string field, int row)
        => map.Fields.TryGetValue(field, out var col) ? sheet.GetCellNumber(row, col) : null;

    // ----- Сроки (строки 764-864) -----

    /// <summary>Результат распознавания срока (аналог возвращаемого объекта Convert-LeadTime).</summary>
    public sealed record LeadInfo(double? Days, string Total, string Warning);

    /// <summary>Аналог Convert-LeadTime.</summary>
    public static LeadInfo ConvertLeadTime(string leadTime)
    {
        var raw = (leadTime ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(raw))
        {
            return new LeadInfo(null, string.Empty, string.Empty);
        }

        var s = raw.ToLowerInvariant().Replace('\u00A0', ' ');
        if (Regex.IsMatch(s, "stock|склад"))
        {
            return new LeadInfo(0.0, "stock", string.Empty);
        }

        var matches = NumberInText.Matches(s);
        if (matches.Count == 0)
        {
            return new LeadInfo(null, string.Empty, "Не распознан срок: " + raw);
        }

        var numbers = matches.Select(m => double.Parse(m.Value.Replace(',', '.'), CultureInfo.InvariantCulture)).ToList();
        var min = numbers.Min();
        var max = numbers.Max();

        if (Regex.IsMatch(s, @"week|weeks|w\b|нед"))
        {
            var days = max * 7;
            var totalMin = (int)Math.Ceiling(min) + 1;
            var totalMax = (int)Math.Ceiling(max) + 2;
            return new LeadInfo(days, FormatWeeks(totalMin, totalMax), string.Empty);
        }

        if (Regex.IsMatch(s, @"\bd\b|day|days|д\b|дн"))
        {
            var days = max;
            if (days <= 3)
            {
                return new LeadInfo(days, "1 week", string.Empty);
            }

            var baseWeeks = (int)Math.Ceiling(days / 7.0);
            return new LeadInfo(days, FormatWeeks(baseWeeks + 1, baseWeeks + 2), string.Empty);
        }

        return new LeadInfo(null, string.Empty, "Не распознан срок: " + raw);
    }

    /// <summary>Аналог Test-BenSupplier.</summary>
    public static bool IsBenSupplier(string supplier)
        => Regex.IsMatch(NormalizeKey(supplier), @"^(BEN|БЕН)(_[0-9]+)?$");

    /// <summary>Аналог Test-CompelSupplier.</summary>
    public static bool IsCompelSupplier(string supplier)
        => Regex.IsMatch(NormalizeKey(supplier), @"^(COMPEL|КОМПЭЛ|КОМПЕЛ)(_[0-9]+)?$");

    /// <summary>Аналог Get-LeadTimeTotalForSupplier.</summary>
    public static string GetLeadTimeTotalForSupplier(string supplier, string leadTime, LeadInfo leadInfo)
    {
        if (IsCompelSupplier(supplier))
        {
            if (leadInfo.Days is null)
            {
                return leadInfo.Total;
            }

            if (leadInfo.Days.Value <= 0)
            {
                return "stock";
            }

            var weeks = (int)Math.Ceiling(leadInfo.Days.Value / 5.0);
            return FormatWeeks(weeks, weeks);
        }

        if (IsBenSupplier(supplier))
        {
            return leadTime;
        }

        return leadInfo.Total;
    }

    /// <summary>Аналог Format-Weeks.</summary>
    public static string FormatWeeks(int minWeeks, int maxWeeks)
    {
        if (minWeeks <= 1 && maxWeeks <= 1)
        {
            return "1 week";
        }

        if (minWeeks == maxWeeks)
        {
            return minWeeks + " weeks";
        }

        return minWeeks + "-" + maxWeeks + " weeks";
    }

    // ----- Строки RFQ и квоты (строки 866-1045) -----

    /// <summary>Аналог New-RfqRow; null для пустых ключей и итоговых строк.</summary>
    private static RfqRow? NewRfqRow(SheetData sheet, string workbookPath, HeaderMap map, int row)
    {
        var key = GetFieldText(sheet, map, "Value", row);
        if (string.IsNullOrWhiteSpace(key))
        {
            return null;
        }

        if (Regex.IsMatch(NormalizeKey(key), "^(TOTAL|ИТОГО)$"))
        {
            return null;
        }

        return new RfqRow
        {
            Id = $"{sheet.Name}::{row}::{NormalizeKey(key)}",
            WorkbookPath = workbookPath,
            SheetName = sheet.Name,
            SheetIndex = sheet.Index,
            Format = map.Type,
            HeaderRow = map.HeaderRow,
            Row = row,
            Key = key,
            KeyNorm = NormalizeKey(key),
            RussianRemark = GetFieldText(sheet, map, "RussianRemark", row),
            ChinaRemark = GetFieldText(sheet, map, "ChinaRemark", row),
            PN = GetFieldText(sheet, map, "PN", row),
            DC = GetFieldText(sheet, map, "DC", row),
            MfgRussia = GetFieldText(sheet, map, "MfgRussia", row),
            MfgChina = GetFieldText(sheet, map, "MfgChina", row),
            QtyPacking = GetFieldText(sheet, map, "QtyPacking", row),
            QtyToBuy = GetFieldText(sheet, map, "QtyToBuy", row),
        };
    }

    /// <summary>Аналог New-Quote.</summary>
    private static Quote NewQuote(SheetData sheet, string workbookPath, string supplier, HeaderMap map, int row, string forcedKey, double vatDivisor)
    {
        var sheetSupplier = GetFieldText(sheet, map, "Supplier", row);
        var quoteSupplier = string.IsNullOrWhiteSpace(sheetSupplier) ? supplier : sheetSupplier.Trim();
        var key = string.IsNullOrWhiteSpace(forcedKey) ? GetFieldText(sheet, map, "Value", row) : forcedKey;
        var pn = GetFieldText(sheet, map, "PN", row);
        if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(pn)
            && string.Equals(key.Trim(), pn.Trim(), StringComparison.OrdinalIgnoreCase))
        {
            pn = string.Empty;
        }

        var lead = GetFieldText(sheet, map, "LeadTime", row);
        var currency = GetFieldText(sheet, map, "Currency", row);

        var unitPrice = GetFieldNumber(sheet, map, "UnitPrice", row);
        var unitPriceVat = GetFieldNumber(sheet, map, "UnitPriceVat", row);
        if (map.Type == "Compel")
        {
            unitPrice = unitPriceVat is null ? null : unitPriceVat.Value / vatDivisor;
        }

        var leadInfo = ConvertLeadTime(lead);
        var leadTotal = GetLeadTimeTotalForSupplier(quoteSupplier, lead, leadInfo);
        var isCurrencyComparable = true;
        var warning = string.Empty;
        if (!string.IsNullOrWhiteSpace(currency) && NormalizeKey(currency) != "USD")
        {
            isCurrencyComparable = false;
            warning = "Валюта не USD: " + currency;
        }

        if (!string.IsNullOrWhiteSpace(leadInfo.Warning))
        {
            warning = string.IsNullOrWhiteSpace(warning) ? leadInfo.Warning : warning + "; " + leadInfo.Warning;
        }

        return new Quote
        {
            WorkbookPath = workbookPath,
            FileName = Path.GetFileName(workbookPath),
            SheetName = sheet.Name,
            SheetIndex = sheet.Index,
            Format = map.Type,
            Row = row,
            Supplier = quoteSupplier,
            Key = key,
            KeyNorm = NormalizeKey(key),
            RussianRemark = GetFieldText(sheet, map, "RussianRemark", row),
            ChinaRemark = GetFieldText(sheet, map, "ChinaRemark", row),
            PN = pn,
            DC = GetFieldText(sheet, map, "DC", row),
            MfgRussia = GetFieldText(sheet, map, "MfgRussia", row),
            MfgChina = GetFieldText(sheet, map, "MfgChina", row),
            QtyPacking = GetFieldText(sheet, map, "QtyPacking", row),
            QtyToBuy = GetFieldText(sheet, map, "QtyToBuy", row),
            UnitPrice = unitPrice,
            UnitPriceVat = unitPriceVat,
            LeadTime = lead,
            LeadDays = leadInfo.Days,
            LeadTimeTotal = leadTotal,
            Currency = currency,
            IsPriceComparable = isCurrencyComparable,
            Warning = warning,
            TargetId = null,
            MatchStatus = "Unmatched",
        };
    }

    /// <summary>Аналог New-ManualQuote (используется ручной квотой из блока 7б).</summary>
    public static Quote NewManualQuote(
        Decision decision,
        string supplier,
        string chinaRemark,
        string pn,
        string dc,
        string mfgRussia,
        string mfgChina,
        string qtyPacking,
        string qtyToBuy,
        string unitPrice,
        string leadTime)
    {
        double? priceNumber = null;
        if (!string.IsNullOrWhiteSpace(unitPrice))
        {
            priceNumber = ConvertToNumber(unitPrice);
            if (priceNumber is null)
            {
                throw new InvalidOperationException("Не удалось распознать цену: " + unitPrice);
            }
        }

        var supplierName = string.IsNullOrWhiteSpace(supplier) ? "Др." : supplier.Trim();
        var leadInfo = ConvertLeadTime(leadTime);

        return new Quote
        {
            WorkbookPath = string.Empty,
            FileName = "Ручной ввод",
            SheetName = decision.SheetName,
            SheetIndex = 0,
            Format = "Manual",
            Row = decision.Row,
            Supplier = supplierName,
            Key = decision.Value,
            KeyNorm = NormalizeKey(decision.Value),
            RussianRemark = decision.RussianRemark,
            ChinaRemark = chinaRemark,
            PN = pn,
            DC = dc,
            MfgRussia = mfgRussia,
            MfgChina = mfgChina,
            QtyPacking = qtyPacking,
            QtyToBuy = qtyToBuy,
            UnitPrice = priceNumber,
            UnitPriceVat = null,
            LeadTime = leadTime,
            LeadDays = leadInfo.Days,
            LeadTimeTotal = GetLeadTimeTotalForSupplier(supplierName, leadTime, leadInfo),
            Currency = "USD",
            IsPriceComparable = true,
            Warning = string.IsNullOrWhiteSpace(leadInfo.Warning) ? "Ручной ввод" : "Ручной ввод; " + leadInfo.Warning,
            TargetId = decision.Id,
            MatchStatus = "ManualQuote",
        };
    }

    /// <summary>Аналог Test-QuoteHasAnyData.</summary>
    private static bool QuoteHasAnyData(Quote quote)
        => quote.UnitPrice is not null || quote.UnitPriceVat is not null;

    /// <summary>Аналог Add-QuoteWarning.</summary>
    public static void AddQuoteWarning(Quote quote, string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(quote.Warning))
        {
            quote.Warning = message;
        }
        else if (!quote.Warning.Contains(message, StringComparison.Ordinal))
        {
            quote.Warning = quote.Warning + "; " + message;
        }
    }

    // ----- Победители и решения (строки 1069-1427) -----

    /// <summary>Аналог Add-SuspiciousPriceWarnings (правило 10x).</summary>
    private static void AddSuspiciousPriceWarnings(List<Quote> quotesForRow)
    {
        var priced = quotesForRow.Where(q => IsQuoteSelectable(q) && q.UnitPrice is > 0).ToList();
        if (priced.Count < 2)
        {
            return;
        }

        foreach (var quote in priced)
        {
            var others = priced.Where(q => q.Id != quote.Id).ToList();
            if (others.Count == 0)
            {
                continue;
            }

            var price = quote.UnitPrice!.Value;
            var minOther = others.Min(q => q.UnitPrice!.Value);
            var maxOther = others.Max(q => q.UnitPrice!.Value);
            if ((minOther > 0 && price > minOther * 10) || (price > 0 && maxOther > price * 10))
            {
                AddQuoteWarning(quote, "Подозрительная цена: отличается более чем в 10 раз от других поставщиков");
            }
        }
    }

    /// <summary>Аналог Get-WinnerReason.</summary>
    private static string GetWinnerReason(Quote? winner, List<Quote> selectable, string priorityMode, bool isManualWinner)
    {
        if (winner is null)
        {
            return string.Empty;
        }

        if (isManualWinner)
        {
            return "ручной выбор";
        }

        var items = selectable.Where(IsQuoteSelectable).ToList();
        if (items.Count <= 1)
        {
            return priorityMode == "LeadTime" ? "лучший срок" : "минимальная цена";
        }

        if (priorityMode == "LeadTime")
        {
            var winnerLead = winner.LeadDays ?? double.PositiveInfinity;
            var sameLead = items.Count(q => Math.Abs((q.LeadDays ?? double.PositiveInfinity) - winnerLead) < 0.0001);
            return sameLead > 1 ? "срок равен, выбрана меньшая цена" : "лучший срок";
        }

        var winnerPrice = winner.UnitPrice!.Value;
        var samePrice = items.Count(q => q.UnitPrice is not null && Math.Abs(q.UnitPrice.Value - winnerPrice) < 0.0000001);
        return samePrice > 1 ? "цена равна, выбран меньший срок" : "минимальная цена";
    }

    /// <summary>Аналог Test-QuoteSelectable.</summary>
    public static bool IsQuoteSelectable(Quote quote)
        => quote.UnitPrice is not null && quote.IsPriceComparable;

    /// <summary>Аналог Compare-QuoteForWinner.</summary>
    private static Quote CompareQuoteForWinner(Quote? a, Quote? b, string priorityMode)
    {
        if (a is null)
        {
            return b!;
        }

        if (b is null)
        {
            return a;
        }

        if (priorityMode == "LeadTime")
        {
            var aLead = a.LeadDays ?? double.PositiveInfinity;
            var bLead = b.LeadDays ?? double.PositiveInfinity;
            if (aLead < bLead)
            {
                return a;
            }

            if (bLead < aLead)
            {
                return b;
            }

            return a.UnitPrice!.Value <= b.UnitPrice!.Value ? a : b;
        }

        if (a.UnitPrice!.Value < b.UnitPrice!.Value)
        {
            return a;
        }

        if (b.UnitPrice!.Value < a.UnitPrice!.Value)
        {
            return b;
        }

        var aLeadPrice = a.LeadDays ?? double.PositiveInfinity;
        var bLeadPrice = b.LeadDays ?? double.PositiveInfinity;
        return aLeadPrice <= bLeadPrice ? a : b;
    }

    /// <summary>Аналог Build-Decisions.</summary>
    public static List<Decision> BuildDecisions(List<RfqRow> rfqRows, List<Quote> quotes, string priorityMode)
    {
        var quotesByTarget = new Dictionary<string, List<Quote>>(StringComparer.Ordinal);
        foreach (var quote in quotes)
        {
            if (string.IsNullOrWhiteSpace(quote.TargetId))
            {
                continue;
            }

            if (!quotesByTarget.TryGetValue(quote.TargetId!, out var list))
            {
                list = new List<Quote>();
                quotesByTarget[quote.TargetId!] = list;
            }

            list.Add(quote);
        }

        var decisions = new List<Decision>();
        foreach (var row in rfqRows)
        {
            var quotesForRow = quotesByTarget.TryGetValue(row.Id, out var found) ? found : new List<Quote>();

            AddSuspiciousPriceWarnings(quotesForRow);
            var selectable = quotesForRow.Where(IsQuoteSelectable).ToList();
            Quote? winner = null;
            foreach (var quote in selectable)
            {
                winner = CompareQuoteForWinner(winner, quote, priorityMode);
            }

            var manualWinnerId = row.ManualWinnerId ?? string.Empty;
            Quote? manualWinner = null;
            if (!string.IsNullOrWhiteSpace(manualWinnerId))
            {
                manualWinner = quotesForRow.FirstOrDefault(q => q.Id == manualWinnerId);
                if (manualWinner is not null && IsQuoteSelectable(manualWinner))
                {
                    winner = manualWinner;
                }
            }

            var status = "Нет квот";
            if (quotesForRow.Count > 0 && selectable.Count == 0)
            {
                status = "Нет сравнимой цены";
            }
            else if (winner is not null)
            {
                status = "OK";
            }

            var warnings = new List<string>();
            foreach (var quote in quotesForRow)
            {
                if (!string.IsNullOrWhiteSpace(quote.Warning))
                {
                    warnings.Add($"{quote.Supplier}: {quote.Warning}");
                }

                if (quote.UnitPrice is null)
                {
                    warnings.Add($"{quote.Supplier}: нет цены");
                }
            }

            if (manualWinner is not null && IsQuoteSelectable(manualWinner))
            {
                warnings.Add("Ручной выбор победителя: " + manualWinner.Supplier);
            }
            else if (!string.IsNullOrWhiteSpace(manualWinnerId))
            {
                warnings.Add("Ручной победитель не найден или без сравнимой цены");
            }

            var isManualWinner = manualWinner is not null && winner is not null && winner.Id == manualWinner.Id;
            var winnerReason = GetWinnerReason(winner, selectable, priorityMode, isManualWinner);

            var summary = string.Join(" | ", quotesForRow.Select(q =>
            {
                var priceText = q.UnitPrice is null
                    ? "-"
                    : q.UnitPrice.Value.ToString("N6", CultureInfo.GetCultureInfo("ru-RU"));
                return $"{q.Supplier}: {priceText}; {q.LeadTime}";
            }));

            var resultMfgChina = string.Empty;
            if (winner is not null)
            {
                if (!string.IsNullOrWhiteSpace(winner.MfgChina))
                {
                    resultMfgChina = winner.MfgChina;
                }
                else if (!string.IsNullOrWhiteSpace(winner.MfgRussia) && string.IsNullOrWhiteSpace(row.MfgRussia))
                {
                    resultMfgChina = winner.MfgRussia;
                }
            }

            decisions.Add(new Decision
            {
                Id = row.Id,
                Include = winner is not null,
                Status = status,
                SheetName = row.SheetName,
                Row = row.Row,
                Key = row.Key,
                Value = row.Key,
                RfqRow = row,
                Quotes = quotesForRow.ToList(),
                Winner = winner,
                ManualWinnerId = manualWinnerId,
                IsManualWinner = isManualWinner,
                WinnerSupplier = winner?.Supplier ?? string.Empty,
                WinnerReason = winnerReason,
                WinnerPrice = winner?.UnitPrice,
                WinnerLead = winner?.LeadTime ?? string.Empty,
                WinnerLeadTotal = winner?.LeadTimeTotal ?? string.Empty,
                WinnerPN = winner?.PN ?? string.Empty,
                RussianRemark = row.RussianRemark,
                ChinaRemark = winner is null || string.IsNullOrWhiteSpace(winner.ChinaRemark) ? row.ChinaRemark : winner.ChinaRemark,
                DC = winner is null || string.IsNullOrWhiteSpace(winner.DC) ? row.DC : winner.DC,
                MfgRussia = row.MfgRussia,
                MfgChina = string.IsNullOrWhiteSpace(resultMfgChina) ? row.MfgChina : resultMfgChina,
                QtyPacking = winner is null || string.IsNullOrWhiteSpace(winner.QtyPacking) ? row.QtyPacking : winner.QtyPacking,
                QtyToBuy = winner is null || string.IsNullOrWhiteSpace(winner.QtyToBuy) ? row.QtyToBuy : winner.QtyToBuy,
                QuotesSummary = summary,
                Warning = string.Join("; ", warnings),
            });
        }

        return decisions;
    }

    /// <summary>
    /// Аналог Rebuild-DecisionsAfterManualMatch (RrfqUi.ps1, строки 223-237):
    /// пересборка решений после ручных действий с сохранением Include по Id.
    /// </summary>
    public static void RebuildDecisionsAfterManualMatch(RrfqAnalysis analysis)
    {
        var oldIncludes = analysis.Decisions.ToDictionary(d => d.Id, d => d.Include, StringComparer.Ordinal);
        analysis.Decisions = BuildDecisions(analysis.RfqRows, analysis.Quotes, analysis.Priority);
        foreach (var decision in analysis.Decisions)
        {
            if (oldIncludes.TryGetValue(decision.Id, out var include))
            {
                decision.Include = include;
            }
        }
    }

    // ----- Чтение книг (строки 1113-1217) -----

    /// <summary>Аналог Get-RfqRowsFromWorkbook.</summary>
    private static List<RfqRow> GetRfqRowsFromWorkbook(dynamic excel, string path)
    {
        var rows = new List<RfqRow>();
        dynamic? wb = null;
        try
        {
            wb = RrfqExcel.OpenWorkbookReadOnly(excel, path);
            var sheetCount = (int)wb.Worksheets.Count;
            for (var i = 1; i <= sheetCount; i++)
            {
                dynamic ws = wb.Worksheets.Item(i);
                try
                {
                    var sheet = RrfqExcel.ReadSheetData(ws);
                    var map = FindTableMap(sheet, "Standard", "Delivery");
                    if (map is null)
                    {
                        continue;
                    }

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

                        var rfqRow = NewRfqRow(sheet, path, map, row);
                        if (rfqRow is not null)
                        {
                            rows.Add(rfqRow);
                        }
                    }
                }
                finally
                {
                    RrfqExcel.Release(ws);
                }
            }
        }
        finally
        {
            if (wb is not null)
            {
                try { wb.Close(false); } catch { /* Книга уже могла быть закрыта. */ }
                RrfqExcel.Release(wb);
            }
        }

        return rows;
    }

    /// <summary>Аналог Get-QuotesFromWorkbook.</summary>
    private static List<Quote> GetQuotesFromWorkbook(dynamic excel, string path, string supplier, double vatDivisor)
    {
        var quotes = new List<Quote>();
        dynamic? wb = null;
        try
        {
            wb = RrfqExcel.OpenWorkbookReadOnly(excel, path);
            var sheetCount = (int)wb.Worksheets.Count;
            for (var i = 1; i <= sheetCount; i++)
            {
                dynamic ws = wb.Worksheets.Item(i);
                try
                {
                    var sheet = RrfqExcel.ReadSheetData(ws);
                    var map = FindTableMap(sheet, "Standard", "Delivery", "Compel");
                    if (map is null)
                    {
                        continue;
                    }

                    var lastKey = string.Empty;
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

                        var rowKey = GetFieldText(sheet, map, "Value", row);
                        if (Regex.IsMatch(NormalizeKey(rowKey), "^(TOTAL|ИТОГО)$"))
                        {
                            continue;
                        }

                        var forcedKey = string.Empty;
                        if (map.Type == "Delivery" && string.IsNullOrWhiteSpace(rowKey) && !string.IsNullOrWhiteSpace(lastKey))
                        {
                            forcedKey = lastKey;
                        }

                        var quote = NewQuote(sheet, path, supplier, map, row, forcedKey, vatDivisor);
                        if (!QuoteHasAnyData(quote))
                        {
                            continue;
                        }

                        if (!string.IsNullOrWhiteSpace(quote.Key))
                        {
                            lastKey = quote.Key;
                        }

                        if (string.IsNullOrWhiteSpace(quote.Key))
                        {
                            quote.MatchStatus = "NoKey";
                        }

                        quotes.Add(quote);
                    }
                }
                finally
                {
                    RrfqExcel.Release(ws);
                }
            }
        }
        finally
        {
            if (wb is not null)
            {
                try { wb.Close(false); } catch { /* Книга уже могла быть закрыта. */ }
                RrfqExcel.Release(wb);
            }
        }

        return quotes;
    }

    // ----- Сопоставление (строки 1219-1270) -----

    /// <summary>Аналог Resolve-QuoteTargets.</summary>
    public static void ResolveQuoteTargets(List<RfqRow> rfqRows, List<Quote> quotes)
    {
        var byExactRow = new Dictionary<string, RfqRow>(StringComparer.Ordinal);
        var byKey = new Dictionary<string, List<RfqRow>>(StringComparer.Ordinal);
        foreach (var row in rfqRows)
        {
            byExactRow[$"{row.SheetName}::{row.Row}"] = row;
            if (!string.IsNullOrWhiteSpace(row.KeyNorm))
            {
                if (!byKey.TryGetValue(row.KeyNorm, out var list))
                {
                    list = new List<RfqRow>();
                    byKey[row.KeyNorm] = list;
                }

                list.Add(row);
            }
        }

        foreach (var quote in quotes)
        {
            quote.TargetId = null;
            quote.CandidateIds = new List<string>();

            if (quote.Format == "Standard")
            {
                var rowKey = $"{quote.SheetName}::{quote.Row}";
                if (byExactRow.TryGetValue(rowKey, out var rowCandidate))
                {
                    if (string.IsNullOrWhiteSpace(quote.KeyNorm) || quote.KeyNorm == rowCandidate.KeyNorm)
                    {
                        quote.TargetId = rowCandidate.Id;
                        quote.MatchStatus = "AutoRow";
                        if (string.IsNullOrWhiteSpace(quote.KeyNorm))
                        {
                            AddQuoteWarning(quote, "Value пустой, сопоставлено только по номеру строки");
                        }

                        continue;
                    }

                    AddQuoteWarning(
                        quote,
                        $"Value не совпадает со строкой RFQ {quote.Row}: в квоте '{quote.Key}', в RFQ '{rowCandidate.Key}'. Сопоставление по строке пропущено.");
                }
            }

            if (!string.IsNullOrWhiteSpace(quote.KeyNorm) && byKey.TryGetValue(quote.KeyNorm, out var candidates))
            {
                quote.CandidateIds = candidates.Select(c => c.Id).ToList();
                if (candidates.Count == 1)
                {
                    quote.TargetId = candidates[0].Id;
                    quote.MatchStatus = "AutoKey";
                }
                else
                {
                    quote.MatchStatus = "Ambiguous";
                }
            }
            else if (quote.MatchStatus != "NoKey")
            {
                quote.MatchStatus = "Unmatched";
            }
        }
    }

    /// <summary>Аналог Get-UnresolvedQuotes.</summary>
    public static List<Quote> GetUnresolvedQuotes(RrfqAnalysis analysis)
        => analysis.Quotes.Where(q => string.IsNullOrWhiteSpace(q.TargetId) || q.MatchStatus == "Ambiguous").ToList();

    // ----- Полный анализ (строки 1429-1470, без записи в базу) -----

    /// <summary>
    /// Аналог Invoke-Analysis: читает RFQ и файлы поставщиков, сопоставляет
    /// квоты и строит решения. Запись истории квот в БД намеренно исключена.
    /// </summary>
    public static RrfqAnalysis Analyze(string rfqPath, IReadOnlyList<SupplierFileEntry> suppliers, string priorityMode)
    {
        if (!File.Exists(rfqPath))
        {
            throw new InvalidOperationException("RFQ не найден: " + rfqPath);
        }

        var vatDivisor = RrfqConfigService.Current.VatDivisor;
        dynamic? excel = null;
        try
        {
            excel = RrfqExcel.OpenExcel();
            var rfqRows = GetRfqRowsFromWorkbook(excel, rfqPath);
            var quotes = new List<Quote>();
            foreach (var supplierFile in suppliers)
            {
                if (!File.Exists(supplierFile.Path))
                {
                    continue;
                }

                quotes.AddRange(GetQuotesFromWorkbook(excel, supplierFile.Path, supplierFile.Supplier, vatDivisor));
            }

            ResolveQuoteTargets(rfqRows, quotes);
            var decisions = BuildDecisions(rfqRows, quotes, priorityMode);

            return new RrfqAnalysis
            {
                RfqPath = rfqPath,
                Suppliers = suppliers.ToList(),
                RfqRows = rfqRows,
                Quotes = quotes,
                Decisions = decisions,
                Priority = priorityMode,
            };
        }
        finally
        {
            RrfqExcel.CloseExcel(excel);
        }
    }
}
