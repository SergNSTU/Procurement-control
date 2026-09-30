using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Разбор HTML-файла расчёта SDS Компэла — перенос модуля
/// app/modules/CompelParser.ps1 (Read-CompelHtmlText, Parse-CompelHtml,
/// New-CompelSummary и вспомогательные функции). База данных не используется.
/// </summary>
public static class CompelParserService
{
    /// <summary>Читает HTML: UTF-8 без BOM, при charset=windows-1251 — перечитывает в 1251.</summary>
    public static string ReadHtmlText(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("HTML-файл не найден: " + path, path);
        }

        var bytes = File.ReadAllBytes(path);
        var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: false);
        var text = utf8.GetString(bytes);
        if (Regex.IsMatch(text, @"charset\s*=\s*""?windows-1251", RegexOptions.IgnoreCase))
        {
            text = Encoding.GetEncoding(1251).GetString(bytes);
        }
        return text;
    }

    /// <summary>Перенос Parse-CompelHtml: позиции + сводка.</summary>
    public static CompelParseResult Parse(string path)
    {
        var text = ReadHtmlText(path);
        var chunks = Regex.Split(text, "(?=<tr data-is=\"ac-row\" class=\"ac-row)");
        var rows = new List<CompelRow>();

        foreach (var chunk in chunks)
        {
            if (!chunk.StartsWith("<tr data-is=\"ac-row\"", StringComparison.Ordinal))
            {
                continue;
            }

            var matchedText = CleanText(RegexFirst(@"<a[^>]*class=""ac-item-pn-link""[^>]*>(.*?)</a>", chunk));
            var (part, manufacturer) = SplitMatchedPart(matchedText);

            var quantityText = CleanText(RegexFirst(@"<span[^>]*class=""ac-item-offer-qty""[^>]*>(.*?)</span>", chunk));
            quantityText = Regex.Replace(quantityText, @"\s*x\s*$", "");

            double? unitPrice = null;
            var currency = string.Empty;
            var unitMatch = Regex.Match(
                chunk,
                @"<span class=""ac-item-offer-price"">.*?<price2\s+val=""([^""]+)""\s+cur=""([^""]+)""",
                RegexOptions.Singleline);
            if (unitMatch.Success)
            {
                unitPrice = ParseFloat(unitMatch.Groups[1].Value);
                currency = unitMatch.Groups[2].Value;
            }

            double? total = null;
            var totalCurrency = string.Empty;
            var totalBlock = RegexFirst(@"<div class=""ac-row-total-sum[^""]*""[^>]*>(.*?)</div>", chunk);
            if (!string.IsNullOrWhiteSpace(totalBlock))
            {
                var totalMatch = Regex.Match(totalBlock, @"<price2\s+val=""([^""]+)""\s+cur=""([^""]+)""", RegexOptions.Singleline);
                if (totalMatch.Success)
                {
                    total = ParseFloat(totalMatch.Groups[1].Value);
                    totalCurrency = totalMatch.Groups[2].Value;
                }
            }

            rows.Add(new CompelRow
            {
                Number = ParseInt(CleanText(RegexFirst(@"<div class=""ac-row-num"">(.*?)</div>", chunk))),
                SourceName = CleanText(RegexFirst(@"<span[^>]*class=""ac-row-name[^>]*>(.*?)</span>", chunk)),
                MatchedPart = part,
                Manufacturer = manufacturer,
                Package = CleanText(RegexFirst(@"<span[^>]*class=""ac-part-mpq""[^>]*>(.*?)</span>", chunk)),
                StockCount = ParseInt(CleanText(RegexFirst(@"<span[^>]*class=""ac-item-stocks-key""[^>]*>(.*?)</span>", chunk))),
                LeadTime = CleanText(RegexFirst(@"<span[^>]*class=""ac-item-offer-dlv""[^>]*>(.*?)</span>", chunk)),
                Quantity = ParseInt(quantityText),
                UnitPrice = unitPrice,
                Currency = currency,
                Total = total,
                TotalCurrency = totalCurrency,
            });
        }

        return new CompelParseResult
        {
            Rows = rows,
            Summary = BuildSummary(text, rows),
        };
    }

    /// <summary>Перенос New-CompelSummary: подсчёты + итоги сайта (дешевле/быстрее).</summary>
    public static CompelSummary BuildSummary(string htmlText, IReadOnlyList<CompelRow> rows)
    {
        var summary = new CompelSummary { TotalRows = rows.Count };

        foreach (var row in rows)
        {
            if (row.Total is not null)
            {
                summary.PricedRows++;
                summary.SelectedTotalUsd += row.Total.Value;
            }
            else if (row.Number is not null)
            {
                summary.MissingRows.Add(row.Number.Value);
            }

            var days = GetLeadDays(row.LeadTime);
            if (days is not null && (summary.LongestLeadDays is null || days > summary.LongestLeadDays))
            {
                summary.LongestLeadDays = days;
                summary.LongestLeadTime = row.LeadTime;
            }
        }

        var cheap = MatchSiteTotal(htmlText, "cheap");
        summary.CheapTotalUsd = cheap.TotalUsd;
        summary.CheapLeadTime = CleanText(cheap.LeadTime);

        var optimal = MatchSiteTotal(htmlText, "optimal");
        summary.OptimalTotalUsd = optimal.TotalUsd;
        summary.OptimalLeadTime = CleanText(optimal.LeadTime);

        return summary;
    }

    private static (double? TotalUsd, string LeadTime) MatchSiteTotal(string htmlText, string dataSet)
    {
        var pattern = "<div data-set=\"" + dataSet + "\"[^>]*>.*?<price2\\s+val=\"([^\"]+)\"\\s+cur=\"USD\".*?"
            + "<span class=\"bom-settings-offer-type-dlv\">([^<]+)</span>";
        var match = Regex.Match(htmlText, pattern, RegexOptions.Singleline);
        if (!match.Success)
        {
            return (null, string.Empty);
        }
        return (ParseFloat(match.Groups[1].Value), match.Groups[2].Value);
    }

    /// <summary>Перенос Get-CompelRegexFirst: первая группа первого совпадения.</summary>
    private static string RegexFirst(string pattern, string text)
    {
        var match = Regex.Match(text, pattern, RegexOptions.Singleline);
        return match.Success && match.Groups.Count > 1 ? match.Groups[1].Value : string.Empty;
    }

    /// <summary>Перенос ConvertTo-CompelCleanText: убрать теги, декодировать, сжать пробелы.</summary>
    public static string CleanText(string? value)
    {
        if (value is null)
        {
            return string.Empty;
        }
        var text = Regex.Replace(value, "<[^>]+>", " ");
        text = WebUtility.HtmlDecode(text);
        if (text is null)
        {
            return string.Empty;
        }
        text = text.Replace('\u00A0', ' ');
        return Regex.Replace(text, @"\s+", " ").Trim();
    }

    /// <summary>Перенос ConvertTo-CompelInt.</summary>
    public static int? ParseInt(string? value)
    {
        if (value is null)
        {
            return null;
        }
        var text = value.Replace(" ", "").Trim();
        return Regex.IsMatch(text, @"^-?\d+$") ? int.Parse(text, CultureInfo.InvariantCulture) : null;
    }

    /// <summary>Перенос ConvertTo-CompelFloat: запятая -> точка, invariant.</summary>
    public static double? ParseFloat(string? value)
    {
        if (value is null)
        {
            return null;
        }
        var text = value.Replace(" ", "").Replace(',', '.').Trim();
        if (text.Length == 0)
        {
            return null;
        }
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number : null;
    }

    /// <summary>Перенос Split-CompelMatchedPart: «Часть (Производитель)» -> часть и производитель.</summary>
    public static (string Part, string Manufacturer) SplitMatchedPart(string value)
    {
        var text = value.Trim();
        var match = Regex.Match(text, @"^(.+?)\s*\(([^()]*)\)$");
        return match.Success
            ? (match.Groups[1].Value.Trim(), match.Groups[2].Value.Trim())
            : (text, string.Empty);
    }

    /// <summary>Перенос Get-CompelLeadDays: первое число из текста срока.</summary>
    public static int? GetLeadDays(string value)
    {
        var match = Regex.Match(value, @"\d+");
        return match.Success ? int.Parse(match.Value, CultureInfo.InvariantCulture) : null;
    }

    /// <summary>Перенос Get-CompelMaxLeadTime: самый долгий срок по позициям.</summary>
    public static string GetMaxLeadTime(IReadOnlyList<CompelRow> rows)
    {
        int? maxDays = null;
        var maxText = string.Empty;
        foreach (var row in rows)
        {
            var days = GetLeadDays(row.LeadTime);
            if (days is not null && (maxDays is null || days > maxDays))
            {
                maxDays = days;
                maxText = row.LeadTime;
            }
        }
        return maxText;
    }
}
