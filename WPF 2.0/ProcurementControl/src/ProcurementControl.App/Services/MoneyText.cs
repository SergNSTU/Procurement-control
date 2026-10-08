using System.Globalization;
using System.Text.RegularExpressions;

namespace ProcurementControl.Services;

/// <summary>Разбор «грязных» денежных строк из БД (пробелы, символы валют, запятые) в decimal.</summary>
internal static class MoneyText
{
    public static decimal ParseUsd(string? value)
    {
        var text = value?.Trim();
        if (string.IsNullOrWhiteSpace(text)) return 0m;
        text = Regex.Replace(text, @"[^0-9,\.\-]", string.Empty);
        if (text.Count(c => c == ',') > 0 && text.Count(c => c == '.') > 0)
        {
            var decimalSeparator = Math.Max(text.LastIndexOf(','), text.LastIndexOf('.'));
            text = text[..decimalSeparator].Replace(",", string.Empty).Replace(".", string.Empty) + "." + text[(decimalSeparator + 1)..];
        }
        else if (text.Count(c => c == ',') == 1)
        {
            text = text.Replace(',', '.');
        }
        return decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var result) ? result : 0m;
    }
}
