using System.Globalization;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Data.Sqlite;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Порт app/modules/PriceSearch.ps1: источники цен (свои квоты, Globalist,
/// Компэл, LCSC, Промэлектроника), нормализация чисел, общий сбор
/// результатов и текст запроса для GPT. HTTP-запросы выполняются асинхронно,
/// поэтому страница не блокирует интерфейс.
/// </summary>
public static class PriceSearchService
{
    private const string PromelecSource = "Промэлектроника";

    /// <summary>Аналог Convert-PriceSearchNumber: разбор чисел с пробелами и запятой.</summary>
    public static double? ConvertNumber(string? value)
    {
        if (value is null)
        {
            return null;
        }

        var text = value.Replace('\u00A0', ' ').Replace(" ", "").Trim().Replace(',', '.');
        return double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var number)
            ? number
            : null;
    }

    /// <summary>Аналог Get-PriceSearchGptPrompt.</summary>
    public static string GetGptPrompt(string pn, int quantity)
        => $"Найди цену, производителя, MOQ, Lead Time и наличие для PN: {pn}, количество: {quantity}. По возможности верни ссылки, сроки поставки и валюту.";

    /// <summary>Аналог Get-PriceSearchResults: последовательный обход всех источников.</summary>
    public static async Task<List<PriceSearchRow>> GetResultsAsync(SqliteConnection connection, string pn, int quantity)
    {
        var items = new List<PriceSearchRow>();

        foreach (var source in new Func<Task<List<PriceSearchRow>>>[]
        {
            () => Task.FromResult(GetOurQuotes(connection, pn)),
            () => Task.FromResult(GetGlobalistQuotes(connection, pn, quantity)),
            () => Task.FromResult(GetCompelQuotes(connection, pn)),
            () => GetLcscQuotesAsync(pn, quantity),
            () => GetPromelecQuotesAsync(pn, quantity),
        })
        {
            try
            {
                items.AddRange(await source().ConfigureAwait(false));
            }
            catch
            {
                // Отдельный источник не должен ронять общий поиск.
            }
        }

        return items;
    }

    /// <summary>Аналог Get-PriceSearchOurQuotes: свои квоты из quote_history.</summary>
    private static List<PriceSearchRow> GetOurQuotes(SqliteConnection connection, string pn)
    {
        var result = new List<PriceSearchRow>();
        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT q.pn, q.supplier, q.unit_price, q.lead_time, q.quote_date, q.mfg
FROM quote_history q
WHERE lower(trim(IFNULL(q.pn,''))) LIKE lower(@needle)
ORDER BY q.quote_date DESC, q.id DESC
LIMIT 200";
        command.Parameters.AddWithValue("@needle", "%" + pn.Trim() + "%");

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var price = NullableDouble(reader, 2);
            if (price is null)
            {
                continue;
            }

            result.Add(NewRow("Наш прайс", Text(reader, 0), Text(reader, 5), price, "", "", "", Text(reader, 3), Text(reader, 4), ""));
        }
        return result;
    }

    /// <summary>Аналог Get-PriceSearchGlobalistQuotes: Globalist с фильтром по количеству.</summary>
    private static List<PriceSearchRow> GetGlobalistQuotes(SqliteConnection connection, string pn, int quantity)
    {
        var result = new List<PriceSearchRow>();
        var repository = new QuoteBaseRepository(connection);
        foreach (var row in repository.GetGlobalistQuotes(pn, 500))
        {
            if (row.UnitPrice is null)
            {
                continue;
            }

            if (quantity > 0)
            {
                var available = ConvertNumber(row.Qty);
                if (available is not null && available.Value < quantity)
                {
                    continue;
                }
            }

            result.Add(NewRow("Globalist", row.PN, row.Brand, row.UnitPrice, "", row.Moq, row.Stock, row.LeadTime, row.ImportedAt, ""));
        }
        return result;
    }

    /// <summary>
    /// Аналог Get-PriceSearchCompelQuotes. Внимание: оригинал выбирает квоты
    /// поставщиков БЕЗ «компэл» в названии, но помечает их источником «Компэл»
    /// (поведение перенесено один в один).
    /// </summary>
    private static List<PriceSearchRow> GetCompelQuotes(SqliteConnection connection, string pn)
    {
        var result = new List<PriceSearchRow>();
        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT q.pn, q.supplier, q.unit_price, q.lead_time, q.quote_date, q.mfg
FROM quote_history q
WHERE lower(trim(IFNULL(q.pn,''))) LIKE lower(@needle)
  AND lower(IFNULL(q.supplier,'')) NOT LIKE '%компэл%'
ORDER BY q.quote_date DESC, q.id DESC
LIMIT 200";
        command.Parameters.AddWithValue("@needle", "%" + pn.Trim() + "%");

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var price = NullableDouble(reader, 2);
            if (price is null)
            {
                continue;
            }

            result.Add(NewRow("Компэл", Text(reader, 0), Text(reader, 5), price, "", "", "", Text(reader, 3), Text(reader, 4), ""));
        }
        return result;
    }

    /// <summary>Аналог Get-PriceSearchLcscQuotes: поиск по каталогу LCSC.</summary>
    private static async Task<List<PriceSearchRow>> GetLcscQuotesAsync(string pn, int quantity)
    {
        var body = JsonSerializer.Serialize(new
        {
            keyword = pn.Trim(),
            catalogIdList = Array.Empty<string>(),
            brandIdList = Array.Empty<string>(),
            encapValueList = Array.Empty<string>(),
            isStock = false,
            isOtherSuppliers = false,
            isAsianBrand = false,
            isDeals = false,
            isEnvironment = false,
            paramNameValueMap = new Dictionary<string, string>(),
            currentPage = 1,
            pageSize = 50,
        });

        try
        {
            using var client = CreateClient(TimeSpan.FromSeconds(12));
            using var request = new HttpRequestMessage(HttpMethod.Post, "https://wmsc.lcsc.com/ftps/wm/product/query/list")
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json"),
            };
            request.Headers.Add("User-Agent", "Mozilla/5.0");
            request.Headers.Add("Referer", "https://www.lcsc.com/products");

            using var response = await client.SendAsync(request).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                return new List<PriceSearchRow>();
            }

            using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            var root = document.RootElement;
            if (!root.TryGetProperty("code", out var code) || code.GetInt32() != 200)
            {
                return new List<PriceSearchRow>();
            }

            if (!root.TryGetProperty("result", out var resultElement)
                || !resultElement.TryGetProperty("dataList", out var dataList))
            {
                return new List<PriceSearchRow>();
            }

            var items = new List<PriceSearchRow>();
            foreach (var product in dataList.EnumerateArray())
            {
                double? price = null;
                if (product.TryGetProperty("productPriceList", out var priceList))
                {
                    JsonElement? selected = null;
                    foreach (var priceBreak in priceList.EnumerateArray())
                    {
                        var ladder = ConvertNumber(GetStringOrDefault(priceBreak, "ladder"));
                        if (ladder is null)
                        {
                            continue;
                        }

                        if (selected is null
                            || (quantity >= ladder.Value && ladder.Value >= ConvertNumber(GetStringOrDefault(selected.Value, "ladder"))!.Value))
                        {
                            selected = priceBreak;
                        }
                    }

                    if (selected is null && priceList.GetArrayLength() > 0)
                    {
                        selected = priceList[0];
                    }

                    if (selected is not null)
                    {
                        price = ConvertNumber(GetStringOrDefault(selected.Value, "usdPrice"))
                            ?? ConvertNumber(GetStringOrDefault(selected.Value, "productPrice"));
                    }
                }

                var lead = "Нет в наличии";
                if (product.TryGetProperty("flashSaleProductPO", out var flashSale)
                    && flashSale.ValueKind == JsonValueKind.Object
                    && flashSale.TryGetProperty("arriveDays", out var arriveDays)
                    && arriveDays.ValueKind == JsonValueKind.Number)
                {
                    lead = arriveDays.GetInt32() + " дн.";
                }

                if (price is null)
                {
                    continue;
                }

                items.Add(NewRow(
                    "LCSC",
                    GetStringOrDefault(product, "productModel"),
                    GetStringOrDefault(product, "brandNameEn"),
                    price,
                    "USD",
                    GetStringOrDefault(product, "minBuyNumber"),
                    GetStringOrDefault(product, "stockNumber"),
                    lead,
                    NowText(),
                    GetStringOrDefault(product, "url")));
            }
            return items;
        }
        catch
        {
            return new List<PriceSearchRow>();
        }
    }

    /// <summary>Аналог Get-PriceSearchPromelecQuotes: разбор страниц Промэлектроники.</summary>
    private static async Task<List<PriceSearchRow>> GetPromelecQuotesAsync(string pn, int quantity)
    {
        var searchUrl = "https://www.promelec.ru/search/?query=" + Uri.EscapeDataString(pn.Trim());
        try
        {
            using var client = CreateClient(TimeSpan.FromSeconds(15));
            var html = await client.GetStringAsync(searchUrl).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(html))
            {
                return new List<PriceSearchRow>();
            }

            var usdRate = await GetCbrUsdRateAsync().ConfigureAwait(false);
            if (usdRate is null || usdRate.Value <= 0)
            {
                return new List<PriceSearchRow>();
            }

            var items = new List<PriceSearchRow>();
            var pnUpper = pn.Trim().ToUpperInvariant();
            var itemBlocks = html.Split("table-list__item");

            for (var i = 1; i < itemBlocks.Length; i++)
            {
                var block = itemBlocks[i];

                var itemPn = MatchGroup(block, @"product-preview__title[^>]*>\s*([^<]+?)\s*</a", 1);
                if (itemPn.ToUpperInvariant() != pnUpper)
                {
                    continue;
                }

                var mfg = MatchGroup(block, "Производитель[^<]*<[^>]*>\\s*<[^>]*>\\s*([^<]+?)\\s*<", 1);
                var productLink = MatchGroup(block, "href=\"(https?://[^\"]*?/product/\\d+/)\"", 1);
                if (string.IsNullOrWhiteSpace(productLink))
                {
                    continue;
                }

                try
                {
                    var productHtml = await client.GetStringAsync(productLink).ConfigureAwait(false);
                    var warehouseBlocks = productHtml.Split("js-accordion-wrap");
                    for (var w = 1; w < warehouseBlocks.Length; w++)
                    {
                        var warehouse = warehouseBlocks[w];
                        var listItems = warehouse.Split("table-popup-list__item");
                        for (var j = 1; j < listItems.Length; j++)
                        {
                            var item = listItems[j];
                            ParsePromelecListItem(item, itemPn, mfg, productLink, quantity, usdRate.Value, items);
                        }
                    }
                }
                catch
                {
                    // Карточка товара может не открыться — пропускаем.
                }
            }
            return items;
        }
        catch
        {
            return new List<PriceSearchRow>();
        }
    }

    /// <summary>Разбор одной складской позиции Промэлектроники (аналог внутреннего цикла по $li).</summary>
    private static void ParsePromelecListItem(
        string item,
        string itemPn,
        string mfg,
        string productLink,
        int quantity,
        double usdRate,
        List<PriceSearchRow> items)
    {
        string lead;
        var leadDays = Regex.Match(item, "(\\d+)\\s*<[^>]*>\\s*дн");
        if (!leadDays.Success)
        {
            leadDays = Regex.Match(item, "(\\d+)\\s*дн");
        }

        if (leadDays.Success)
        {
            lead = leadDays.Groups[1].Value + " дн.";
        }
        else if (Regex.IsMatch(item, "В\\s+наличи"))
        {
            lead = "В наличии";
        }
        else
        {
            lead = string.Empty;
        }

        var tierMatches = Regex.Matches(item, "data-min-qty=\"(\\d+)\"");
        var priceMatches = Regex.Matches(item, "([\\d]+[,\\d]*,\\d+)\\s*<span[^>]*>₽</span>");
        var tierCount = Math.Min(tierMatches.Count, priceMatches.Count);
        if (tierCount == 0)
        {
            return;
        }

        (int Qty, double Price)? selected = null;
        for (var t = 0; t < tierCount; t++)
        {
            var qty = int.Parse(tierMatches[t].Groups[1].Value, CultureInfo.InvariantCulture);
            var priceText = priceMatches[t].Groups[1].Value.Replace(',', '.');
            if (double.TryParse(priceText, NumberStyles.Any, CultureInfo.InvariantCulture, out var tierPrice)
                && quantity >= qty
                && (selected is null || qty > selected.Value.Qty))
            {
                selected = (qty, tierPrice);
            }
        }

        if (selected is null)
        {
            var firstPriceText = priceMatches[0].Groups[1].Value.Replace(',', '.');
            if (double.TryParse(firstPriceText, NumberStyles.Any, CultureInfo.InvariantCulture, out var firstPrice))
            {
                selected = (0, firstPrice);
            }
        }

        if (selected is null || selected.Value.Price <= 0)
        {
            return;
        }

        var stock = string.Empty;
        var stockMatch = Regex.Match(item, ">\\s*(\\d[\\d\\s]*)\\s*<span[^>]*>шт");
        if (stockMatch.Success)
        {
            stock = Regex.Replace(stockMatch.Groups[1].Value, "\\s+", "").Trim();
        }

        var priceUsd = Math.Round(selected.Value.Price / usdRate / 1.22, 4);

        // Дедупликация: одинаковые срок+цена уже могли попасть из другого склада.
        foreach (var existing in items)
        {
            if (existing.Source == PromelecSource
                && existing.PN == itemPn
                && existing.LeadTime == lead
                && existing.Price == priceUsd)
            {
                return;
            }
        }

        items.Add(NewRow(PromelecSource, itemPn, mfg, priceUsd, "USD", "", stock, lead, NowText(), productLink));
    }

    /// <summary>Аналог Get-CbrUsdRate: курс USD ЦБ РФ.</summary>
    private static async Task<double?> GetCbrUsdRateAsync()
    {
        try
        {
            using var client = CreateClient(TimeSpan.FromSeconds(10));
            var content = await client.GetStringAsync("https://www.cbr.ru/scripts/XML_daily.asp").ConfigureAwait(false);
            var document = XDocument.Parse(content);
            foreach (var currency in document.Descendants("Valute"))
            {
                if ((string?)currency.Element("CharCode") == "USD")
                {
                    return ConvertNumber((string?)currency.Element("Value"));
                }
            }
        }
        catch
        {
            // Курс недоступен — источник на его основе просто пуст.
        }
        return null;
    }

    private static PriceSearchRow NewRow(
        string source, string? pn, string? mfg, double? price, string? currency,
        string? moq, string? stock, string? lead, string date, string? link)
        => new()
        {
            Source = source,
            PN = pn ?? string.Empty,
            Manufacturer = mfg ?? string.Empty,
            Price = price,
            Currency = currency ?? string.Empty,
            Moq = moq ?? string.Empty,
            Stock = stock ?? string.Empty,
            LeadTime = lead ?? string.Empty,
            Date = date,
            Link = link ?? string.Empty,
        };

    private static HttpClient CreateClient(TimeSpan timeout)
    {
        var handler = new SocketsHttpHandler
        {
            AllowAutoRedirect = true,
        };
        var client = new HttpClient(handler) { Timeout = timeout };
        client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
        return client;
    }

    private static string? GetStringOrDefault(JsonElement element, string name)
        => element.TryGetProperty(name, out var value) && value.ValueKind is JsonValueKind.String or JsonValueKind.Number
            ? value.ToString()
            : null;

    private static string MatchGroup(string input, string pattern, int group)
    {
        var match = Regex.Match(input, pattern);
        return match.Success ? match.Groups[group].Value.Trim() : string.Empty;
    }

    private static string Text(SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? string.Empty : Convert.ToString(reader.GetValue(ordinal)) ?? string.Empty;

    private static double? NullableDouble(SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? null : Convert.ToDouble(reader.GetValue(ordinal), CultureInfo.InvariantCulture);

    private static string NowText()
        => DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
}
