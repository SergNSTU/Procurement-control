using System.Globalization;
using System.Net.Http;
using System.Text.Json;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Подбор аналогов через внутренний JSON-API кросс-инструмента LCSC
/// (wmsc.lcsc.com, тот же хост, что и в PriceSearchService). Схема работы:
/// 1) GET .../alternate/part/sku?keyword=&lt;MPN&gt; — MPN → код позиции LCSC;
/// 2) GET .../alternate/part/list?productCode=&lt;код&gt; — таблица аналогов.
/// API неофициальный, но не требует ключей; обязателен заголовок Referer
/// с lcsc.com, иначе Akamai отдаёт Access Denied.
/// </summary>
public static class LcscCrossService
{
    /// <summary>Сколько аналогов запрашивать за раз (у LCSC максимум страницы).</summary>
    private const int PageSize = 50;

    public static async Task<CrossReferenceResult> SearchAsync(string mpn, bool inStockOnly, CancellationToken ct = default)
    {
        var query = mpn.Trim();
        using var client = CreateClient();

        var candidate = await ResolveProductAsync(client, query, ct).ConfigureAwait(false);
        if (candidate is null)
        {
            return new CrossReferenceResult { Query = query };
        }

        var (model, code, candidateCount) = candidate.Value;
        var rows = new List<CrossReferenceRow>();

        var listUrl = "https://wmsc.lcsc.com/ftps/wm/product/alternate/part/list"
                      + "?productCode=" + Uri.EscapeDataString(code)
                      + "&inStockOnly=" + (inStockOnly ? "true" : "false")
                      + "&currentPage=1&pageSize=" + PageSize;

        using (var response = await client.GetAsync(listUrl, ct).ConfigureAwait(false))
        {
            response.EnsureSuccessStatusCode();
            using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
            var root = document.RootElement;
            RequireOk(root);

            if (root.TryGetProperty("result", out var result) && result.ValueKind == JsonValueKind.Object)
            {
                if (result.TryGetProperty("rawMaterial", out var original) && original.ValueKind == JsonValueKind.Object)
                {
                    rows.Add(ToRow(original, "Оригинал"));
                }

                if (result.TryGetProperty("pageInfo", out var pageInfo)
                    && pageInfo.TryGetProperty("dataList", out var dataList)
                    && dataList.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in dataList.EnumerateArray())
                    {
                        rows.Add(ToRow(item, MatchTypeLabel(GetStringOrDefault(item, "matchType"))));
                    }
                }
            }
        }

        return new CrossReferenceResult
        {
            Query = query,
            ExactMatch = string.Equals(model, query, StringComparison.OrdinalIgnoreCase),
            CandidateCount = candidateCount,
            ResolvedModel = model,
            ResolvedCode = code,
            Rows = rows,
        };
    }

    /// <summary>MPN → код позиции LCSC. При неоднозначном запросе берётся точное совпадение, иначе первый кандидат.</summary>
    private static async Task<(string Model, string Code, int Count)?> ResolveProductAsync(
        HttpClient client, string query, CancellationToken ct)
    {
        var url = "https://wmsc.lcsc.com/ftps/wm/product/alternate/part/sku?keyword="
                  + Uri.EscapeDataString(query);
        using var response = await client.GetAsync(url, ct).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();
        using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
        var root = document.RootElement;
        RequireOk(root);

        if (!root.TryGetProperty("result", out var result)
            || result.ValueKind != JsonValueKind.Object
            || !result.TryGetProperty("dataList", out var dataList)
            || dataList.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        string? fallbackModel = null, fallbackCode = null;
        var count = 0;
        foreach (var item in dataList.EnumerateArray())
        {
            var model = GetStringOrDefault(item, "productModel");
            var code = GetStringOrDefault(item, "productCode");
            if (model is null || code is null)
            {
                continue;
            }

            count++;
            if (string.Equals(model, query, StringComparison.OrdinalIgnoreCase))
            {
                return (model, code, count);
            }

            fallbackModel ??= model;
            fallbackCode ??= code;
        }

        return fallbackCode is null ? null : (fallbackModel ?? query, fallbackCode, count);
    }

    private static CrossReferenceRow ToRow(JsonElement product, string type)
    {
        var stock = GetStringOrDefault(product, "stockNumber");
        return new CrossReferenceRow
        {
            Type = type,
            Mpn = GetStringOrDefault(product, "productModel") ?? string.Empty,
            LcscCode = GetStringOrDefault(product, "productCode") ?? string.Empty,
            Brand = GetStringOrDefault(product, "brandNameEn") ?? string.Empty,
            Package = GetStringOrDefault(product, "encapStandard") ?? string.Empty,
            Category = GetStringOrDefault(product, "catalogName") ?? string.Empty,
            Stock = stock is null or "0" ? "—" : stock,
            Moq = GetStringOrDefault(product, "minBuyNumber") ?? string.Empty,
            PriceText = FormatPrices(product),
            Description = GetStringOrDefault(product, "productIntroEn") ?? string.Empty,
            Link = GetStringOrDefault(product, "url") ?? string.Empty,
            Datasheet = GetStringOrDefault(product, "pdfUrl") ?? string.Empty,
        };
    }

    /// <summary>Первые ступени лесенки цен: «1+ $2.17 · 10+ $1.83».</summary>
    private static string FormatPrices(JsonElement product)
    {
        if (!product.TryGetProperty("productPriceList", out var priceList)
            || priceList.ValueKind != JsonValueKind.Array
            || priceList.GetArrayLength() == 0)
        {
            return string.Empty;
        }

        var tiers = new List<string>();
        foreach (var priceBreak in priceList.EnumerateArray().Take(3))
        {
            var ladder = ConvertNumber(GetStringOrDefault(priceBreak, "ladder"));
            var usd = ConvertNumber(GetStringOrDefault(priceBreak, "usdPrice")
                                    ?? GetStringOrDefault(priceBreak, "productPrice"));
            if (ladder is null || usd is null)
            {
                continue;
            }

            tiers.Add(string.Create(CultureInfo.InvariantCulture, $"{ladder}+ ${usd:0.00##}"));
        }

        return string.Join(" · ", tiers);
    }

    /// <summary>Подписи типов замены; значение matchType у LCSC строковое.</summary>
    private static string MatchTypeLabel(string? matchType) => matchType switch
    {
        "2" => "Другая упаковка",
        "4" => "Похожий",
        null or "" => "Аналог",
        _ => "Аналог (тип " + matchType + ")",
    };

    private static void RequireOk(JsonElement root)
    {
        if (root.TryGetProperty("code", out var code)
            && code.ValueKind == JsonValueKind.Number
            && code.GetInt32() != 200)
        {
            var message = root.TryGetProperty("msg", out var msg) && msg.ValueKind == JsonValueKind.String
                ? msg.GetString()
                : null;
            throw new InvalidOperationException("LCSC: " + (message ?? "ошибка запроса " + code.GetInt32()));
        }
    }

    private static HttpClient CreateClient()
    {
        var handler = new SocketsHttpHandler
        {
            AllowAutoRedirect = true,
        };
        var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(15) };
        client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
        client.DefaultRequestHeaders.Add("Referer", "https://www.lcsc.com/tools/cross-reference");
        return client;
    }

    private static string? GetStringOrDefault(JsonElement element, string name)
        => element.TryGetProperty(name, out var value) && value.ValueKind is JsonValueKind.String or JsonValueKind.Number
            ? value.ToString()
            : null;

    private static double? ConvertNumber(string? value)
        => double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var number)
            ? number
            : null;
}
