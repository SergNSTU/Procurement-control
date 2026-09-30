namespace ProcurementControl.Services;

/// <summary>
/// Общий справочник «email → короткое имя поставщика». Используется в обе
/// стороны: почтовый импорт привязывает письма по адресу отправителя,
/// рассылка RFQ находит адрес по имени поставщика. Ключи в нижнем регистре.
/// </summary>
public static class SupplierDirectory
{
    private static readonly Dictionary<string, string> SupplierByEmail = new(StringComparer.OrdinalIgnoreCase)
    {
        ["sales5@chipera.net"] = "CPR",
        ["huananshengke@163.com"] = "HNN",
        ["vivi@bomzon.com"] = "BMZ",
        ["kotakoti@optechic.com"] = "OPT",
        ["sales12@allpartsic.cn"] = "CRL",
        ["sales17@allpartsic.cn"] = "CRL",
        ["xinkeweiltd@126.com"] = "TIM",
        ["purp-d@moscentech.com"] = "Ben",
        ["purp-e@moscentech.com"] = "Ben",
        ["purp-b@moscentech.com"] = "Ben",
        ["benjamin401@163.com"] = "Ben",
        ["anna@3-k.com"] = "3K",
    };

    /// <summary>Короткое имя поставщика по адресу отправителя; null — вне справочника.</summary>
    public static string? FindSupplierByEmail(string? email)
        => email is not null && SupplierByEmail.TryGetValue(email.Trim(), out var supplier) ? supplier : null;

    /// <summary>
    /// Адрес поставщика по короткому имени (первое совпадение в справочнике).
    /// Сравнение без учёта регистра; null — адрес неизвестен.
    /// </summary>
    public static string? FindEmailBySupplier(string? supplier)
    {
        if (string.IsNullOrWhiteSpace(supplier))
        {
            return null;
        }

        foreach (var pair in SupplierByEmail)
        {
            if (string.Equals(pair.Value, supplier.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return pair.Key;
            }
        }

        return null;
    }
}
