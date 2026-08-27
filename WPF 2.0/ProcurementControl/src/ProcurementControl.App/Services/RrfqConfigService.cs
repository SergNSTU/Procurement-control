using System.IO;
using System.Text.Json;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Конфигурация раздела Сравнение RRFQ — перенос Get-AppConfig и
/// Get-SupplierFromFileName из app/modules/AppConfig.ps1. Читает общий файл
/// настроек оригинала (только чтение); при отсутствии файла использует тот же
/// встроенный дефолт, что и оригинальный Get-AppConfig.
/// </summary>
public static class RrfqConfigService
{
    private static readonly Lazy<RrfqConfig> Config = new(LoadConfig);

    /// <summary>Текущая конфигурация (общая на приложение).</summary>
    public static RrfqConfig Current => Config.Value;

    /// <summary>Имена поставщиков для ручного ввода (аналог Get-SupplierNames).</summary>
    public static IReadOnlyList<string> GetSupplierNames()
        => Current.Suppliers.Select(s => s.Name).ToList();

    /// <summary>Аналог Get-SupplierFromFileName: поиск алиаса в имени файла.</summary>
    public static string GetSupplierFromFileName(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path).ToLowerInvariant();
        foreach (var supplier in Current.Suppliers)
        {
            foreach (var alias in supplier.Aliases)
            {
                if (!string.IsNullOrEmpty(alias) && name.Contains(alias.ToLowerInvariant(), StringComparison.Ordinal))
                {
                    return supplier.Name;
                }
            }
        }

        return Current.OtherSupplier;
    }

    private static RrfqConfig LoadConfig()
    {
        try
        {
            var configPath = Path.Combine(AppPaths.AppRoot, "config", "suppliers.json");
            if (File.Exists(configPath))
            {
                var text = File.ReadAllText(configPath);
                var parsed = JsonSerializer.Deserialize<RrfqConfig>(text, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                });
                if (parsed is not null && parsed.Suppliers.Count > 0)
                {
                    return parsed;
                }
            }
        }
        catch
        {
            // Повреждённый конфиг — используем встроенный дефолт, как оригинал.
        }

        return BuildDefault();
    }

    /// <summary>Встроенный дефолт из Get-AppConfig (AppConfig.ps1, строки 10-22).</summary>
    private static RrfqConfig BuildDefault() => new()
    {
        Suppliers =
        {
            new SupplierInfo { Name = "Ben", Aliases = new List<string> { "ben", "бен" } },
            new SupplierInfo { Name = "BMZ", Aliases = new List<string> { "bmz", "бмз" } },
            new SupplierInfo { Name = "HNN", Aliases = new List<string> { "hnn" } },
            new SupplierInfo { Name = "CPR", Aliases = new List<string> { "cpr" } },
            new SupplierInfo { Name = "Компэл", Aliases = new List<string> { "compel", "компэл", "компел" } },
            new SupplierInfo { Name = "Промэлектроника", Aliases = new List<string> { "promelec", "promelektronika", "промэлектроника" } },
            new SupplierInfo { Name = "Др.", Aliases = new List<string> { "other", "др", "другой" } },
        },
        OtherSupplier = "Др.",
        VatDivisor = 1.22,
    };
}
