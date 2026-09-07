namespace ProcurementControl.Services;

/// <summary>Передаёт сгенерированные файлы в страницу сравнения RRFQ в рамках сеанса.</summary>
public static class RrfqImportQueue
{
    private static readonly List<(string Path, string Supplier)> Pending = new();
    public static event Action<string, string>? Added;

    public static void Add(string path, string supplier)
    {
        if (Added is { } handler)
        {
            handler(path, supplier);
            return;
        }
        if (!Pending.Any(item => string.Equals(item.Path, path, StringComparison.OrdinalIgnoreCase)))
            Pending.Add((path, supplier));
    }

    public static IReadOnlyList<(string Path, string Supplier)> TakePending()
    {
        var items = Pending.ToList(); Pending.Clear(); return items;
    }
}
