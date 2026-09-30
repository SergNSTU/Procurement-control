namespace ProcurementControl.Services;

/// <summary>Передаёт сгенерированные файлы в страницу сравнения RRFQ в рамках сеанса.</summary>
public static class RrfqImportQueue
{
    private static readonly List<(string Path, string Supplier)> Pending = new();
    public static event Action<string, string>? Added;
    public static event Func<RrfqDealComparisonContext, bool>? DealComparisonRequested;
    private static RrfqDealComparisonContext? PendingDealComparison;

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

    /// <summary>Передаёт в сравнение документы одной сделки вместе с их идентификаторами.</summary>
    public static bool OpenDealComparison(RrfqDealComparisonContext context)
    {
        if (DealComparisonRequested is { } handler)
            return handler.GetInvocationList().Cast<Func<RrfqDealComparisonContext, bool>>().All(action => action(context));
        PendingDealComparison = context;
        return true;
    }

    public static RrfqDealComparisonContext? TakePendingDealComparison()
    {
        var context = PendingDealComparison;
        PendingDealComparison = null;
        return context;
    }
}

public sealed record RrfqDealSourceDocument(long DocumentId, long SupplierId, string Supplier, string Path);

public sealed record RrfqDealComparisonContext(
    long DealId,
    string DealNumber,
    long RfqDocumentId,
    string RfqPath,
    IReadOnlyList<RrfqDealSourceDocument> SupplierDocuments);
