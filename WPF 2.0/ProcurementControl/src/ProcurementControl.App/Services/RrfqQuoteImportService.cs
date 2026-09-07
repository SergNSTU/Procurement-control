using System.Threading;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>Импорт квот из прикреплённых RRFQ-документов в общую базу квот.</summary>
public static class RrfqQuoteImportService
{
    public static List<Quote> ReadQuotes(IReadOnlyList<(string Path, string Supplier)> sources)
        => RunSta(() =>
        {
            var quotes = new List<Quote>();
            foreach (var source in sources.Where(s => !string.IsNullOrWhiteSpace(s.Path)))
            {
                var supplier = string.IsNullOrWhiteSpace(source.Supplier)
                    ? RrfqConfigService.GetSupplierFromFileName(source.Path)
                    : source.Supplier;
                quotes.AddRange(RrfqEngine.ReadQuotesFromWorkbook(source.Path, supplier));
            }
            return quotes;
        });

    public static int SaveQuotes(IReadOnlyList<Quote> quotes, IEnumerable<string> paths, string comment)
        => PurchaseWriteRepository.SaveQuoteHistory(quotes, string.Join("; ", paths), comment);

    private static T RunSta<T>(Func<T> action)
    {
        T? result = default;
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try { result = action(); }
            catch (Exception ex) { error = ex; }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (error is not null) throw new InvalidOperationException("Не удалось прочитать RRFQ: " + error.Message, error);
        return result!;
    }
}
