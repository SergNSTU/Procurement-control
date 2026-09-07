using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>Чтение статистики и управление справочником производителей.</summary>
public sealed class SupplierAnalyticsService
{
    private const double HalfLifeDays = 180.0;
    private readonly SqliteConnection _connection;

    public SupplierAnalyticsService(SqliteConnection connection) => _connection = connection;

    public static string NormalizeManufacturer(string? value)
    {
        var text = (value ?? string.Empty).Trim().ToUpperInvariant();
        text = text.Replace('Ё', 'Е');
        text = Regex.Replace(text, @"[\s\._\-/\\]+", string.Empty);
        text = Regex.Replace(text, @"(LIMITED|LTD|INCORPORATED|INC|CORP|CORPORATION|ООО|ЗАО|ОАО|АО)$", string.Empty);
        return text;
    }

    public static double? ParseLeadTimeTotalDays(string? value)
    {
        var text = (value ?? string.Empty).Trim().ToLowerInvariant();
        if (text.Length == 0) return null;
        if (text.Contains("stock") || text.Contains("склад")) return 0;
        var numbers = Regex.Matches(text, @"\d+(?:[\.,]\d+)?")
            .Select(match => double.TryParse(match.Value.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number : (double?)null)
            .Where(number => number is not null).Select(number => number!.Value).ToList();
        if (numbers.Count == 0) return null;
        var max = numbers.Max();
        return text.Contains("day") || text.Contains("дн") || Regex.IsMatch(text, @"\bd\b") ? max : max * 7;
    }

    public List<AnalyticsManufacturerOption> GetManufacturerOptions()
    {
        var result = new List<AnalyticsManufacturerOption>();
        using var command = _connection.CreateCommand();
        command.CommandText = @"SELECT m.id, m.canonical_name, IFNULL(a.alias, '')
FROM manufacturers m LEFT JOIN manufacturer_aliases a ON a.manufacturer_id=m.id
ORDER BY m.canonical_name, a.alias;";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new AnalyticsManufacturerOption
            {
                Id = reader.GetInt64(0), Name = reader.GetString(1), Alias = reader.GetString(2),
            });
        }
        return result;
    }

    public List<ManufacturerAliasRow> GetAliasRows()
    {
        var result = new List<ManufacturerAliasRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = @"SELECT a.id, IFNULL(a.manufacturer_id, 0), IFNULL(m.canonical_name, ''), a.alias
FROM manufacturer_aliases a LEFT JOIN manufacturers m ON m.id=a.manufacturer_id
ORDER BY CASE WHEN a.manufacturer_id IS NULL THEN 0 ELSE 1 END, a.alias;";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new ManufacturerAliasRow
            {
                Id = reader.GetInt64(0), ManufacturerId = reader.GetInt64(1),
                Manufacturer = reader.GetString(2), Alias = reader.GetString(3),
            });
        }
        return result;
    }

    public long CreateManufacturer(string name)
    {
        var canonical = name.Trim();
        var normalized = NormalizeManufacturer(canonical);
        if (normalized.Length == 0) throw new ArgumentException("Введите название производителя.");
        return PurchaseWriteSession.Execute(connection =>
        {
            using var insert = connection.CreateCommand();
            insert.CommandText = @"INSERT OR IGNORE INTO manufacturers(canonical_name, normalized_name, created_at, updated_at)
VALUES(@name, @normalized, @now, @now);";
            PurchaseWriteRepository.Add(insert, "@name", canonical);
            PurchaseWriteRepository.Add(insert, "@normalized", normalized);
            PurchaseWriteRepository.Add(insert, "@now", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            insert.ExecuteNonQuery();
            var id = GetManufacturerId(connection, normalized);
            using var alias = connection.CreateCommand();
            alias.CommandText = "UPDATE manufacturer_aliases SET manufacturer_id=@id, updated_at=@now WHERE normalized_alias=@normalized AND manufacturer_id IS NULL;";
            PurchaseWriteRepository.Add(alias, "@id", id);
            PurchaseWriteRepository.Add(alias, "@normalized", normalized);
            PurchaseWriteRepository.Add(alias, "@now", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            alias.ExecuteNonQuery();
            return id;
        });
    }

    public void AssignAlias(long aliasId, long manufacturerId)
    {
        if (aliasId <= 0 || manufacturerId <= 0) throw new ArgumentException("Выберите алиас и производителя.");
        PurchaseWriteSession.Execute(connection =>
        {
            using var command = connection.CreateCommand();
            command.CommandText = "UPDATE manufacturer_aliases SET manufacturer_id=@manufacturer_id, updated_at=@now WHERE id=@id;";
            PurchaseWriteRepository.Add(command, "@manufacturer_id", manufacturerId);
            PurchaseWriteRepository.Add(command, "@now", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            PurchaseWriteRepository.Add(command, "@id", aliasId);
            command.ExecuteNonQuery();
            return 0;
        });
    }

    public void MergeManufacturers(long sourceId, long targetId)
    {
        if (sourceId <= 0 || targetId <= 0 || sourceId == targetId) throw new ArgumentException("Выберите двух разных производителей.");
        PurchaseWriteSession.Execute(connection =>
        {
            using var aliases = connection.CreateCommand();
            aliases.CommandText = "UPDATE manufacturer_aliases SET manufacturer_id=@target, updated_at=@now WHERE manufacturer_id=@source;";
            PurchaseWriteRepository.Add(aliases, "@target", targetId);
            PurchaseWriteRepository.Add(aliases, "@source", sourceId);
            PurchaseWriteRepository.Add(aliases, "@now", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            aliases.ExecuteNonQuery();
            using var delete = connection.CreateCommand();
            delete.CommandText = "DELETE FROM manufacturers WHERE id=@source;";
            PurchaseWriteRepository.Add(delete, "@source", sourceId);
            delete.ExecuteNonQuery();
            return 0;
        });
    }

    public List<SupplierRecommendationRow> GetRecommendations(string manufacturer, string dateFrom = "", string dateTo = "")
    {
        var aliases = LoadAliases();
        var targetId = ResolveManufacturerId(manufacturer, aliases);
        var targetNorm = NormalizeManufacturer(manufacturer);
        var all = LoadObservations(dateFrom, dateTo);
        var scoped = targetId > 0
            ? all.Where(item => ResolveManufacturerId(item.ManufacturerRaw, aliases) == targetId).ToList()
            : all.Where(item => NormalizeManufacturer(item.ManufacturerRaw) == targetNorm).ToList();

        var global = BuildRanking(all, null, aliases);
        var local = BuildRanking(scoped, targetId > 0 ? targetId : null, aliases);
        var result = local.Count == 0 ? global.Select(row => CopyRow(row, row.Score, true)).ToList() : BlendWithGlobal(local, global);
        return result.OrderByDescending(row => row.Score).ThenByDescending(row => row.Comparisons).Take(3).ToList();
    }

    public List<SupplierRecommendationRow> GetStatistics(string manufacturer, string dateFrom = "", string dateTo = "")
    {
        var aliases = LoadAliases();
        var all = LoadObservations(dateFrom, dateTo);
        if (!string.IsNullOrWhiteSpace(manufacturer))
        {
            var targetId = ResolveManufacturerId(manufacturer, aliases);
            all = targetId > 0
                ? all.Where(item => ResolveManufacturerId(item.ManufacturerRaw, aliases) == targetId).ToList()
                : all.Where(item => NormalizeManufacturer(item.ManufacturerRaw) == NormalizeManufacturer(manufacturer)).ToList();
        }
        return BuildRanking(all, null, aliases).ToList();
    }

    public List<AnalyticsMonthlyRow> GetMonthlyTrend(string manufacturer, string supplier, string dateFrom = "", string dateTo = "")
    {
        var aliases = LoadAliases();
        var observations = LoadObservations(dateFrom, dateTo).Where(item => item.SourceKind == "full_analysis" && item.Price is > 0 && item.Comparable).ToList();
        if (!string.IsNullOrWhiteSpace(manufacturer))
        {
            var targetId = ResolveManufacturerId(manufacturer, aliases);
            observations = targetId > 0
                ? observations.Where(item => ResolveManufacturerId(item.ManufacturerRaw, aliases) == targetId).ToList()
                : observations.Where(item => NormalizeManufacturer(item.ManufacturerRaw) == NormalizeManufacturer(manufacturer)).ToList();
        }
        var sourceObservations = observations;
        if (!string.IsNullOrWhiteSpace(supplier))
            observations = observations.Where(item => item.Supplier.Contains(supplier.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();

        var result = new List<AnalyticsMonthlyRow>();
        foreach (var month in observations.GroupBy(item => item.Date.ToString("yyyy-MM")))
        {
            foreach (var supplierGroup in month.GroupBy(item => item.Supplier, StringComparer.OrdinalIgnoreCase))
            {
                var comparisonGroups = sourceObservations.Where(item => item.Date.ToString("yyyy-MM") == month.Key)
                    .GroupBy(item => new { item.BatchId, item.PositionKey })
                    .Select(group => group.Where(item => item.Price is > 0).ToList())
                    .Where(group => group.Select(item => item.Supplier).Distinct(StringComparer.OrdinalIgnoreCase).Count() >= 2);
                var ratios = new List<(double Ratio, bool Best)>();
                foreach (var group in comparisonGroups)
                {
                    var min = group.Min(item => item.Price!.Value);
                    var own = group.Where(item => string.Equals(item.Supplier, supplierGroup.Key, StringComparison.OrdinalIgnoreCase)).OrderBy(item => item.Price).FirstOrDefault();
                    if (own is not null) ratios.Add((min / own.Price!.Value, own.Price.Value <= min * 1.01));
                }
                if (ratios.Count == 0) continue;
                result.Add(new AnalyticsMonthlyRow
                {
                    Month = month.Key, Supplier = supplierGroup.Key, Comparisons = ratios.Count,
                    PriceCompetitiveness = ratios.Average(item => item.Ratio), BestPriceRate = ratios.Count(item => item.Best) / (double)ratios.Count,
                });
            }
        }
        return result.OrderByDescending(item => item.Month).ThenByDescending(item => item.PriceCompetitiveness).ToList();
    }

    public List<AnalyticsBatchRow> GetBatches(int limit = 200)
    {
        var result = new List<AnalyticsBatchRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = @"SELECT b.id, IFNULL(b.created_at,''), IFNULL(b.source_kind,''), IFNULL(b.data_quality,''),
IFNULL(b.rfq_path,''), IFNULL(b.result_path,''),
(SELECT COUNT(*) FROM quote_batch_positions p WHERE p.batch_id=b.id),
(SELECT COUNT(*) FROM quote_history q WHERE q.batch_id=b.id)
FROM quote_batches b ORDER BY b.id DESC LIMIT @limit;";
        PurchaseWriteRepository.Add(command, "@limit", Math.Max(1, limit));
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new AnalyticsBatchRow
            {
                Id = reader.GetInt64(0), CreatedAt = reader.GetString(1), SourceKind = reader.GetString(2),
                DataQuality = reader.GetString(3), RfqPath = reader.GetString(4), ResultPath = reader.GetString(5),
                PositionCount = Convert.ToInt32(reader.GetValue(6)), QuoteCount = Convert.ToInt32(reader.GetValue(7)),
            });
        }
        return result;
    }

    public List<AnalyticsBatchQuoteRow> GetBatchQuotes(long batchId)
    {
        var result = new List<AnalyticsBatchQuoteRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = @"SELECT q.batch_id, IFNULL(b.created_at,''), IFNULL(q.rfq_value,''), IFNULL(q.pn,''),
COALESCE(NULLIF(q.manufacturer_raw,''), q.mfg, ''), IFNULL(q.supplier,''), q.unit_price, IFNULL(q.lead_time,''),
IFNULL(q.is_winner,0), IFNULL(b.data_quality,'')
FROM quote_history q JOIN quote_batches b ON b.id=q.batch_id WHERE q.batch_id=@batch_id
ORDER BY q.sheet_name, q.row_number, q.supplier;";
        PurchaseWriteRepository.Add(command, "@batch_id", batchId);
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new AnalyticsBatchQuoteRow
            {
                BatchId = reader.GetInt64(0), CreatedAt = reader.GetString(1), RfqValue = reader.GetString(2),
                PN = reader.GetString(3), Manufacturer = reader.GetString(4), Supplier = reader.GetString(5),
                UnitPrice = reader.IsDBNull(6) ? null : Convert.ToDouble(reader.GetValue(6)), LeadTime = reader.GetString(7),
                IsWinner = Convert.ToInt64(reader.GetValue(8)) != 0, DataQuality = reader.GetString(9),
            });
        }
        return result;
    }

    private List<Observation> LoadObservations(string dateFrom, string dateTo)
    {
        var result = new List<Observation>();
        using var command = _connection.CreateCommand();
        command.CommandText = @"SELECT q.batch_id, b.created_at, IFNULL(b.source_kind,'legacy'), IFNULL(b.data_quality,'legacy'),
IFNULL(q.position_key, q.sheet_name || ':' || q.row_number || ':' || q.rfq_value), IFNULL(q.supplier,''), q.unit_price,
IFNULL(q.is_price_comparable,1), q.lead_total_days, COALESCE(NULLIF(q.manufacturer_raw,''), q.mfg, ''), IFNULL(q.is_winner,0)
FROM quote_history q JOIN quote_batches b ON b.id=q.batch_id
WHERE (@from='' OR substr(q.quote_date,1,10)>=@from) AND (@to='' OR substr(q.quote_date,1,10)<=@to);";
        PurchaseWriteRepository.Add(command, "@from", dateFrom ?? string.Empty);
        PurchaseWriteRepository.Add(command, "@to", dateTo ?? string.Empty);
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new Observation
            {
                BatchId = reader.GetInt64(0), Date = ParseDate(reader.GetString(1)), SourceKind = reader.GetString(2),
                DataQuality = reader.GetString(3), PositionKey = reader.GetString(4), Supplier = reader.GetString(5),
                Price = reader.IsDBNull(6) ? null : Convert.ToDouble(reader.GetValue(6)), Comparable = Convert.ToInt64(reader.GetValue(7)) != 0,
                LeadDays = reader.IsDBNull(8) ? null : Convert.ToDouble(reader.GetValue(8)), ManufacturerRaw = reader.GetString(9),
                IsWinner = Convert.ToInt64(reader.GetValue(10)) != 0,
            });
        }
        return result;
    }

    private List<SupplierRecommendationRow> BuildRanking(List<Observation> observations, long? targetManufacturerId, Dictionary<string, long> aliases)
    {
        var priced = observations.Where(item => item.Price is > 0 && item.Comparable).ToList();
        var valid = priced.GroupBy(item => new { item.BatchId, item.PositionKey })
            .Where(group => group.Any(item => item.SourceKind == "full_analysis") || group.Count() >= 2)
            .SelectMany(group => group).ToList();
        var stats = new Dictionary<string, MutableStat>(StringComparer.OrdinalIgnoreCase);
        foreach (var group in valid.GroupBy(item => new { item.BatchId, item.PositionKey }))
        {
            var bySupplier = group.GroupBy(item => item.Supplier, StringComparer.OrdinalIgnoreCase)
                .Select(items => items.OrderBy(item => item.Price).First()).Where(item => item.Price is > 0).ToList();
            if (bySupplier.Count < 2) continue;
            var minimum = bySupplier.Min(item => item.Price!.Value);
            var minLead = bySupplier.Where(item => item.LeadDays is not null).Select(item => item.LeadDays!.Value).DefaultIfEmpty().Min();
            var age = Math.Max(0, (DateTime.Today - bySupplier[0].Date.Date).TotalDays);
            var weight = Math.Pow(0.5, age / HalfLifeDays);
            foreach (var item in bySupplier)
            {
                if (!stats.TryGetValue(item.Supplier, out var stat)) stats[item.Supplier] = stat = new MutableStat();
                var ratio = minimum / item.Price!.Value;
                stat.Weight += weight;
                stat.PriceSum += weight * ratio;
                stat.BestWeight += weight * (item.Price.Value <= minimum * 1.01 ? 1 : 0);
                stat.Comparisons++;
                if (item.LeadDays is not null && bySupplier.Any(candidate => candidate.LeadDays is not null))
                {
                    var leadScore = minLead == 0 ? (item.LeadDays.Value == 0 ? 1 : 0) : Math.Min(1, minLead / item.LeadDays.Value);
                    stat.LeadSum += weight * leadScore;
                    stat.LeadDaysSum += weight * item.LeadDays.Value;
                    stat.LeadWeight += weight;
                }
            }
        }

        var participants = LoadParticipants();
        var positions = LoadPositions();
        var responses = observations.Where(item => item.SourceKind == "full_analysis")
            .GroupBy(item => $"{item.BatchId}\u001f{item.PositionKey}\u001f{item.Supplier.Trim().ToUpperInvariant()}")
            .ToDictionary(group => group.Key, group => group.Any(), StringComparer.Ordinal);
        foreach (var position in positions.Where(item => item.SourceKind == "full_analysis"))
        {
            if (targetManufacturerId is not null && ResolveManufacturerId(position.ManufacturerRaw, aliases) != targetManufacturerId) continue;
            if (!participants.TryGetValue(position.BatchId, out var supplierSet)) continue;
            foreach (var supplier in supplierSet)
            {
                if (!stats.TryGetValue(supplier, out var stat)) stats[supplier] = stat = new MutableStat();
                stat.ResponseDenominator++;
                if (responses.TryGetValue($"{position.BatchId}\u001f{position.PositionKey}\u001f{supplier.Trim().ToUpperInvariant()}", out var responded) && responded)
                    stat.ResponseNumerator++;
            }
        }

        return stats.Select(pair =>
        {
            var stat = pair.Value;
            var price = stat.Weight == 0 ? 0.5 : stat.PriceSum / stat.Weight;
            var best = stat.Weight == 0 ? 0 : stat.BestWeight / stat.Weight;
            var response = stat.ResponseDenominator == 0 ? 0 : stat.ResponseNumerator / (double)stat.ResponseDenominator;
            var lead = stat.LeadWeight == 0 ? 0.5 : stat.LeadSum / stat.LeadWeight;
            var raw = price * 0.60 + best * 0.20 + response * 0.10 + lead * 0.10;
            var confidence = Math.Min(1, stat.Comparisons / 10.0);
            var confidenceText = stat.Comparisons < 5 ? "Низкое" : stat.Comparisons < 15 ? "Среднее" : "Высокое";
            return new SupplierRecommendationRow
            {
                Supplier = pair.Key, Score = raw * confidence * 100, Confidence = confidenceText,
                Comparisons = stat.Comparisons, PriceCompetitiveness = price, BestPriceRate = best,
                ResponseRate = response, AverageLeadDays = stat.LeadWeight == 0 ? null : stat.LeadDaysSum / stat.LeadWeight,
                Explanation = $"Лучшая цена: {best:P0}; средняя цена относительно минимума: {price:P0}; ответов: {response:P0}.",
            };
        }).OrderByDescending(row => row.Score).ToList();
    }

    private static List<SupplierRecommendationRow> BlendWithGlobal(List<SupplierRecommendationRow> local, List<SupplierRecommendationRow> global)
    {
        var globalBySupplier = global.ToDictionary(row => row.Supplier, StringComparer.OrdinalIgnoreCase);
        return local.Select(row =>
        {
            var confidence = row.Comparisons < 5 ? row.Comparisons / 10.0 : 1.0;
            if (!globalBySupplier.TryGetValue(row.Supplier, out var prior)) return row;
            var score = row.Score * confidence + prior.Score * (1 - confidence);
            return CopyRow(row, score, false, row.Explanation + $" Сравнений по производителю: {row.Comparisons}.");
        }).ToList();
    }

    private static SupplierRecommendationRow CopyRow(SupplierRecommendationRow row, double score, bool fallback, string? explanation = null)
        => new()
        {
            Supplier = row.Supplier, Score = score, Confidence = row.Confidence, Comparisons = row.Comparisons,
            PriceCompetitiveness = row.PriceCompetitiveness, BestPriceRate = row.BestPriceRate,
            ResponseRate = row.ResponseRate, AverageLeadDays = row.AverageLeadDays,
            Explanation = explanation ?? row.Explanation, IsFallback = fallback,
        };

    private Dictionary<string, long> LoadAliases()
    {
        var result = new Dictionary<string, long>(StringComparer.Ordinal);
        using var command = _connection.CreateCommand();
        command.CommandText = "SELECT normalized_alias, IFNULL(manufacturer_id,0) FROM manufacturer_aliases;";
        using var reader = command.ExecuteReader();
        while (reader.Read()) result[reader.GetString(0)] = reader.GetInt64(1);
        using var canonical = _connection.CreateCommand();
        canonical.CommandText = "SELECT normalized_name, id FROM manufacturers;";
        using var canonicalReader = canonical.ExecuteReader();
        while (canonicalReader.Read()) result[canonicalReader.GetString(0)] = canonicalReader.GetInt64(1);
        return result;
    }

    private static long ResolveManufacturerId(string value, Dictionary<string, long> aliases)
        => aliases.TryGetValue(NormalizeManufacturer(value), out var id) ? id : 0;

    private static long GetManufacturerId(SqliteConnection connection, string normalized)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT id FROM manufacturers WHERE normalized_name=@normalized;";
        PurchaseWriteRepository.Add(command, "@normalized", normalized);
        return Convert.ToInt64(command.ExecuteScalar());
    }

    private Dictionary<long, HashSet<string>> LoadParticipants()
    {
        var result = new Dictionary<long, HashSet<string>>();
        using var command = _connection.CreateCommand();
        command.CommandText = "SELECT batch_id, supplier FROM quote_batch_suppliers;";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var batchId = reader.GetInt64(0);
            if (!result.TryGetValue(batchId, out var suppliers)) result[batchId] = suppliers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            suppliers.Add(reader.GetString(1));
        }
        return result;
    }

    private List<PositionInfo> LoadPositions()
    {
        var result = new List<PositionInfo>();
        using var command = _connection.CreateCommand();
        command.CommandText = "SELECT p.batch_id, p.position_key, IFNULL(p.manufacturer_raw,''), IFNULL(b.source_kind,'') FROM quote_batch_positions p JOIN quote_batches b ON b.id=p.batch_id;";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new PositionInfo
            {
                BatchId = reader.GetInt64(0), PositionKey = reader.GetString(1), ManufacturerRaw = reader.GetString(2), SourceKind = reader.GetString(3),
            });
        }
        return result;
    }

    private static DateTime ParseDate(string value)
        => DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var result) ? result : DateTime.Today;

    private sealed class MutableStat
    {
        public double Weight;
        public double PriceSum;
        public double BestWeight;
        public int Comparisons;
        public int ResponseNumerator;
        public int ResponseDenominator;
        public double LeadSum;
        public double LeadWeight;
        public double LeadDaysSum;
    }

    private sealed class Observation
    {
        public long BatchId;
        public DateTime Date;
        public string SourceKind = string.Empty;
        public string DataQuality = string.Empty;
        public string PositionKey = string.Empty;
        public string Supplier = string.Empty;
        public double? Price;
        public bool Comparable;
        public double? LeadDays;
        public string ManufacturerRaw = string.Empty;
        public bool IsWinner;
    }

    private sealed class PositionInfo
    {
        public long BatchId;
        public string PositionKey = string.Empty;
        public string ManufacturerRaw = string.Empty;
        public string SourceKind = string.Empty;
    }
}
