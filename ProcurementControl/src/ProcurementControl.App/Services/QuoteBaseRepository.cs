using Microsoft.Data.Sqlite;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Read-only запросы раздела «База квот» — перенос Get-QuoteHistory и
/// Get-GlobalistQuotes из app/modules/QuoteHistory.ps1 (строки 164-197).
/// Только SELECT: запись и удаление квот в порт не входят.
/// </summary>
public sealed class QuoteBaseRepository
{
    private readonly SqliteConnection _connection;

    public QuoteBaseRepository(SqliteConnection connection)
    {
        _connection = connection;
    }

    /// <summary>
    /// Аналог Get-QuoteHistory: квоты с данными пакета, поиск по
    /// rfq_value/pn/supplier/mfg и диапазон дат по первым 10 символам даты.
    /// Даты приходят в формате «yyyy-MM-dd» (аналог Convert-PurchaseDateForSort),
    /// пустая строка — условие не применяется.
    /// </summary>
    public List<QuoteHistoryRow> GetQuoteHistory(string search, string dateFrom, string dateTo, int limit)
    {
        var sql = @"
SELECT
    q.id,
    q.quote_date,
    IFNULL(q.rfq_value, '') AS rfq_value,
    IFNULL(q.pn, '') AS pn,
    IFNULL(q.requested_qty, '') AS requested_qty,
    IFNULL(q.supplier, '') AS supplier,
    q.unit_price,
    IFNULL(q.lead_time, '') AS lead_time,
    IFNULL(q.lead_time_total, '') AS lead_time_total,
    IFNULL(q.mfg, '') AS mfg,
    IFNULL(q.is_winner, 0) AS is_winner,
    IFNULL(q.winner_reason, '') AS winner_reason,
    IFNULL(q.warning, '') AS warning,
    IFNULL(b.rfq_path, '') AS rfq_path,
    IFNULL(b.priority, '') AS priority
FROM quote_history q
JOIN quote_batches b ON b.id = q.batch_id
WHERE (@search = '' OR IFNULL(q.rfq_value, '') LIKE @needle OR IFNULL(q.pn, '') LIKE @needle OR IFNULL(q.supplier, '') LIKE @needle OR IFNULL(q.mfg, '') LIKE @needle)";

        if (!string.IsNullOrWhiteSpace(dateFrom))
        {
            sql += " AND substr(q.quote_date, 1, 10) >= @date_from";
        }

        if (!string.IsNullOrWhiteSpace(dateTo))
        {
            sql += " AND substr(q.quote_date, 1, 10) <= @date_to";
        }

        sql += @"
ORDER BY q.quote_date DESC, q.id DESC
LIMIT @limit";

        var result = new List<QuoteHistoryRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("@search", search ?? string.Empty);
        command.Parameters.AddWithValue("@needle", "%" + (search ?? string.Empty) + "%");
        command.Parameters.AddWithValue("@date_from", dateFrom ?? string.Empty);
        command.Parameters.AddWithValue("@date_to", dateTo ?? string.Empty);
        command.Parameters.AddWithValue("@limit", Math.Max(1, limit));

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new QuoteHistoryRow
            {
                Id = reader.GetInt64(0),
                QuoteDate = Text(reader, 1),
                RfqValue = Text(reader, 2),
                PN = Text(reader, 3),
                RequestedQuantity = Text(reader, 4),
                Supplier = Text(reader, 5),
                UnitPrice = NullableDouble(reader, 6),
                LeadTime = Text(reader, 7),
                LeadTimeTotal = Text(reader, 8),
                Mfg = Text(reader, 9),
                IsWinner = !reader.IsDBNull(10) && Convert.ToInt64(reader.GetValue(10)) != 0,
                WinnerReason = Text(reader, 11),
                Warning = Text(reader, 12),
                RfqPath = Text(reader, 13),
                Priority = Text(reader, 14),
            });
        }
        return result;
    }

    /// <summary>Аналог Get-GlobalistQuotes: поиск по pn/brand/factory/pi_number, новые сверху.</summary>
    public List<GlobalistRow> GetGlobalistQuotes(string search, int limit)
    {
        const string sql = @"
SELECT
    id,
    IFNULL(imported_at, '') AS imported_at,
    IFNULL(factory, '') AS factory,
    IFNULL(pn, '') AS pn,
    IFNULL(comment, '') AS comment,
    IFNULL(pi_number, '') AS pi_number,
    IFNULL(replacement, '') AS replacement,
    IFNULL(chinese_remark, '') AS chinese_remark,
    IFNULL(package, '') AS package,
    IFNULL(brand, '') AS brand,
    IFNULL(datacode, '') AS datacode,
    IFNULL(moq, '') AS moq,
    IFNULL(qty, '') AS qty,
    IFNULL(stock, '') AS stock,
    IFNULL(need_spq, '') AS need_spq,
    IFNULL(spq, '') AS spq,
    unit_price,
    total_amount,
    IFNULL(lead_time, '') AS lead_time,
    IFNULL(weight, '') AS weight,
    IFNULL(target, '') AS target,
    IFNULL(supplier_quote_id, '') AS supplier_quote_id,
    IFNULL(sheet_name, '') AS sheet_name
FROM globalist_quotes
WHERE (@search = '' OR IFNULL(pn, '') LIKE @needle OR IFNULL(brand, '') LIKE @needle OR IFNULL(factory, '') LIKE @needle OR IFNULL(pi_number, '') LIKE @needle)
ORDER BY id DESC
LIMIT @limit";

        var result = new List<GlobalistRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("@search", search ?? string.Empty);
        command.Parameters.AddWithValue("@needle", "%" + (search ?? string.Empty) + "%");
        command.Parameters.AddWithValue("@limit", Math.Max(1, limit));

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new GlobalistRow
            {
                Id = reader.GetInt64(0),
                ImportedAt = Text(reader, 1),
                Factory = Text(reader, 2),
                PN = Text(reader, 3),
                Comment = Text(reader, 4),
                PiNumber = Text(reader, 5),
                Replacement = Text(reader, 6),
                ChineseRemark = Text(reader, 7),
                Package = Text(reader, 8),
                Brand = Text(reader, 9),
                Datacode = Text(reader, 10),
                Moq = Text(reader, 11),
                Qty = Text(reader, 12),
                Stock = Text(reader, 13),
                NeedSpq = Text(reader, 14),
                Spq = Text(reader, 15),
                UnitPrice = NullableDouble(reader, 16),
                TotalAmount = NullableDouble(reader, 17),
                LeadTime = Text(reader, 18),
                Weight = Text(reader, 19),
                Target = Text(reader, 20),
                SupplierQuoteId = Text(reader, 21),
                SheetName = Text(reader, 22),
            });
        }
        return result;
    }

    private static string Text(SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? string.Empty : Convert.ToString(reader.GetValue(ordinal)) ?? string.Empty;

    private static double? NullableDouble(SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? null : Convert.ToDouble(reader.GetValue(ordinal));
}
