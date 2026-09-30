using Microsoft.Data.Sqlite;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Read-only запросы раздела «История» — перенос Get-ActivityLog из
/// app/modules/PurchaseStore.ps1 (строки 880-902). Только SELECT: очистка
/// журнала выполняется через <see cref="PurchaseWriteRepository"/>.
/// </summary>
public sealed class HistoryRepository
{
    private readonly SqliteConnection _connection;

    public HistoryRepository(SqliteConnection connection)
    {
        _connection = connection;
    }

    /// <summary>
    /// Аналог Get-ActivityLog: свежие записи сверху, фильтр по сделке и
    /// поиск по действию, деталям, объекту и идентификаторам.
    /// </summary>
    public List<ActivityLogRow> GetActivityLog(long dealId, int limit, string search)
    {
        var where = new List<string>();
        using var command = _connection.CreateCommand();
        command.Parameters.AddWithValue("@limit", Math.Max(1, limit));

        if (dealId > 0)
        {
            where.Add("deal_id = @deal_id");
            command.Parameters.AddWithValue("@deal_id", dealId);
        }

        if (!string.IsNullOrWhiteSpace(search))
        {
            where.Add("(IFNULL(action, '') LIKE @needle OR IFNULL(details, '') LIKE @needle OR IFNULL(entity_type, '') LIKE @needle OR CAST(IFNULL(entity_id, '') AS TEXT) LIKE @needle OR CAST(IFNULL(deal_id, '') AS TEXT) LIKE @needle OR CAST(IFNULL(supplier_id, '') AS TEXT) LIKE @needle)");
            command.Parameters.AddWithValue("@needle", "%" + search.Trim() + "%");
        }

        var whereText = where.Count > 0 ? "WHERE " + string.Join(" AND ", where) : string.Empty;
        command.CommandText = $@"
SELECT id, entity_type, entity_id, deal_id, supplier_id, action, IFNULL(details, '') AS details, created_at
FROM activity_log
{whereText}
ORDER BY id DESC
LIMIT @limit";

        var result = new List<ActivityLogRow>();
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new ActivityLogRow
            {
                Id = reader.GetInt64(0),
                EntityType = Text(reader, 1),
                DealId = reader.IsDBNull(3) ? null : Convert.ToInt64(reader.GetValue(3)),
                Action = Text(reader, 5),
                Details = Text(reader, 6),
                CreatedAt = Text(reader, 7),
            });
        }
        return result;
    }

    private static string Text(SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? string.Empty : Convert.ToString(reader.GetValue(ordinal)) ?? string.Empty;
}
