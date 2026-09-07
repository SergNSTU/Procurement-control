using Microsoft.Data.Sqlite;
using ProcurementControl.Models;

namespace ProcurementControl.Services;

/// <summary>
/// Read-only запросы к снапшоту базы закупок. Переносит read-часть
/// app/modules/PurchaseStore.ps1 (Invoke-PurchaseScalar, Get-PurchaseSetting).
/// Никаких INSERT/UPDATE/DELETE — только SELECT.
/// </summary>
public sealed class PurchaseRepository
{
    private readonly SqliteConnection _connection;

    public PurchaseRepository(SqliteConnection connection)
    {
        _connection = connection;
    }

    /// <summary>Аналог Get-PurchaseSetting: SELECT value FROM settings WHERE key=@key.</summary>
    public string? GetSetting(string key)
    {
        return ScalarString(
            "SELECT value FROM settings WHERE key = @key",
            new Dictionary<string, object?> { ["@key"] = key });
    }

    /// <summary>Годы, встречающиеся в датах дашборда по обоим направлениям.</summary>
    public IReadOnlyList<int> GetDashboardYears()
    {
        const string sql = @"
SELECT DISTINCT year FROM (
    SELECT CAST(SUBSTR(created_at, 1, 4) AS INTEGER) AS year FROM deals WHERE LENGTH(IFNULL(created_at, '')) >= 4
    UNION
    SELECT CAST(SUBSTR(ordered_at, 1, 4) AS INTEGER) FROM deals WHERE LENGTH(IFNULL(ordered_at, '')) >= 4
    UNION
    SELECT CAST(SUBSTR(created_at, 1, 4) AS INTEGER) FROM component_deals WHERE LENGTH(IFNULL(created_at, '')) >= 4
    UNION
    SELECT CAST(SUBSTR(ordered_at, 1, 4) AS INTEGER) FROM component_deals WHERE LENGTH(IFNULL(ordered_at, '')) >= 4
) WHERE year BETWEEN 2000 AND 2100 ORDER BY year DESC;";

        var years = new List<int>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;
        using var reader = command.ExecuteReader();
        while (reader.Read()) years.Add(reader.GetInt32(0));
        return years;
    }

    /// <summary>
    /// Аналог Get-ComponentDeals из PurchaseStore.ps1: список задач (component_deals)
    /// с поиском по номеру сделки и описанию, новые сверху.
    /// </summary>
    public List<TaskRow> GetTasks(string search)
    {
        var needle = string.IsNullOrWhiteSpace(search) ? "%" : "%" + search.Trim() + "%";
        const string sql = @"
SELECT
    id,
    entry_date,
    IFNULL(deal_number, '') AS deal_number,
    IFNULL(status, 'В работе') AS status,
    IFNULL(stage, 'Запросил поставщиков') AS stage,
    IFNULL(description, '') AS description,
    IFNULL(next_action, '') AS next_action,
    IFNULL(reminder_date, '') AS reminder_date,
    IFNULL(deadline_date, '') AS deadline_date,
    IFNULL(priority, '3') AS priority,
    IFNULL(period, '') AS period,
    IFNULL(order_amount, '') AS order_amount,
    IFNULL(notes, '') AS notes,
    updated_at
FROM component_deals
WHERE IFNULL(deal_number, '') LIKE @needle
   OR IFNULL(description, '') LIKE @needle
ORDER BY id DESC";

        var result = new List<TaskRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("@needle", needle);

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new TaskRow
            {
                Id = reader.GetInt64(0),
                EntryDate = reader.GetString(1),
                DealNumber = reader.GetString(2),
                Status = reader.GetString(3),
                Stage = reader.GetString(4),
                Description = reader.GetString(5),
                NextAction = reader.GetString(6),
                ReminderDate = reader.GetString(7),
                DeadlineDate = reader.GetString(8),
                Priority = reader.GetString(9),
                Period = reader.GetString(10),
                OrderAmount = reader.GetString(11),
                Notes = reader.GetString(12),
                UpdatedAt = reader.IsDBNull(13) ? string.Empty : reader.GetString(13),
            });
        }
        return result;
    }

    /// <summary>
    /// Аналог Get-PurchaseDeals с режимом AllRecords: все сделки со статистикой
    /// поставщиков, без фильтров и поиска (фильтрация выполняется в памяти).
    /// </summary>
    public List<PurchaseDealRow> GetDeals()
    {
        const string sql = @"
SELECT
    d.id,
    d.deal_number,
    IFNULL(d.board_count, '') AS board_count,
    IFNULL(d.client, '') AS client,
    IFNULL(d.period, '') AS period,
    IFNULL(d.executor, '') AS executor,
    IFNULL(d.reminder_date, '') AS reminder_date,
    IFNULL(d.assembly_location, '') AS assembly_location,
    COALESCE(NULLIF(TRIM(d.priority), ''), '3') AS priority,
    IFNULL(d.tracking_status, 'Ожидание') AS tracking_status,
    IFNULL(d.status, 'RFQ') AS status,
    IFNULL(d.masks, 0) AS masks,
    IFNULL(d.archived, 0) AS archived,
    d.title,
    d.comment,
    d.updated_at,
    IFNULL(stats.supplier_count, 0) AS supplier_count,
    IFNULL(stats.invoice_count, 0) AS invoice_count,
    IFNULL(stats.payment_submitted_count, 0) AS payment_submitted_count,
    IFNULL(stats.paid_count, 0) AS paid_count,
    IFNULL(stats.done_count, 0) AS done_count,
    IFNULL(stats.invoice_confirmed_count, 0) AS invoice_confirmed_count,
    IFNULL(stats.receipt_count, 0) AS receipt_count,
    IFNULL(stats.max_receipt_date, '') AS completion_receipt_date
FROM deals d
LEFT JOIN (
    SELECT
        deal_id,
        COUNT(*) AS supplier_count,
        SUM(invoice_received) AS invoice_count,
        SUM(payment_submitted) AS payment_submitted_count,
        SUM(paid) AS paid_count,
        SUM(CASE WHEN invoice_received = 1 AND invoice_confirmed = 1 THEN 1 ELSE 0 END) AS invoice_confirmed_count,
        SUM(CASE WHEN invoice_received = 1 AND invoice_confirmed = 1 AND ((supplier LIKE 'Компэл%' OR supplier LIKE 'Компел%' OR supplier LIKE 'ЧиД%' OR supplier LIKE 'ЧИД%' OR supplier LIKE 'чид%' OR supplier LIKE 'Промэлектроника%' OR supplier LIKE 'промэлектроника%' OR supplier LIKE 'Promelec%' OR supplier LIKE 'promelec%') OR (erp_supplier_sent = 1 AND erp_roger_sent = 1)) THEN 1 ELSE 0 END) AS done_count,
        SUM(CASE WHEN IFNULL(actual_receipt_date, '') <> '' THEN 1 ELSE 0 END) AS receipt_count,
        MAX(CASE
            WHEN IFNULL(actual_receipt_date, '') LIKE '__.__.____' THEN substr(actual_receipt_date, 7, 4) || '-' || substr(actual_receipt_date, 4, 2) || '-' || substr(actual_receipt_date, 1, 2)
            WHEN IFNULL(actual_receipt_date, '') LIKE '____-__-__%' THEN substr(actual_receipt_date, 1, 10)
            ELSE ''
        END) AS max_receipt_date
    FROM deal_suppliers
    GROUP BY deal_id
) stats ON stats.deal_id = d.id
-- Порядок списка фиксирован по «Этапу»: правка сделки обновляет updated_at,
-- но не переставляет её среди остальных строк.
ORDER BY d.status COLLATE NOCASE ASC, d.id ASC";

        var result = new List<PurchaseDealRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new PurchaseDealRow
            {
                Id = reader.GetInt64(0),
                DealNumber = Text(reader, 1),
                BoardCount = Text(reader, 2),
                Client = Text(reader, 3),
                Period = Text(reader, 4),
                Executor = Text(reader, 5),
                ReminderDate = Text(reader, 6),
                AssemblyLocation = Text(reader, 7),
                Priority = Text(reader, 8),
                TrackingStatus = Text(reader, 9),
                Status = Text(reader, 10),
                Masks = reader.IsDBNull(11) ? 0 : Convert.ToInt32(reader.GetValue(11)),
                Archived = reader.IsDBNull(12) ? 0 : Convert.ToInt32(reader.GetValue(12)),
                Title = reader.IsDBNull(13) ? null : reader.GetString(13),
                Comment = reader.IsDBNull(14) ? null : reader.GetString(14),
                UpdatedAt = Text(reader, 15),
                SupplierCount = Convert.ToInt32(reader.GetValue(16)),
                InvoiceCount = Convert.ToInt32(reader.GetValue(17)),
                PaymentSubmittedCount = Convert.ToInt32(reader.GetValue(18)),
                PaidCount = Convert.ToInt32(reader.GetValue(19)),
                DoneCount = Convert.ToInt32(reader.GetValue(20)),
                InvoiceConfirmedCount = Convert.ToInt32(reader.GetValue(21)),
                ReceiptCount = Convert.ToInt32(reader.GetValue(22)),
                CompletionReceiptDate = Text(reader, 23),
            });
        }
        return result;
    }

    /// <summary>Аналог Get-PurchaseSuppliers: поставщики выбранной сделки.</summary>
    public List<SupplierRow> GetSuppliers(long dealId)
    {
        const string sql = @"
SELECT
    id,
    deal_id,
    supplier,
    invoice_received,
    invoice_confirmed,
    supplier_order_created,
    erp_supplier_sent,
    erp_roger_sent,
    pi_amount_usd,
    pi_amount_cny,
    pi_amount_rub,
    IFNULL(paid_amount, '') AS paid_amount,
    delivery_weeks,
    payment_submitted,
    paid,
    invoice_confirmed_date,
    components_receipt_date,
    IFNULL(actual_receipt_date, '') AS actual_receipt_date,
    comment
FROM deal_suppliers
WHERE deal_id = @deal_id
ORDER BY supplier";

        var result = new List<SupplierRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("@deal_id", dealId);

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new SupplierRow
            {
                Id = reader.GetInt64(0),
                DealId = reader.GetInt64(1),
                Supplier = Text(reader, 2),
                InvoiceReceived = Bool(reader, 3),
                InvoiceConfirmed = Bool(reader, 4),
                SupplierOrderCreated = Bool(reader, 5),
                ErpSupplierSent = Bool(reader, 6),
                ErpRogerSent = Bool(reader, 7),
                PiAmountUsd = PurchaseFormatting.Amount(NullableDouble(reader, 8)),
                PiAmountCny = PurchaseFormatting.Amount(NullableDouble(reader, 9)),
                PiAmountRub = PurchaseFormatting.Amount(NullableDouble(reader, 10)),
                PaidAmount = Text(reader, 11),
                DeliveryWeeks = reader.IsDBNull(12) ? string.Empty : Convert.ToString(reader.GetValue(12)) ?? string.Empty,
                PaymentSubmitted = Bool(reader, 13),
                Paid = Bool(reader, 14),
                InvoiceConfirmedDate = PurchaseFormatting.Date(Text(reader, 15)),
                ComponentsReceiptDate = PurchaseFormatting.Date(Text(reader, 16)),
                ActualReceiptDate = PurchaseFormatting.Date(Text(reader, 17)),
                Comment = reader.IsDBNull(18) ? null : reader.GetString(18),
            });
        }
        return result;
    }

    /// <summary>
    /// Запрос поставщиков активных сделок для авто-пунктов «Напоминаний».
    /// Точный перенос SQL из Get-PurchaseActionItems (PurchaseStore.ps1, строки 1231-1252).
    /// </summary>
    public List<SupplierActionRow> GetSupplierActionRows()
    {
        const string sql = @"
SELECT
    d.id AS deal_id,
    d.deal_number,
    ds.id AS supplier_id,
    ds.supplier,
    ds.invoice_received,
    ds.invoice_confirmed,
    ds.supplier_order_created,
    ds.erp_supplier_sent,
    ds.erp_roger_sent,
    ds.payment_submitted,
    ds.paid,
    IFNULL(ds.pi_amount_usd, '') AS pi_amount_usd,
    ds.invoice_confirmed_date,
    ds.components_receipt_date,
    IFNULL(ds.actual_receipt_date, '') AS actual_receipt_date,
    ds.delivery_weeks
FROM deal_suppliers ds
JOIN deals d ON d.id = ds.deal_id
WHERE IFNULL(d.archived, 0) = 0
ORDER BY d.updated_at DESC, d.deal_number, ds.supplier";

        var result = new List<SupplierActionRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new SupplierActionRow
            {
                DealId = reader.GetInt64(0),
                DealNumber = Text(reader, 1),
                SupplierId = reader.GetInt64(2),
                Supplier = Text(reader, 3),
                InvoiceReceived = Bool(reader, 4),
                InvoiceConfirmed = Bool(reader, 5),
                SupplierOrderCreated = Bool(reader, 6),
                ErpSupplierSent = Bool(reader, 7),
                ErpRogerSent = Bool(reader, 8),
                PaymentSubmitted = Bool(reader, 9),
                Paid = Bool(reader, 10),
                PiAmountUsd = Text(reader, 11),
                InvoiceConfirmedDate = Text(reader, 12),
                ComponentsReceiptDate = Text(reader, 13),
                ActualReceiptDate = Text(reader, 14),
                DeliveryWeeks = Text(reader, 15),
            });
        }
        return result;
    }

    /// <summary>USD-суммы поставщиков только сделок, заказанных в выбранный период.</summary>
    public IReadOnlyList<DashboardSupplierAmountRow> GetOrderedDealSupplierAmounts(DateTime start, DateTime end)
    {
        const string sql = @"
SELECT ds.deal_id, IFNULL(ds.pi_amount_usd, '')
FROM deals d JOIN deal_suppliers ds ON ds.deal_id = d.id
WHERE IFNULL(d.ordered_at, '') <> '' AND d.ordered_at >= @start AND d.ordered_at < @end;";
        var rows = new List<DashboardSupplierAmountRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("@start", start.ToString("yyyy-MM-dd HH:mm:ss"));
        command.Parameters.AddWithValue("@end", end.ToString("yyyy-MM-dd HH:mm:ss"));
        using var reader = command.ExecuteReader();
        while (reader.Read()) rows.Add(new DashboardSupplierAmountRow(reader.GetInt64(0), Text(reader, 1)));
        return rows;
    }

    /// <summary>Суммы компонентов, заказанных в выбранный период.</summary>
    public IReadOnlyList<string> GetOrderedComponentAmounts(DateTime start, DateTime end)
    {
        const string sql = @"
SELECT IFNULL(order_amount, '') FROM component_deals
WHERE IFNULL(ordered_at, '') <> '' AND ordered_at >= @start AND ordered_at < @end;";
        var amounts = new List<string>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;
        command.Parameters.AddWithValue("@start", start.ToString("yyyy-MM-dd HH:mm:ss"));
        command.Parameters.AddWithValue("@end", end.ToString("yyyy-MM-dd HH:mm:ss"));
        using var reader = command.ExecuteReader();
        while (reader.Read()) amounts.Add(Text(reader, 0));
        return amounts;
    }

    /// <summary>
    /// Напоминания сделок (поле reminder_date) — перенос запроса из строк 1293-1300.
    /// </summary>
    public List<DealReminderRow> GetDealReminderRows()
    {
        const string sql = @"
SELECT id, deal_number, reminder_date
FROM deals
WHERE IFNULL(archived, 0) = 0
  AND IFNULL(status, '') NOT IN ('Выполнено', 'Не актуально')
  AND IFNULL(reminder_date, '') <> ''
ORDER BY reminder_date, id DESC";

        var result = new List<DealReminderRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new DealReminderRow
            {
                Id = reader.GetInt64(0),
                DealNumber = Text(reader, 1),
                ReminderDate = Text(reader, 2),
            });
        }
        return result;
    }

    /// <summary>
    /// Ручные напоминания (таблица reminders, status <> 'Done') — перенос строк 1315-1323.
    /// </summary>
    public List<ManualReminderRow> GetManualReminderRows()
    {
        const string sql = @"
SELECT id, deal_id, supplier_id, title, due_date, status, source
FROM reminders
WHERE status <> 'Done'
ORDER BY
    CASE WHEN IFNULL(due_date, '') = '' THEN 1 ELSE 0 END,
    due_date,
    id DESC";

        var result = new List<ManualReminderRow>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new ManualReminderRow
            {
                Id = reader.GetInt64(0),
                DealId = reader.IsDBNull(1) ? 0 : reader.GetInt64(1),
                SupplierId = reader.IsDBNull(2) ? 0 : reader.GetInt64(2),
                Title = Text(reader, 3),
                DueDate = Text(reader, 4),
                Source = Text(reader, 6),
            });
        }
        return result;
    }

    /// <summary>Ключи автоматических напоминаний, удалённых пользователем.</summary>
    public HashSet<string> GetAutomaticReminderSuppressions()
    {
        using var command = _connection.CreateCommand();
        command.CommandText = "SELECT reminder_key FROM reminder_suppressions;";
        using var reader = command.ExecuteReader();
        var result = new HashSet<string>(StringComparer.Ordinal);
        while (reader.Read())
        {
            result.Add(reader.GetString(0));
        }

        return result;
    }

    /// <summary>Задачи в работе для счётчика «Активные» (аналог части Get-ActiveCockpitDealCount).</summary>
    public List<CockpitTaskKey> GetCockpitTaskKeys()
    {
        const string sql = @"
SELECT id, IFNULL(deal_number, '') AS deal_number, IFNULL(status, '') AS status
FROM component_deals
WHERE IFNULL(status, '') IN ('В работе', 'Подано в оплату')";

        var result = new List<CockpitTaskKey>();
        using var command = _connection.CreateCommand();
        command.CommandText = sql;

        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new CockpitTaskKey
            {
                Id = reader.GetInt64(0),
                DealNumber = Text(reader, 1),
                Status = Text(reader, 2),
            });
        }
        return result;
    }

    private static string Text(SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? string.Empty : Convert.ToString(reader.GetValue(ordinal)) ?? string.Empty;

    private static bool Bool(SqliteDataReader reader, int ordinal)
        => !reader.IsDBNull(ordinal) && Convert.ToInt64(reader.GetValue(ordinal)) != 0;

    private static double? NullableDouble(SqliteDataReader reader, int ordinal)
        => reader.IsDBNull(ordinal) ? null : Convert.ToDouble(reader.GetValue(ordinal));

    /// <summary>Аналог Invoke-PurchaseScalar для целых чисел.</summary>
    public int ScalarInt(string sql, IDictionary<string, object?>? parameters = null)
    {
        var value = ExecuteScalar(sql, parameters);
        return value is null || value is DBNull ? 0 : Convert.ToInt32(value);
    }

    /// <summary>Аналог Invoke-PurchaseScalar для строк.</summary>
    public string? ScalarString(string sql, IDictionary<string, object?>? parameters = null)
    {
        var value = ExecuteScalar(sql, parameters);
        return value is null || value is DBNull ? null : Convert.ToString(value);
    }

    private object? ExecuteScalar(string sql, IDictionary<string, object?>? parameters)
    {
        using var command = _connection.CreateCommand();
        command.CommandText = sql;
        if (parameters is not null)
        {
            foreach (var (name, value) in parameters)
            {
                var p = command.CreateParameter();
                p.ParameterName = name;
                p.Value = value ?? DBNull.Value;
                command.Parameters.Add(p);
            }
        }
        return command.ExecuteScalar();
    }
}
