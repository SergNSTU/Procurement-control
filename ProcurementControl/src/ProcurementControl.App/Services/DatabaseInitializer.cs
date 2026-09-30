using System.IO;
using Microsoft.Data.Sqlite;

namespace ProcurementControl.Services;

/// <summary>
/// Инициализация данных «с нуля».
///
/// WPF-приложение работает с общей SQLite-базой закупок
/// (data/purchase_control/purchase_control.sqlite), которую в оригинале создаёт
/// PowerShell-часть портативной сборки. Если программу перенесли на другой
/// компьютер без папки data, файла базы нет, и <see cref="PurchaseSnapshot"/>,
/// <see cref="PurchaseWriteSession"/> и <see cref="PurchaseDocumentsStore"/> падают
/// с FileNotFoundException ещё до показа окна.
///
/// Этот класс создаёт каталоги данных и, если база отсутствует, создаёт её с
/// полной схемой (все таблицы, которые читает/пишет приложение) и минимальными
/// стартовыми данными. Метод идемпотентен: существующая база не трогается.
/// Вызывается один раз при старте приложения до создания первого окна.
/// </summary>
public static class DatabaseInitializer
{
    private static readonly object SyncRoot = new();

    /// <summary>Создаёт каталоги данных и, при отсутствии, саму базу закупок со схемой.</summary>
    public static void EnsureCreated()
    {
        lock (SyncRoot)
        {
            EnsureDirectories();

            var databasePath = AppPaths.PurchaseDatabasePath;
            if (File.Exists(databasePath))
            {
                // База уже есть (оригинальная или созданная ранее). Выполняем
                // небольшие идемпотентные миграции для новых таблиц.
                EnsureCompatibility(databasePath);
                return;
            }

            var directory = Path.GetDirectoryName(databasePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var builder = new SqliteConnectionStringBuilder
            {
                DataSource = databasePath,
                Mode = SqliteOpenMode.ReadWriteCreate,
                Pooling = false,
            };

            using var connection = new SqliteConnection(builder.ConnectionString);
            connection.Open();

            // Оригинал работает в режиме WAL — снапшот копирует файлы .sqlite и -wal.
            using (var wal = connection.CreateCommand())
            {
                wal.CommandText = "PRAGMA journal_mode=WAL;";
                wal.ExecuteNonQuery();
            }

            using (var schema = connection.CreateCommand())
            {
                schema.CommandText = SchemaScript;
                schema.ExecuteNonQuery();
            }

            using (var seed = connection.CreateCommand())
            {
                seed.CommandText = SeedScript;
                seed.ExecuteNonQuery();
            }
            EnsureAnalyticsIndexes(connection);

            //Pooling=false, но снимем пул явно, чтобы файл не остался открытым.
            SqliteConnection.ClearPool(connection);
        }
    }

    private static void EnsureCompatibility(string databasePath)
    {
        var builder = new SqliteConnectionStringBuilder
        {
            DataSource = databasePath,
            Mode = SqliteOpenMode.ReadWrite,
            Pooling = false,
        };

        using var connection = new SqliteConnection(builder.ConnectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = @"
CREATE TABLE IF NOT EXISTS manufacturers(
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_name  TEXT NOT NULL,
    normalized_name TEXT NOT NULL UNIQUE,
    created_at      TEXT,
    updated_at      TEXT
);

CREATE TABLE IF NOT EXISTS manufacturer_aliases(
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    manufacturer_id INTEGER,
    alias           TEXT NOT NULL,
    normalized_alias TEXT NOT NULL UNIQUE,
    created_at      TEXT,
    updated_at      TEXT
);

CREATE TABLE IF NOT EXISTS quote_batch_positions(
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id          INTEGER NOT NULL,
    position_key      TEXT NOT NULL,
    sheet_name        TEXT,
    row_number        INTEGER,
    rfq_value         TEXT,
    pn                TEXT,
    manufacturer_raw  TEXT,
    manufacturer_norm TEXT,
    requested_qty     TEXT,
    component_type    TEXT,
    description       TEXT,
    UNIQUE(batch_id, position_key)
);

CREATE TABLE IF NOT EXISTS quote_batch_suppliers(
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id   INTEGER NOT NULL,
    supplier   TEXT NOT NULL,
    source_path TEXT,
    UNIQUE(batch_id, supplier)
);

CREATE TABLE IF NOT EXISTS reminder_suppressions(
    reminder_key TEXT PRIMARY KEY,
    title        TEXT,
    deal_id      INTEGER,
    supplier_id  INTEGER,
    created_at   TEXT
);";
        command.ExecuteNonQuery();
        EnsureAnalyticsIndexes(connection);

        EnsureColumn(connection, "quote_batches", "result_path", "TEXT");
        EnsureColumn(connection, "quote_batches", "source_kind", "TEXT DEFAULT 'legacy'");
        EnsureColumn(connection, "quote_batches", "fingerprint", "TEXT");
        EnsureColumn(connection, "quote_batches", "data_quality", "TEXT DEFAULT 'legacy'");
        EnsureColumn(connection, "quote_batches", "deal_id", "INTEGER");
        EnsureColumn(connection, "quote_batches", "summary_document_id", "INTEGER");
        EnsureColumn(connection, "quote_batches", "deal_number", "TEXT");
        EnsureColumn(connection, "quote_history", "position_key", "TEXT");
        EnsureColumn(connection, "quote_history", "quote_id", "TEXT");
        EnsureColumn(connection, "quote_history", "currency", "TEXT");
        EnsureColumn(connection, "quote_history", "requested_qty", "TEXT");
        EnsureColumn(connection, "quote_history", "manufacturer_raw", "TEXT");
        EnsureColumn(connection, "quote_history", "manufacturer_norm", "TEXT");
        EnsureColumn(connection, "quote_history", "lead_total_days", "REAL");
        EnsureColumn(connection, "quote_history", "is_price_comparable", "INTEGER DEFAULT 1");
        EnsureColumn(connection, "notification_state", "last_shown_at", "TEXT");
        BackfillRequestedQuantities(connection);
        ClassifyLegacyPiImports(connection);
        ClassifyLegacyComparisonResults(connection);
        BackfillBatchDealLinks(connection);
        ConsolidateDealBatches(connection);
        BackfillDocumentImportPositionKeys(connection);
        SqliteConnection.ClearPool(connection);
    }

    /// <summary>
    /// Старые PI лежали в базе как legacy, хотя содержат только итоговые
    /// победившие позиции, а не ответы на весь RFQ. Такие пакеты нельзя
    /// использовать в частоте ответа и доле побед.
    /// </summary>
    private static void ClassifyLegacyPiImports(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = @"
UPDATE quote_batches
SET source_kind='winner-only', data_quality='incomplete',
    comment=CASE WHEN IFNULL(comment,'')='' THEN 'winner-only legacy PI' ELSE comment END
WHERE IFNULL(source_kind,'legacy')='legacy'
  AND IFNULL(rfq_path,'') LIKE '%\Общий_PI\%';";
        command.ExecuteNonQuery();
    }

    /// <summary>Итоговый файл сравнения содержит уже отобранные и сведённые строки, а не независимые RRFQ поставщиков.</summary>
    private static void ClassifyLegacyComparisonResults(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = @"
UPDATE quote_batches
SET source_kind='comparison-result', data_quality='derived',
    comment=CASE WHEN IFNULL(comment,'')='' THEN 'legacy comparison result' ELSE comment END
WHERE IFNULL(source_kind,'legacy')='legacy'
  AND lower(IFNULL(rfq_path,'')) LIKE '%_result.xlsx';";
        command.ExecuteNonQuery();
    }

    private static void EnsureAnalyticsIndexes(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = @"
CREATE INDEX IF NOT EXISTS ix_quote_history_batch_position_supplier ON quote_history(batch_id, position_key, supplier);
CREATE INDEX IF NOT EXISTS ix_quote_history_pn ON quote_history(pn);
CREATE INDEX IF NOT EXISTS ix_quote_history_manufacturer ON quote_history(manufacturer_raw);
CREATE INDEX IF NOT EXISTS ix_quote_batches_deal_number ON quote_batches(deal_number);
CREATE INDEX IF NOT EXISTS ix_quote_batches_deal_id ON quote_batches(deal_id);
CREATE INDEX IF NOT EXISTS ix_quote_batch_positions_batch_position ON quote_batch_positions(batch_id, position_key);
CREATE INDEX IF NOT EXISTS ix_quote_batch_suppliers_batch_supplier ON quote_batch_suppliers(batch_id, supplier);";
        command.ExecuteNonQuery();
    }

    /// <summary>
    /// Заполняет количество у старых строк, только когда в сохранённой когорте
    /// есть точное совпадение позиции. Сопоставление лишь по PN небезопасно:
    /// один и тот же PN может фигурировать с разным количеством.
    /// </summary>
    private static void BackfillRequestedQuantities(SqliteConnection connection)
    {
        using var command = connection.CreateCommand();
        command.CommandText = @"
UPDATE quote_history
SET requested_qty = (
    SELECT p.requested_qty
    FROM quote_batch_positions p
    WHERE p.batch_id = quote_history.batch_id
      AND p.position_key = quote_history.position_key
)
WHERE IFNULL(TRIM(requested_qty), '') = ''
  AND IFNULL(TRIM(position_key), '') <> ''
  AND EXISTS (
      SELECT 1
      FROM quote_batch_positions p
      WHERE p.batch_id = quote_history.batch_id
        AND p.position_key = quote_history.position_key
        AND IFNULL(TRIM(p.requested_qty), '') <> ''
  );";
        command.ExecuteNonQuery();
    }

    /// <summary>Связывает старые пакеты только с точно распознанным номером существующей сделки.</summary>
    private static void BackfillBatchDealLinks(SqliteConnection connection)
    {
        var deals = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        using (var readDeals = connection.CreateCommand())
        {
            readDeals.CommandText = "SELECT id, deal_number FROM deals WHERE IFNULL(TRIM(deal_number), '') <> '';";
            using var reader = readDeals.ExecuteReader();
            while (reader.Read()) deals[reader.GetString(1).Trim()] = reader.GetInt64(0);
        }
        if (deals.Count == 0) return;
        var rows = new List<(long Id, string Path)>();
        using (var readBatches = connection.CreateCommand())
        {
            readBatches.CommandText = "SELECT id, IFNULL(rfq_path, '') FROM quote_batches WHERE IFNULL(deal_id, 0)=0 AND IFNULL(TRIM(deal_number), '')='';";
            using var reader = readBatches.ExecuteReader();
            while (reader.Read()) rows.Add((reader.GetInt64(0), reader.GetString(1)));
        }
        foreach (var row in rows)
        {
            var number = deals.Keys.FirstOrDefault(number => row.Path.Contains(number, StringComparison.OrdinalIgnoreCase));
            if (string.IsNullOrWhiteSpace(number)) continue;
            using var update = connection.CreateCommand();
            update.CommandText = "UPDATE quote_batches SET deal_id=@deal_id, deal_number=@deal_number WHERE id=@id;";
            update.Parameters.AddWithValue("@deal_id", deals[number]);
            update.Parameters.AddWithValue("@deal_number", number);
            update.Parameters.AddWithValue("@id", row.Id);
            update.ExecuteNonQuery();
        }
    }

    /// <summary>Для старых пакетов одной сделки оставляет одну сборку и переносит в неё квоты.</summary>
    private static void ConsolidateDealBatches(SqliteConnection connection)
    {
        var groups = new List<(long DealId, long KeepId)>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText = @"SELECT deal_id, MAX(id) FROM quote_batches WHERE IFNULL(deal_id,0)>0 GROUP BY deal_id HAVING COUNT(*)>1;";
            using var reader = command.ExecuteReader();
            while (reader.Read()) groups.Add((reader.GetInt64(0), reader.GetInt64(1)));
        }
        foreach (var group in groups)
        {
            using (var move = connection.CreateCommand())
            {
                move.CommandText = "UPDATE quote_history SET batch_id=@keep WHERE batch_id IN (SELECT id FROM quote_batches WHERE deal_id=@deal AND id<>@keep);";
                move.Parameters.AddWithValue("@keep", group.KeepId); move.Parameters.AddWithValue("@deal", group.DealId); move.ExecuteNonQuery();
            }
            // Оставляем последнюю запись одинакового предложения; более новые строки имеют больший id.
            using (var deduplicate = connection.CreateCommand())
            {
                deduplicate.CommandText = @"DELETE FROM quote_history WHERE batch_id=@keep AND id NOT IN (
SELECT MAX(id) FROM quote_history WHERE batch_id=@keep GROUP BY lower(trim(IFNULL(supplier,''))), IFNULL(rfq_value,''), IFNULL(pn,''),
IFNULL(unit_price,-999999999), IFNULL(requested_qty,''), IFNULL(currency,''), IFNULL(lead_time,''), IFNULL(lead_time_total,''), IFNULL(mfg,''));";
                deduplicate.Parameters.AddWithValue("@keep", group.KeepId); deduplicate.ExecuteNonQuery();
            }
            using (var delete = connection.CreateCommand())
            {
                delete.CommandText = "DELETE FROM quote_batches WHERE deal_id=@deal AND id<>@keep;";
                delete.Parameters.AddWithValue("@deal", group.DealId); delete.Parameters.AddWithValue("@keep", group.KeepId); delete.ExecuteNonQuery();
            }
        }
    }

    /// <summary>Починка старых импортов: ключ позиции не должен зависеть от строки Excel.</summary>
    private static void BackfillDocumentImportPositionKeys(SqliteConnection connection)
    {
        var rows = new List<(long Id, string Key, string Pn)>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText = @"SELECT q.id, IFNULL(q.rfq_value,''), IFNULL(q.pn,'') FROM quote_history q
JOIN quote_batches b ON b.id=q.batch_id WHERE b.source_kind='document_import'
AND (IFNULL(q.position_key,'')='' OR q.position_key NOT LIKE 'document:%');";
            using var reader = command.ExecuteReader();
            while (reader.Read()) rows.Add((reader.GetInt64(0), reader.GetString(1), reader.GetString(2)));
        }
        foreach (var row in rows)
        {
            var normalized = RrfqEngine.NormalizeKey(string.IsNullOrWhiteSpace(row.Key) ? row.Pn : row.Key);
            if (string.IsNullOrWhiteSpace(normalized)) continue;
            using var update = connection.CreateCommand();
            update.CommandText = "UPDATE quote_history SET position_key=@key WHERE id=@id;";
            update.Parameters.AddWithValue("@key", "document:" + normalized);
            update.Parameters.AddWithValue("@id", row.Id);
            update.ExecuteNonQuery();
        }
    }

    private static void EnsureColumn(SqliteConnection connection, string table, string column, string definition)
    {
        using var check = connection.CreateCommand();
        check.CommandText = $"PRAGMA table_info({table});";
        using var reader = check.ExecuteReader();
        while (reader.Read())
        {
            if (string.Equals(reader.GetString(1), column, StringComparison.OrdinalIgnoreCase)) return;
        }

        using var alter = connection.CreateCommand();
        alter.CommandText = $"ALTER TABLE {table} ADD COLUMN {column} {definition};";
        alter.ExecuteNonQuery();
    }

    /// <summary>
    /// Каталог данных портативной сборки: база, файлы документов, папки задач и заметки.
    /// Пути соответствуют <see cref="AppPaths"/> и дефолтам остальных сервисов.
    /// </summary>
    private static void EnsureDirectories()
    {
        Directory.CreateDirectory(AppPaths.DataRoot);
        Directory.CreateDirectory(AppPaths.PurchaseDataDirectory);
        Directory.CreateDirectory(Path.Combine(AppPaths.PurchaseDataDirectory, "files"));
        Directory.CreateDirectory(AppPaths.NotesDirectory);
        Directory.CreateDirectory(Path.Combine(AppPaths.DataRoot, "components", "files"));
    }

    /// <summary>
    /// Полная схема базы закупок. Набор таблиц и колонок восстановлен по всем
    /// SQL-запросам приложения. Ограничения UNIQUE соответствуют используемым в
    /// коде конструкциям ON CONFLICT(...) и INSERT OR IGNORE. Внешние ключи
    /// намеренно не объявлены: оригинальная схема их не принуждает, а удаление
    /// связанных строк код выполняет сам в нужном порядке.
    /// </summary>
    private const string SchemaScript = @"
CREATE TABLE IF NOT EXISTS deals(
    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_number          TEXT    NOT NULL UNIQUE,
    client               TEXT    DEFAULT '',
    status               TEXT    DEFAULT 'RFQ',
    title                TEXT,
    comment              TEXT,
    workflow_template_id INTEGER,
    masks                INTEGER DEFAULT 2,
    board_count          TEXT,
    period               TEXT,
    priority             TEXT    DEFAULT '3',
    tracking_status      TEXT    DEFAULT 'Ожидание',
    executor             TEXT,
    reminder_date        TEXT,
    assembly_location    TEXT,
    archived             INTEGER DEFAULT 0,
    created_at           TEXT,
    ordered_at           TEXT,
    updated_at           TEXT
);

CREATE TABLE IF NOT EXISTS deal_suppliers(
    id                     INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_id                INTEGER NOT NULL,
    supplier               TEXT    NOT NULL,
    invoice_received       INTEGER DEFAULT 0,
    invoice_confirmed      INTEGER DEFAULT 0,
    supplier_order_created INTEGER DEFAULT 0,
    erp_supplier_sent      INTEGER DEFAULT 0,
    erp_roger_sent         INTEGER DEFAULT 0,
    pi_amount_usd          REAL,
    pi_amount_cny          REAL,
    pi_amount_rub          REAL,
    paid_amount            TEXT,
    delivery_weeks         TEXT,
    payment_submitted      INTEGER DEFAULT 0,
    paid                   INTEGER DEFAULT 0,
    invoice_confirmed_date TEXT,
    components_receipt_date TEXT,
    actual_receipt_date    TEXT,
    comment                TEXT,
    created_at             TEXT,
    updated_at             TEXT,
    UNIQUE(deal_id, supplier)
);

CREATE TABLE IF NOT EXISTS component_deals(
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date   TEXT,
    deal_number  TEXT,
    status       TEXT    DEFAULT 'В работе',
    stage        TEXT    DEFAULT 'Запросил поставщиков',
    description  TEXT,
    next_action  TEXT,
    reminder_date TEXT,
    deadline_date TEXT,
    priority     TEXT    DEFAULT '3',
    period       TEXT,
    order_amount TEXT,
    notes        TEXT,
    source       TEXT,
    source_text  TEXT,
    folder_path  TEXT,
    created_at   TEXT,
    ordered_at   TEXT,
    updated_at   TEXT
);

CREATE TABLE IF NOT EXISTS documents(
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_id       INTEGER,
    supplier_id   INTEGER,
    document_type TEXT,
    original_name TEXT,
    stored_path   TEXT,
    file_hash     TEXT,
    uploaded_at   TEXT
);

CREATE TABLE IF NOT EXISTS settings(
    key   TEXT PRIMARY KEY,
    value TEXT
);

CREATE TABLE IF NOT EXISTS reminders(
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_id      INTEGER,
    supplier_id  INTEGER,
    component_id INTEGER,
    title        TEXT,
    due_date     TEXT,
    status       TEXT    DEFAULT 'Open',
    source       TEXT,
    created_at   TEXT,
    updated_at   TEXT
);

CREATE TABLE IF NOT EXISTS reminder_suppressions(
    reminder_key TEXT PRIMARY KEY,
    title        TEXT,
    deal_id      INTEGER,
    supplier_id  INTEGER,
    created_at   TEXT
);

CREATE TABLE IF NOT EXISTS activity_log(
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_type TEXT,
    entity_id   INTEGER,
    deal_id     INTEGER,
    supplier_id INTEGER,
    action      TEXT,
    details     TEXT,
    created_at  TEXT
);

CREATE TABLE IF NOT EXISTS trash_items(
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_type  TEXT,
    entity_id    INTEGER,
    title        TEXT,
    payload_json TEXT,
    files_json   TEXT,
    deleted_at   TEXT
);

CREATE TABLE IF NOT EXISTS notification_state(
    source        TEXT    NOT NULL,
    source_id     INTEGER NOT NULL,
    due_kind      TEXT    NOT NULL,
    due_date      TEXT    NOT NULL,
    handled       INTEGER DEFAULT 0,
    snooze_until  TEXT,
    last_shown_at TEXT,
    created_at    TEXT,
    updated_at    TEXT,
    UNIQUE(source, source_id, due_kind, due_date)
);

CREATE TABLE IF NOT EXISTS quote_batches(
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT,
    rfq_path   TEXT,
    priority   TEXT,
    comment    TEXT,
    result_path TEXT,
    source_kind TEXT DEFAULT 'legacy',
    fingerprint TEXT,
    data_quality TEXT DEFAULT 'legacy'
    ,deal_id INTEGER
    ,summary_document_id INTEGER
    ,deal_number TEXT
);

CREATE TABLE IF NOT EXISTS quote_history(
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id        INTEGER,
    quote_date      TEXT,
    rfq_value       TEXT,
    pn              TEXT,
    supplier        TEXT,
    unit_price      REAL,
    lead_time       TEXT,
    lead_time_total TEXT,
    mfg             TEXT,
    is_winner       INTEGER DEFAULT 0,
    winner_reason   TEXT,
    warning         TEXT,
    sheet_name      TEXT,
    row_number      INTEGER,
    match_status    TEXT,
    position_key    TEXT,
    quote_id        TEXT,
    currency        TEXT,
    requested_qty   TEXT,
    manufacturer_raw TEXT,
    manufacturer_norm TEXT,
    lead_total_days REAL,
    is_price_comparable INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS manufacturers(
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    canonical_name  TEXT NOT NULL,
    normalized_name TEXT NOT NULL UNIQUE,
    created_at      TEXT,
    updated_at      TEXT
);

CREATE TABLE IF NOT EXISTS manufacturer_aliases(
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    manufacturer_id INTEGER,
    alias           TEXT NOT NULL,
    normalized_alias TEXT NOT NULL UNIQUE,
    created_at      TEXT,
    updated_at      TEXT
);

CREATE TABLE IF NOT EXISTS quote_batch_positions(
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id          INTEGER NOT NULL,
    position_key      TEXT NOT NULL,
    sheet_name        TEXT,
    row_number        INTEGER,
    rfq_value         TEXT,
    pn                TEXT,
    manufacturer_raw  TEXT,
    manufacturer_norm TEXT,
    requested_qty     TEXT,
    component_type    TEXT,
    description       TEXT,
    UNIQUE(batch_id, position_key)
);

CREATE TABLE IF NOT EXISTS quote_batch_suppliers(
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id   INTEGER NOT NULL,
    supplier   TEXT NOT NULL,
    source_path TEXT,
    UNIQUE(batch_id, supplier)
);

CREATE TABLE IF NOT EXISTS globalist_quotes(
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    imported_at       TEXT,
    source_file       TEXT,
    sheet_name        TEXT,
    row_number        INTEGER,
    factory           TEXT,
    pn                TEXT,
    comment           TEXT,
    pi_number         TEXT,
    replacement       TEXT,
    chinese_remark    TEXT,
    package           TEXT,
    brand             TEXT,
    datacode          TEXT,
    moq               TEXT,
    qty               TEXT,
    stock             TEXT,
    need_spq          TEXT,
    spq               TEXT,
    unit_price        REAL,
    total_amount      REAL,
    lead_time         TEXT,
    weight            TEXT,
    target            TEXT,
    supplier_quote_id TEXT
);

CREATE TABLE IF NOT EXISTS bitrix_blocked_tasks(
    bitrix_task_id INTEGER PRIMARY KEY,
    title          TEXT,
    blocked_at     TEXT
);

CREATE TABLE IF NOT EXISTS bitrix_task_links(
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_id        INTEGER,
    bitrix_task_id INTEGER,
    created_at     TEXT
);

CREATE TABLE IF NOT EXISTS workflow_templates(
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT,
    is_default INTEGER DEFAULT 0
);
";

    /// <summary>
    /// Минимальные стартовые данные: шаблон рабочего процесса по умолчанию,
    /// на который ссылается создание сделки (workflow_template_id).
    /// </summary>
    private const string SeedScript = @"
INSERT INTO workflow_templates(name, is_default)
SELECT 'Основной', 1
WHERE NOT EXISTS (SELECT 1 FROM workflow_templates WHERE is_default = 1);
";
}
