# Purchase data store, document storage, and workflow operations.
# Dot-sourced by RRFQComparer.ps1; no application startup code belongs here.

function Load-SqliteProvider {
    if ($script:SqliteLoaded) {
        return
    }

    $sqliteDir = Join-Path $script:AppRoot 'app\lib\sqlite'
    $sqliteDll = Join-Path $sqliteDir 'System.Data.SQLite.dll'
    if (-not (Test-Path -LiteralPath $sqliteDll)) {
        throw "SQLite-компонент не найден: $sqliteDll"
    }

    $archDir = Join-Path $sqliteDir ($(if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }))
    if (Test-Path -LiteralPath $archDir) {
        $env:PATH = "$archDir;$sqliteDir;$env:PATH"
    }

    try {
        Add-Type -Path $sqliteDll
    } catch {
        throw "Не удалось загрузить встроенный SQLite-компонент. Установите Microsoft Visual C++ Redistributable 2015–2022 (x64 и x86), затем перезапустите приложение. Технические детали: $($_.Exception.Message)"
    }
    $script:SqliteLoaded = $true
}

function Get-PurchaseDataRoot {
    return (Join-Path $script:AppRoot 'data\purchase_control')
}

function Get-DefaultPurchaseDocumentsRoot {
    return (Join-Path (Get-PurchaseDataRoot) 'files')
}

function Get-ComponentDataRoot {
    return (Join-Path $script:AppRoot 'data\components')
}

function Get-DefaultComponentFilesRoot {
    return (Join-Path (Get-ComponentDataRoot) 'files')
}

function Get-NotesDataRoot {
    return (Join-Path $script:AppRoot 'data\notes')
}

function Get-DefaultNotesFilePath {
    return (Join-Path (Get-NotesDataRoot) 'notes.txt')
}

function Resolve-PortablePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $text = ([string]$Path).Trim()
    if ([IO.Path]::IsPathRooted($text)) { return $text }
    try {
        return [IO.Path]::GetFullPath((Join-Path $script:AppRoot $text))
    } catch {
        return (Join-Path $script:AppRoot $text)
    }
}

function Convert-ToPortablePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    try {
        $full = [IO.Path]::GetFullPath((Resolve-PortablePath $Path))
        $root = [IO.Path]::GetFullPath($script:AppRoot)
        if (-not $root.EndsWith('\')) { $root = "$root\" }
        if ($full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
            return $full.Substring($root.Length)
        }
    } catch {
    }
    return [string]$Path
}

function Get-MovedPortableTreeCandidate {
    param([string]$StoredPath, [string]$PortableTree)

    if ([string]::IsNullOrWhiteSpace($StoredPath) -or [string]::IsNullOrWhiteSpace($PortableTree)) { return '' }
    $separator = [string][IO.Path]::DirectorySeparatorChar
    $normalizedPath = ([string]$StoredPath).Replace([string][IO.Path]::AltDirectorySeparatorChar, $separator)
    $normalizedTree = ([string]$PortableTree).Trim([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar).Replace([string][IO.Path]::AltDirectorySeparatorChar, $separator)
    if ([string]::IsNullOrWhiteSpace($normalizedTree)) { return '' }

    $marker = "$normalizedTree$separator"
    $index = $normalizedPath.IndexOf($marker, [StringComparison]::OrdinalIgnoreCase)
    if ($index -ge 0) {
        $relative = $normalizedPath.Substring($index)
        return (Join-Path $script:AppRoot $relative)
    }
    if ($normalizedPath.EndsWith($normalizedTree, [StringComparison]::OrdinalIgnoreCase)) {
        return (Join-Path $script:AppRoot $normalizedTree)
    }
    return ''
}

function Resolve-PurchaseStoredPath {
    param([string]$StoredPath, [switch]$Directory)

    if ([string]::IsNullOrWhiteSpace($StoredPath)) { return '' }
    $pathType = if ($Directory) { 'Container' } else { 'Leaf' }
    $path = Resolve-PortablePath $StoredPath
    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType $pathType)) {
        try { return [IO.Path]::GetFullPath($path) } catch { return $path }
    }

    $candidate = Get-MovedPortableTreeCandidate $StoredPath 'data\purchase_control\files'
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType $pathType)) {
        try { return [IO.Path]::GetFullPath($candidate) } catch { return $candidate }
    }
    return $path
}

function Resolve-ComponentStoredPath {
    param([string]$StoredPath)

    if ([string]::IsNullOrWhiteSpace($StoredPath)) { return '' }
    $path = Resolve-PortablePath $StoredPath
    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType Container)) {
        try { return [IO.Path]::GetFullPath($path) } catch { return $path }
    }

    $candidate = Get-MovedPortableTreeCandidate $StoredPath 'data\components\files'
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType Container)) {
        try { return [IO.Path]::GetFullPath($candidate) } catch { return $candidate }
    }
    return $path
}

function Get-PurchaseDocumentsRoot {
    $root = Get-PurchaseSetting 'documents_root'
    if ([string]::IsNullOrWhiteSpace($root)) {
        $root = Convert-ToPortablePath (Get-DefaultPurchaseDocumentsRoot)
        Set-PurchaseSetting 'documents_root' $root
    }

    $movedRoot = Get-MovedPortableTreeCandidate $root 'data\purchase_control\files'
    $resolved = if (-not [string]::IsNullOrWhiteSpace($movedRoot)) { $movedRoot } else { Resolve-PortablePath $root }
    if ([string]::IsNullOrWhiteSpace($resolved)) { $resolved = Get-DefaultPurchaseDocumentsRoot }
    return $resolved
}

function Set-PurchaseDocumentsRoot {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Documents folder is required.' }
    $resolved = Resolve-PortablePath $Path
    New-Item -ItemType Directory -Force -Path $resolved | Out-Null
    Set-PurchaseSetting 'documents_root' (Convert-ToPortablePath $resolved)
    $script:PurchaseDocumentsRoot = $resolved
    return $resolved
}

function Repair-PortablePurchasePaths {
    try {
        [void](Set-PurchaseDocumentsRoot (Get-PurchaseDocumentsRoot))
    } catch {
    }

    $docs = Invoke-PurchaseQuery "SELECT id, stored_path FROM documents WHERE IFNULL(stored_path, '') <> ''" @{}
    foreach ($row in $docs.Rows) {
        $stored = [string]$row.stored_path
        $resolved = Resolve-PurchaseStoredPath $stored
        if ([string]::IsNullOrWhiteSpace($resolved) -or -not (Test-Path -LiteralPath $resolved -PathType Leaf)) { continue }
        $portable = Convert-ToPortablePath $resolved
        if ($portable -ne $stored) {
            Invoke-PurchaseNonQuery 'UPDATE documents SET stored_path = @stored_path WHERE id = @id' @{
                '@stored_path' = $portable
                '@id' = [int]$row.id
            }
        }
    }
}

function Repair-PortableComponentPaths {
    $rows = Invoke-PurchaseQuery "SELECT id, folder_path FROM component_deals WHERE IFNULL(folder_path, '') <> ''" @{}
    foreach ($row in $rows.Rows) {
        $stored = [string]$row.folder_path
        $resolved = Resolve-ComponentStoredPath $stored
        if ([string]::IsNullOrWhiteSpace($resolved) -or -not (Test-Path -LiteralPath $resolved -PathType Container)) { continue }
        $portable = Convert-ToPortablePath $resolved
        if ($portable -ne $stored) {
            Invoke-PurchaseNonQuery 'UPDATE component_deals SET folder_path = @folder_path WHERE id = @id' @{
                '@folder_path' = $portable
                '@id' = [int]$row.id
            }
        }
    }
}

function Get-NotesFilePath {
    if (-not [string]::IsNullOrWhiteSpace([string]$script:CurrentNotesFilePath)) {
        return [string]$script:CurrentNotesFilePath
    }
    return Get-DefaultNotesFilePath
}

function Get-NoteDisplayName {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    return [IO.Path]::GetFileNameWithoutExtension($Path)
}

function Get-SafeNoteFileName {
    param([string]$Title)

    $name = if ([string]::IsNullOrWhiteSpace($Title)) { 'Новая заметка' } else { $Title.Trim() }
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) {
        $name = $name.Replace([string]$char, '_')
    }
    $name = ($name -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'Новая заметка' }
    if (-not $name.EndsWith('.txt', [StringComparison]::OrdinalIgnoreCase)) { $name = "$name.txt" }
    return $name
}

function Get-ProjectNoteFiles {
    $root = Get-NotesDataRoot
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $files = @(Get-ChildItem -LiteralPath $root -Filter '*.txt' -File -ErrorAction SilentlyContinue | Sort-Object BaseName)
    if ($files.Count -eq 0) {
        $defaultPath = Join-Path $root 'Общие инструкции.txt'
        Set-Content -LiteralPath $defaultPath -Value '' -Encoding UTF8 -NoNewline
        $files = @(Get-ChildItem -LiteralPath $root -Filter '*.txt' -File -ErrorAction SilentlyContinue | Sort-Object BaseName)
    }
    return $files
}

function New-ProjectNoteFile {
    param([string]$Title)

    $root = Get-NotesDataRoot
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $fileName = Get-SafeNoteFileName $Title
    $stem = [IO.Path]::GetFileNameWithoutExtension($fileName)
    $extension = [IO.Path]::GetExtension($fileName)
    $path = Join-Path $root $fileName
    for ($index = 2; (Test-Path -LiteralPath $path); $index++) {
        $path = Join-Path $root ('{0} ({1}){2}' -f $stem, $index, $extension)
    }
    Set-Content -LiteralPath $path -Value '' -Encoding UTF8 -NoNewline
    $script:CurrentNotesFilePath = $path
    return $path
}

function Read-ProjectNotes {
    $path = Get-NotesFilePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return '' }
    return Get-Content -LiteralPath $path -Raw -Encoding UTF8
}

function Save-ProjectNotes {
    param([string]$Text)

    $root = Get-NotesDataRoot
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $path = Get-NotesFilePath
    if ([string]::IsNullOrWhiteSpace($path)) { $path = Get-DefaultNotesFilePath }
    Set-Content -LiteralPath $path -Value $Text -Encoding UTF8 -NoNewline
    $script:CurrentNotesFilePath = $path
}
function Open-PurchaseConnection {
    Load-SqliteProvider
    $dataRoot = Get-PurchaseDataRoot
    New-Item -ItemType Directory -Force -Path $dataRoot | Out-Null
    if ([string]::IsNullOrWhiteSpace($script:PurchaseDbPath)) {
        $script:PurchaseDbPath = Join-Path $dataRoot 'purchase_control.sqlite'
    }

    $connection = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$script:PurchaseDbPath;Version=3;Foreign Keys=True;Pooling=True;Journal Mode=WAL;Synchronous=Normal;")
    $connection.Open()
    return $connection
}

function Add-SqliteParameters {
    param($Command, [hashtable]$Parameters)

    foreach ($key in @($Parameters.Keys)) {
        $parameter = $Command.CreateParameter()
        $parameter.ParameterName = $key
        $value = $Parameters[$key]
        $parameter.Value = if ($null -eq $value) { [DBNull]::Value } else { $value }
        [void]$Command.Parameters.Add($parameter)
    }
}

function Invoke-PurchaseNonQuery {
    param([string]$Sql, [hashtable]$Parameters = @{})

    $connection = Open-PurchaseConnection
    try {
        $command = $connection.CreateCommand()
        $command.CommandText = $Sql
        Add-SqliteParameters $command $Parameters
        [void]$command.ExecuteNonQuery()
    } finally {
        if ($null -ne $connection) { $connection.Dispose() }
    }
}

function Invoke-PurchaseScalar {
    param([string]$Sql, [hashtable]$Parameters = @{})

    $connection = Open-PurchaseConnection
    try {
        $command = $connection.CreateCommand()
        $command.CommandText = $Sql
        Add-SqliteParameters $command $Parameters
        $value = $command.ExecuteScalar()
        if ($value -is [DBNull]) { return $null }
        return $value
    } finally {
        if ($null -ne $connection) { $connection.Dispose() }
    }
}

function Invoke-PurchaseQuery {
    param([string]$Sql, [hashtable]$Parameters = @{})

    $connection = Open-PurchaseConnection
    try {
        $command = $connection.CreateCommand()
        $command.CommandText = $Sql
        Add-SqliteParameters $command $Parameters
        $reader = $command.ExecuteReader()
        $table = New-Object System.Data.DataTable
        $table.Load($reader)
        return ,$table
    } finally {
        if ($null -ne $connection) { $connection.Dispose() }
    }
}

function Test-PurchaseColumnExists {
    param([string]$TableName, [string]$ColumnName)

    if ($TableName -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
        throw "Некорректное имя таблицы: $TableName"
    }

    $table = Invoke-PurchaseQuery "PRAGMA table_info($TableName)"
    foreach ($row in $table.Rows) {
        if ([string]$row.name -eq $ColumnName) {
            return $true
        }
    }

    return $false
}

function Ensure-PurchaseColumn {
    param([string]$TableName, [string]$ColumnName, [string]$Definition)

    if (-not (Test-PurchaseColumnExists $TableName $ColumnName)) {
        Invoke-PurchaseNonQuery "ALTER TABLE $TableName ADD COLUMN $Definition"
    }
}

function Initialize-PurchaseStore {
    $dataRoot = Get-PurchaseDataRoot
    New-Item -ItemType Directory -Force -Path $dataRoot | Out-Null
    $script:PurchaseDbPath = Join-Path $dataRoot 'purchase_control.sqlite'

    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS bitrix_blocked_tasks (
    bitrix_task_id INTEGER PRIMARY KEY,
    title TEXT,
    blocked_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS deals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_number TEXT NOT NULL UNIQUE,
    board_count TEXT,
    client TEXT,
    period TEXT,
    executor TEXT,
    reminder_date TEXT,
    assembly_location TEXT,
    status TEXT NOT NULL DEFAULT 'RFQ',
    tracking_status TEXT NOT NULL DEFAULT 'Ожидание',
    title TEXT,
    comment TEXT,
    created_at TEXT NOT NULL,
    ordered_at TEXT,
    updated_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS bitrix_task_links (
    bitrix_task_id INTEGER PRIMARY KEY,
    deal_id INTEGER,
    bitrix_deal_id INTEGER,
    bitrix_status INTEGER,
    is_manual_deleted INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    linked_at TEXT,
    last_seen_at TEXT,
    deleted_at TEXT,
    FOREIGN KEY(deal_id) REFERENCES deals(id) ON DELETE SET NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS bitrix_rfq_files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bitrix_task_id INTEGER NOT NULL,
    file_key TEXT NOT NULL,
    bitrix_file_id TEXT,
    file_name TEXT,
    file_hash TEXT,
    local_path TEXT,
    upload_state TEXT,
    error_text TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(bitrix_task_id, file_key),
    FOREIGN KEY(bitrix_task_id) REFERENCES bitrix_task_links(bitrix_task_id) ON DELETE CASCADE
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS deal_suppliers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_id INTEGER NOT NULL,
    supplier TEXT NOT NULL,
    invoice_received INTEGER NOT NULL DEFAULT 0,
    invoice_confirmed INTEGER NOT NULL DEFAULT 0,
    supplier_order_created INTEGER NOT NULL DEFAULT 0,
    erp_supplier_sent INTEGER NOT NULL DEFAULT 0,
    erp_roger_sent INTEGER NOT NULL DEFAULT 0,
    pi_amount_usd REAL,
    pi_amount_cny REAL,
    paid_amount TEXT,
    delivery_weeks INTEGER,
    payment_submitted INTEGER NOT NULL DEFAULT 0,
    paid INTEGER NOT NULL DEFAULT 0,
    invoice_confirmed_date TEXT,
    components_receipt_date TEXT,
    actual_receipt_date TEXT,
    comment TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(deal_id, supplier),
    FOREIGN KEY(deal_id) REFERENCES deals(id) ON DELETE CASCADE
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_id INTEGER NOT NULL,
    supplier_id INTEGER,
    document_type TEXT NOT NULL,
    original_name TEXT NOT NULL,
    stored_path TEXT NOT NULL,
    uploaded_at TEXT NOT NULL,
    FOREIGN KEY(deal_id) REFERENCES deals(id) ON DELETE CASCADE,
    FOREIGN KEY(supplier_id) REFERENCES deal_suppliers(id) ON DELETE SET NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS component_deals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT NOT NULL,
    deal_number TEXT,
    status TEXT NOT NULL DEFAULT 'В работе',
    stage TEXT NOT NULL DEFAULT 'Запросил поставщиков',
    description TEXT,
    next_action TEXT,
    reminder_date TEXT,
    deadline_date TEXT,
    priority TEXT NOT NULL DEFAULT '3',
    period TEXT,
    folder_path TEXT,
    notes TEXT,
    created_at TEXT NOT NULL,
    ordered_at TEXT,
    updated_at TEXT NOT NULL
);
'@
    Ensure-PurchaseColumn 'component_deals' 'source' 'source TEXT'
    Ensure-PurchaseColumn 'component_deals' 'source_text' 'source_text TEXT'
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS activity_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_type TEXT NOT NULL,
    entity_id INTEGER,
    deal_id INTEGER,
    supplier_id INTEGER,
    action TEXT NOT NULL,
    details TEXT,
    created_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS reminders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_id INTEGER,
    supplier_id INTEGER,
    component_id INTEGER,
    title TEXT NOT NULL,
    due_date TEXT,
    status TEXT NOT NULL DEFAULT 'Open',
    source TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS trash_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_type TEXT NOT NULL,
    entity_id INTEGER,
    title TEXT NOT NULL,
    payload_json TEXT NOT NULL,
    files_json TEXT,
    deleted_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS notification_state (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source TEXT NOT NULL,
    source_id INTEGER NOT NULL,
    due_kind TEXT NOT NULL,
    due_date TEXT NOT NULL,
    handled INTEGER NOT NULL DEFAULT 0,
    snooze_until TEXT,
    last_shown_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(source, source_id, due_kind, due_date)
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS workflow_templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    is_default INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS workflow_steps (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id INTEGER NOT NULL,
    step_order INTEGER NOT NULL,
    name TEXT NOT NULL,
    FOREIGN KEY(template_id) REFERENCES workflow_templates(id) ON DELETE CASCADE
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS quote_batches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT NOT NULL,
    rfq_path TEXT,
    priority TEXT,
    comment TEXT
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS quote_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER NOT NULL,
    quote_date TEXT NOT NULL,
    rfq_value TEXT,
    pn TEXT,
    supplier TEXT,
    unit_price REAL,
    lead_time TEXT,
    lead_time_total TEXT,
    mfg TEXT,
    is_winner INTEGER NOT NULL DEFAULT 0,
    warning TEXT,
    sheet_name TEXT,
    row_number INTEGER,
    match_status TEXT,
    winner_reason TEXT,
    FOREIGN KEY(batch_id) REFERENCES quote_batches(id) ON DELETE CASCADE
);
'@
    $globalistSchema = @(
        'CREATE TABLE IF NOT EXISTS globalist_quotes (',
        '    id INTEGER PRIMARY KEY AUTOINCREMENT,',
        '    imported_at TEXT NOT NULL,', '    source_file TEXT,', '    sheet_name TEXT,', '    row_number INTEGER,',
        '    factory TEXT,', '    pn TEXT,', '    comment TEXT,', '    pi_number TEXT,', '    replacement TEXT,',
        '    chinese_remark TEXT,', '    package TEXT,', '    brand TEXT,', '    datacode TEXT,', '    moq TEXT,',
        '    qty TEXT,', '    stock TEXT,', '    need_spq TEXT,', '    spq TEXT,', '    unit_price REAL,',
        '    total_amount REAL,', '    lead_time TEXT,', '    weight TEXT,', '    target TEXT,', '    supplier_quote_id TEXT',
        ');'
    ) -join "`n"
    Invoke-PurchaseNonQuery $globalistSchema @{}

    Ensure-PurchaseColumn 'quote_batches' 'source_type' 'source_type TEXT NOT NULL DEFAULT ''Our'''
    Ensure-PurchaseColumn 'deals' 'client' 'client TEXT'
    Ensure-PurchaseColumn 'deals' 'status' "status TEXT NOT NULL DEFAULT 'RFQ'"
    Ensure-PurchaseColumn 'deals' 'tracking_status' "tracking_status TEXT NOT NULL DEFAULT 'Ожидание'"
    Ensure-PurchaseColumn 'deals' 'board_count' 'board_count TEXT'
    Ensure-PurchaseColumn 'deals' 'period' 'period TEXT'
    Ensure-PurchaseColumn 'deals' 'executor' 'executor TEXT'
    Ensure-PurchaseColumn 'deals' 'reminder_date' 'reminder_date TEXT'
    Ensure-PurchaseColumn 'deals' 'assembly_location' 'assembly_location TEXT'
    Ensure-PurchaseColumn 'deals' 'bitrix_deal_id' 'bitrix_deal_id INTEGER'
    Ensure-PurchaseColumn 'deals' 'priority' "priority TEXT NOT NULL DEFAULT '3'"
    Ensure-PurchaseColumn 'deals' 'archived' 'archived INTEGER NOT NULL DEFAULT 0'
    Ensure-PurchaseColumn 'deals' 'workflow_template_id' 'workflow_template_id INTEGER'
    Ensure-PurchaseColumn 'deals' 'pi_sent' 'pi_sent INTEGER NOT NULL DEFAULT 0'
    Ensure-PurchaseColumn 'deals' 'masks' 'masks INTEGER NOT NULL DEFAULT 0'
    Ensure-PurchaseColumn 'deals' 'ordered_at' 'ordered_at TEXT'
    Ensure-PurchaseColumn 'documents' 'file_hash' 'file_hash TEXT'
    Ensure-PurchaseColumn 'deal_suppliers' 'pi_amount_usd' 'pi_amount_usd REAL'
    Ensure-PurchaseColumn 'deal_suppliers' 'pi_amount_cny' 'pi_amount_cny REAL'
    Ensure-PurchaseColumn 'deal_suppliers' 'pi_amount_rub' 'pi_amount_rub REAL'
    Ensure-PurchaseColumn 'deal_suppliers' 'paid_amount' 'paid_amount TEXT'
    Ensure-PurchaseColumn 'deal_suppliers' 'delivery_weeks' 'delivery_weeks INTEGER'
    Ensure-PurchaseColumn 'deal_suppliers' 'payment_submitted' 'payment_submitted INTEGER NOT NULL DEFAULT 0'
    Ensure-PurchaseColumn 'deal_suppliers' 'paid' 'paid INTEGER NOT NULL DEFAULT 0'
    Ensure-PurchaseColumn 'deal_suppliers' 'actual_receipt_date' 'actual_receipt_date TEXT'
    Ensure-PurchaseColumn 'component_deals' 'folder_path' 'folder_path TEXT'
    Ensure-PurchaseColumn 'component_deals' 'notes' 'notes TEXT'
    Ensure-PurchaseColumn 'component_deals' 'ordered_at' 'ordered_at TEXT'

    if ([string]::IsNullOrWhiteSpace((Get-PurchaseSetting 'dashboard.orders_tracking_started_at'))) {
        Set-PurchaseSetting 'dashboard.orders_tracking_started_at' (Get-NowText)
    }

    Initialize-DefaultWorkflowTemplates

    $documentsRoot = Get-PurchaseDocumentsRoot
    $script:PurchaseDocumentsRoot = $documentsRoot
    New-Item -ItemType Directory -Force -Path $script:PurchaseDocumentsRoot | Out-Null
    New-Item -ItemType Directory -Force -Path (Get-DefaultComponentFilesRoot) | Out-Null
    New-Item -ItemType Directory -Force -Path (Get-NotesDataRoot) | Out-Null
    Repair-PortablePurchasePaths
    Repair-PortableComponentPaths
}

function Get-PurchaseSetting {
    param([string]$Key)

    return [string](Invoke-PurchaseScalar 'SELECT value FROM settings WHERE key = @key' @{ '@key' = $Key })
}

function Set-PurchaseSetting {
    param([string]$Key, [string]$Value)

    Invoke-PurchaseNonQuery 'INSERT INTO settings(key, value) VALUES(@key, @value) ON CONFLICT(key) DO UPDATE SET value = excluded.value' @{
        '@key' = $Key
        '@value' = $Value
    }
}

function Get-BitrixBlockedTaskIds {
    return @(Invoke-PurchaseQuery 'SELECT bitrix_task_id FROM bitrix_blocked_tasks ORDER BY bitrix_task_id' @{} | ForEach-Object { [int]$_.bitrix_task_id })
}

function Block-BitrixTask {
    param([int]$TaskId, [string]$Title)

    if ($TaskId -le 0) { return }
    Invoke-PurchaseNonQuery 'INSERT OR IGNORE INTO bitrix_blocked_tasks(bitrix_task_id, title, blocked_at) VALUES(@task_id, @title, @blocked_at)' @{
        '@task_id' = $TaskId
        '@title' = $Title
        '@blocked_at' = Get-NowText
    }
}

function Get-NowText {
    return (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}

function Convert-PurchaseRowsToPayload {
    param($Table)

    $items = New-Object System.Collections.ArrayList
    if ($null -eq $Table) { return @() }
    foreach ($row in $Table.Rows) {
        $item = [ordered]@{}
        foreach ($column in $Table.Columns) {
            $value = $row[$column.ColumnName]
            $item[$column.ColumnName] = if ($null -eq $value -or $value -is [DBNull]) { $null } else { $value }
        }
        [void]$items.Add([pscustomobject]$item)
    }
    return @($items)
}

function Get-PurchaseTrashFilesRoot {
    return (Join-Path (Get-PurchaseDataRoot) 'trash_files')
}

function New-PurchaseTrashItem {
    param(
        [string]$EntityType,
        [int]$EntityId,
        [string]$Title,
        [hashtable]$Payload,
        [string[]]$FilePaths = @(),
        [string]$ComponentFolder = ''
    )

    $json = ($Payload | ConvertTo-Json -Depth 8 -Compress)
    $now = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO trash_items(entity_type, entity_id, title, payload_json, files_json, deleted_at)
VALUES(@entity_type, @entity_id, @title, @payload_json, '', @deleted_at)
'@ @{
        '@entity_type' = $EntityType; '@entity_id' = $EntityId; '@title' = $Title; '@payload_json' = $json; '@deleted_at' = $now
    }
    $trashId = [int](Invoke-PurchaseScalar 'SELECT id FROM trash_items ORDER BY id DESC LIMIT 1' @{})
    $itemRoot = Join-Path (Get-PurchaseTrashFilesRoot) $trashId
    $fileRecords = New-Object System.Collections.ArrayList
    try {
        New-Item -ItemType Directory -Force -Path $itemRoot | Out-Null
        $number = 0
        foreach ($filePath in @($FilePaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
            if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) { continue }
            $number++
            $target = Join-Path $itemRoot ('file_{0:D4}_{1}' -f $number, [IO.Path]::GetFileName($filePath))
            Move-Item -LiteralPath $filePath -Destination $target -Force -ErrorAction Stop
            [void]$fileRecords.Add([pscustomobject]@{ OriginalPath = $filePath; TrashPath = $target; IsFolder = $false })
        }
        if (-not [string]::IsNullOrWhiteSpace($ComponentFolder) -and (Test-Path -LiteralPath $ComponentFolder -PathType Container)) {
            $folderTarget = Join-Path $itemRoot 'component_folder'
            Move-Item -LiteralPath $ComponentFolder -Destination $folderTarget -Force -ErrorAction Stop
            [void]$fileRecords.Add([pscustomobject]@{ OriginalPath = $ComponentFolder; TrashPath = $folderTarget; IsFolder = $true })
        }
        Invoke-PurchaseNonQuery 'UPDATE trash_items SET files_json = @files_json WHERE id = @id' @{
            '@files_json' = ($fileRecords | ConvertTo-Json -Depth 4 -Compress); '@id' = $trashId
        }
        return $trashId
    } catch {
        # The database is intentionally not changed by the caller until this succeeds.
        throw "Не удалось переместить данные в корзину: $($_.Exception.Message)"
    }
}

function Get-PurchaseTrashItems {
    return Invoke-PurchaseQuery @'
SELECT id, entity_type, entity_id, title, deleted_at,
       CASE WHEN IFNULL(files_json, '') = '' THEN 0 ELSE 1 END AS has_files
FROM trash_items
ORDER BY deleted_at DESC, id DESC
'@ @{}
}

function Restore-PurchaseTrashRows {
    param([string]$TableName, $Rows)

    $allowed = @('deals', 'deal_suppliers', 'documents', 'component_deals', 'reminders', 'bitrix_task_links', 'quote_batches', 'quote_history', 'activity_log')
    if ($TableName -notin $allowed -or $null -eq $Rows) { return }
    foreach ($row in @($Rows)) {
        $properties = @($row.PSObject.Properties | Where-Object { $_.Name -match '^[A-Za-z_][A-Za-z0-9_]*$' })
        if ($properties.Count -eq 0) { continue }
        $columns = @($properties | ForEach-Object { $_.Name })
        $parameters = @{}
        $names = New-Object System.Collections.ArrayList
        foreach ($column in $columns) {
            $name = '@p_' + $column
            $parameters[$name] = $row.$column
            [void]$names.Add($name)
        }
        $sql = 'INSERT OR REPLACE INTO {0} ({1}) VALUES ({2})' -f $TableName, ($columns -join ', '), (@($names) -join ', ')
        Invoke-PurchaseNonQuery $sql $parameters
    }
}

function Restore-PurchaseTrashItem {
    param([int]$TrashId)

    if ($TrashId -le 0) { throw 'Выберите запись в корзине.' }
    $table = Invoke-PurchaseQuery 'SELECT * FROM trash_items WHERE id = @id' @{ '@id' = $TrashId }
    if ($table.Rows.Count -eq 0) { throw 'Запись корзины не найдена.' }
    $item = $table.Rows[0]
    $payload = ([string]$item.payload_json | ConvertFrom-Json)
    foreach ($tableName in @('deals', 'component_deals', 'deal_suppliers', 'reminders', 'documents', 'quote_batches', 'quote_history', 'activity_log', 'bitrix_task_links')) {
        if ($null -ne $payload.PSObject.Properties[$tableName]) {
            Restore-PurchaseTrashRows $tableName $payload.$tableName
        }
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$item.files_json)) {
        foreach ($record in @([string]$item.files_json | ConvertFrom-Json)) {
            $source = [string]$record.TrashPath
            $target = [string]$record.OriginalPath
            if (-not (Test-Path -LiteralPath $source)) { continue }
            $parent = Split-Path -Parent $target
            if (-not [string]::IsNullOrWhiteSpace($parent)) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }
            Move-Item -LiteralPath $source -Destination $target -Force -ErrorAction Stop
        }
    }
    $itemRoot = Join-Path (Get-PurchaseTrashFilesRoot) $TrashId
    if (Test-Path -LiteralPath $itemRoot -PathType Container) { Remove-Item -LiteralPath $itemRoot -Recurse -Force }
    Invoke-PurchaseNonQuery 'DELETE FROM trash_items WHERE id = @id' @{ '@id' = $TrashId }
    Write-ActivityLog 'trash' $TrashId 'Восстановлено из корзины' ([string]$item.title) 0 0
}

function Remove-PurchaseTrashItem {
    param([int]$TrashId)

    if ($TrashId -le 0) { throw 'Выберите запись в корзине.' }
    $itemRoot = Join-Path (Get-PurchaseTrashFilesRoot) $TrashId
    if (Test-Path -LiteralPath $itemRoot -PathType Container) { Remove-Item -LiteralPath $itemRoot -Recurse -Force }
    Invoke-PurchaseNonQuery 'DELETE FROM trash_items WHERE id = @id' @{ '@id' = $TrashId }
}

function Get-PurchaseOrderedStatus {
    return -join @([char]0x0417, [char]0x0430, [char]0x043A, [char]0x0430, [char]0x0437, [char]0x0430, [char]0x043D, [char]0x043E)
}

function Test-PurchaseOrderedStatus {
    param([string]$Status)

    return (([string]$Status).Trim() -eq (Get-PurchaseOrderedStatus))
}

function Initialize-DefaultWorkflowTemplates {
    $exists = [int](Invoke-PurchaseScalar "SELECT COUNT(*) FROM workflow_templates WHERE name = 'Базовая закупка'" @{})
    if ($exists -gt 0) { return }

    $now = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO workflow_templates(name, is_default, created_at)
VALUES('Базовая закупка', 1, @created_at)
'@ @{ '@created_at' = $now }
    $templateId = [int](Invoke-PurchaseScalar "SELECT id FROM workflow_templates WHERE name = 'Базовая закупка'" @{})
    $order = 1
    foreach ($step in @('RFQ', 'RRFQ', 'PO', 'PI', 'Подано в оплату', 'Оплачено', 'ERP', 'Поступление')) {
        Invoke-PurchaseNonQuery @'
INSERT INTO workflow_steps(template_id, step_order, name)
VALUES(@template_id, @step_order, @name)
'@ @{
            '@template_id' = $templateId
            '@step_order' = $order
            '@name' = $step
        }
        $order++
    }
}

function Write-ActivityLog {
    param(
        [string]$EntityType,
        [int]$EntityId = 0,
        [string]$Action,
        [string]$Details = '',
        [int]$DealId = 0,
        [int]$SupplierId = 0
    )

    if ([string]::IsNullOrWhiteSpace($Action)) { return }
    try {
        Invoke-PurchaseNonQuery @'
INSERT INTO activity_log(entity_type, entity_id, deal_id, supplier_id, action, details, created_at)
VALUES(@entity_type, @entity_id, @deal_id, @supplier_id, @action, @details, @created_at)
'@ @{
            '@entity_type' = if ([string]::IsNullOrWhiteSpace($EntityType)) { 'system' } else { $EntityType }
            '@entity_id' = if ($EntityId -gt 0) { $EntityId } else { $null }
            '@deal_id' = if ($DealId -gt 0) { $DealId } else { $null }
            '@supplier_id' = if ($SupplierId -gt 0) { $SupplierId } else { $null }
            '@action' = $Action
            '@details' = $Details
            '@created_at' = Get-NowText
        }
    } catch {
        # The log should never block the user's main workflow.
    }
}

function Get-ActivityLog {
    param([int]$DealId = 0, [int]$Limit = 200, [string]$Search = '')

    $whereParts = New-Object System.Collections.ArrayList
    $params = @{ '@limit' = [Math]::Max(1, $Limit) }
    if ($DealId -gt 0) {
        [void]$whereParts.Add('deal_id = @deal_id')
        $params['@deal_id'] = $DealId
    }
    if (-not [string]::IsNullOrWhiteSpace($Search)) {
        [void]$whereParts.Add("(IFNULL(action, '') LIKE @needle OR IFNULL(details, '') LIKE @needle OR IFNULL(entity_type, '') LIKE @needle OR CAST(IFNULL(entity_id, '') AS TEXT) LIKE @needle OR CAST(IFNULL(deal_id, '') AS TEXT) LIKE @needle OR CAST(IFNULL(supplier_id, '') AS TEXT) LIKE @needle)")
        $params['@needle'] = '%' + $Search.Trim() + '%'
    }
    $where = if ($whereParts.Count -gt 0) { 'WHERE ' + [string]::Join(' AND ', [string[]]$whereParts.ToArray()) } else { '' }

    return Invoke-PurchaseQuery @"
SELECT id, entity_type, entity_id, deal_id, supplier_id, action, IFNULL(details, '') AS details, created_at
FROM activity_log
$where
ORDER BY id DESC
LIMIT @limit
"@ $params
}

function Get-DefaultWorkflowTemplateId {
    $id = Invoke-PurchaseScalar 'SELECT id FROM workflow_templates WHERE is_default = 1 ORDER BY id LIMIT 1' @{}
    if ($null -eq $id -or $id -is [DBNull]) { return $null }
    return [int]$id
}

function Get-WorkflowTemplates {
    return Invoke-PurchaseQuery 'SELECT id, name, is_default, created_at FROM workflow_templates ORDER BY is_default DESC, name' @{}
}

function Get-WorkflowSteps {
    param([int]$TemplateId = 0)
    if ($TemplateId -le 0) {
        $TemplateId = Get-DefaultWorkflowTemplateId
    }
    if ($null -eq $TemplateId -or $TemplateId -le 0) {
        return New-Object System.Data.DataTable
    }
    return Invoke-PurchaseQuery @'
SELECT id, template_id, step_order, name
FROM workflow_steps
WHERE template_id = @template_id
ORDER BY step_order
'@ @{ '@template_id' = $TemplateId }
}

function Set-PurchaseDealArchived {
    param([int]$DealId, [bool]$Archived)

    if ($DealId -le 0) { throw 'Выберите сделку.' }
    Invoke-PurchaseNonQuery 'UPDATE deals SET archived = @archived, updated_at = @updated_at WHERE id = @id' @{
        '@archived' = Convert-ToDbBool $Archived
        '@updated_at' = Get-NowText
        '@id' = $DealId
    }
    Write-ActivityLog 'deal' $DealId $(if ($Archived) { 'Сделка отправлена в архив' } else { 'Сделка возвращена из архива' }) '' $DealId
}

function Set-PurchaseDealWorkflowTemplate {
    param([int]$DealId, [int]$TemplateId)

    if ($DealId -le 0) { throw 'Выберите сделку.' }
    if ($TemplateId -le 0) { throw 'Выберите шаблон этапов.' }
    $templateName = Invoke-PurchaseScalar 'SELECT name FROM workflow_templates WHERE id = @id' @{ '@id' = $TemplateId }
    if ($null -eq $templateName -or $templateName -is [DBNull]) { throw 'Шаблон этапов не найден.' }
    Invoke-PurchaseNonQuery 'UPDATE deals SET workflow_template_id = @template_id, updated_at = @updated_at WHERE id = @deal_id' @{
        '@template_id' = $TemplateId
        '@updated_at' = Get-NowText
        '@deal_id' = $DealId
    }
    Write-ActivityLog 'deal' $DealId 'Изменен шаблон этапов' ([string]$templateName) $DealId
}

function Save-Reminder {
    param(
        [string]$Title,
        [string]$DueDate = '',
        [int]$DealId = 0,
        [int]$SupplierId = 0,
        [int]$ComponentId = 0,
        [string]$Source = 'manual'
    )

    if ([string]::IsNullOrWhiteSpace($Title)) { throw 'Введите текст напоминания.' }
    $now = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO reminders(deal_id, supplier_id, component_id, title, due_date, status, source, created_at, updated_at)
VALUES(@deal_id, @supplier_id, @component_id, @title, @due_date, 'Open', @source, @created_at, @updated_at)
'@ @{
        '@deal_id' = if ($DealId -gt 0) { $DealId } else { $null }
        '@supplier_id' = if ($SupplierId -gt 0) { $SupplierId } else { $null }
        '@component_id' = if ($ComponentId -gt 0) { $ComponentId } else { $null }
        '@title' = $Title
        '@due_date' = $DueDate
        '@source' = $Source
        '@created_at' = $now
        '@updated_at' = $now
    }
    Write-ActivityLog 'reminder' 0 'Создано напоминание' $Title $DealId $SupplierId
}

function Set-ReminderDone {
    param([int]$ReminderId, [bool]$Done)

    if ($ReminderId -le 0) { throw 'Выберите напоминание.' }
    Invoke-PurchaseNonQuery 'UPDATE reminders SET status = @status, updated_at = @updated_at WHERE id = @id' @{
        '@status' = if ($Done) { 'Done' } else { 'Open' }
        '@updated_at' = Get-NowText
        '@id' = $ReminderId
    }
}

function Get-NotificationState {
    param(
        [string]$Source,
        [int]$SourceId,
        [string]$DueKind,
        [string]$DueDate
    )

    return Invoke-PurchaseQuery @'
SELECT id, handled, IFNULL(snooze_until, '') AS snooze_until
FROM notification_state
WHERE source = @source AND source_id = @source_id
  AND due_kind = @due_kind AND due_date = @due_date
'@ @{
        '@source' = $Source
        '@source_id' = $SourceId
        '@due_kind' = $DueKind
        '@due_date' = $DueDate
    }
}

function Get-PurchaseBooleanSetting {
    param(
        [string]$Key,
        [bool]$DefaultValue
    )

    $rawValue = ([string](Get-PurchaseSetting $Key)).Trim().ToLowerInvariant()
    if ($rawValue -in @('1', 'true', 'yes')) { return $true }
    if ($rawValue -in @('0', 'false', 'no')) { return $false }

    Set-PurchaseSetting $Key ($(if ($DefaultValue) { '1' } else { '0' }))
    return $DefaultValue
}

function Set-NotificationState {
    param(
        [string]$Source,
        [int]$SourceId,
        [string]$DueKind,
        [string]$DueDate,
        [bool]$Handled = $false,
        [string]$SnoozeUntil = '',
        [bool]$Shown = $false
    )

    $now = Get-NowText
    $shownAt = if ($Shown) { $now } else { $null }
    Invoke-PurchaseNonQuery @'
INSERT INTO notification_state(source, source_id, due_kind, due_date, handled, snooze_until, last_shown_at, created_at, updated_at)
VALUES(@source, @source_id, @due_kind, @due_date, @handled, @snooze_until, @last_shown_at, @created_at, @updated_at)
ON CONFLICT(source, source_id, due_kind, due_date) DO UPDATE SET
    handled = @handled,
    snooze_until = @snooze_until,
    last_shown_at = CASE WHEN @last_shown_at IS NULL THEN notification_state.last_shown_at ELSE @last_shown_at END,
    updated_at = @updated_at
'@ @{
        '@source' = $Source
        '@source_id' = $SourceId
        '@due_kind' = $DueKind
        '@due_date' = $DueDate
        '@handled' = Convert-ToDbBool $Handled
        '@snooze_until' = if ([string]::IsNullOrWhiteSpace($SnoozeUntil)) { $null } else { $SnoozeUntil }
        '@last_shown_at' = $shownAt
        '@created_at' = $now
        '@updated_at' = $now
    }
}

function Get-DueNotificationCandidates {
    $today = (Get-Date).Date
    $currentTime = Get-Date
    $items = New-Object System.Collections.ArrayList
    $stateByKey = @{}
    $states = Invoke-PurchaseQuery @'
SELECT source, source_id, due_kind, due_date, handled, IFNULL(snooze_until, '') AS snooze_until
FROM notification_state
'@ @{}
    foreach ($state in $states.Rows) {
        $stateKey = '{0}|{1}|{2}|{3}' -f [string]$state.source, [int]$state.source_id, [string]$state.due_kind, [string]$state.due_date
        $stateByKey[$stateKey] = $state
    }

    $components = Invoke-PurchaseQuery @'
SELECT id, IFNULL(deal_number, '') AS deal_number,
       IFNULL(description, '') AS description,
       IFNULL(next_action, '') AS next_action,
       IFNULL(status, '') AS status,
       IFNULL(stage, '') AS stage,
       IFNULL(reminder_date, '') AS reminder_date,
       IFNULL(deadline_date, '') AS deadline_date
FROM component_deals
WHERE IFNULL(status, '') NOT IN ('Выполнено', 'Не актуально')
  AND (IFNULL(reminder_date, '') <> '' OR IFNULL(deadline_date, '') <> '')
ORDER BY id DESC
'@ @{}

    foreach ($row in $components.Rows) {
        $due = New-Object System.Collections.ArrayList
        foreach ($entry in @(@('reminder', 'Напоминание', [string]$row.reminder_date), @('deadline', 'Дедлайн', [string]$row.deadline_date))) {
            $date = Convert-PurchaseDateText $entry[2]
            if ($null -eq $date -or $date.Date -gt $today) { continue }
            $dateText = $date.ToString('yyyy-MM-dd')
            $stateKey = 'component|{0}|{1}|{2}' -f [int]$row.id, [string]$entry[0], $dateText
            $stateRow = if ($stateByKey.ContainsKey($stateKey)) { $stateByKey[$stateKey] } else { $null }
            $eligible = $null -eq $stateRow
            if ($null -ne $stateRow) {
                $snooze = [DateTime]::MinValue
                $eligible = (-not (Convert-DbBool $stateRow.handled)) -and
                    ([string]::IsNullOrWhiteSpace([string]$stateRow.snooze_until) -or
                     ([DateTime]::TryParse([string]$stateRow.snooze_until, [ref]$snooze) -and $snooze -le $currentTime))
            }
            if ($eligible) {
                [void]$due.Add([pscustomobject]@{ Kind = [string]$entry[0]; Label = [string]$entry[1]; Date = $date; DateText = $dateText })
            }
        }
        if ($due.Count -eq 0) { continue }
        $name = ([string]$row.deal_number).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { $name = ([string]$row.description).Trim() }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = 'Задача #' + [string]$row.id }
        [void]$items.Add([pscustomobject]@{
            Source = 'component'; SourceId = [int]$row.id; Title = $name
            Description = ([string]$row.description).Trim()
            NextAction = ([string]$row.next_action).Trim()
            Detail = ([string]$row.next_action).Trim(); Due = @($due)
            Status = [string]$row.status; Stage = [string]$row.stage
        })
    }

    $dealReminders = Invoke-PurchaseQuery @'
SELECT id, IFNULL(deal_number, '') AS deal_number,
       IFNULL(comment, '') AS comment,
       IFNULL(reminder_date, '') AS reminder_date
FROM deals
WHERE IFNULL(archived, 0) = 0
  AND IFNULL(status, '') NOT IN ('Выполнено', 'Не актуально')
  AND IFNULL(reminder_date, '') <> ''
ORDER BY reminder_date, id
'@ @{}
    foreach ($row in $dealReminders.Rows) {
        $date = Convert-PurchaseDateText ([string]$row.reminder_date)
        if ($null -eq $date -or $date.Date -gt $today) { continue }
        $dateText = $date.ToString('yyyy-MM-dd')
        $stateKey = 'deal|{0}|reminder|{1}' -f [int]$row.id, $dateText
        $stateRow = if ($stateByKey.ContainsKey($stateKey)) { $stateByKey[$stateKey] } else { $null }
        $eligible = $null -eq $stateRow
        if ($null -ne $stateRow) {
            $snooze = [DateTime]::MinValue
            $eligible = (-not (Convert-DbBool $stateRow.handled)) -and
                ([string]::IsNullOrWhiteSpace([string]$stateRow.snooze_until) -or
                 ([DateTime]::TryParse([string]$stateRow.snooze_until, [ref]$snooze) -and $snooze -le $currentTime))
        }
        if ($eligible) {
            [void]$items.Add([pscustomobject]@{
                Source = 'deal'; SourceId = [int]$row.id; Title = [string]$row.deal_number
                Detail = ([string]$row.comment).Trim(); Due = @([pscustomobject]@{ Kind = 'reminder'; Label = 'Напоминание'; Date = $date; DateText = $dateText })
                DealId = [int]$row.id; SupplierId = 0
            })
        }
    }

    $manual = Invoke-PurchaseQuery @'
SELECT id, IFNULL(title, '') AS title, IFNULL(due_date, '') AS due_date,
       IFNULL(deal_id, 0) AS deal_id, IFNULL(supplier_id, 0) AS supplier_id
FROM reminders
WHERE status <> 'Done' AND IFNULL(due_date, '') <> ''
ORDER BY due_date, id
'@ @{}
    foreach ($row in $manual.Rows) {
        $date = Convert-PurchaseDateText ([string]$row.due_date)
        if ($null -eq $date -or $date.Date -gt $today) { continue }
        $dateText = $date.ToString('yyyy-MM-dd')
        $stateKey = 'reminder|{0}|manual|{1}' -f [int]$row.id, $dateText
        $stateRow = if ($stateByKey.ContainsKey($stateKey)) { $stateByKey[$stateKey] } else { $null }
        $eligible = $null -eq $stateRow
        if ($null -ne $stateRow) {
            $snooze = [DateTime]::MinValue
            $eligible = (-not (Convert-DbBool $stateRow.handled)) -and
                ([string]::IsNullOrWhiteSpace([string]$stateRow.snooze_until) -or
                 ([DateTime]::TryParse([string]$stateRow.snooze_until, [ref]$snooze) -and $snooze -le $currentTime))
        }
        if ($eligible) {
            [void]$items.Add([pscustomobject]@{
                Source = 'reminder'; SourceId = [int]$row.id; Title = [string]$row.title
                Detail = ''; Due = @([pscustomobject]@{ Kind = 'manual'; Label = 'Напоминание'; Date = $date; DateText = $dateText })
                DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id
            })
        }
    }
    return @($items)
}

function Mark-NotificationHandled {
    param($Notification)
    foreach ($due in @($Notification.Due)) {
        Set-NotificationState $Notification.Source ([int]$Notification.SourceId) ([string]$due.Kind) ([string]$due.DateText) $true '' $true
    }
}

function Snooze-Notification {
    param($Notification, [DateTime]$Until)
    $untilText = $Until.ToString('yyyy-MM-dd HH:mm:ss')
    foreach ($due in @($Notification.Due)) {
        Set-NotificationState $Notification.Source ([int]$Notification.SourceId) ([string]$due.Kind) ([string]$due.DateText) $false $untilText $true
    }
}

function Convert-PurchaseDateForSort {
    param([string]$Value)
    $date = Convert-PurchaseDateText $Value
    if ($null -eq $date) { return '' }
    return $date.ToString('yyyy-MM-dd')
}

function Test-PurchaseSupplierErpNotRequired {
    param([string]$Supplier)
    $name = (([string]$Supplier).Trim().ToLowerInvariant()) -replace '\s+', ''
    if ([string]::IsNullOrWhiteSpace($name)) { return $false }
    return (
        $name -match '^(компэл|compel)(_\d+)?$' -or
        $name -match '^(чид|chid)(_\d+)?$' -or
        $name -match '^(промэлектроника|promelec|promelektronika)(_\d+)?$'
    )
}

function Get-PurchaseActionItems {
    param([int]$DealId = 0)

    $where = 'WHERE IFNULL(d.archived, 0) = 0'
    $params = @{}
    if ($DealId -gt 0) {
        $where += ' AND d.id = @deal_id'
        $params['@deal_id'] = $DealId
    }

    $suppliers = Invoke-PurchaseQuery @"
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
    ds.invoice_confirmed_date,
    ds.components_receipt_date,
    IFNULL(ds.actual_receipt_date, '') AS actual_receipt_date,
    ds.delivery_weeks
FROM deal_suppliers ds
JOIN deals d ON d.id = ds.deal_id
$where
ORDER BY d.updated_at DESC, d.deal_number, ds.supplier
"@ $params

    $items = New-Object System.Collections.ArrayList
    $today = (Get-Date).Date
    foreach ($row in $suppliers.Rows) {
        $dealNumber = [string]$row.deal_number
        $supplier = [string]$row.supplier
        $prefix = "$dealNumber / $supplier"
        if (-not (Convert-DbBool $row.invoice_received)) {
            [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Warn'; Title = "${prefix}: нет PI"; DueDate = ''; Source = 'auto' })
            continue
        }
        if (-not (Convert-DbBool $row.payment_submitted)) {
            [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Info'; Title = "${prefix}: PI есть, но не подан в оплату"; DueDate = ''; Source = 'auto' })
        } elseif (-not (Convert-DbBool $row.paid)) {
            [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Attention'; Title = "${prefix}: подан в оплату, но не оплачен"; DueDate = ''; Source = 'auto' })
        }
        if ([string]::IsNullOrWhiteSpace([string]$row.invoice_confirmed_date)) {
            [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Warn'; Title = "${prefix}: нет даты подтверждения инвойса"; DueDate = ''; Source = 'auto' })
        }
        if (-not (Test-PurchaseSupplierErpNotRequired ([string]$row.supplier))) {
            if (-not (Convert-DbBool $row.erp_supplier_sent)) {
                [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Info'; Title = "${prefix}: Не отправлен ERP поставщику"; DueDate = ''; Source = 'auto' })
            }
            if (-not (Convert-DbBool $row.erp_roger_sent)) {
                [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Info'; Title = "${prefix}: Не отправлен ERP заказ"; DueDate = ''; Source = 'auto' })
            }
        }

        $actualReceiptDate = Convert-PurchaseDateText ([string]$row.actual_receipt_date)
        $receiptDate = Convert-PurchaseDateText ([string]$row.components_receipt_date)
        if ($null -eq $actualReceiptDate -and $null -ne $receiptDate) {
            $days = [int]($receiptDate.Date - $today).TotalDays
            if ($days -lt 0) {
                [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Danger'; Title = "${prefix}: поступление просрочено"; DueDate = $receiptDate.ToString('dd.MM.yyyy'); Source = 'auto' })
            } elseif ($days -ge 0 -and $days -le 7) {
                [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Attention'; Title = "${prefix}: скоро поступление"; DueDate = $receiptDate.ToString('dd.MM.yyyy'); Source = 'auto' })
            }
        }
    }

    $dealReminders = Invoke-PurchaseQuery @'
SELECT id, deal_number, reminder_date
FROM deals
WHERE IFNULL(archived, 0) = 0
  AND IFNULL(status, '') NOT IN ('Выполнено', 'Не актуально')
  AND IFNULL(reminder_date, '') <> ''
ORDER BY reminder_date, id DESC
'@ @{}
    foreach ($row in $dealReminders.Rows) {
        $due = Format-PurchaseDate $row.reminder_date
        $severity = 'Info'
        $date = Convert-PurchaseDateText $due
        if ($null -eq $date) {
            $severity = 'Warn'
        } elseif ($date.Date -lt $today) {
            $severity = 'Danger'
        } elseif ([int]($date.Date - $today).TotalDays -le 7) {
            $severity = 'Attention'
        }
        [void]$items.Add([pscustomobject]@{ DealId = [int]$row.id; SupplierId = 0; ReminderId = 0; Severity = $severity; Title = [string]$row.deal_number; DueDate = $due; Source = 'deal' })
    }

    $manual = Invoke-PurchaseQuery @'
SELECT id, deal_id, supplier_id, title, due_date, status, source
FROM reminders
WHERE status <> 'Done'
ORDER BY
    CASE WHEN IFNULL(due_date, '') = '' THEN 1 ELSE 0 END,
    due_date,
    id DESC
'@ @{}
    foreach ($row in $manual.Rows) {
        $due = Format-PurchaseDate $row.due_date
        $severity = 'Info'
        $date = Convert-PurchaseDateText $due
        if ($null -eq $date) {
            $severity = 'Warn'
        } elseif ($date.Date -lt $today) {
            $severity = 'Danger'
        } elseif ([int]($date.Date - $today).TotalDays -le 7) {
            $severity = 'Attention'
        }
        [void]$items.Add([pscustomobject]@{ DealId = if ($null -eq $row.deal_id -or $row.deal_id -is [DBNull]) { 0 } else { [int]$row.deal_id }; SupplierId = if ($null -eq $row.supplier_id -or $row.supplier_id -is [DBNull]) { 0 } else { [int]$row.supplier_id }; ReminderId = [int]$row.id; Severity = $severity; Title = [string]$row.title; DueDate = $due; Source = [string]$row.source })
    }
    return @($items)
}

# Entity management and document workflow.

function Convert-DbBool {
    param($Value)
    if ($null -eq $Value -or $Value -is [DBNull]) { return $false }
    return ([int]$Value -ne 0)
}

function Convert-ToDbBool {
    param($Value)
    if ([bool]$Value) { return 1 }
    return 0
}

function Convert-DealYesNoToDbValue {
    param($Value)

    if ($null -eq $Value -or $Value -is [DBNull]) { return $null }
    if ($Value -is [bool]) { return (Convert-ToDbBool $Value) }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    if ($text -eq 'Да' -or $text -eq 'Yes' -or $text -eq 'True' -or $text -eq '1') { return 1 }
    return 0
}

function Convert-MasksToDbValue {
    param($Value)

    if ($null -eq $Value -or $Value -is [DBNull]) { return 2 }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return 2 }
    if ($text -eq 'Да' -or $text -eq 'Yes' -or $text -eq 'True' -or $text -eq '1') { return 1 }
    return 0
}

function Convert-PurchaseDocumentType {
    param([string]$DocumentType)

    $text = if ([string]::IsNullOrWhiteSpace($DocumentType)) { 'Other' } else { $DocumentType.Trim() }
    if ($text -eq 'ERP supplier' -or $text -eq 'ERP Roger') { return 'ERP' }
    return $text
}

function Test-DealLevelDocumentType {
    param([string]$DocumentType)

    $text = if ([string]::IsNullOrWhiteSpace($DocumentType)) { '' } else { $DocumentType.Trim() }
    return ($text -eq 'Общий PI' -or $text -eq 'RFQ' -or $text -eq 'RRFQ' -or $text -eq 'PO' -or $text -eq 'Other')
}

function Get-UniquePurchaseDocumentPath {
    param([string]$Directory, [string]$FileName)

    $safeName = if ([string]::IsNullOrWhiteSpace($FileName)) { 'document' } else { $FileName.Trim() }
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) {
        $safeName = $safeName.Replace([string]$char, '_')
    }
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = 'document' }
    $candidate = Join-Path $Directory $safeName
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }

    $stem = [IO.Path]::GetFileNameWithoutExtension($safeName)
    $extension = [IO.Path]::GetExtension($safeName)
    for ($index = 2; $index -lt 1000; $index++) {
        $candidate = Join-Path $Directory ('{0} ({1}){2}' -f $stem, $index, $extension)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }

    throw "Не удалось подобрать свободное имя файла: $FileName"
}

function Get-SafePathPart {
    param([string]$Value)

    $text = if ([string]::IsNullOrWhiteSpace($Value)) { 'empty' } else { $Value.Trim() }
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) {
        $text = $text.Replace([string]$char, '_')
    }
    return ($text -replace '\s+', '_')
}

function Get-ComponentDealFolderPath {
    param([int]$Id, [string]$DealNumber = '')

    $root = Get-DefaultComponentFilesRoot
    $name = if ([string]::IsNullOrWhiteSpace($DealNumber)) {
        'component_{0:D5}' -f $Id
    } else {
        '{0:D5}_{1}' -f $Id, (Get-SafePathPart $DealNumber)
    }
    return (Join-Path $root $name)
}

function Ensure-ComponentDealFolder {
    param([int]$Id, [string]$DealNumber = '')

    if ($Id -le 0) { return '' }
    $root = Get-DefaultComponentFilesRoot
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $target = Get-ComponentDealFolderPath $Id $DealNumber
    $currentStored = [string](Invoke-PurchaseScalar 'SELECT IFNULL(folder_path, '''') FROM component_deals WHERE id = @id' @{ '@id' = $Id })
    $current = Resolve-ComponentStoredPath $currentStored
    $chosen = $target
    $samePath = $false
    if (-not [string]::IsNullOrWhiteSpace($current)) {
        try {
            $samePath = ([IO.Path]::GetFullPath($current).TrimEnd('\') -eq [IO.Path]::GetFullPath($target).TrimEnd('\'))
        } catch {
            $samePath = ($current -eq $target)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($current) -and -not $samePath) {
        try {
            if ((Test-Path -LiteralPath $current -PathType Container) -and -not (Test-Path -LiteralPath $target)) {
                Move-Item -LiteralPath $current -Destination $target -Force
                $chosen = $target
            } elseif (Test-Path -LiteralPath $current -PathType Container) {
                $chosen = $current
            }
        } catch {
            $chosen = $current
        }
    }

    if ([string]::IsNullOrWhiteSpace($chosen)) { $chosen = $target }
    New-Item -ItemType Directory -Force -Path $chosen | Out-Null
    Invoke-PurchaseNonQuery 'UPDATE component_deals SET folder_path = @folder_path WHERE id = @id' @{
        '@folder_path' = Convert-ToPortablePath $chosen
        '@id' = $Id
    }
    return $chosen
}

function Remove-ComponentDealFolder {
    param([int]$Id, [string]$FolderPath = '')

    if ($Id -le 0) { return }
    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        $FolderPath = [string](Invoke-PurchaseScalar 'SELECT IFNULL(folder_path, '''') FROM component_deals WHERE id = @id' @{ '@id' = $Id })
    }
    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return }

    $root = [IO.Path]::GetFullPath((Get-DefaultComponentFilesRoot))
    $resolvedFolderPath = Resolve-ComponentStoredPath $FolderPath
    $target = [IO.Path]::GetFullPath($resolvedFolderPath)
    if (-not $target.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { return }
    if (Test-Path -LiteralPath $target -PathType Container) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}

function New-PurchaseDeal {
    param(
        [string]$DealNumber,
        [string]$Title = '',
        [string]$Comment = '',
        [string]$Client = '',
        [string]$Status = ''
    )

    if ([string]::IsNullOrWhiteSpace($DealNumber)) {
        throw 'Введите номер сделки.'
    }

    $dealNumberParam = $DealNumber.Trim()
    $now = Get-NowText
    $existingStatus = [string](Invoke-PurchaseScalar 'SELECT IFNULL(status, '''') FROM deals WHERE deal_number = @deal_number' @{ '@deal_number' = $dealNumberParam })
    $orderedAt = if (-not (Test-PurchaseOrderedStatus $existingStatus) -and (Test-PurchaseOrderedStatus $Status)) { $now } else { $null }
    $workflowTemplateId = Get-DefaultWorkflowTemplateId
    Invoke-PurchaseNonQuery @'
INSERT INTO deals(deal_number, client, status, title, comment, workflow_template_id, masks, created_at, ordered_at, updated_at)
VALUES(@deal_number, @client, CASE WHEN @status <> '' THEN @status ELSE 'RFQ' END, @title, @comment, @workflow_template_id, 2, @created_at, @ordered_at, @updated_at)
ON CONFLICT(deal_number) DO UPDATE SET
    client = CASE WHEN excluded.client <> '' THEN excluded.client ELSE client END,
    status = CASE WHEN @status <> '' THEN @status ELSE status END,
    ordered_at = CASE
        WHEN IFNULL(ordered_at, '') <> '' THEN ordered_at
        WHEN @ordered_at IS NOT NULL THEN @ordered_at
        ELSE ordered_at
    END,
    title = CASE WHEN excluded.title <> '' THEN excluded.title ELSE title END,
    comment = CASE WHEN excluded.comment <> '' THEN excluded.comment ELSE comment END,
    archived = 0,
    updated_at = excluded.updated_at;
'@ @{
        '@deal_number' = $dealNumberParam
        '@client' = $Client
        '@status' = $Status
        '@title' = $Title
        '@comment' = $Comment
        '@workflow_template_id' = $workflowTemplateId
        '@created_at' = $now
        '@ordered_at' = $orderedAt
        '@updated_at' = $now
    }
    $id = [int](Invoke-PurchaseScalar 'SELECT id FROM deals WHERE deal_number = @deal_number' @{ '@deal_number' = $dealNumberParam })
    Write-ActivityLog 'deal' $id 'Создана/обновлена сделка' $dealNumberParam $id
    return $id
}

function Update-PurchaseDeal {
    param(
        [int]$DealId,
        [string]$Client,
        [string]$Status,
        [string]$Comment,
        [AllowNull()][string]$BoardCount = $null,
        [AllowNull()][string]$DealNumber = $null,
        [AllowNull()][int]$WorkflowTemplateId = $null,
        [AllowNull()][string]$Period = $null,
        [AllowNull()]$Masks = $null,
        [AllowNull()]$PiSent = $null,
        [AllowNull()][string]$Priority = $null,
        [AllowNull()][string]$TrackingStatus = $null,
        [AllowNull()][string]$Executor = $null,
        [AllowNull()][string]$ReminderDate = $null,
        [AllowNull()][string]$AssemblyLocation = $null
    )

    if ($DealId -le 0) { throw 'Выберите сделку.' }
    $previousStatus = [string](Invoke-PurchaseScalar 'SELECT IFNULL(status, '''') FROM deals WHERE id = @id' @{ '@id' = $DealId })
    $orderedAt = if (-not (Test-PurchaseOrderedStatus $previousStatus) -and (Test-PurchaseOrderedStatus $Status)) { Get-NowText } else { $null }
    $dealNumberParam = $null
    if (-not [string]::IsNullOrWhiteSpace($DealNumber)) {
        $dealNumberParam = $DealNumber.Trim()
    }
    Invoke-PurchaseNonQuery @'
UPDATE deals SET
    deal_number = CASE WHEN @deal_number IS NULL THEN deal_number ELSE @deal_number END,
    board_count = CASE WHEN @board_count IS NULL THEN board_count ELSE @board_count END,
    workflow_template_id = CASE WHEN @workflow_template_id IS NULL THEN workflow_template_id ELSE @workflow_template_id END,
    period = CASE WHEN @period IS NULL THEN period ELSE @period END,
    priority = CASE WHEN @priority IS NULL THEN priority ELSE @priority END,
    masks = CASE WHEN @masks IS NULL THEN masks ELSE @masks END,
    pi_sent = CASE WHEN @pi_sent IS NULL THEN pi_sent ELSE @pi_sent END,
    client = @client,
    status = @status,
    ordered_at = CASE
        WHEN IFNULL(ordered_at, '') <> '' THEN ordered_at
        WHEN @ordered_at IS NOT NULL THEN @ordered_at
        ELSE ordered_at
    END,
    tracking_status = CASE WHEN @tracking_status IS NULL THEN tracking_status ELSE @tracking_status END,
    -- NULL means "do not change" for callers that omit the optional field;
    -- an empty string is an intentional request to clear the executor.
    executor = CASE WHEN @executor IS NULL THEN executor ELSE @executor END,
    reminder_date = CASE WHEN @reminder_date IS NULL THEN reminder_date ELSE @reminder_date END,
    assembly_location = CASE WHEN @assembly_location IS NULL THEN assembly_location ELSE @assembly_location END,
    comment = @comment,
    updated_at = @updated_at
WHERE id = @id
'@ @{
        '@deal_number' = $dealNumberParam
        '@board_count' = $BoardCount
        '@workflow_template_id' = if ($null -eq $WorkflowTemplateId -or $WorkflowTemplateId -le 0) { $null } else { $WorkflowTemplateId }
        '@period' = $Period
        '@priority' = if ($null -eq $Priority) { $null } else { $Priority.Trim() }
        '@masks' = Convert-MasksToDbValue $Masks
        '@pi_sent' = Convert-DealYesNoToDbValue $PiSent
        '@client' = $Client
        '@status' = $Status
        '@ordered_at' = $orderedAt
        '@tracking_status' = if ([string]::IsNullOrWhiteSpace($TrackingStatus)) { $null } else { $TrackingStatus.Trim() }
        '@executor' = if ($null -eq $Executor) { $null } else { ([string]$Executor).Trim() }
        '@reminder_date' = if ($null -eq $ReminderDate) { $null } else { ([string]$ReminderDate).Trim() }
        '@assembly_location' = if ($null -eq $AssemblyLocation) { $null } else { ([string]$AssemblyLocation).Trim() }
        '@comment' = $Comment
        '@updated_at' = Get-NowText
        '@id' = $DealId
    }
    Write-ActivityLog 'deal' $DealId 'Изменена сделка' "Этап: $Status; Статус: $TrackingStatus; Клиент: $Client; Период: $Period; Исполнитель: $Executor; Место сборки: $AssemblyLocation" $DealId
}

function Touch-PurchaseDeal {
    param([int]$DealId)

    if ($DealId -le 0) { return }
    Invoke-PurchaseNonQuery 'UPDATE deals SET updated_at = @updated_at WHERE id = @id' @{
        '@updated_at' = Get-NowText
        '@id' = $DealId
    }
}

function Clear-PurchaseDealReminder {
    param([int]$DealId)

    if ($DealId -le 0) { return }
    Invoke-PurchaseNonQuery 'UPDATE deals SET reminder_date = @reminder_date, updated_at = @updated_at WHERE id = @id' @{
        '@reminder_date' = ''
        '@updated_at' = Get-NowText
        '@id' = $DealId
    }
}

function Touch-PurchaseDealBySupplier {
    param([int]$SupplierId)

    if ($SupplierId -le 0) { return }
    Invoke-PurchaseNonQuery @'
UPDATE deals SET updated_at = @updated_at
WHERE id = (SELECT deal_id FROM deal_suppliers WHERE id = @supplier_id)
'@ @{
        '@updated_at' = Get-NowText
        '@supplier_id' = $SupplierId
    }
}

function Add-PurchaseSupplier {
    param([int]$DealId, [string]$Supplier)

    if ($DealId -le 0) { throw 'Выберите сделку.' }
    if ([string]::IsNullOrWhiteSpace($Supplier)) { throw 'Введите поставщика.' }

    $now = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO deal_suppliers(deal_id, supplier, created_at, updated_at)
VALUES(@deal_id, @supplier, @created_at, @updated_at)
ON CONFLICT(deal_id, supplier) DO UPDATE SET updated_at = excluded.updated_at;
'@ @{
        '@deal_id' = $DealId
        '@supplier' = $Supplier.Trim()
        '@created_at' = $now
        '@updated_at' = $now
    }
    Touch-PurchaseDeal $DealId
    Write-ActivityLog 'supplier' 0 'Добавлен/обновлен поставщик' $Supplier.Trim() $DealId
}

function Delete-PurchaseSupplier {
    param([int]$SupplierId)

    if ($SupplierId -le 0) { throw 'Выберите поставщика.' }
    $table = Invoke-PurchaseQuery 'SELECT * FROM documents WHERE supplier_id = @supplier_id' @{ '@supplier_id' = $SupplierId }
    $supplierRows = Invoke-PurchaseQuery 'SELECT * FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId }
    $supplier = [string](Invoke-PurchaseScalar 'SELECT supplier FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId })
    $dealId = [int](Invoke-PurchaseScalar 'SELECT deal_id FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId })
    if ([string]::IsNullOrWhiteSpace($supplier) -or $dealId -le 0) { throw 'Поставщик не найден.' }

    $filePaths = @($table.Rows | ForEach-Object { Resolve-PurchaseStoredPath ([string]$_.stored_path) })
    [void](New-PurchaseTrashItem 'supplier' $SupplierId ('Поставщик: ' + $supplier) @{
        deal_suppliers = Convert-PurchaseRowsToPayload $supplierRows
        documents = Convert-PurchaseRowsToPayload $table
    } $filePaths)

    foreach ($row in $table.Rows) {
        $storedPath = [string]$row.stored_path
        # Files have already been moved to the lightweight basket.
    }
    Invoke-PurchaseNonQuery 'DELETE FROM documents WHERE supplier_id = @supplier_id' @{ '@supplier_id' = $SupplierId }
    Invoke-PurchaseNonQuery 'DELETE FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId }
    Touch-PurchaseDeal $dealId
    Write-ActivityLog 'supplier' $SupplierId 'Поставщик удален' $supplier $dealId
}

function Get-PurchaseDeals {
    param([string]$Search = '', [string]$Filter = 'All')

    $where = 'WHERE (@search = '''' OR d.deal_number LIKE @needle OR IFNULL(d.title, '''') LIKE @needle OR IFNULL(d.client, '''') LIKE @needle OR IFNULL(d.status, '''') LIKE @needle OR IFNULL(d.tracking_status, '''') LIKE @needle OR IFNULL(d.period, '''') LIKE @needle OR IFNULL(d.executor, '''') LIKE @needle OR IFNULL(d.assembly_location, '''') LIKE @needle OR IFNULL(d.board_count, '''') LIKE @needle)'
    if ($Filter -eq 'Archived') {
        $where += ' AND IFNULL(d.archived, 0) = 1'
    } elseif ($Filter -eq 'AllRecords') {
    } elseif ($Filter -eq 'China') {
        $where += " AND trim(IFNULL(d.assembly_location, '')) = 'Китай'"
    } else {
        $where += ' AND IFNULL(d.archived, 0) = 0'
        if ($Filter -eq 'Active') {
$where += " AND trim(IFNULL(d.executor, '')) = '' AND (IFNULL(d.status, '') IN ('RFQ', 'PO', 'Закупка', 'В работе') OR (IFNULL(d.status, '') = 'Заказано' AND (IFNULL(stats.supplier_count, 0) = 0 OR IFNULL(stats.invoice_confirmed_count, 0) < IFNULL(stats.supplier_count, 0))))"
        }
    }

    # AllRecords is used when navigating from reminders and must really include
    # every deal, including deals assembled in China. The regular/default views
    # keep their existing non-China restriction.
    if ($Filter -ne 'China' -and $Filter -ne 'Tracking' -and $Filter -ne 'AllRecords') {
        $where += " AND trim(IFNULL(d.assembly_location, '')) <> 'Китай'"
    }

    if ($Filter -eq 'Tracking') {
$where += " AND trim(IFNULL(d.executor, '')) <> ''"
    } elseif ($Filter -eq 'China') {
        $where += " AND trim(IFNULL(d.assembly_location, '')) = 'Китай'"
    } elseif ($Filter -eq 'Purchase') {
        $where += ' AND IFNULL(d.status, '''') = ''Закупка'''
    } elseif ($Filter -eq 'Pause') {
        $where += ' AND IFNULL(d.status, '''') IN (''RRFQ'', ''PI'')'
    } elseif ($Filter -eq 'NoInvoices') {
        $where += ' AND IFNULL(stats.invoice_count, 0) = 0'
    } elseif ($Filter -eq 'WaitingReceipt') {
        $where += ' AND IFNULL(stats.supplier_count, 0) > 0 AND IFNULL(stats.receipt_count, 0) < IFNULL(stats.supplier_count, 0)'
    } elseif ($Filter -eq 'Done') {
        $where += " AND IFNULL(d.status, '') = 'Заказано' AND IFNULL(stats.supplier_count, 0) > 0 AND IFNULL(stats.invoice_confirmed_count, 0) = IFNULL(stats.supplier_count, 0)"
    } elseif ($Filter -eq 'Open') {
        $where += ' AND (IFNULL(stats.supplier_count, 0) = 0 OR IFNULL(stats.done_count, 0) < IFNULL(stats.supplier_count, 0))'
    }

    $sql = @"
SELECT
    d.id,
    d.deal_number,
    IFNULL(d.board_count, '') AS board_count,
    IFNULL(d.client, '') AS client,
    IFNULL(d.period, '') AS period,
    IFNULL(d.executor, '') AS executor,
    IFNULL(d.reminder_date, '') AS reminder_date,
    IFNULL(d.assembly_location, '') AS assembly_location,
    IFNULL(d.priority, '3') AS priority,
    IFNULL(d.tracking_status, 'Ожидание') AS tracking_status,
    IFNULL(d.status, 'RFQ') AS status,
    IFNULL(d.pi_sent, 0) AS pi_sent,
    IFNULL(d.masks, 0) AS masks,
    IFNULL(d.archived, 0) AS archived,
    d.workflow_template_id,
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
$where
ORDER BY d.updated_at DESC, d.id DESC
"@

    return Invoke-PurchaseQuery $sql @{
        '@search' = $Search
        '@needle' = "%$Search%"
    }
}

function Get-PurchaseSuppliers {
    param([int]$DealId)

    if ($DealId -le 0) {
        return New-Object System.Data.DataTable
    }

    return Invoke-PurchaseQuery @'
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
ORDER BY supplier
'@ @{ '@deal_id' = $DealId }
}

function Update-PurchaseSupplier {
    param(
        [int]$SupplierId,
        [bool]$InvoiceReceived,
        [bool]$InvoiceConfirmed,
        [bool]$SupplierOrderCreated,
        [bool]$ErpSupplierSent,
        [bool]$ErpRogerSent,
        $PiAmountUsd,
        $PiAmountCny,
        $PiAmountRub,
        $PaidAmount,
        $DeliveryWeeks,
        [bool]$PaymentSubmitted,
        [bool]$Paid,
        [string]$InvoiceConfirmedDate,
        [string]$ComponentsReceiptDate,
        [string]$ActualReceiptDate,
        [string]$Comment
    )

    Invoke-PurchaseNonQuery @'
UPDATE deal_suppliers SET
    invoice_received = @invoice_received,
    invoice_confirmed = @invoice_confirmed,
    supplier_order_created = @supplier_order_created,
    erp_supplier_sent = @erp_supplier_sent,
    erp_roger_sent = @erp_roger_sent,
    pi_amount_usd = @pi_amount_usd,
    pi_amount_cny = @pi_amount_cny,
    pi_amount_rub = @pi_amount_rub,
    paid_amount = @paid_amount,
    delivery_weeks = @delivery_weeks,
    payment_submitted = @payment_submitted,
    paid = @paid,
    invoice_confirmed_date = @invoice_confirmed_date,
    components_receipt_date = @components_receipt_date,
    actual_receipt_date = @actual_receipt_date,
    comment = @comment,
    updated_at = @updated_at
WHERE id = @id
'@ @{
        '@invoice_received' = Convert-ToDbBool $InvoiceReceived
        '@invoice_confirmed' = Convert-ToDbBool $InvoiceConfirmed
        '@supplier_order_created' = Convert-ToDbBool $SupplierOrderCreated
        '@erp_supplier_sent' = Convert-ToDbBool $ErpSupplierSent
        '@erp_roger_sent' = Convert-ToDbBool $ErpRogerSent
        '@pi_amount_usd' = Convert-ToNullableAmount $PiAmountUsd
        '@pi_amount_cny' = Convert-ToNullableAmount $PiAmountCny
        '@pi_amount_rub' = Convert-ToNullableAmount $PiAmountRub
        '@paid_amount' = $PaidAmount
        '@delivery_weeks' = Convert-ToNullableInt $DeliveryWeeks
        '@payment_submitted' = Convert-ToDbBool $PaymentSubmitted
        '@paid' = Convert-ToDbBool $Paid
        '@invoice_confirmed_date' = $InvoiceConfirmedDate
        '@components_receipt_date' = $ComponentsReceiptDate
        '@actual_receipt_date' = $ActualReceiptDate
        '@comment' = $Comment
        '@updated_at' = Get-NowText
        '@id' = $SupplierId
    }
    Touch-PurchaseDealBySupplier $SupplierId
    $dealId = Invoke-PurchaseScalar 'SELECT deal_id FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId }
    Write-ActivityLog 'supplier' $SupplierId 'Изменен поставщик' "Оплата: $Paid; PI: $InvoiceReceived; срок: $DeliveryWeeks" $(if ($null -eq $dealId -or $dealId -is [DBNull]) { 0 } else { [int]$dealId }) $SupplierId
}

function Set-PurchaseSupplierActualReceiptDate {
    param([int]$SupplierId, [string]$ActualReceiptDate)

    if ($SupplierId -le 0) { throw 'Choose supplier.' }
    $actualText = ''
    if (-not [string]::IsNullOrWhiteSpace($ActualReceiptDate)) {
        $date = Convert-PurchaseDateText $ActualReceiptDate
        if ($null -eq $date) { throw 'Фактическая дата поступления должна быть в формате дд.мм.гггг.' }
        $actualText = $date.ToString('dd.MM.yyyy')
    }

    Invoke-PurchaseNonQuery 'UPDATE deal_suppliers SET actual_receipt_date = @actual_receipt_date, updated_at = @updated_at WHERE id = @id' @{
        '@actual_receipt_date' = $actualText
        '@updated_at' = Get-NowText
        '@id' = $SupplierId
    }
    Touch-PurchaseDealBySupplier $SupplierId
    $dealId = Invoke-PurchaseScalar 'SELECT deal_id FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId }
    Write-ActivityLog 'supplier' $SupplierId 'Поступление подтверждено на склад' ('Факт: {0}' -f $actualText) $(if ($null -eq $dealId -or $dealId -is [DBNull]) { 0 } else { [int]$dealId }) $SupplierId
}
function Update-PurchaseSupplierPiAmounts {
    param(
        [int]$SupplierId,
        $PiAmountUsd,
        $PiAmountCny,
        [bool]$MarkInvoiceReceived = $true
    )

    if ($SupplierId -le 0) { return }
    Invoke-PurchaseNonQuery @'
UPDATE deal_suppliers SET
    invoice_received = CASE WHEN @mark_invoice_received = 1 THEN 1 ELSE invoice_received END,
    pi_amount_usd = CASE WHEN @pi_amount_usd IS NULL THEN pi_amount_usd ELSE @pi_amount_usd END,
    pi_amount_cny = CASE WHEN @pi_amount_cny IS NULL THEN pi_amount_cny ELSE @pi_amount_cny END,
    updated_at = @updated_at
WHERE id = @id
'@ @{
        '@mark_invoice_received' = Convert-ToDbBool $MarkInvoiceReceived
        '@pi_amount_usd' = Convert-ToNullableAmount $PiAmountUsd
        '@pi_amount_cny' = Convert-ToNullableAmount $PiAmountCny
        '@updated_at' = Get-NowText
        '@id' = $SupplierId
    }
    Touch-PurchaseDealBySupplier $SupplierId
    $dealId = Invoke-PurchaseScalar 'SELECT deal_id FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId }
    Write-ActivityLog 'supplier' $SupplierId 'Обновлены суммы PI' "USD=$(Format-PurchaseAmount $PiAmountUsd); CNY=$(Format-PurchaseAmount $PiAmountCny)" $(if ($null -eq $dealId -or $dealId -is [DBNull]) { 0 } else { [int]$dealId }) $SupplierId
}

function Get-PurchaseDocuments {
    param([int]$DealId)

    if ($DealId -le 0) {
        return New-Object System.Data.DataTable
    }

    return Invoke-PurchaseQuery @'
SELECT
    doc.id,
    doc.document_type,
    IFNULL(ds.supplier, '') AS supplier,
    doc.original_name,
    doc.stored_path,
    IFNULL(doc.file_hash, '') AS file_hash,
    doc.uploaded_at
FROM documents doc
LEFT JOIN deal_suppliers ds ON ds.id = doc.supplier_id
WHERE doc.deal_id = @deal_id
ORDER BY doc.uploaded_at DESC, doc.id DESC
'@ @{ '@deal_id' = $DealId }
}

function Get-PurchaseDocumentFileHash {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    try {
        return [string](Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    } catch {
        return ''
    }
}

function Update-PurchaseDocumentFileReference {
    param([int]$DocumentId, [string]$Path, [string]$FileHash = '')

    if ($DocumentId -le 0 -or [string]::IsNullOrWhiteSpace($Path)) { return }
    $resolvedPath = Resolve-PurchaseStoredPath $Path
    if ([string]::IsNullOrWhiteSpace($resolvedPath)) { $resolvedPath = $Path }
    $name = [IO.Path]::GetFileName($resolvedPath)
    if ([string]::IsNullOrWhiteSpace($FileHash)) {
        $FileHash = Get-PurchaseDocumentFileHash $resolvedPath
    }

    Invoke-PurchaseNonQuery 'UPDATE documents SET stored_path = @stored_path, original_name = @original_name, file_hash = CASE WHEN @file_hash <> '''' THEN @file_hash ELSE file_hash END WHERE id = @id' @{
        '@stored_path' = Convert-ToPortablePath $resolvedPath
        '@original_name' = $name
        '@file_hash' = $FileHash
        '@id' = $DocumentId
    }
}

function Resolve-PurchaseDocumentFile {
    param(
        [int]$DocumentId,
        [string]$StoredPath,
        [string]$OriginalName = '',
        [string]$FileHash = ''
    )

    $storedPath = [string]$StoredPath
    $path = Resolve-PurchaseStoredPath $storedPath
    $name = if ([string]::IsNullOrWhiteSpace($OriginalName)) { [IO.Path]::GetFileName($path) } else { $OriginalName }
    $hash = [string]$FileHash

    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType Leaf)) {
        $currentName = [IO.Path]::GetFileName($path)
        if ([string]::IsNullOrWhiteSpace($hash)) { $hash = Get-PurchaseDocumentFileHash $path }
        $portablePath = Convert-ToPortablePath $path
        if ($DocumentId -gt 0 -and ($currentName -ne $name -or [string]::IsNullOrWhiteSpace($FileHash) -or $portablePath -ne $storedPath)) {
            Update-PurchaseDocumentFileReference $DocumentId $path $hash
        }
        return [pscustomobject]@{ Id = $DocumentId; Path = $path; Name = $currentName; FileHash = $hash; Found = $true; Changed = ($currentName -ne $name -or $portablePath -ne $storedPath) }
    }

    if ([string]::IsNullOrWhiteSpace($path)) {
        return [pscustomobject]@{ Id = $DocumentId; Path = $path; Name = $name; FileHash = $hash; Found = $false; Changed = $false }
    }

    try {
        $directory = [IO.Path]::GetDirectoryName($path)
    } catch {
        $directory = ''
    }
    if ([string]::IsNullOrWhiteSpace($directory) -or -not (Test-Path -LiteralPath $directory -PathType Container)) {
        return [pscustomobject]@{ Id = $DocumentId; Path = $path; Name = $name; FileHash = $hash; Found = $false; Changed = $false }
    }

    $candidates = @(Get-ChildItem -LiteralPath $directory -File -ErrorAction SilentlyContinue)
    $chosen = $null
    if (-not [string]::IsNullOrWhiteSpace($hash)) {
        $hashMatches = New-Object System.Collections.ArrayList
        foreach ($candidate in $candidates) {
            if ((Get-PurchaseDocumentFileHash $candidate.FullName) -eq $hash) {
                [void]$hashMatches.Add($candidate)
            }
        }
        if ($hashMatches.Count -eq 1) { $chosen = $hashMatches[0] }
    }

    if ($null -eq $chosen) {
        $knownPaths = @{}
        $otherDocs = Invoke-PurchaseQuery 'SELECT stored_path FROM documents WHERE id <> @id' @{ '@id' = $DocumentId }
        foreach ($row in $otherDocs.Rows) {
            $otherPath = Resolve-PurchaseStoredPath ([string]$row.stored_path)
            if ([string]::IsNullOrWhiteSpace($otherPath)) { continue }
            try { $knownPaths[[IO.Path]::GetFullPath($otherPath).ToLowerInvariant()] = $true } catch { }
        }
        $available = @($candidates | Where-Object {
            try { -not $knownPaths.ContainsKey([IO.Path]::GetFullPath($_.FullName).ToLowerInvariant()) } catch { $false }
        })
        $extension = [IO.Path]::GetExtension($path)
        if (-not [string]::IsNullOrWhiteSpace($extension)) {
            $sameExtension = @($available | Where-Object { [IO.Path]::GetExtension($_.FullName).Equals($extension, [StringComparison]::OrdinalIgnoreCase) })
            if ($sameExtension.Count -eq 1) { $chosen = $sameExtension[0] }
        }
        if ($null -eq $chosen -and $available.Count -eq 1) { $chosen = $available[0] }
    }

    if ($null -ne $chosen) {
        $newPath = [string]$chosen.FullName
        $newHash = Get-PurchaseDocumentFileHash $newPath
        Update-PurchaseDocumentFileReference $DocumentId $newPath $newHash
        return [pscustomobject]@{ Id = $DocumentId; Path = $newPath; Name = $chosen.Name; FileHash = $newHash; Found = $true; Changed = $true }
    }

    return [pscustomobject]@{ Id = $DocumentId; Path = $path; Name = $name; FileHash = $hash; Found = $false; Changed = $false }
}
function Get-PurchaseSupplierPiDocuments {
    param([int]$SupplierId)

    if ($SupplierId -le 0) {
        return New-Object System.Data.DataTable
    }

    return Invoke-PurchaseQuery @'
SELECT
    doc.id,
    doc.original_name,
    doc.stored_path,
    IFNULL(doc.file_hash, '') AS file_hash,
    doc.uploaded_at
FROM documents doc
WHERE doc.supplier_id = @supplier_id
  AND doc.document_type = 'PI/Invoice'
ORDER BY doc.uploaded_at DESC, doc.id DESC
'@ @{ '@supplier_id' = $SupplierId }
}

function Get-PurchaseSupplierName {
    param([int]$SupplierId)

    if ($SupplierId -le 0) { return '' }
    return [string](Invoke-PurchaseScalar 'SELECT supplier FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId })
}

function Import-PiAmountsFromPdfForSupplier {
    param([int]$SupplierId, [string]$Supplier, [string]$Path)

    if ($SupplierId -le 0) { throw 'Выберите поставщика.' }
    if ([string]::IsNullOrWhiteSpace($Supplier)) {
        $Supplier = Get-PurchaseSupplierName $SupplierId
    }

    if (Test-SkipPiPdfAmountExtraction $Supplier) {
        Update-PurchaseSupplierPiAmounts $SupplierId $null $null $true
        return [pscustomobject]@{
            Updated = $false
            Skipped = $true
            Message = "Для $Supplier авто-чтение суммы PI отключено."
        }
    }

    if ([IO.Path]::GetExtension($Path).ToLowerInvariant() -ne '.pdf') {
        Update-PurchaseSupplierPiAmounts $SupplierId $null $null $true
        return [pscustomobject]@{
            Updated = $false
            Skipped = $true
            Message = 'Файл не PDF, сумму нужно заполнить вручную.'
        }
    }

    $amounts = Get-PiAmountsFromPdf $Path
    Update-PurchaseSupplierPiAmounts $SupplierId $amounts.Usd $amounts.Cny $true
    $updated = ($null -ne $amounts.Usd -or $null -ne $amounts.Cny)
    $message = if ($updated) {
        "Суммы найдены: `$=$(Format-PurchaseAmount $amounts.Usd), ?=$(Format-PurchaseAmount $amounts.Cny)"
    } else {
        $amounts.Warning
    }

    return [pscustomobject]@{
        Updated = $updated
        Skipped = $false
        Message = $message
    }
}

function Add-PurchaseDocument {
    param(
        [int]$DealId,
        [int]$SupplierId,
        [string]$DocumentType,
        [string]$SourcePath,
        [string]$FolderName = ''
    )

    if ($DealId -le 0) { throw 'Выберите сделку.' }
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { throw "Файл не найден: $SourcePath" }
    if ([string]::IsNullOrWhiteSpace($DocumentType)) { $DocumentType = 'Other' }
    $DocumentType = Convert-PurchaseDocumentType $DocumentType
    if (Test-DealLevelDocumentType $DocumentType) { $SupplierId = 0 }

    $dealNumber = [string](Invoke-PurchaseScalar 'SELECT deal_number FROM deals WHERE id = @id' @{ '@id' = $DealId })
    $supplier = ''
    if ($SupplierId -gt 0) {
        $supplier = [string](Invoke-PurchaseScalar 'SELECT supplier FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId })
    }

    $root = Get-PurchaseDocumentsRoot

    $folderPart = if (-not [string]::IsNullOrWhiteSpace($FolderName)) {
        $FolderName
    } elseif (Test-DealLevelDocumentType $DocumentType) {
        $DocumentType
    } else {
        $supplier
    }
    $targetDir = Join-Path (Join-Path $root (Get-SafePathPart $dealNumber)) (Get-SafePathPart $folderPart)
    New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
    $originalName = [IO.Path]::GetFileName($SourcePath)
    $targetPath = Get-UniquePurchaseDocumentPath $targetDir $originalName
    Copy-Item -LiteralPath $SourcePath -Destination $targetPath

    Invoke-PurchaseNonQuery @'
INSERT INTO documents(deal_id, supplier_id, document_type, original_name, stored_path, file_hash, uploaded_at)
VALUES(@deal_id, @supplier_id, @document_type, @original_name, @stored_path, @file_hash, @uploaded_at)
'@ @{
        '@deal_id' = $DealId
        '@supplier_id' = if ($SupplierId -gt 0) { $SupplierId } else { $null }
        '@document_type' = $DocumentType
        '@original_name' = $originalName
        '@stored_path' = Convert-ToPortablePath $targetPath
        '@file_hash' = Get-PurchaseDocumentFileHash $targetPath
        '@uploaded_at' = Get-NowText
    }
    Touch-PurchaseDeal $DealId
    Write-ActivityLog 'document' 0 'Загружен документ' "${DocumentType}: $originalName" $DealId $SupplierId
    return $targetPath
}

function Delete-PurchaseDocument {
    param([int]$DocumentId)

    if ($DocumentId -le 0) {
        throw 'Выберите документ.'
    }

    $table = Invoke-PurchaseQuery 'SELECT * FROM documents WHERE id = @id' @{ '@id' = $DocumentId }
    if ($table.Rows.Count -eq 0) {
        throw 'Документ не найден.'
    }

    $dealId = [int]$table.Rows[0].deal_id
    $resolvedDoc = Resolve-PurchaseDocumentFile $DocumentId ([string]$table.Rows[0].stored_path) ([string]$table.Rows[0].original_name) ([string]$table.Rows[0].file_hash)
    $path = [string]$resolvedDoc.Path
    [void](New-PurchaseTrashItem 'document' $DocumentId ('Документ: ' + [string]$table.Rows[0].original_name) @{
        documents = Convert-PurchaseRowsToPayload $table
    } @($path))

    Invoke-PurchaseNonQuery 'DELETE FROM documents WHERE id = @id' @{ '@id' = $DocumentId }
    Touch-PurchaseDeal $dealId
    Write-ActivityLog 'document' $DocumentId 'Удален документ' $path $dealId
}

function Delete-PurchaseDeal {
    param([int]$DealId)

    if ($DealId -le 0) {
        throw 'Выберите сделку.'
    }

    $dealNumber = [string](Invoke-PurchaseScalar 'SELECT deal_number FROM deals WHERE id = @id' @{ '@id' = $DealId })
    if ([string]::IsNullOrWhiteSpace($dealNumber)) { throw 'Сделка не найдена.' }
    $safeDealPart = Get-SafePathPart $dealNumber
    $dealDirs = New-Object System.Collections.ArrayList
    $dealRows = Invoke-PurchaseQuery 'SELECT * FROM deals WHERE id = @id' @{ '@id' = $DealId }
    $supplierRows = Invoke-PurchaseQuery 'SELECT * FROM deal_suppliers WHERE deal_id = @deal_id' @{ '@deal_id' = $DealId }
    $docs = Invoke-PurchaseQuery 'SELECT * FROM documents WHERE deal_id = @deal_id' @{ '@deal_id' = $DealId }
    $bitrixRows = Invoke-PurchaseQuery 'SELECT * FROM bitrix_task_links WHERE deal_id = @deal_id' @{ '@deal_id' = $DealId }
    $filePaths = @($docs.Rows | ForEach-Object { Resolve-PurchaseStoredPath ([string]$_.stored_path) })
    [void](New-PurchaseTrashItem 'deal' $DealId ('Сделка: ' + $dealNumber) @{
        deals = Convert-PurchaseRowsToPayload $dealRows
        deal_suppliers = Convert-PurchaseRowsToPayload $supplierRows
        documents = Convert-PurchaseRowsToPayload $docs
        bitrix_task_links = Convert-PurchaseRowsToPayload $bitrixRows
    } $filePaths)
    foreach ($row in $docs.Rows) {
        $path = Resolve-PurchaseStoredPath ([string]$row.stored_path)
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType Leaf)) {
            try {
                $parent = [IO.Directory]::GetParent([IO.Path]::GetFullPath($path))
                if ($null -ne $parent -and $null -ne $parent.Parent -and $parent.Parent.Name -eq $safeDealPart) {
                    [void]$dealDirs.Add($parent.Parent.FullName)
                }
            } catch {
            }
            # The file has been moved to the basket above.
        }
    }

    if (Get-Command -Name Save-BitrixTaskLink -ErrorAction SilentlyContinue) {
        $bitrixTaskLinks = Invoke-PurchaseQuery 'SELECT bitrix_task_id FROM bitrix_task_links WHERE deal_id = @deal_id' @{ '@deal_id' = $DealId }
        foreach ($row in $bitrixTaskLinks.Rows) {
            Save-BitrixTaskLink -BitrixTaskId ([int]$row.bitrix_task_id) -DealId $null -IsManualDeleted:$true
        }
    }

    Invoke-PurchaseNonQuery 'DELETE FROM deals WHERE id = @id' @{ '@id' = $DealId }
    Write-ActivityLog 'deal' $DealId 'Удалена сделка' $dealNumber $DealId

    $currentRoot = Get-PurchaseDocumentsRoot
    if (-not [string]::IsNullOrWhiteSpace($currentRoot)) {
        [void]$dealDirs.Add((Join-Path $currentRoot $safeDealPart))
    }
    [void]$dealDirs.Add((Join-Path (Get-DefaultPurchaseDocumentsRoot) $safeDealPart))

    foreach ($dir in @($dealDirs | Sort-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        try {
            $fullDir = [IO.Path]::GetFullPath($dir)
            if ((Split-Path -Leaf $fullDir) -ne $safeDealPart) { continue }
            if (Test-Path -LiteralPath $fullDir -PathType Container) {
                Remove-Item -LiteralPath $fullDir -Recurse -Force
            }
        } catch {
        }
    }
}

function New-ComponentDeal {
    $now = Get-NowText
    $entryDate = (Get-Date).ToString('dd.MM.yyyy')
    Invoke-PurchaseNonQuery @'
INSERT INTO component_deals(entry_date, status, stage, priority, created_at, updated_at)
VALUES(@entry_date, 'В работе', 'Запросил поставщиков', '3', @created_at, @updated_at)
'@ @{
        '@entry_date' = $entryDate
        '@created_at' = $now
        '@updated_at' = $now
    }
    $id = [int](Invoke-PurchaseScalar 'SELECT id FROM component_deals ORDER BY id DESC LIMIT 1')
    $folder = Ensure-ComponentDealFolder $id ''
    Write-ActivityLog 'component' $id 'Создана задача' $folder 0 0
    return $id
}

function Get-ComponentDeals {
    param([string]$Search = '')

    $needle = if ([string]::IsNullOrWhiteSpace($Search)) { '%' } else { '%' + $Search.Trim() + '%' }
    return Invoke-PurchaseQuery @'
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
    IFNULL(folder_path, '') AS folder_path,
    IFNULL(notes, '') AS notes,
    updated_at
FROM component_deals
WHERE IFNULL(deal_number, '') LIKE @needle
   OR IFNULL(description, '') LIKE @needle
ORDER BY id DESC
'@ @{ '@needle' = $needle }
}

function Get-ActiveCockpitDealCount {
    # The cockpit combines operational purchase queues with the Tasks page.
    # Use the deal number as the shared key where it is available so a deal
    # shown in more than one queue is counted only once.
    $keys = @{}
    $purchaseTable = Get-PurchaseDeals '' 'AllRecords'
    foreach ($row in $purchaseTable.Rows) {
        if (Convert-DbBool $row.archived) { continue }

        $status = ([string]$row.status).Trim()
        $assemblyLocation = ([string]$row.assembly_location).Trim()
        $executor = ([string]$row.executor).Trim()
        $isOperationalStage = $status -in @('PO', 'RFQ', 'Закупка')
        $isActiveFilterDeal =
            $assemblyLocation -ne 'Китай' -and
            [string]::IsNullOrWhiteSpace($executor) -and
            ($status -in @('RFQ', 'PO', 'Закупка', 'В работе') -or
             ($status -eq 'Заказано' -and ([int]$row.supplier_count -eq 0 -or [int]$row.invoice_confirmed_count -lt [int]$row.supplier_count)))
        $isTrackingFilterDeal = (-not [string]::IsNullOrWhiteSpace($executor) -and $isOperationalStage)
        $isChinaFilterDeal = ($assemblyLocation -eq 'Китай' -and $isOperationalStage)

        if (-not ($isActiveFilterDeal -or $isTrackingFilterDeal -or $isChinaFilterDeal)) {
            continue
        }

        $dealNumber = ([string]$row.deal_number).Trim()
        $key = if ([string]::IsNullOrWhiteSpace($dealNumber)) { 'purchase:' + [string]$row.id } else { 'deal:' + (Normalize-Key $dealNumber) }
        $keys[$key] = $true
    }

    $taskRows = Invoke-PurchaseQuery @'
SELECT id, IFNULL(deal_number, '') AS deal_number
FROM component_deals
WHERE IFNULL(status, '') IN ('В работе', 'Подано в оплату')
'@ @{}
    foreach ($row in $taskRows.Rows) {
        $dealNumber = ([string]$row.deal_number).Trim()
        $key = if ([string]::IsNullOrWhiteSpace($dealNumber)) { 'task:' + [string]$row.id } else { 'deal:' + (Normalize-Key $dealNumber) }
        $keys[$key] = $true
    }

    return [int]$keys.Count
}


function Update-ComponentDeal {
    param(
        [int]$Id,
        [string]$DealNumber,
        [string]$Status,
        [string]$Stage,
        [string]$Description,
        [string]$NextAction,
        [string]$ReminderDate,
        [string]$DeadlineDate,
        [string]$Priority,
        [string]$Period
    )

    if ($Id -le 0) { return }
    $previousStatus = [string](Invoke-PurchaseScalar 'SELECT IFNULL(status, '''') FROM component_deals WHERE id = @id' @{ '@id' = $Id })
    $orderedAt = if (-not (Test-PurchaseOrderedStatus $previousStatus) -and (Test-PurchaseOrderedStatus $Status)) { Get-NowText } else { $null }
    $dealNumberParam = if ([string]::IsNullOrWhiteSpace($DealNumber)) { '' } else { $DealNumber.Trim() }
    Invoke-PurchaseNonQuery @'
UPDATE component_deals SET
    deal_number = @deal_number,
    status = @status,
    ordered_at = CASE
        WHEN IFNULL(ordered_at, '') <> '' THEN ordered_at
        WHEN @ordered_at IS NOT NULL THEN @ordered_at
        ELSE ordered_at
    END,
    stage = @stage,
    description = @description,
    next_action = @next_action,
    reminder_date = @reminder_date,
    deadline_date = @deadline_date,
    priority = @priority,
    period = @period,
    updated_at = @updated_at
WHERE id = @id
'@ @{
        '@deal_number' = $dealNumberParam
        '@status' = $Status
        '@ordered_at' = $orderedAt
        '@stage' = $Stage
        '@description' = $Description
        '@next_action' = $NextAction
        '@reminder_date' = $ReminderDate
        '@deadline_date' = $DeadlineDate
        '@priority' = $Priority
        '@period' = $Period
        '@updated_at' = Get-NowText
        '@id' = $Id
    }
    [void](Ensure-ComponentDealFolder $Id $dealNumberParam)
    Write-ActivityLog 'component' $Id 'Изменена задача' $dealNumberParam 0 0
}

function Get-ComponentDealNotes {
    param([int]$Id)

    if ($Id -le 0) { return '' }
    $value = Invoke-PurchaseScalar 'SELECT IFNULL(notes, '''') FROM component_deals WHERE id = @id' @{ '@id' = $Id }
    if ($null -eq $value -or $value -is [DBNull]) { return '' }
    return [string]$value
}

function Set-ComponentDealNotes {
    param([int]$Id, [string]$Notes)

    if ($Id -le 0) { return }
    $current = Get-ComponentDealNotes $Id
    if ($current -eq $Notes) { return }

    Invoke-PurchaseNonQuery 'UPDATE component_deals SET notes = @notes, updated_at = @updated_at WHERE id = @id' @{
        '@notes' = $Notes
        '@updated_at' = Get-NowText
        '@id' = $Id
    }
    Write-ActivityLog 'component' $Id 'Обновлены записи по задаче' '' 0 0
}
function Delete-ComponentDeal {
    param([int]$Id)
    if ($Id -le 0) { return }
    $componentRows = Invoke-PurchaseQuery 'SELECT * FROM component_deals WHERE id = @id' @{ '@id' = $Id }
    if ($componentRows.Rows.Count -eq 0) { throw 'Задача не найдена.' }
    $componentName = [string]$componentRows.Rows[0].deal_number
    $folderPath = [string](Invoke-PurchaseScalar 'SELECT IFNULL(folder_path, '''') FROM component_deals WHERE id = @id' @{ '@id' = $Id })
    $resolvedFolder = Resolve-ComponentStoredPath $folderPath
    [void](New-PurchaseTrashItem 'component' $Id ('Задача: ' + $componentName) @{
        component_deals = Convert-PurchaseRowsToPayload $componentRows
    } @() $resolvedFolder)
    Invoke-PurchaseNonQuery 'DELETE FROM component_deals WHERE id = @id' @{ '@id' = $Id }
    Remove-ComponentDealFolder $Id $folderPath
    Write-ActivityLog 'component' $Id 'Удалена задача' '' 0 0
}

# Reminder deletion.

function Remove-Reminder {
    param([int]$ReminderId)

    if ($ReminderId -le 0) { throw 'Автоматическое напоминание нельзя удалить вручную.' }
    $reminder = Invoke-PurchaseQuery 'SELECT * FROM reminders WHERE id = @id' @{ '@id' = $ReminderId }
    if ($reminder.Rows.Count -eq 0) { throw 'Напоминание не найдено.' }
    $row = $reminder.Rows[0]
    [void](New-PurchaseTrashItem 'reminder' $ReminderId ('Напоминание: ' + [string]$row.title) @{
        reminders = Convert-PurchaseRowsToPayload $reminder
    })
    Invoke-PurchaseNonQuery 'DELETE FROM reminders WHERE id = @id' @{ '@id' = $ReminderId }
    $dealId = if ($null -eq $row.deal_id -or $row.deal_id -is [DBNull]) { 0 } else { [int]$row.deal_id }
    $supplierId = if ($null -eq $row.supplier_id -or $row.supplier_id -is [DBNull]) { 0 } else { [int]$row.supplier_id }
    Write-ActivityLog 'reminder' $ReminderId 'Удалено напоминание' ([string]$row.title) $dealId $supplierId
}
