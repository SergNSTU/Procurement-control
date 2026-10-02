param(
    [int]$ParentPid = 0,
    [int]$Port = 8765,
    [switch]$SelfTest
)

$ErrorActionPreference = 'Stop'
$script:AppRoot = Split-Path -Parent $PSScriptRoot
$script:DbPath = Join-Path (Join-Path (Join-Path $script:AppRoot 'data') 'purchase_control') 'purchase_control.sqlite'
$script:Port = $Port
$script:Http = $null
$script:Listener = $null
$script:Form = $null
$script:LastHotkeyState = @{}
$script:LogPath = Join-Path (Join-Path $script:AppRoot 'data\logs') 'quick_capture_agent.log'
$script:InstanceMutex = $null
$script:ClipboardSnapshot = $null
$script:CaptureDialogOpen = $false

try { [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false) } catch { }

function Write-QuickCaptureLog {
    param([string]$Message)
    try {
        $logRoot = Split-Path -Parent $script:LogPath
        New-Item -ItemType Directory -Force -Path $logRoot | Out-Null
        Add-Content -LiteralPath $script:LogPath -Value ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Message) -Encoding UTF8
    } catch { }
}

trap {
    Write-QuickCaptureLog ('FATAL: ' + $_.Exception.Message)
    break
}

if (-not ('QuickCapture.Native' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace QuickCapture {
    public static class Native {
        [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vKey);
        public static bool IsDown(int vKey) { return (GetAsyncKeyState(vKey) & 0x8000) != 0; }
        [DllImport("user32.dll", SetLastError=true)] private static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, UIntPtr extraInfo);
        [DllImport("user32.dll")] public static extern uint GetClipboardSequenceNumber();
        public static void SendCtrlC() {
            keybd_event(0x11, 0, 0, UIntPtr.Zero); keybd_event(0x43, 0, 0, UIntPtr.Zero);
            keybd_event(0x43, 0, 2, UIntPtr.Zero); keybd_event(0x11, 0, 2, UIntPtr.Zero);
        }
    }
}
'@
}

if (-not ('QuickCapture.HotkeyWindow' -as [type])) {
    Add-Type -ReferencedAssemblies @('System.Windows.Forms.dll') -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
namespace QuickCapture {
    public sealed class HotkeyWindow : NativeWindow, IDisposable {
        private const int WmHotkey = 0x0312;
        private const uint ModControl = 0x0002;
        private const uint ModShift = 0x0004;
        private const uint ModNoRepeat = 0x4000;
        [DllImport("user32.dll", SetLastError=true)] private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint modifiers, uint key);
        [DllImport("user32.dll", SetLastError=true)] private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        public event EventHandler TaskPressed;
        public event EventHandler ReminderPressed;
        public string Mode { get; private set; }
        public HotkeyWindow() { CreateHandle(new CreateParams()); }
        public void Register() {
            if (RegisterPair(ModControl | ModShift | ModNoRepeat)) { Mode = "Ctrl+Shift"; return; }
            if (RegisterPair(ModControl | 0x0001 | ModNoRepeat)) { Mode = "Ctrl+Alt"; return; }
            throw new InvalidOperationException("Не удалось зарегистрировать сочетания быстрого захвата.");
        }
        private bool RegisterPair(uint modifiers) {
            if (!RegisterHotKey(Handle, 1, modifiers, (uint)'T')) return false;
            if (!RegisterHotKey(Handle, 2, modifiers, (uint)'R')) { UnregisterHotKey(Handle, 1); return false; }
            return true;
        }
        protected override void WndProc(ref Message m) {
            if (m.Msg == WmHotkey) {
                if (m.WParam.ToInt32() == 1 && TaskPressed != null) TaskPressed(this, EventArgs.Empty);
                if (m.WParam.ToInt32() == 2 && ReminderPressed != null) ReminderPressed(this, EventArgs.Empty);
            }
            base.WndProc(ref m);
        }
        public void Dispose() { if (Handle != IntPtr.Zero) { UnregisterHotKey(Handle, 1); UnregisterHotKey(Handle, 2); DestroyHandle(); } }
    }
}
'@
}

function Load-QuickCaptureSqlite {
    if ('System.Data.SQLite.SQLiteConnection' -as [type]) { return }
    $sqliteDir = Join-Path $script:AppRoot 'app\lib\sqlite'
    $dll = Join-Path $sqliteDir 'System.Data.SQLite.dll'
    if (-not (Test-Path -LiteralPath $dll -PathType Leaf)) { throw "SQLite-компонент не найден: $dll" }
    $archDir = Join-Path $sqliteDir ($(if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }))
    $env:PATH = "$archDir;$sqliteDir;$env:PATH"
    Add-Type -Path $dll
}

function Open-QuickCaptureConnection {
    Load-QuickCaptureSqlite
    $root = Split-Path -Parent $script:DbPath
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $connection = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$script:DbPath;Version=3;Foreign Keys=True;Pooling=True;Journal Mode=WAL;Synchronous=Normal;")
    $connection.Open()
    return $connection
}

function Add-QuickCaptureParameters {
    param($Command, [hashtable]$Parameters)
    foreach ($key in @($Parameters.Keys)) {
        $parameter = $Command.CreateParameter()
        $parameter.ParameterName = $key
        $parameter.Value = if ($null -eq $Parameters[$key]) { [DBNull]::Value } else { $Parameters[$key] }
        [void]$Command.Parameters.Add($parameter)
    }
}

function Invoke-QuickCaptureNonQuery {
    param([string]$Sql, [hashtable]$Parameters = @{})
    $connection = Open-QuickCaptureConnection
    try {
        $command = $connection.CreateCommand(); $command.CommandText = $Sql
        Add-QuickCaptureParameters $command $Parameters
        [void]$command.ExecuteNonQuery()
    } finally { $connection.Dispose() }
}

function Invoke-QuickCaptureScalar {
    param([string]$Sql, [hashtable]$Parameters = @{})
    $connection = Open-QuickCaptureConnection
    try {
        $command = $connection.CreateCommand(); $command.CommandText = $Sql
        Add-QuickCaptureParameters $command $Parameters
        $value = $command.ExecuteScalar()
        if ($value -is [DBNull]) { return $null }
        return $value
    } finally { $connection.Dispose() }
}

function Invoke-QuickCaptureQuery {
    param([string]$Sql, [hashtable]$Parameters = @{})
    $connection = Open-QuickCaptureConnection
    try {
        $command = $connection.CreateCommand(); $command.CommandText = $Sql
        Add-QuickCaptureParameters $command $Parameters
        $reader = $command.ExecuteReader()
        $table = New-Object System.Data.DataTable
        $table.Load($reader)
        return ,$table
    } finally { $connection.Dispose() }
}

function Ensure-QuickCaptureSchema {
    Invoke-QuickCaptureNonQuery @'
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS component_deals (
 id INTEGER PRIMARY KEY AUTOINCREMENT, entry_date TEXT NOT NULL, deal_number TEXT,
 status TEXT NOT NULL DEFAULT 'В работе', stage TEXT NOT NULL DEFAULT 'Запросил поставщиков',
 description TEXT, next_action TEXT, reminder_date TEXT, deadline_date TEXT,
 priority TEXT NOT NULL DEFAULT '3', period TEXT, folder_path TEXT, notes TEXT,
 created_at TEXT NOT NULL, updated_at TEXT NOT NULL, source TEXT, source_text TEXT
);
CREATE TABLE IF NOT EXISTS reminders (
 id INTEGER PRIMARY KEY AUTOINCREMENT, deal_id INTEGER, supplier_id INTEGER, component_id INTEGER,
 title TEXT NOT NULL, due_date TEXT, status TEXT NOT NULL DEFAULT 'Open', source TEXT,
 created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS activity_log (
 id INTEGER PRIMARY KEY AUTOINCREMENT, entity_type TEXT NOT NULL, entity_id INTEGER,
 deal_id INTEGER, supplier_id INTEGER, action TEXT NOT NULL, details TEXT, created_at TEXT NOT NULL
);
'@
    foreach ($definition in @(
        @('component_deals', 'source', 'TEXT'),
        @('component_deals', 'source_text', 'TEXT')
    )) {
        $exists = Invoke-QuickCaptureScalar "SELECT COUNT(*) FROM pragma_table_info('$($definition[0])') WHERE name = @name" @{ '@name' = $definition[1] }
        if ([int]$exists -eq 0) { Invoke-QuickCaptureNonQuery "ALTER TABLE $($definition[0]) ADD COLUMN $($definition[1]) $($definition[2])" }
    }
}

function Invoke-QuickCaptureLegacyMigration {
    $legacyPath = Join-Path (Join-Path $script:AppRoot 'data') 'purchase_control.sqlite'
    if (-not (Test-Path -LiteralPath $legacyPath -PathType Leaf)) { return }
    if ((Get-QuickCaptureSetting 'quick_capture.legacy_migrated' '') -eq '1') { return }
    $legacy = $null
    try {
        $legacy = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$legacyPath;Version=3;Read Only=True;")
        $legacy.Open()
        $command = $legacy.CreateCommand()
        $command.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name IN ('reminders','component_deals')"
        $reader = $command.ExecuteReader(); $tables = @(); while ($reader.Read()) { $tables += [string]$reader.GetValue(0) }; $reader.Close()
        if ($tables -contains 'reminders') {
            $command.CommandText = "SELECT title, due_date, status, source, created_at, updated_at FROM reminders WHERE IFNULL(source, '') = 'quick_capture'"
            $rows = New-Object System.Data.DataTable; $rows.Load($command.ExecuteReader())
            foreach ($row in $rows.Rows) {
                $exists = Invoke-QuickCaptureScalar 'SELECT COUNT(*) FROM reminders WHERE source = @source AND title = @title AND created_at = @created_at' @{ '@source'='quick_capture'; '@title'=[string]$row.title; '@created_at'=[string]$row.created_at }
                if ([int]$exists -eq 0) {
                    Invoke-QuickCaptureNonQuery @'
INSERT INTO reminders(title, due_date, status, source, created_at, updated_at)
VALUES(@title, @due_date, @status, @source, @created_at, @updated_at)
'@ @{ '@title'=[string]$row.title; '@due_date'=[string]$row.due_date; '@status'=[string]$row.status; '@source'='quick_capture'; '@created_at'=[string]$row.created_at; '@updated_at'=[string]$row.updated_at }
                }
            }
        }
        if ($tables -contains 'component_deals') {
            $command.CommandText = "SELECT entry_date, deal_number, status, stage, description, next_action, reminder_date, deadline_date, priority, period, folder_path, notes, created_at, updated_at, source, source_text FROM component_deals WHERE IFNULL(source, '') = 'quick_capture'"
            $rows = New-Object System.Data.DataTable; $rows.Load($command.ExecuteReader())
            foreach ($row in $rows.Rows) {
                $exists = Invoke-QuickCaptureScalar 'SELECT COUNT(*) FROM component_deals WHERE source = @source AND source_text = @source_text AND created_at = @created_at' @{ '@source'='quick_capture'; '@source_text'=[string]$row.source_text; '@created_at'=[string]$row.created_at }
                if ([int]$exists -eq 0) {
                    Invoke-QuickCaptureNonQuery @'
INSERT INTO component_deals(entry_date, deal_number, status, stage, description, next_action,
 reminder_date, deadline_date, priority, period, folder_path, notes, created_at, updated_at, source, source_text)
VALUES(@entry_date, @deal_number, @status, @stage, @description, @next_action,
 @reminder_date, @deadline_date, @priority, @period, @folder_path, @notes, @created_at, @updated_at, @source, @source_text)
'@ @{
                        '@entry_date'=[string]$row.entry_date; '@deal_number'=[string]$row.deal_number; '@status'=[string]$row.status; '@stage'=[string]$row.stage
                        '@description'=[string]$row.description; '@next_action'=[string]$row.next_action; '@reminder_date'=[string]$row.reminder_date; '@deadline_date'=[string]$row.deadline_date
                        '@priority'=[string]$row.priority; '@period'=[string]$row.period; '@folder_path'=[string]$row.folder_path; '@notes'=[string]$row.notes
                        '@created_at'=[string]$row.created_at; '@updated_at'=[string]$row.updated_at; '@source'='quick_capture'; '@source_text'=[string]$row.source_text
                    }
                }
            }
        }
        Invoke-QuickCaptureNonQuery 'INSERT INTO settings(key, value) VALUES(@key, @value) ON CONFLICT(key) DO UPDATE SET value = excluded.value' @{ '@key'='quick_capture.legacy_migrated'; '@value'='1' }
        Write-QuickCaptureLog ("Migrated legacy quick-capture database: {0}" -f $legacyPath)
    } catch { Write-QuickCaptureLog ('Legacy quick-capture migration failed: ' + $_.Exception.Message) }
    finally { if ($null -ne $legacy) { $legacy.Dispose() } }
}

function Get-QuickCaptureSetting {
    param([string]$Key, [string]$Default = '')
    $value = Invoke-QuickCaptureScalar 'SELECT value FROM settings WHERE key = @key' @{ '@key' = $Key }
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return $Default }
    return [string]$value
}

function Normalize-QuickCaptureText {
    param([string]$Text)
    $value = ([string]$Text) -replace "`0", ''
    $value = $value.Trim()
    return $value
}

function New-QuickCaptureRecord {
    param([ValidateSet('task','reminder')][string]$Kind, [string]$Title, [string]$SourceText, [string]$DueDate, [string]$DealNumber = '')
    if (Get-QuickCaptureSetting 'quick_capture.enabled' '1' -eq '0') {
        Invoke-QuickCaptureNonQuery 'INSERT INTO settings(key, value) VALUES(@key, @value) ON CONFLICT(key) DO UPDATE SET value = excluded.value' @{ '@key' = 'quick_capture.enabled'; '@value' = '1' }
        Write-QuickCaptureLog 'quick_capture.enabled was 0; enabled automatically after user requested a capture.'
    }
    $title = Normalize-QuickCaptureText $Title
    $source = Normalize-QuickCaptureText $SourceText
    $dealNumber = Normalize-QuickCaptureText $DealNumber
    if ([string]::IsNullOrWhiteSpace($title)) { throw 'Введите заголовок.' }
    if ($title.Length -gt 240) { $title = $title.Substring(0, 237) + '...' }
    if ($Kind -eq 'reminder' -and [string]::IsNullOrWhiteSpace($DueDate)) { throw 'Для напоминания укажите дату и время.' }
    $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    if ($Kind -eq 'reminder') {
        Invoke-QuickCaptureNonQuery @'
INSERT INTO reminders(title, due_date, status, source, created_at, updated_at)
VALUES(@title, @due_date, 'Open', 'quick_capture', @created_at, @updated_at)
'@ @{ '@title'=$title; '@due_date'=$DueDate; '@created_at'=$now; '@updated_at'=$now }
        $id = [int](Invoke-QuickCaptureScalar 'SELECT id FROM reminders ORDER BY id DESC LIMIT 1')
        Invoke-QuickCaptureNonQuery @'
INSERT INTO activity_log(entity_type, entity_id, action, details, created_at)
VALUES('reminder', @id, 'Создано напоминание', @details, @created_at)
'@ @{ '@id'=$id; '@details'=$title; '@created_at'=$now }
        return [pscustomobject]@{ Id=$id; Kind=$Kind; Title=$title; DueDate=$DueDate }
    }
    $deadline = if ([string]::IsNullOrWhiteSpace($DueDate)) { '' } else { $DueDate }
    Invoke-QuickCaptureNonQuery @'
INSERT INTO component_deals(entry_date, deal_number, status, stage, description, next_action, deadline_date,
 priority, created_at, updated_at, source, source_text)
VALUES(@entry_date, @deal_number, 'В работе', 'Запросил поставщиков', @description, @next_action, @deadline_date,
 '3', @created_at, @updated_at, 'quick_capture', @source_text)
'@ @{
        '@entry_date'=(Get-Date).ToString('dd.MM.yyyy'); '@deal_number'=$dealNumber; '@description'=if ($source) { $source } else { $title }
        '@next_action'=$title; '@deadline_date'=$deadline; '@created_at'=$now; '@updated_at'=$now; '@source_text'=if ($source) { $source } else { $title }
    }
    $id = [int](Invoke-QuickCaptureScalar 'SELECT id FROM component_deals ORDER BY id DESC LIMIT 1')
    Invoke-QuickCaptureNonQuery @'
INSERT INTO activity_log(entity_type, entity_id, action, details, created_at)
VALUES('component', @id, 'Создана задача из выделенного текста', @details, @created_at)
'@ @{ '@id'=$id; '@details'=$title; '@created_at'=$now }
    return [pscustomobject]@{ Id=$id; Kind=$Kind; Title=$title; DueDate=$deadline; DealNumber=$dealNumber }
}

function ConvertTo-JsonResponse { param($Value) return ($Value | ConvertTo-Json -Depth 8 -Compress) }

function Send-QuickCaptureResponse {
    param($Stream, [int]$StatusCode, $Body)
    $bytes = [Text.Encoding]::UTF8.GetBytes((ConvertTo-JsonResponse $Body))
    $status = if ($StatusCode -eq 202) { '202 Accepted' } elseif ($StatusCode -eq 204) { '204 No Content' } elseif ($StatusCode -eq 404) { '404 Not Found' } else { '400 Bad Request' }
    $header = "HTTP/1.1 $status`r`nContent-Type: application/json; charset=utf-8`r`nAccess-Control-Allow-Origin: *`r`nAccess-Control-Allow-Methods: POST, OPTIONS`r`nAccess-Control-Allow-Headers: Content-Type`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`n`r`n"
    $headerBytes = [Text.Encoding]::ASCII.GetBytes($header)
    $Stream.Write($headerBytes, 0, $headerBytes.Length)
    if ($bytes.Length -gt 0) { $Stream.Write($bytes, 0, $bytes.Length) }
    $Stream.Flush()
}

function New-QuickCaptureRoundedPath {
    param([int]$Width, [int]$Height, [int]$Radius = 14)
    $path = New-Object Drawing.Drawing2D.GraphicsPath
    $diameter = $Radius * 2
    $path.AddArc(0, 0, $diameter, $diameter, 180, 90)
    $path.AddArc($Width - $diameter, 0, $diameter, $diameter, 270, 90)
    $path.AddArc($Width - $diameter, $Height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc(0, $Height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function Set-QuickCaptureButtonLook {
    param($Button, [bool]$Primary = $false, [Drawing.Color]$Accent = [Drawing.Color]::Empty)
    $Button.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = if ($Primary) { 0 } else { 1 }
    $Button.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(226, 232, 240)
    $Button.BackColor = if ($Primary) { $Accent } else { [Drawing.Color]::White }
    $Button.ForeColor = if ($Primary) { [Drawing.Color]::White } else { [Drawing.Color]::FromArgb(51, 65, 85) }
    $Button.Font = New-Object Drawing.Font('Segoe UI', 9, $(if ($Primary) { [Drawing.FontStyle]::Bold } else { [Drawing.FontStyle]::Regular }))
    $Button.Cursor = [Windows.Forms.Cursors]::Hand
    $Button.Region = New-QuickCaptureRoundedPath ([Math]::Max(1, $Button.Width)) ([Math]::Max(1, $Button.Height)) 8
    $Button.Add_Resize({
        try { $this.Region = New-QuickCaptureRoundedPath ([Math]::Max(1, $this.Width)) ([Math]::Max(1, $this.Height)) 8 } catch { }
    })
}

function Show-QuickCaptureDialog {
    param([string]$Kind, [string]$Text)
    $isReminder = $Kind -eq 'reminder'
    $accent = if ($isReminder) { [Drawing.Color]::FromArgb(217, 119, 6) } else { [Drawing.Color]::FromArgb(37, 99, 235) }
    $accentLight = if ($isReminder) { [Drawing.Color]::FromArgb(255, 247, 237) } else { [Drawing.Color]::FromArgb(239, 246, 255) }
    $dialog = New-Object Windows.Forms.Form
    $dialog.Text = if ($isReminder) { 'Создать напоминание' } else { 'Создать задачу' }
    $dialog.Width = 560; $dialog.Height = if ($isReminder) { 330 } else { 400 }; $dialog.StartPosition = 'CenterScreen'
    $dialog.FormBorderStyle = [Windows.Forms.FormBorderStyle]::None
    $dialog.MaximizeBox = $false; $dialog.MinimizeBox = $false; $dialog.ShowInTaskbar = $false
    $dialog.TopMost = $true; $dialog.BackColor = [Drawing.Color]::White
    $dialog.Font = New-Object Drawing.Font('Segoe UI', 9)
    $dialog.Padding = New-Object Windows.Forms.Padding(1)
    $dialog.Region = New-QuickCaptureRoundedPath $dialog.Width $dialog.Height 14
    $dialog.Add_Resize({ try { $this.Region = New-QuickCaptureRoundedPath $this.Width $this.Height 14 } catch { } })

    $accentBar = New-Object Windows.Forms.Panel
    $accentBar.Left = 0; $accentBar.Top = 0; $accentBar.Width = 6; $accentBar.Height = $dialog.Height
    $accentBar.BackColor = $accent; $dialog.Controls.Add($accentBar)

    $icon = New-Object Windows.Forms.Panel
    $icon.Left = 24; $icon.Top = 22; $icon.Width = 54; $icon.Height = 54; $icon.BackColor = $accentLight
    $icon.Add_Paint({
        param($sender, $eventArgs)
        $g = $eventArgs.Graphics; $g.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $pen = New-Object Drawing.Pen($accent, 2.2); $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
        if ($isReminder) {
            $g.DrawEllipse($pen, 16, 14, 22, 22); $g.DrawLine($pen, 27, 20, 27, 26); $g.DrawLine($pen, 27, 26, 33, 29)
        } else {
            $g.DrawRectangle($pen, 15, 15, 24, 24); $g.DrawLine($pen, 15, 22, 39, 22); $g.DrawLine($pen, 21, 11, 21, 18); $g.DrawLine($pen, 33, 11, 33, 18)
        }
        $pen.Dispose()
    })
    $dialog.Controls.Add($icon)

    $heading = New-Object Windows.Forms.Label
    $heading.Text = if ($isReminder) { 'Новое напоминание' } else { 'Новая задача' }
    $heading.Left = 98; $heading.Top = 22; $heading.Width = 380; $heading.Height = 25
    $heading.Font = New-Object Drawing.Font('Segoe UI Semibold', 11, [Drawing.FontStyle]::Bold); $heading.ForeColor = [Drawing.Color]::FromArgb(15, 23, 42)
    $dialog.Controls.Add($heading)
    $subheading = New-Object Windows.Forms.Label
    $subheading.Text = 'Быстрый захват из выделенного текста'; $subheading.Left = 98; $subheading.Top = 49; $subheading.Width = 390; $subheading.Height = 20
    $subheading.ForeColor = [Drawing.Color]::FromArgb(100, 116, 139); $dialog.Controls.Add($subheading)

    $close = New-Object Windows.Forms.Button
    $close.Text = '×'; $close.Left = 500; $close.Top = 17; $close.Width = 34; $close.Height = 30
    $close.FlatStyle = [Windows.Forms.FlatStyle]::Flat; $close.FlatAppearance.BorderSize = 0; $close.BackColor = [Drawing.Color]::White
    $close.ForeColor = [Drawing.Color]::FromArgb(100, 116, 139); $close.Font = New-Object Drawing.Font('Segoe UI', 14)
    $close.Cursor = [Windows.Forms.Cursors]::Hand; $dialog.Controls.Add($close)
    $close.Add_Click({ $dialog.Close() })

    $divider = New-Object Windows.Forms.Panel
    $divider.Left = 24; $divider.Top = 92; $divider.Width = 510; $divider.Height = 1; $divider.BackColor = [Drawing.Color]::FromArgb(226, 232, 240); $dialog.Controls.Add($divider)

    $titleCaption = New-Object Windows.Forms.Label
    $titleCaption.Text = 'Что нужно сделать'; $titleCaption.Left = 24; $titleCaption.Top = 111; $titleCaption.Width = 150; $titleCaption.Height = 20
    $titleCaption.ForeColor = [Drawing.Color]::FromArgb(100, 116, 139); $dialog.Controls.Add($titleCaption)
    $titleBox = New-Object Windows.Forms.TextBox
    $titleBox.Left = 24; $titleBox.Top = 133; $titleBox.Width = 510; $titleBox.Height = 30; $titleBox.Text = Normalize-QuickCaptureText $Text
    $titleBox.Font = New-Object Drawing.Font('Segoe UI', 10); $titleBox.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle; $titleBox.BackColor = [Drawing.Color]::FromArgb(248, 250, 252); $dialog.Controls.Add($titleBox)
    if ($titleBox.Text.Length -gt 240) { $titleBox.Text = $titleBox.Text.Substring(0, 237) + '...' }

    $dealBox = $null
    if (-not $isReminder) {
        $dealCaption = New-Object Windows.Forms.Label
        $dealCaption.Text = 'Сделка'; $dealCaption.Left = 24; $dealCaption.Top = 181; $dealCaption.Width = 150; $dealCaption.Height = 20
        $dealCaption.ForeColor = [Drawing.Color]::FromArgb(100, 116, 139); $dialog.Controls.Add($dealCaption)
        $dealBox = New-Object Windows.Forms.TextBox
        $dealBox.Left = 24; $dealBox.Top = 203; $dealBox.Width = 510; $dealBox.Height = 30
        $dealBox.Font = New-Object Drawing.Font('Segoe UI', 10); $dealBox.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle; $dealBox.BackColor = [Drawing.Color]::FromArgb(248, 250, 252); $dialog.Controls.Add($dealBox)
    }

    $dateTop = if ($isReminder) { 181 } else { 251 }
    $dateCaption = New-Object Windows.Forms.Label
    $dateCaption.Text = if ($isReminder) { 'Напомнить' } else { 'Срок задачи' }; $dateCaption.Left = 24; $dateCaption.Top = $dateTop; $dateCaption.Width = 150; $dateCaption.Height = 20
    $dateCaption.ForeColor = [Drawing.Color]::FromArgb(100, 116, 139); $dialog.Controls.Add($dateCaption)
    $datePicker = New-Object Windows.Forms.DateTimePicker
    $datePicker.Left = 24; $datePicker.Top = $dateTop + 22; $datePicker.Width = 210; $datePicker.Height = 30; $datePicker.Format = 'Custom'; $datePicker.CustomFormat = 'dd.MM.yyyy  HH:mm'; $datePicker.ShowUpDown = $true; $datePicker.Value = (Get-Date).AddHours(1)
    $datePicker.Font = New-Object Drawing.Font('Segoe UI', 9); $datePicker.CalendarMonthBackground = [Drawing.Color]::White; $dialog.Controls.Add($datePicker)
    $dateCheck = $null
    if ($isReminder) {
        $dateHint = New-Object Windows.Forms.Label
        $dateHint.Text = 'Уведомление появится в указанное время'; $dateHint.Left = 250; $dateHint.Top = 208; $dateHint.Width = 280; $dateHint.Height = 20; $dateHint.ForeColor = [Drawing.Color]::FromArgb(100, 116, 139); $dialog.Controls.Add($dateHint)
    } else {
        $dateCheck = New-Object Windows.Forms.CheckBox
        $dateCheck.Text = 'Указать срок'; $dateCheck.Left = 250; $dateCheck.Top = $dateTop + 25; $dateCheck.Width = 150; $dateCheck.Height = 24; $dateCheck.ForeColor = [Drawing.Color]::FromArgb(51, 65, 85); $dateCheck.Checked = $false; $dialog.Controls.Add($dateCheck)
        $datePicker.Enabled = $false; $dateCheck.Add_CheckedChanged({ $datePicker.Enabled = $dateCheck.Checked })
    }

    $sourceTop = if ($isReminder) { 246 } else { 316 }
    $buttonTop = if ($isReminder) { 278 } else { 348 }
    $sourceHint = New-Object Windows.Forms.Label
    $sourceHint.Text = 'Исходный текст сохранится полностью'; $sourceHint.Left = 24; $sourceHint.Top = $sourceTop; $sourceHint.Width = 330; $sourceHint.Height = 20; $sourceHint.ForeColor = [Drawing.Color]::FromArgb(100, 116, 139); $dialog.Controls.Add($sourceHint)
    $cancel = New-Object Windows.Forms.Button
    $cancel.Text = 'Отмена'; $cancel.Left = 348; $cancel.Top = $buttonTop; $cancel.Width = 88; $cancel.Height = 32; Set-QuickCaptureButtonLook $cancel $false $accent; $dialog.Controls.Add($cancel)
    $ok = New-Object Windows.Forms.Button
    $ok.Text = if ($isReminder) { 'Создать напоминание' } else { 'Создать задачу' }; $ok.Left = 408; $ok.Top = $buttonTop; $ok.Width = 126; $ok.Height = 32; Set-QuickCaptureButtonLook $ok $true $accent; $dialog.Controls.Add($ok)
    $dialog.AcceptButton = $ok; $dialog.CancelButton = $cancel; $dialog.Tag = $null
    $cancel.Add_Click({ $dialog.Close() })
    $ok.Add_Click({
        try {
            $due = if ($isReminder -or ($null -ne $dateCheck -and $dateCheck.Checked)) { $datePicker.Value.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            $dealNumber = if ($null -ne $dealBox) { [string]$dealBox.Text } else { '' }
            $dialog.Tag = New-QuickCaptureRecord $Kind $titleBox.Text $Text $due $dealNumber
            $dialog.DialogResult = 'OK'; $dialog.Close()
        } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Быстрый захват') | Out-Null }
    })
    $dialog.Add_Shown({ $titleBox.Focus(); $titleBox.SelectionStart = $titleBox.Text.Length })
    [void]$dialog.ShowDialog()
    $dialog.Dispose()
    return $dialog.Tag
}

function Get-SelectedText {
    Add-Type -AssemblyName System.Windows.Forms
    $clipboard = $null
    try { $clipboard = [Windows.Forms.Clipboard]::GetDataObject() } catch { Write-QuickCaptureLog ('Could not snapshot clipboard: ' + $_.Exception.Message) }
    $script:ClipboardSnapshot = $clipboard
    $beforeSequence = [QuickCapture.Native]::GetClipboardSequenceNumber()
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        if (-not [QuickCapture.Native]::IsDown(0x11) -and -not [QuickCapture.Native]::IsDown(0x10)) { break }
        Start-Sleep -Milliseconds 25
    }
    [QuickCapture.Native]::SendCtrlC()
    for ($attempt = 0; $attempt -lt 15; $attempt++) {
        Start-Sleep -Milliseconds 100
        $value = [string][Windows.Forms.Clipboard]::GetText()
        $sequence = [QuickCapture.Native]::GetClipboardSequenceNumber()
        if ($sequence -ne $beforeSequence -and -not [string]::IsNullOrWhiteSpace($value)) { return $value }
    }
    Write-QuickCaptureLog 'Ctrl+C did not change the clipboard sequence or produce text.'
    return ''
}

function Restore-QuickCaptureClipboard {
    if ($null -eq $script:ClipboardSnapshot) { return }
    $snapshot = $script:ClipboardSnapshot
    $script:ClipboardSnapshot = $null
    for ($attempt = 0; $attempt -lt 8; $attempt++) {
        try { [Windows.Forms.Clipboard]::SetDataObject($snapshot, $true); return } catch { Start-Sleep -Milliseconds 100 }
    }
    Write-QuickCaptureLog 'Could not restore the original clipboard contents.'
}

function Invoke-QuickCaptureFromHotkey {
    param([ValidateSet('task','reminder')][string]$Kind)
    if ($script:CaptureDialogOpen) {
        Write-QuickCaptureLog 'Ignored hotkey while a Quick Capture dialog is already open.'
        return
    }
    $script:CaptureDialogOpen = $true
    try {
        Write-QuickCaptureLog ("Hotkey {0} received." -f $Kind)
        $text = Get-SelectedText
        [void](Show-QuickCaptureDialog $Kind $text)
    } catch {
        Write-QuickCaptureLog ("{0} hotkey error: {1}" -f $Kind, $_.Exception.Message)
    } finally {
        Restore-QuickCaptureClipboard
        $script:CaptureDialogOpen = $false
    }
}

function Start-QuickCaptureHttp {
    $script:Listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, $script:Port)
    $script:Listener.Start()
}

function Process-QuickCaptureHttp {
    if ($null -eq $script:Listener -or -not $script:Listener.Pending()) { return }
    $client = $null; $stream = $null
    try {
        $client = $script:Listener.AcceptTcpClient()
        $stream = $client.GetStream()
        $buffer = New-Object byte[] 65536
        $count = $stream.Read($buffer, 0, $buffer.Length)
        $raw = [Text.Encoding]::UTF8.GetString($buffer, 0, $count)
        $parts = $raw -split "`r`n`r`n", 2
        $requestLine = ([string]$parts[0] -split "`r`n")[0]
        if ($requestLine -match '^OPTIONS ') { Send-QuickCaptureResponse $stream 204 @{}; return }
        if ($requestLine -notmatch '^POST /quick-capture/') { Send-QuickCaptureResponse $stream 404 @{ok=$false;error='Not found'}; return }
        $headers = [string]$parts[0]; $contentLength = 0
        if ($headers -match '(?im)^Content-Length:\s*(\d+)') { $contentLength = [int]$Matches[1] }
        $body = if ($parts.Count -gt 1) { [string]$parts[1] } else { '' }
        while ([Text.Encoding]::UTF8.GetByteCount($body) -lt $contentLength) {
            $n = $stream.Read($buffer, 0, $buffer.Length); if ($n -le 0) { break }
            $body += [Text.Encoding]::UTF8.GetString($buffer, 0, $n)
        }
        $request = if ([string]::IsNullOrWhiteSpace($body)) { $null } else { $body | ConvertFrom-Json }
        $kind = [string]$request.kind; $text = Normalize-QuickCaptureText ([string]$request.text)
        if ($kind -notin @('task','reminder')) { throw 'Неизвестный тип быстрого захвата.' }
        if ([string]::IsNullOrWhiteSpace($text)) { throw 'Выделенный текст пуст.' }
        [void](Show-QuickCaptureDialog $kind $text)
        Send-QuickCaptureResponse $stream 202 @{ok=$true;accepted=$true}
    } catch { try { if ($null -ne $stream) { Send-QuickCaptureResponse $stream 400 @{ok=$false;error=$_.Exception.Message} } } catch { } }
    finally { if ($null -ne $client) { $client.Close() } }
}

function Test-ParentAlive {
    if ($ParentPid -le 0) { return $true }
    try { return $null -ne (Get-Process -Id $ParentPid -ErrorAction Stop) } catch { return $false }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
if ($SelfTest) {
    $sample = Normalize-QuickCaptureText "  запрос коллеги`r`n  "
    if ($sample -ne 'запрос коллеги') { throw 'Normalize-QuickCaptureText failed.' }
    if ((Normalize-QuickCaptureText "`0abc") -ne 'abc') { throw 'NUL cleanup failed.' }
    Write-Host 'Quick Capture self-test OK'
    exit 0
}
Ensure-QuickCaptureSchema
Invoke-QuickCaptureLegacyMigration
$script:InstanceMutex = New-Object Threading.Mutex($false, 'Local\ProcurementControl.QuickCaptureAgent')
if (-not $script:InstanceMutex.WaitOne(0, $false)) {
    Write-QuickCaptureLog 'Another Quick Capture Agent instance is already running; this instance is exiting.'
    exit 0
}
$script:Form = New-Object Windows.Forms.Form
$script:Form.ShowInTaskbar = $false; $script:Form.WindowState = 'Minimized'; $script:Form.FormBorderStyle='FixedToolWindow'; $script:Form.Width=1; $script:Form.Height=1
$script:Form.Opacity = 0
$script:Form.Add_Shown({ $script:Form.Hide() })
$hotkeyWindow = New-Object QuickCapture.HotkeyWindow
$hotkeyWindow.Add_TaskPressed({
    Invoke-QuickCaptureFromHotkey 'task'
})
$hotkeyWindow.Add_ReminderPressed({
    Invoke-QuickCaptureFromHotkey 'reminder'
})
$hotkeyWindow.Register()
Write-QuickCaptureLog ("Agent started. ParentPid={0}; Port={1}; Hotkeys={2}+T,{2}+R" -f $ParentPid, $Port, $hotkeyWindow.Mode)
Start-QuickCaptureHttp
$timer = New-Object Windows.Forms.Timer; $timer.Interval=120
$timer.Add_Tick({
    try { Process-QuickCaptureHttp } catch { }
    if (-not (Test-ParentAlive)) { $script:Form.Close(); return }
})
$timer.Start()
[Windows.Forms.Application]::Run($script:Form)
try { $timer.Stop(); $script:Listener.Stop(); $hotkeyWindow.Dispose(); if ($null -ne $script:InstanceMutex) { $script:InstanceMutex.ReleaseMutex(); $script:InstanceMutex.Dispose() } } catch { }
