# Recovery snapshots for procurement data. Dot-sourced by RRFQComparer.ps1.
# A snapshot is a self-contained ZIP with the SQLite database and internal data folders.

function Get-PurchaseRecoveryRoot {
    return (Join-Path $script:AppRoot 'data\recovery')
}

function Get-PurchaseRecoveryFolder {
    param([ValidateSet('Backup', 'Trash')][string]$Kind = 'Backup')

    $name = if ($Kind -eq 'Trash') { 'trash' } else { 'backups' }
    return (Join-Path (Get-PurchaseRecoveryRoot) $name)
}

function Initialize-PurchaseRecoveryStore {
    foreach ($folder in @((Get-PurchaseRecoveryFolder 'Backup'), (Get-PurchaseRecoveryFolder 'Trash'))) {
        New-Item -ItemType Directory -Force -Path $folder | Out-Null
    }
}

function Get-ReadableFileSize {
    param([long]$Bytes)

    if ($Bytes -ge 1GB) { return ('{0:N1} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N1} KB' -f ($Bytes / 1KB)) }
    return "$Bytes B"
}

function Copy-RecoveryDirectory {
    param([string]$Source, [string]$Destination)

    if (-not (Test-Path -LiteralPath $Source -PathType Container)) { return }
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    foreach ($item in @(Get-ChildItem -LiteralPath $Source -Force -ErrorAction Stop)) {
        Copy-Item -LiteralPath $item.FullName -Destination $Destination -Recurse -Force -ErrorAction Stop
    }
}

function Export-PurchaseDatabaseForRecovery {
    param([string]$Destination)

    if ([string]::IsNullOrWhiteSpace($script:PurchaseDbPath)) {
        $script:PurchaseDbPath = Join-Path (Get-PurchaseDataRoot) 'purchase_control.sqlite'
    }
    if (-not (Test-Path -LiteralPath $script:PurchaseDbPath -PathType Leaf)) {
        throw 'База данных закупки не найдена.'
    }

    Load-SqliteProvider
    $connection = $null
    try {
        $connection = Open-PurchaseConnection
        $command = $connection.CreateCommand()
        $command.CommandText = 'PRAGMA wal_checkpoint(TRUNCATE);'
        [void]$command.ExecuteScalar()
    } finally {
        if ($null -ne $connection) { $connection.Dispose() }
    }
    [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
    Copy-Item -LiteralPath $script:PurchaseDbPath -Destination $Destination -Force -ErrorAction Stop
}

function New-PurchaseRecoverySnapshot {
    param(
        [ValidateSet('Backup', 'Trash')][string]$Kind = 'Backup',
        [string]$Reason = 'Резервная копия'
    )

    Initialize-PurchaseRecoveryStore
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $createdAt = Get-Date
    $stamp = $createdAt.ToString('yyyyMMdd_HHmmss_fff')
    $safeReason = (($Reason -replace '[^0-9A-Za-zА-Яа-я_-]+', '_').Trim('_'))
    if ([string]::IsNullOrWhiteSpace($safeReason)) { $safeReason = 'snapshot' }
    if ($safeReason.Length -gt 45) { $safeReason = $safeReason.Substring(0, 45) }
    $prefix = if ($Kind -eq 'Trash') { 'trash' } else { 'backup' }
    $archivePath = Join-Path (Get-PurchaseRecoveryFolder $Kind) ('{0}_{1}_{2}.zip' -f $prefix, $stamp, $safeReason)
    $stage = Join-Path ([IO.Path]::GetTempPath()) ('procurement-recovery-' + [Guid]::NewGuid().ToString('N'))

    try {
        New-Item -ItemType Directory -Force -Path $stage | Out-Null
        $snapshotDbDir = Join-Path $stage 'purchase_control'
        New-Item -ItemType Directory -Force -Path $snapshotDbDir | Out-Null
        Export-PurchaseDatabaseForRecovery (Join-Path $snapshotDbDir 'purchase_control.sqlite')
        Copy-RecoveryDirectory (Join-Path (Get-PurchaseDataRoot) 'purchase_control\files') (Join-Path $snapshotDbDir 'files')
        Copy-RecoveryDirectory (Join-Path (Get-ComponentDataRoot) 'files') (Join-Path $stage 'components\files')
        Copy-RecoveryDirectory (Get-NotesDataRoot) (Join-Path $stage 'notes')

        $manifest = [pscustomobject]@{
            Format = 1
            Kind = $Kind
            Reason = $Reason
            CreatedAt = $createdAt.ToString('o')
            Includes = @('purchase_control.sqlite', 'purchase_control/files', 'components/files', 'notes')
            Note = 'Внешняя папка документов не входит в резервную копию.'
        }
        [IO.File]::WriteAllText((Join-Path $stage 'manifest.json'), ($manifest | ConvertTo-Json -Depth 4), [Text.UTF8Encoding]::new($false))
        [IO.Compression.ZipFile]::CreateFromDirectory($stage, $archivePath, [IO.Compression.CompressionLevel]::Optimal, $false)
    } finally {
        if (Test-Path -LiteralPath $stage -PathType Container) {
            Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if (-not (Test-Path -LiteralPath $archivePath -PathType Leaf)) { throw 'Не удалось создать резервную копию.' }
    if ($Kind -eq 'Backup') {
        # Manual and pre-restore backups stay useful without letting the folder grow forever.
        $oldBackups = @(Get-PurchaseRecoverySnapshots | Where-Object { $_.Kind -eq 'Backup' } | Sort-Object CreatedAt -Descending | Select-Object -Skip 20)
        foreach ($oldBackup in $oldBackups) {
            try { Remove-PurchaseRecoverySnapshot ([string]$oldBackup.Path) } catch { }
        }
    }
    return $archivePath
}

function Get-RecoverySnapshotManifest {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($Path)
        $entry = $archive.GetEntry('manifest.json')
        if ($null -eq $entry) { throw 'В архиве нет manifest.json.' }
        $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
        try { return ($reader.ReadToEnd() | ConvertFrom-Json) } finally { $reader.Dispose() }
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
    }
}

function Get-PurchaseRecoverySnapshots {
    Initialize-PurchaseRecoveryStore
    $items = New-Object System.Collections.ArrayList
    foreach ($kind in @('Trash', 'Backup')) {
        $folder = Get-PurchaseRecoveryFolder $kind
        foreach ($file in @(Get-ChildItem -LiteralPath $folder -Filter '*.zip' -File -ErrorAction SilentlyContinue)) {
            try {
                $manifest = Get-RecoverySnapshotManifest $file.FullName
                [void]$items.Add([pscustomobject]@{
                    Path = $file.FullName
                    Name = $file.Name
                    Kind = [string]$manifest.Kind
                    Reason = [string]$manifest.Reason
                    CreatedAt = [datetime]$manifest.CreatedAt
                    SizeBytes = [long]$file.Length
                    Size = Get-ReadableFileSize $file.Length
                })
            } catch {
                [void]$items.Add([pscustomobject]@{
                    Path = $file.FullName; Name = $file.Name; Kind = $kind; Reason = 'Поврежденный или устаревший архив'; CreatedAt = $file.LastWriteTime
                    SizeBytes = [long]$file.Length; Size = Get-ReadableFileSize $file.Length
                })
            }
        }
    }
    return @($items | Sort-Object CreatedAt -Descending)
}

function Test-PurchaseRecoveryArchivePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    try {
        $fullPath = [IO.Path]::GetFullPath($Path)
        $root = [IO.Path]::GetFullPath((Get-PurchaseRecoveryRoot))
        if (-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)) { $root += [IO.Path]::DirectorySeparatorChar }
        return $fullPath.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)
    } catch { return $false }
}

function Restore-PurchaseRecoverySnapshot {
    param([string]$Path)

    if (-not (Test-PurchaseRecoveryArchivePath $Path)) { throw 'Выберите архив из списка резервных копий.' }
    $manifest = Get-RecoverySnapshotManifest $Path
    $stage = Join-Path ([IO.Path]::GetTempPath()) ('procurement-restore-' + [Guid]::NewGuid().ToString('N'))
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [IO.Compression.ZipFile]::ExtractToDirectory($Path, $stage)
        $snapshotDb = Join-Path $stage 'purchase_control\purchase_control.sqlite'
        if (-not (Test-Path -LiteralPath $snapshotDb -PathType Leaf)) { throw 'Архив не содержит базу данных закупки.' }

        # Preserve the current state before replacing it, so restore itself is reversible.
        $restoreBackup = New-PurchaseRecoverySnapshot -Kind 'Backup' -Reason ('Перед восстановлением: ' + [string]$manifest.Reason)
        [System.Data.SQLite.SQLiteConnection]::ClearAllPools()
        $targets = @(
            (Join-Path (Get-PurchaseDataRoot) 'purchase_control'),
            (Get-ComponentDataRoot),
            (Get-NotesDataRoot)
        )
        foreach ($target in $targets) {
            if (Test-Path -LiteralPath $target -PathType Container) {
                Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction Stop
            }
        }
        Copy-RecoveryDirectory (Join-Path $stage 'purchase_control') (Join-Path (Get-PurchaseDataRoot) 'purchase_control')
        Copy-RecoveryDirectory (Join-Path $stage 'components') (Get-ComponentDataRoot)
        Copy-RecoveryDirectory (Join-Path $stage 'notes') (Get-NotesDataRoot)
        $script:PurchaseDbPath = Join-Path (Get-PurchaseDataRoot) 'purchase_control.sqlite'
        Initialize-PurchaseStore
        return $restoreBackup
    } finally {
        if (Test-Path -LiteralPath $stage -PathType Container) {
            Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Remove-PurchaseRecoverySnapshot {
    param([string]$Path)

    if (-not (Test-PurchaseRecoveryArchivePath $Path)) { throw 'Можно удалить только архив из списка восстановления.' }
    Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
}
