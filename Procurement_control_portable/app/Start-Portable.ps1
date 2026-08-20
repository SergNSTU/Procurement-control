param(
    [switch]$UiSmokeTest,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$RemainingArgs
)

$ErrorActionPreference = 'Stop'

$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PortableRoot = Split-Path -Parent $AppDir
$MainScript = Join-Path $AppDir 'RRFQComparer.ps1'
$LogDir = Join-Path $PortableRoot 'data\logs'

function Write-StartupLog {
    param([string]$Message, $ErrorRecord)

    try { New-Item -ItemType Directory -Force -Path $LogDir | Out-Null } catch { }
    $logPath = Join-Path $LogDir ('startup_error_{0}.txt' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $details = New-Object System.Collections.ArrayList
    [void]$details.Add('Procurement control startup error')
    [void]$details.Add(('Time: {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    [void]$details.Add(('PortableRoot: {0}' -f $PortableRoot))
    [void]$details.Add(('MainScript: {0}' -f $MainScript))
    [void]$details.Add('')
    [void]$details.Add($Message)
    if ($null -ne $ErrorRecord) {
        [void]$details.Add('')
        [void]$details.Add(($ErrorRecord | Out-String))
        if ($ErrorRecord.ScriptStackTrace) {
            [void]$details.Add('Stack trace:')
            [void]$details.Add($ErrorRecord.ScriptStackTrace)
        }
    }
    try { Set-Content -LiteralPath $logPath -Value ([string[]]$details) -Encoding UTF8 } catch { }
    return $logPath
}

try {
    if ($PSVersionTable.PSVersion -lt [Version]'5.1') {
        throw 'Требуется Windows PowerShell 5.1 или новее.'
    }

    $netFrameworkRelease = $null
    try {
        $netFrameworkRelease = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction Stop).Release
    } catch {
        $netFrameworkRelease = $null
    }
    if ($null -eq $netFrameworkRelease -or [int]$netFrameworkRelease -lt 393295) {
        throw '.NET Framework 4.6 или новее не найден. Установите .NET Framework и повторите запуск.'
    }

    if (-not (Test-Path -LiteralPath $MainScript -PathType Leaf)) {
        throw "Main script was not found: $MainScript"
    }

    Set-Location -LiteralPath $PortableRoot

    $dataDir = Join-Path $PortableRoot 'data'
    New-Item -ItemType Directory -Force -Path $dataDir | Out-Null
    $writeTest = Join-Path $dataDir '.portable_write_test'
    Set-Content -LiteralPath $writeTest -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -Encoding ASCII
    Remove-Item -LiteralPath $writeTest -Force -ErrorAction SilentlyContinue

    $pathsToUnblock = New-Object System.Collections.ArrayList
    foreach ($path in @(
        (Join-Path $PortableRoot 'Start.bat'),
        (Join-Path $PortableRoot 'Start.vbs'),
        $MainScript,
        $MyInvocation.MyCommand.Path
    )) {
        if (Test-Path -LiteralPath $path -PathType Leaf) { [void]$pathsToUnblock.Add($path) }
    }
    $sqliteDir = Join-Path $AppDir 'lib\sqlite'
    if (Test-Path -LiteralPath $sqliteDir -PathType Container) {
        Get-ChildItem -LiteralPath $sqliteDir -Recurse -File -Include *.dll -ErrorAction SilentlyContinue | ForEach-Object { [void]$pathsToUnblock.Add($_.FullName) }
    }
    foreach ($path in @($pathsToUnblock)) {
        try { Unblock-File -LiteralPath $path -ErrorAction SilentlyContinue } catch { }
    }

    if ($UiSmokeTest) {
        & $MainScript -UiSmokeTest @RemainingArgs
    } else {
        & $MainScript @RemainingArgs
    }
    exit $LASTEXITCODE
} catch {
    $message = $_.Exception.Message
    $logPath = Write-StartupLog $message $_
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("Procurement control could not start.`r`n`r`n$message`r`n`r`nLog:`r`n$logPath", 'Procurement control') | Out-Null
    } catch {
    }
    exit 1
}
