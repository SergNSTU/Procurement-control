param(
    [string]$CurrentPath = (Join-Path $PSScriptRoot '..\RRFQComparer.ps1'),
    [string]$BackupPath = (Join-Path $PSScriptRoot '..\RRFQComparer.ps1.bak_20260812_165215')
)

$utf8 = [System.Text.UTF8Encoding]::new($false, $false)
$cp1251 = [System.Text.Encoding]::GetEncoding(1251)
$replacementChar = [char]0xFFFD

function Read-TextExact {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][System.Text.Encoding]$Encoding
    )

    $resolved = Resolve-Path $Path
    $bytes = [System.IO.File]::ReadAllBytes($resolved)
    return $Encoding.GetString($bytes)
}

function Get-ReplacementCount {
    param([Parameter(Mandatory = $true)][string]$Text)

    return ($Text.ToCharArray() | Where-Object { $_ -eq $replacementChar }).Count
}

function Get-AssignmentValue {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$VariableName
    )

    $escapedName = [regex]::Escape($VariableName)
    $match = [regex]::Match(
        $Text,
        "(?m)^\s*" + $escapedName + "\.Text\s*=\s*'([^']*)'"
    )
    if ($match.Success) {
        return $match.Groups[1].Value
    }

    return $null
}

$currentText = Read-TextExact -Path $CurrentPath -Encoding $utf8
$backupText = Read-TextExact -Path $BackupPath -Encoding $cp1251

$labels = @(
    '$btnRfq',
    '$btnSave',
    '$btnNavRrfq',
    '$btnImportRrfq',
    '$historyHint'
)

$labelRows = foreach ($label in $labels) {
    [pscustomobject]@{
        Label = $label
        Current = Get-AssignmentValue -Text $currentText -VariableName $label
        Backup = Get-AssignmentValue -Text $backupText -VariableName $label
    }
}

[pscustomobject]@{
    CurrentPath = (Resolve-Path $CurrentPath).Path
    BackupPath = (Resolve-Path $BackupPath).Path
    CurrentReplacementCount = Get-ReplacementCount -Text $currentText
    BackupReplacementCount = Get-ReplacementCount -Text $backupText
    CurrentContainsReplacementRuns = [regex]::IsMatch($currentText, [string]$replacementChar + '{4,}')
    BackupContainsReplacementRuns = [regex]::IsMatch($backupText, [string]$replacementChar + '{4,}')
} | Format-List

''
'Key labels:'
$labelRows | Format-Table -AutoSize
