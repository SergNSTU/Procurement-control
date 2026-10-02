$ErrorActionPreference = 'Stop'
$agent = Join-Path (Split-Path -Parent $PSScriptRoot) 'QuickCaptureAgent.ps1'

function Assert-True { param([bool]$Condition, [string]$Message); if (-not $Condition) { throw $Message } }
function Assert-Equal { param($Actual, $Expected, [string]$Message); if ($Actual -ne $Expected) { throw "$Message Expected=[$Expected] Actual=[$Actual]" } }

$text = & powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File $agent -SelfTest 2>&1
Assert-True ($LASTEXITCODE -eq 0) "QuickCaptureAgent self-test failed: $text"
Assert-True (($text -join "`n") -match 'Quick Capture self-test OK') 'Self-test output is missing.'
Write-Host '[PASS] Quick Capture agent text normalization and startup self-test'
Write-Host 'Quick Capture tests passed.'
