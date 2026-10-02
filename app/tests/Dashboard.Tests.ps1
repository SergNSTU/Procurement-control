#requires -Version 5.1
$ErrorActionPreference = 'Stop'

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message)
    if ($Expected -ne $Actual) { throw "$Message Expected=[$Expected] Actual=[$Actual]" }
}

. (Join-Path $PSScriptRoot '..\RRFQComparer.ps1') -ImportOnly

$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('procurement-dashboard-' + [Guid]::NewGuid().ToString('N'))
$script:DashboardTestDataRoot = Join-Path $testRoot 'data\purchase_control'
$script:DashboardTestComponentsRoot = Join-Path $testRoot 'data\components'

function Get-PurchaseDataRoot { return $script:DashboardTestDataRoot }
function Get-DefaultPurchaseDocumentsRoot { return (Join-Path (Get-PurchaseDataRoot) 'files') }
function Get-ComponentDataRoot { return $script:DashboardTestComponentsRoot }
function Get-DefaultComponentFilesRoot { return (Join-Path (Get-ComponentDataRoot) 'files') }
function Open-PurchaseConnection {
    Load-SqliteProvider
    $dataRoot = Get-PurchaseDataRoot
    New-Item -ItemType Directory -Force -Path $dataRoot | Out-Null
    if ([string]::IsNullOrWhiteSpace($script:PurchaseDbPath)) {
        $script:PurchaseDbPath = Join-Path $dataRoot 'purchase_control.sqlite'
    }
    $connection = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$script:PurchaseDbPath;Version=3;Foreign Keys=True;Pooling=False;Journal Mode=Delete;Synchronous=Normal;")
    $connection.Open()
    return $connection
}

try {
    $script:PurchaseDbPath = $null
    $script:PurchaseDocumentsRoot = $null
    Initialize-PurchaseStore
    $orderedStatus = Get-PurchaseOrderedStatus

    Assert-True (Test-PurchaseColumnExists 'deals' 'ordered_at') 'Expected ordered_at on deals.'
    Assert-True (Test-PurchaseColumnExists 'component_deals' 'ordered_at') 'Expected ordered_at on component_deals.'
    Assert-True (-not [string]::IsNullOrWhiteSpace((Get-PurchaseSetting 'dashboard.orders_tracking_started_at'))) 'Expected tracking start setting.'

    $legacyNow = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO deals(deal_number, status, created_at, updated_at)
VALUES(@deal_number, @status, @created_at, @updated_at)
'@ @{
        '@deal_number' = 'LEGACY-ORDERED'
        '@status' = $orderedStatus
        '@created_at' = $legacyNow
        '@updated_at' = $legacyNow
    }
    $legacyId = [int](Invoke-PurchaseScalar 'SELECT id FROM deals WHERE deal_number = @deal_number' @{ '@deal_number' = 'LEGACY-ORDERED' })
    Update-PurchaseDeal $legacyId '' $orderedStatus ''
    Assert-Equal '' ([string](Invoke-PurchaseScalar 'SELECT IFNULL(ordered_at, '''') FROM deals WHERE id = @id' @{ '@id' = $legacyId })) 'Legacy order must not get an inferred order date.'

    $directOrderId = New-PurchaseDeal -DealNumber 'ASSEMBLY-DIRECT-ORDER' -Status $orderedStatus
    $directOrderStatus = [string](Invoke-PurchaseScalar 'SELECT status FROM deals WHERE id = @id' @{ '@id' = $directOrderId })
    $directOrderedAt = [string](Invoke-PurchaseScalar 'SELECT IFNULL(ordered_at, '''') FROM deals WHERE id = @id' @{ '@id' = $directOrderId })
    Assert-True (-not [string]::IsNullOrWhiteSpace($directOrderedAt)) "Expected order date on direct ordered assembly creation. Status=[$directOrderStatus] OrderedAt=[$directOrderedAt]"

    $assemblyId = New-PurchaseDeal 'ASSEMBLY-ORDER' '' '' '' 'RFQ'
    Update-PurchaseDeal $assemblyId '' $orderedStatus ''
    $assemblyOrderedAt = [string](Invoke-PurchaseScalar 'SELECT IFNULL(ordered_at, '''') FROM deals WHERE id = @id' @{ '@id' = $assemblyId })
    Assert-True (-not [string]::IsNullOrWhiteSpace($assemblyOrderedAt)) 'Expected order date on assembly status transition.'
    Invoke-PurchaseNonQuery 'UPDATE deals SET ordered_at = @ordered_at WHERE id = @id' @{ '@ordered_at' = '2099-01-02 03:04:05'; '@id' = $assemblyId }
    Update-PurchaseDeal $assemblyId '' $orderedStatus 'Повторное редактирование'
    Assert-Equal '2099-01-02 03:04:05' ([string](Invoke-PurchaseScalar 'SELECT ordered_at FROM deals WHERE id = @id' @{ '@id' = $assemblyId })) 'Assembly order date must not be overwritten.'

    $componentId = New-ComponentDeal
    Update-ComponentDeal $componentId 'COMPONENT-ORDER' $orderedStatus 'Заказан' '' '' '' '' '3' ''
    $componentOrderedAt = [string](Invoke-PurchaseScalar 'SELECT IFNULL(ordered_at, '''') FROM component_deals WHERE id = @id' @{ '@id' = $componentId })
    Assert-True (-not [string]::IsNullOrWhiteSpace($componentOrderedAt)) 'Expected order date on component status transition.'
    Invoke-PurchaseNonQuery 'UPDATE component_deals SET ordered_at = @ordered_at WHERE id = @id' @{ '@ordered_at' = '2099-01-02 03:04:05'; '@id' = $componentId }
    Update-ComponentDeal $componentId 'COMPONENT-ORDER' $orderedStatus 'Заказан' 'Повторное редактирование' '' '' '' '3' ''
    Assert-Equal '2099-01-02 03:04:05' ([string](Invoke-PurchaseScalar 'SELECT ordered_at FROM component_deals WHERE id = @id' @{ '@id' = $componentId })) 'Component order date must not be overwritten.'

    Set-PurchaseSetting 'dashboard.orders_tracking_started_at' '2025-01-01 00:00:00'
    Invoke-PurchaseNonQuery @'
INSERT INTO deals(deal_number, status, created_at, ordered_at, updated_at)
VALUES('ASSEMBLY-COHORT', @status, '2025-01-15 10:00:00', '2025-03-10 12:00:00', '2025-03-10 12:00:00')
'@ @{ '@status' = $orderedStatus }
    Invoke-PurchaseNonQuery @'
INSERT INTO component_deals(entry_date, deal_number, status, stage, priority, created_at, ordered_at, updated_at)
VALUES('15.01.2025', 'COMPONENT-COHORT', @status, 'Заказан', '3', '2025-01-15 10:00:00', '2025-03-10 12:00:00', '2025-03-10 12:00:00')
'@ @{ '@status' = $orderedStatus }
    $januaryStart = [DateTime]::new(2025, 1, 1)
    $marchStart = [DateTime]::new(2025, 3, 1)
    $trackingStart = [DateTime]::new(2025, 1, 1)
    $aprilNow = [DateTime]::new(2025, 4, 20)
    $assemblyJanuary = Get-PurchaseDashboardDirection 'deals' $januaryStart $januaryStart.AddMonths(1) $trackingStart $aprilNow
    $assemblyMarch = Get-PurchaseDashboardDirection 'deals' $marchStart $marchStart.AddMonths(1) $trackingStart $aprilNow
    $componentJanuary = Get-PurchaseDashboardDirection 'component_deals' $januaryStart $januaryStart.AddMonths(1) $trackingStart $aprilNow
    $componentMarch = Get-PurchaseDashboardDirection 'component_deals' $marchStart $marchStart.AddMonths(1) $trackingStart $aprilNow
    Assert-Equal 1 $assemblyJanuary.AddedThisMonth 'Assembly must be counted in its creation month.'
    Assert-Equal 0 $assemblyJanuary.OrderedThisMonth 'Assembly must not be counted as ordered in its creation month.'
    Assert-Equal 1 $assemblyMarch.OrderedThisMonth 'Assembly must be counted in its order month.'
    Assert-Equal 100 $assemblyJanuary.CurrentConversion 'Assembly cohort conversion must include later orders.'
    Assert-Equal 1 $componentJanuary.AddedThisMonth 'Component must be counted in its creation month.'
    Assert-Equal 0 $componentJanuary.OrderedThisMonth 'Component must not be counted as ordered in its creation month.'
    Assert-Equal 1 $componentMarch.OrderedThisMonth 'Component must be counted in its order month.'
    Assert-Equal 100 $componentJanuary.CurrentConversion 'Component cohort conversion must include later orders.'

    Write-Host 'Dashboard migration and cohort tests passed.'
} finally {
    try { [System.Data.SQLite.SQLiteConnection]::ClearAllPools() } catch { }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    if (Test-Path -LiteralPath $testRoot) {
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}
