param(
    [string]$TestName = 'All'
)

$ErrorActionPreference = 'Stop'

$script:FixtureResults = [ordered]@{
    Passed = 0
    Failed = 0
}

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Equal {
    param(
        $Actual,
        $Expected,
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw ($Message + "`nExpected: {0}`nActual:   {1}" -f $Expected, $Actual)
    }
}

function Invoke-TestCase {
    param(
        [string]$Name,
        [scriptblock]$Script
    )

    try {
        & $Script
        $script:FixtureResults.Passed++
        Write-Host "[PASS] $Name"
    } catch {
        $script:FixtureResults.Failed++
        Write-Host "[FAIL] $Name"
        Write-Host $_.Exception.Message
    }
}

function New-TestAppRoot {
    $root = Join-Path ([IO.Path]::GetTempPath()) ("bitrix-sync-fixture-{0}" -f [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path (Join-Path $root 'data\purchase_control') | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $root 'data\notes') | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $root 'data\components\files') | Out-Null

    return $root
}

function Remove-TestAppRoot {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (Test-Path -LiteralPath $Path) {
        try {
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
        } catch {
        }
    }
}

function Use-TestStore {
    param(
        [string]$AppRoot,
        [scriptblock]$Script
    )

    $oldAppRoot = $script:AppRoot
    $oldDbPath = $script:PurchaseDbPath
    $oldDocsRoot = $script:PurchaseDocumentsRoot
    $oldSqliteLoaded = $script:SqliteLoaded
    $oldDataRootOverride = $script:PurchaseDataRootOverride
    $oldNotesRootOverride = $script:PurchaseNotesRootOverride
    $oldComponentsRootOverride = $script:PurchaseComponentsRootOverride

    try {
        $script:AppRoot = $oldAppRoot
        $script:PurchaseDbPath = $null
        $script:PurchaseDocumentsRoot = $null
        $script:SqliteLoaded = $false
        $script:PurchaseDataRootOverride = Join-Path $AppRoot 'data\purchase_control'
        $script:PurchaseNotesRootOverride = Join-Path $AppRoot 'data\notes'
        $script:PurchaseComponentsRootOverride = Join-Path $AppRoot 'data\components'
        & $Script
    } finally {
        $script:AppRoot = $oldAppRoot
        $script:PurchaseDbPath = $oldDbPath
        $script:PurchaseDocumentsRoot = $oldDocsRoot
        $script:SqliteLoaded = $oldSqliteLoaded
        $script:PurchaseDataRootOverride = $oldDataRootOverride
        $script:PurchaseNotesRootOverride = $oldNotesRootOverride
        $script:PurchaseComponentsRootOverride = $oldComponentsRootOverride
    }
}

. (Join-Path $PSScriptRoot '..\RRFQComparer.ps1') -ImportOnly

function Get-PurchaseDataRoot {
    if (-not [string]::IsNullOrWhiteSpace($script:PurchaseDataRootOverride)) {
        return $script:PurchaseDataRootOverride
    }

    return (Join-Path $script:AppRoot 'data\purchase_control')
}

function Get-DefaultPurchaseDocumentsRoot {
    return (Join-Path (Get-PurchaseDataRoot) 'files')
}

function Get-ComponentDataRoot {
    if (-not [string]::IsNullOrWhiteSpace($script:PurchaseComponentsRootOverride)) {
        return $script:PurchaseComponentsRootOverride
    }

    return (Join-Path $script:AppRoot 'data\components')
}

function Get-DefaultComponentFilesRoot {
    return (Join-Path (Get-ComponentDataRoot) 'files')
}

function Get-NotesDataRoot {
    if (-not [string]::IsNullOrWhiteSpace($script:PurchaseNotesRootOverride)) {
        return $script:PurchaseNotesRootOverride
    }

    return (Join-Path $script:AppRoot 'data\notes')
}

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

function Get-TableSql {
    param([string]$Name)

    return [string](Invoke-PurchaseScalar 'SELECT sql FROM sqlite_master WHERE type = ''table'' AND name = @name' @{
        '@name' = $Name
    })
}

function Get-ColumnNames {
    param([string]$TableName)

    $names = @()
    $table = Invoke-PurchaseQuery "PRAGMA table_info($TableName)"
    foreach ($row in $table.Rows) {
        $names += [string]$row.name
    }
    return $names
}

function Run-SchemaTests {
    Invoke-TestCase 'initializes Task 3 schema idempotently and preserves blocked tasks' {
        $appRoot = New-TestAppRoot
        try {
            Use-TestStore $appRoot {
                Initialize-PurchaseStore
                Block-BitrixTask 901 'Keep me'
                $dealId = New-PurchaseDeal 'AE26081801PK-000001'
                Initialize-PurchaseStore

                $dealColumns = @(Get-ColumnNames 'deals')
                Assert-True ($dealColumns -contains 'bitrix_deal_id') 'Expected deals.bitrix_deal_id column.'
                Assert-Equal (@($dealColumns | Where-Object { $_ -eq 'bitrix_deal_id' }).Count) 1 'Expected deals.bitrix_deal_id to exist once.'

                $taskLinkSql = Get-TableSql 'bitrix_task_links'
                Assert-True (-not [string]::IsNullOrWhiteSpace($taskLinkSql)) 'Expected bitrix_task_links table.'
                Assert-True ($taskLinkSql -match 'bitrix_task_id') 'Expected bitrix_task_links to include bitrix_task_id.'
                Assert-True ($taskLinkSql -match 'deal_id') 'Expected bitrix_task_links to include deal_id.'
                Assert-True ($taskLinkSql -match 'is_manual_deleted') 'Expected bitrix_task_links to include manual deletion tombstone.'

                $rfqSql = Get-TableSql 'bitrix_rfq_files'
                Assert-True (-not [string]::IsNullOrWhiteSpace($rfqSql)) 'Expected bitrix_rfq_files table.'
                Assert-True ($rfqSql -match 'file_key') 'Expected bitrix_rfq_files to include file_key.'

                $blocked = @(Get-BitrixBlockedTaskIds)
                Assert-Equal $blocked.Count 1 'Expected blocked task rows to survive re-initialization.'
                Assert-Equal $blocked[0] 901 'Unexpected blocked task ID after re-initialization.'

                Save-BitrixTaskLink -BitrixTaskId 501 -DealId $dealId -BitrixStatus 2 -BitrixDealId 777 -IsManualDeleted:$true
                $link = Get-BitrixTaskLink 501
                Assert-Equal $link.bitrix_task_id 501 'Expected to load the saved task link.'
                Assert-Equal $link.deal_id $dealId 'Expected saved task link to keep the local deal id.'
                Assert-Equal $link.bitrix_deal_id 777 'Expected saved task link to keep the Bitrix deal id.'
                Assert-Equal $link.bitrix_status 2 'Expected saved task link to keep the Bitrix status.'
                Assert-Equal $link.is_manual_deleted 1 'Expected saved task link to keep the manual deletion tombstone.'

                Save-BitrixTaskLink -BitrixTaskId 501 -DealId $dealId -BitrixStatus 6 -BitrixDealId 778 -IsManualDeleted:$false
                $linkRows = [int](Invoke-PurchaseScalar 'SELECT COUNT(*) FROM bitrix_task_links WHERE bitrix_task_id = @task_id' @{
                    '@task_id' = 501
                })
                Assert-Equal $linkRows 1 'Expected bitrix_task_links to stay unique per task.'

                $updatedLink = Get-BitrixTaskLink 501
                Assert-Equal $updatedLink.bitrix_deal_id 778 'Expected task link upsert to update the Bitrix deal id.'
                Assert-Equal $updatedLink.bitrix_status 6 'Expected task link upsert to update the Bitrix status.'
                Assert-Equal $updatedLink.is_manual_deleted 0 'Expected task link upsert to clear the manual deletion tombstone.'

                $deletedDealId = New-PurchaseDeal 'AE26081801PK-000099'
                Save-BitrixTaskLink -BitrixTaskId 502 -DealId $deletedDealId -BitrixStatus 3 -BitrixDealId 880
                Delete-PurchaseDeal $deletedDealId

                $deletedLink = Get-BitrixTaskLink 502
                Assert-True ($null -ne $deletedLink) 'Expected deleted deals to keep a Bitrix tombstone row.'
                Assert-Equal $deletedLink.bitrix_task_id 502 'Expected deleted deal tombstone to keep the Bitrix task id.'
                Assert-Equal $deletedLink.deal_id ([DBNull]::Value) 'Expected deleted deal tombstone to clear the local deal id.'
                Assert-Equal $deletedLink.bitrix_deal_id 880 'Expected deleted deal tombstone to keep the Bitrix deal id.'
                Assert-Equal $deletedLink.bitrix_status 3 'Expected deleted deal tombstone to keep the Bitrix status.'
                Assert-Equal $deletedLink.is_manual_deleted 1 'Expected deleted deal tombstone to set the manual deletion flag.'

                Save-ProcessedBitrixFile -BitrixTaskId 501 -FileKey 'rfq:latest' -BitrixFileId 'f-1' -FileName 'RFQ_AE26081801PK-000001.xlsx' -FileHash 'hash-1' -LocalPath 'data\purchase_control\files\AE26081801PK-000001\RRFQ\RFQ_AE26081801PK-000001.xlsx' -UploadState 'Uploaded' -ErrorText ''
                $rfqFile = Get-ProcessedBitrixFile 501 'rfq:latest'
                Assert-Equal $rfqFile.bitrix_task_id 501 'Expected to load the saved processed RFQ file.'
                Assert-Equal $rfqFile.file_key 'rfq:latest' 'Expected saved processed RFQ file key.'
                Assert-Equal $rfqFile.bitrix_file_id 'f-1' 'Expected saved processed RFQ Bitrix file id.'
                Assert-Equal $rfqFile.upload_state 'Uploaded' 'Expected saved processed RFQ upload state.'

                Save-ProcessedBitrixFile -BitrixTaskId 501 -FileKey 'rfq:latest' -BitrixFileId 'f-2' -FileName 'RFQ_AE26081801PK-000001_v2.xlsx' -FileHash 'hash-2' -LocalPath 'data\purchase_control\files\AE26081801PK-000001\RRFQ\RFQ_AE26081801PK-000001_v2.xlsx' -UploadState 'Skipped' -ErrorText 'duplicate'
                $rfqRows = [int](Invoke-PurchaseScalar 'SELECT COUNT(*) FROM bitrix_rfq_files WHERE bitrix_task_id = @task_id AND file_key = @file_key' @{
                    '@task_id' = 501
                    '@file_key' = 'rfq:latest'
                })
                Assert-Equal $rfqRows 1 'Expected processed RFQ rows to stay unique per task and file key.'

                $updatedRfqFile = Get-ProcessedBitrixFile 501 'rfq:latest'
                Assert-Equal $updatedRfqFile.bitrix_file_id 'f-2' 'Expected processed RFQ upsert to update the Bitrix file id.'
                Assert-Equal $updatedRfqFile.file_name 'RFQ_AE26081801PK-000001_v2.xlsx' 'Expected processed RFQ upsert to update the file name.'
                Assert-Equal $updatedRfqFile.file_hash 'hash-2' 'Expected processed RFQ upsert to update the file hash.'
                Assert-Equal $updatedRfqFile.upload_state 'Skipped' 'Expected processed RFQ upsert to update the upload state.'
                Assert-Equal $updatedRfqFile.error_text 'duplicate' 'Expected processed RFQ upsert to update the error text.'

                Save-ProcessedBitrixFile -BitrixTaskId 501 -FileKey 'rfq:preserve' -BitrixFileId 'f-9' -FileName 'RFQ_keep.xlsx' -FileHash 'hash-keep' -LocalPath 'data\purchase_control\files\AE26081801PK-000001\RRFQ\RFQ_keep.xlsx' -UploadState 'Uploaded' -ErrorText 'keep-me'
                Save-ProcessedBitrixFile -BitrixTaskId 501 -FileKey 'rfq:preserve' -UploadState 'Skipped'
                $preservedRfqFile = Get-ProcessedBitrixFile 501 'rfq:preserve'
                Assert-Equal $preservedRfqFile.bitrix_file_id 'f-9' 'Expected partial processed RFQ update to preserve the Bitrix file id.'
                Assert-Equal $preservedRfqFile.file_name 'RFQ_keep.xlsx' 'Expected partial processed RFQ update to preserve the file name.'
                Assert-Equal $preservedRfqFile.file_hash 'hash-keep' 'Expected partial processed RFQ update to preserve the file hash.'
                Assert-Equal $preservedRfqFile.local_path 'data\purchase_control\files\AE26081801PK-000001\RRFQ\RFQ_keep.xlsx' 'Expected partial processed RFQ update to preserve the local path.'
                Assert-Equal $preservedRfqFile.upload_state 'Skipped' 'Expected partial processed RFQ update to change only the upload state.'
                Assert-Equal $preservedRfqFile.error_text 'keep-me' 'Expected partial processed RFQ update to preserve the error text.'
            }
        } finally {
            Remove-TestAppRoot $appRoot
        }
    }
}

function New-TaskSyncApiFixture {
    param(
        [int]$CreatedDealId = 7101,
        [int]$ExistingDealId = 0
    )

    $fixture = [pscustomobject]@{
        CreatedDealId = $CreatedDealId
        ExistingDealId = $ExistingDealId
        CreatedDealCalls = 0
        ListDealCalls = 0
        TaskListCalls = 0
        ChecklistCalls = 0
    }

    $taskTitle = [string]::Concat(
        'AE26081801PK-000001/Acme/Controller board - 5',
        [char]0x0448,
        [char]0x0442,
        '.'
    )

    $apiInvoker = {
        param($Method, $Params, $TimeoutSec)

        switch ($Method) {
            'tasks.task.list' {
                $fixture.TaskListCalls++
                return [pscustomobject]@{
                    result = [pscustomobject]@{
                        tasks = @(
                            [pscustomobject]@{
                                id = 3001
                                title = $taskTitle
                                status = 3
                                deadline = '2026-08-21T10:00:00+03:00'
                                groupId = 27
                            }
                        )
                    }
                }
            }
            'task.checklistitem.getlist' {
                $fixture.ChecklistCalls++
                return [pscustomobject]@{
                    result = @()
                }
            }
            'crm.item.list' {
                $fixture.ListDealCalls++
                if ($fixture.ExistingDealId -gt 0) {
                    return [pscustomobject]@{
                        result = [pscustomobject]@{
                            items = @(
                                [pscustomobject]@{
                                    id = $fixture.ExistingDealId
                                    title = $taskTitle
                                }
                            )
                        }
                    }
                }

                return [pscustomobject]@{
                    result = [pscustomobject]@{
                        items = @()
                    }
                }
            }
            'crm.item.add' {
                $fixture.CreatedDealCalls++
                return [pscustomobject]@{
                    result = [pscustomobject]@{
                        item = [pscustomobject]@{
                            id = $fixture.CreatedDealId
                        }
                    }
                }
            }
            default {
                throw "Unexpected Bitrix method: $Method"
            }
        }
    }.GetNewClosure()

    return [pscustomobject]@{
        State = $fixture
        ApiInvoker = $apiInvoker
    }
}

function Get-DealCount {
    return [int](Invoke-PurchaseScalar 'SELECT COUNT(*) FROM deals' @{})
}

function Get-TaskLinkCount {
    return [int](Invoke-PurchaseScalar 'SELECT COUNT(*) FROM bitrix_task_links' @{})
}

function Get-DealRowByNumber {
    param([string]$DealNumber)

    $table = Invoke-PurchaseQuery @'
SELECT
    id,
    deal_number,
    client,
    status,
    tracking_status,
    IFNULL(priority, '') AS priority,
    IFNULL(title, '') AS title,
    bitrix_deal_id
FROM deals
WHERE deal_number = @deal_number
'@ @{
        '@deal_number' = $DealNumber
    }

    if ($table.Rows.Count -eq 0) {
        return $null
    }

    return $table.Rows[0]
}

function Run-TaskSyncTests {
    Invoke-TestCase 'imports a new Bitrix task into one local deal and one Bitrix deal' {
        $appRoot = New-TestAppRoot
        try {
            Use-TestStore $appRoot {
                Initialize-PurchaseStore
                Set-BitrixConfig -BaseUrl 'unirec.bitrix24.ru/' -WebhookUserId 165 -WebhookToken 'secret' -ProjectId 27 -PollIntervalMinutes 10
                Set-BitrixConfig -CrmEntityTypeId 2 -DealCategoryId 9 -DealStageId 'C9:NEW'
                $fixture = New-TaskSyncApiFixture -CreatedDealId 7101

                $summary = Sync-BitrixTasks -ApiInvoker $fixture.ApiInvoker

                Assert-Equal $summary.Imported 1 'Expected one imported procurement.'
                Assert-Equal $summary.Linked 1 'Expected one linked Bitrix task.'
                Assert-Equal $summary.Skipped 0 'Expected no skipped tasks.'
                Assert-Equal $summary.Failed 0 'Expected no failed tasks.'
                Assert-Equal (Get-DealCount) 1 'Expected exactly one local procurement row.'
                Assert-Equal (Get-TaskLinkCount) 1 'Expected exactly one Bitrix task link row.'

                $deal = Get-DealRowByNumber 'AE26081801PK-000001'
                $expectedTrackingStatus = [string]::Concat([char]0x0412, [char]0x0020, [char]0x0440, [char]0x0430, [char]0x0431, [char]0x043E, [char]0x0442, [char]0x0435)
                Assert-True ($null -ne $deal) 'Expected the imported procurement row.'
                Assert-Equal $deal.client 'Acme' 'Expected the imported client.'
                Assert-Equal $deal.status 'RFQ' 'Expected imported stage RFQ.'
                Assert-Equal $deal.tracking_status $expectedTrackingStatus 'Expected imported local status В работе.'
                Assert-Equal $deal.priority '' 'Expected imported priority to stay empty.'
                Assert-Equal $deal.title '' 'Expected no local title to be created from the task title.'
                Assert-Equal $deal.bitrix_deal_id 7101 'Expected the local procurement to store the Bitrix deal id.'

                $link = Get-BitrixTaskLink 3001
                Assert-Equal $link.deal_id $deal.id 'Expected the Bitrix task link to point at the local procurement.'
                Assert-Equal $link.bitrix_deal_id 7101 'Expected the task link to store the Bitrix deal id.'
                Assert-Equal $link.bitrix_status 3 'Expected the task link to store the Bitrix task status.'
                Assert-Equal $fixture.State.CreatedDealCalls 1 'Expected one Bitrix deal creation call.'
            }
        } finally {
            Remove-TestAppRoot $appRoot
        }
    }

    Invoke-TestCase 'reuses an existing local procurement and existing Bitrix deal by deal number' {
        $appRoot = New-TestAppRoot
        try {
            Use-TestStore $appRoot {
                Initialize-PurchaseStore
                Set-BitrixConfig -BaseUrl 'unirec.bitrix24.ru/' -WebhookUserId 165 -WebhookToken 'secret' -ProjectId 27 -PollIntervalMinutes 10
                Set-BitrixConfig -CrmEntityTypeId 2 -DealCategoryId 9 -DealStageId 'C9:NEW'

                $existingDealId = New-PurchaseDeal 'AE26081801PK-000001' '' '' 'Existing Client' 'RFQ'
                $fixture = New-TaskSyncApiFixture -ExistingDealId 7202

                $summary = Sync-BitrixTasks -ApiInvoker $fixture.ApiInvoker

                Assert-Equal $summary.Imported 0 'Expected the existing procurement to be reused.'
                Assert-Equal $summary.Linked 1 'Expected the task to be linked.'
                Assert-Equal (Get-DealCount) 1 'Expected the local deal count to stay at one.'

                $deal = Get-DealRowByNumber 'AE26081801PK-000001'
                Assert-Equal $deal.id $existingDealId 'Expected the same local procurement row to be reused.'
                Assert-Equal $deal.bitrix_deal_id 7202 'Expected the existing Bitrix deal id to be stored locally.'
                Assert-Equal $fixture.State.CreatedDealCalls 0 'Expected no new Bitrix deal creation call.'
                Assert-Equal $fixture.State.ListDealCalls 1 'Expected a Bitrix deal lookup.'

                $link = Get-BitrixTaskLink 3001
                Assert-Equal $link.deal_id $existingDealId 'Expected the task link to reuse the existing procurement.'
                Assert-Equal $link.bitrix_deal_id 7202 'Expected the task link to store the reused Bitrix deal id.'
            }
        } finally {
            Remove-TestAppRoot $appRoot
        }
    }

    Invoke-TestCase 'keeps duplicate reruns idempotent' {
        $appRoot = New-TestAppRoot
        try {
            Use-TestStore $appRoot {
                Initialize-PurchaseStore
                Set-BitrixConfig -BaseUrl 'unirec.bitrix24.ru/' -WebhookUserId 165 -WebhookToken 'secret' -ProjectId 27 -PollIntervalMinutes 10
                Set-BitrixConfig -CrmEntityTypeId 2 -DealCategoryId 9 -DealStageId 'C9:NEW'
                $fixture = New-TaskSyncApiFixture -CreatedDealId 7303

                $first = Sync-BitrixTasks -ApiInvoker $fixture.ApiInvoker
                $second = Sync-BitrixTasks -ApiInvoker $fixture.ApiInvoker

                Assert-Equal $first.Imported 1 'Expected the first run to import the procurement.'
                Assert-Equal $second.Imported 0 'Expected the second run not to import the procurement again.'
                Assert-Equal (Get-DealCount) 1 'Expected one local procurement after two sync runs.'
                Assert-Equal (Get-TaskLinkCount) 1 'Expected one task link after two sync runs.'
                Assert-Equal $fixture.State.CreatedDealCalls 1 'Expected only one Bitrix deal creation across reruns.'
            }
        } finally {
            Remove-TestAppRoot $appRoot
        }
    }

    Invoke-TestCase 'skips a manually deleted Bitrix tombstone without recreating anything' {
        $appRoot = New-TestAppRoot
        try {
            Use-TestStore $appRoot {
                Initialize-PurchaseStore
                Set-BitrixConfig -BaseUrl 'unirec.bitrix24.ru/' -WebhookUserId 165 -WebhookToken 'secret' -ProjectId 27 -PollIntervalMinutes 10
                Set-BitrixConfig -CrmEntityTypeId 2 -DealCategoryId 9 -DealStageId 'C9:NEW'

                $deletedDealId = New-PurchaseDeal 'AE26081801PK-000001'
                Save-BitrixTaskLink -BitrixTaskId 3001 -DealId $deletedDealId -BitrixDealId 7404 -BitrixStatus 3
                Delete-PurchaseDeal $deletedDealId

                $fixture = New-TaskSyncApiFixture -CreatedDealId 7405
                $summary = Sync-BitrixTasks -ApiInvoker $fixture.ApiInvoker

                Assert-Equal $summary.Imported 0 'Expected no procurement import for a deleted tombstone.'
                Assert-Equal $summary.Linked 0 'Expected no task link creation for a deleted tombstone.'
                Assert-Equal $summary.Skipped 1 'Expected the deleted tombstone to be skipped.'
                Assert-Equal (Get-DealCount) 0 'Expected no local procurement row after skip.'
                Assert-Equal $fixture.State.CreatedDealCalls 0 'Expected no Bitrix deal creation for a deleted tombstone.'

                $link = Get-BitrixTaskLink 3001
                Assert-True ($null -ne $link) 'Expected the deleted tombstone row to remain.'
                Assert-Equal $link.is_manual_deleted 1 'Expected the deleted tombstone row to stay flagged as manually deleted.'
                Assert-Equal $link.bitrix_deal_id 7404 'Expected the deleted tombstone row to preserve the Bitrix deal id.'
            }
        } finally {
            Remove-TestAppRoot $appRoot
        }
    }
}

if ($TestName -eq 'All' -or $TestName -eq 'Schema') {
    Run-SchemaTests
}

if ($TestName -eq 'All' -or $TestName -eq 'TaskSync') {
    Run-TaskSyncTests
}

Write-Host ("Passed: {0}  Failed: {1}" -f $script:FixtureResults.Passed, $script:FixtureResults.Failed)
if ($script:FixtureResults.Failed -gt 0) {
    exit 1
}
