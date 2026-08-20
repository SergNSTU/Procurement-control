param(
    [string]$TestName = 'All',
    [switch]$All
)

$ErrorActionPreference = 'Stop'

$script:TestResults = [ordered]@{
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

function Assert-Throws {
    param(
        [scriptblock]$Script,
        [string]$Message
    )

    $thrown = $false
    try {
        & $Script
    } catch {
        $thrown = $true
    }

    if (-not $thrown) {
        throw $Message
    }
}

function New-TestTempDirectory {
    $path = Join-Path ([IO.Path]::GetTempPath()) ('bitrix-rfq-tests-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $path | Out-Null
    return $path
}

function Invoke-TestCase {
    param(
        [string]$Name,
        [scriptblock]$Script
    )

    try {
        & $Script
        $script:TestResults.Passed++
        Write-Host "[PASS] $Name"
    } catch {
        $script:TestResults.Failed++
        Write-Host "[FAIL] $Name"
        Write-Host $_.Exception.Message
    }
}

$script:BitrixTestSettings = @{}
function Get-PurchaseSetting {
    param([string]$Key)
    if ($script:BitrixTestSettings.ContainsKey($Key)) {
        return [string]$script:BitrixTestSettings[$Key]
    }

    return $null
}

function Set-PurchaseSetting {
    param([string]$Key, [string]$Value)
    $script:BitrixTestSettings[$Key] = $Value
}

. (Join-Path $PSScriptRoot '..\BitrixIntegration.ps1')

$script:RfqDocumentsRoot = ''
$script:RfqProcessedFiles = @{}

function Reset-RfqTestState {
    $script:RfqProcessedFiles = @{}
}

function Get-PurchaseDocumentsRoot {
    if ([string]::IsNullOrWhiteSpace($script:RfqDocumentsRoot)) {
        $script:RfqDocumentsRoot = New-TestTempDirectory
    }

    return $script:RfqDocumentsRoot
}

function Get-ProcessedBitrixFile {
    param(
        [int]$BitrixTaskId,
        [string]$FileKey
    )

    $key = '{0}|{1}' -f $BitrixTaskId, $FileKey
    if ($script:RfqProcessedFiles.ContainsKey($key)) {
        return $script:RfqProcessedFiles[$key]
    }

    return $null
}

function Save-ProcessedBitrixFile {
    param(
        [int]$BitrixTaskId,
        [string]$FileKey,
        [AllowNull()][string]$BitrixFileId = $null,
        [AllowNull()][string]$FileName = $null,
        [AllowNull()][string]$FileHash = $null,
        [AllowNull()][string]$LocalPath = $null,
        [AllowNull()][string]$UploadState = $null,
        [AllowNull()][string]$ErrorText = $null
    )

    $key = '{0}|{1}' -f $BitrixTaskId, $FileKey
    $existing = if ($script:RfqProcessedFiles.ContainsKey($key)) { $script:RfqProcessedFiles[$key] } else { [pscustomobject]@{} }
    $record = [ordered]@{
        bitrix_task_id = $BitrixTaskId
        file_key = $FileKey
        bitrix_file_id = $existing.bitrix_file_id
        file_name = $existing.file_name
        file_hash = $existing.file_hash
        local_path = $existing.local_path
        upload_state = $existing.upload_state
        error_text = $existing.error_text
    }

    if ($PSBoundParameters.ContainsKey('BitrixFileId')) { $record.bitrix_file_id = $BitrixFileId }
    if ($PSBoundParameters.ContainsKey('FileName')) { $record.file_name = $FileName }
    if ($PSBoundParameters.ContainsKey('FileHash')) { $record.file_hash = $FileHash }
    if ($PSBoundParameters.ContainsKey('LocalPath')) { $record.local_path = $LocalPath }
    if ($PSBoundParameters.ContainsKey('UploadState')) { $record.upload_state = $UploadState }
    if ($PSBoundParameters.ContainsKey('ErrorText')) { $record.error_text = $ErrorText }

    $script:RfqProcessedFiles[$key] = [pscustomobject]$record
    return $script:RfqProcessedFiles[$key]
}

function Run-ConfigurationTests {
    Invoke-TestCase 'builds a webhook endpoint without adding a second slash' {
        $uri = New-BitrixMethodUri 'https://unirec.bitrix24.ru/' 165 'secret' 'tasks.task.list'
        Assert-Equal $uri 'https://unirec.bitrix24.ru/rest/165/secret/tasks.task.list' 'Unexpected Bitrix endpoint.'
    }

    Invoke-TestCase 'rejects a missing token before making a request' {
        Assert-Throws {
            New-BitrixMethodUri 'https://unirec.bitrix24.ru' 165 '' 'tasks.task.list'
        } 'Expected a missing-token error.'
    }

    Invoke-TestCase 'uses injected transport and round-trips Bitrix config' {
        Set-BitrixConfig -BaseUrl 'unirec.bitrix24.ru/' -WebhookUserId 165 -WebhookToken 'secret' -ProjectId 27 -PollIntervalMinutes 10
        $script:captured = $null
        $result = Invoke-BitrixMethod -Method 'tasks.task.list' -Params @{ 'filter[GROUP_ID]' = 27 } -Transport {
            param($Uri, $JsonBody, $TimeoutSec)
            $script:captured = [pscustomobject]@{
                Uri = $Uri
                Body = $JsonBody
                TimeoutSec = $TimeoutSec
            }
            return [pscustomobject]@{ result = @(@{ id = 1 }) }
        }

        Assert-Equal $script:captured.Uri 'https://unirec.bitrix24.ru/rest/165/secret/tasks.task.list' 'Transport received the wrong URI.'
        Assert-True ($script:captured.Body -match '"filter\[GROUP_ID\]":27') 'Transport did not receive the JSON body.'
        Assert-Equal $script:captured.TimeoutSec 30 'Transport received the wrong timeout.'
        Assert-Equal $result.result[0].id 1 'Invoke-BitrixMethod did not return the transport response.'
    }

    Invoke-TestCase 'keeps the final Bitrix task filter exact in the page handler' {
        $source = Get-Content -Encoding Default (Join-Path $PSScriptRoot '..\RRFQComparer.ps1')
        Assert-True ([bool]($source | Select-String -SimpleMatch 'GROUP_ID = $bitrixConfig.ProjectId')) 'Missing configured GROUP_ID filter.'
        Assert-True ([bool]($source | Select-String -SimpleMatch "'>=REAL_STATUS' = 1")) 'Missing lower status bound.'
        Assert-True ([bool]($source | Select-String -SimpleMatch "'<=REAL_STATUS' = 6")) 'Missing upper status bound.'
        Assert-True ([bool]($source | Select-String -SimpleMatch '$blockedTaskIds = @(Get-BitrixBlockedTaskIds)')) 'Missing blocked-task exclusion.'
        Assert-True ([bool]($source | Select-String -SimpleMatch '$allowedStatuses = @(1, 2, 3, 6)')) 'Missing final local allow-list.'
        Assert-True ([bool]($source | Select-String -SimpleMatch '$tasks = @($tasks | Where-Object { $allowedStatuses -contains [int]$_.status })')) 'Missing final local allow-list filter.'
        Assert-True ([bool]($source | Select-String -SimpleMatch 'if (-not $rfqCompleted) {')) 'Missing completed-RFQ exclusion.'
    }
}

function Run-ParsingTests {
    Invoke-TestCase 'parses the planned Bitrix task title into procurement fields' {
        $parsed = Convert-BitrixTaskTitleToProcurement 'AE26081703PK-001949/БОРЕЦ-НЭО, ООО/плата управления УВФК.112.15.02.02 - 29шт.'

        Assert-Equal $parsed.RequestNumber 'AE26081703PK-001949' 'Unexpected request number.'
        Assert-Equal $parsed.Client 'БОРЕЦ-НЭО, ООО' 'Unexpected client.'
        Assert-Equal $parsed.ItemDescription 'плата управления УВФК.112.15.02.02' 'Unexpected item description.'
        Assert-Equal $parsed.Quantity 29 'Unexpected quantity.'
        Assert-Equal $parsed.OriginalTitle 'AE26081703PK-001949/БОРЕЦ-НЭО, ООО/плата управления УВФК.112.15.02.02 - 29шт.' 'Unexpected original title.'
    }

    Invoke-TestCase 'parses a Bitrix task title with extra spaces' {
        $parsed = Convert-BitrixTaskTitleToProcurement '  AE26081703PK-001949 /  БОРЕЦ-НЭО, ООО / плата управления УВФК.112.15.02.02 - 29 шт.  '

        Assert-Equal $parsed.RequestNumber 'AE26081703PK-001949' 'Unexpected request number.'
        Assert-Equal $parsed.Client 'БОРЕЦ-НЭО, ООО' 'Unexpected client.'
        Assert-Equal $parsed.ItemDescription 'плата управления УВФК.112.15.02.02' 'Unexpected item description.'
        Assert-Equal $parsed.Quantity 29 'Unexpected quantity.'
    }

    Invoke-TestCase 'parses a Bitrix task title with an extra hyphen in the description' {
        $parsed = Convert-BitrixTaskTitleToProcurement 'AE26081703PK-001949/БОРЕЦ-НЭО, ООО/плата управления - модуль УВФК.112.15.02.02 - 29шт.'

        Assert-Equal $parsed.RequestNumber 'AE26081703PK-001949' 'Unexpected request number.'
        Assert-Equal $parsed.Client 'БОРЕЦ-НЭО, ООО' 'Unexpected client.'
        Assert-Equal $parsed.ItemDescription 'плата управления - модуль УВФК.112.15.02.02' 'Unexpected item description.'
        Assert-Equal $parsed.Quantity 29 'Unexpected quantity.'
    }

    Invoke-TestCase 'rejects a Bitrix task title missing the first segment' {
        Assert-Throws {
            Convert-BitrixTaskTitleToProcurement '/БОРЕЦ-НЭО, ООО/плата управления УВФК.112.15.02.02 - 29шт.'
        } 'Expected a validation error for a missing first segment.'
    }

    Invoke-TestCase 'matches an RFQ filename regardless of prefix' {
        $match = Test-RfqFilenameMatchesDeal 'SC26081202VNV-082372' 'Monthly export RFQ_SC26081202VNV-082372.xlsx'

        Assert-True $match.Matches 'Expected the filename to match the deal number.'
        Assert-Equal $match.NormalizedDealNumber 'SC26081202VNV-082372' 'Unexpected normalized deal number.'
        Assert-Equal $match.FileName 'Monthly export RFQ_SC26081202VNV-082372.xlsx' 'Unexpected file name.'
    }

    Invoke-TestCase 'rejects an RFQ filename for a different deal number' {
        $match = Test-RfqFilenameMatchesDeal 'SC26081202VNV-082372' 'Monthly export RFQ_AE26081703PK-001949.xlsx'

        Assert-True (-not $match.Matches) 'Expected the filename not to match the deal number.'
        Assert-Equal $match.NormalizedDealNumber 'SC26081202VNV-082372' 'Unexpected normalized deal number.'
        Assert-Equal $match.FileName 'Monthly export RFQ_AE26081703PK-001949.xlsx' 'Unexpected file name.'
    }
}

function Run-TaskRetrievalTests {
    Invoke-TestCase 'normalizes only the final allowed Bitrix task set' {
        Set-BitrixConfig -BaseUrl 'unirec.bitrix24.ru/' -WebhookUserId 165 -WebhookToken 'secret' -ProjectId 27 -PollIntervalMinutes 10

        function Get-BitrixBlockedTaskIds { @(2002) }

        $calls = New-Object System.Collections.Generic.List[string]
        $tasks = @(Get-BitrixNewProjectTasks -ApiInvoker {
            param($Method, $Params, $TimeoutSec)
            [void]$calls.Add($Method)
            switch ($Method) {
                'tasks.task.list' {
                    return [pscustomobject]@{
                        result = [pscustomobject]@{
                            tasks = @(
                                [pscustomobject]@{ id = 2001; title = 'AE26081801PK-000001/Client/Board - 5шт.'; status = 3; deadline = '2026-08-21T10:00:00+03:00'; groupId = 27 },
                                [pscustomobject]@{ id = 2002; title = 'AE26081801PK-000002/Client/Board - 7шт.'; status = 3; deadline = '2026-08-22T10:00:00+03:00'; groupId = 27 },
                                [pscustomobject]@{ id = 2003; title = 'AE26081801PK-000003/Client/Board - 9шт.'; status = 4; deadline = '2026-08-23T10:00:00+03:00'; groupId = 27 },
                                [pscustomobject]@{ id = 2004; title = 'AE26081801PK-000004/Client/Board - 11шт.'; status = 6; deadline = '2026-08-24T10:00:00+03:00'; groupId = 27 }
                            )
                        }
                    }
                }
                'task.checklistitem.getlist' {
                    if ([int]$Params.TASKID -eq 2004) {
                        return [pscustomobject]@{
                            result = @(
                                [pscustomobject]@{
                                    TITLE = 'Заполнение сводного RFQ и выделение компонентов на расчет (Сенько Е., Сидоров М.)'
                                    IS_COMPLETE = 'Y'
                                }
                            )
                        }
                    }

                    return [pscustomobject]@{
                        result = @(
                            [pscustomobject]@{
                                TITLE = 'Заполнение сводного RFQ и выделение компонентов на расчет (Сенько Е., Сидоров М.)'
                                IS_COMPLETE = 'N'
                            }
                        )
                    }
                }
                default {
                    throw "Unexpected method: $Method"
                }
            }
        })

        Assert-Equal $tasks.Count 1 'Expected only one task after final filtering.'
        Assert-Equal $tasks[0].Id 2001 'Expected the remaining task to keep its Bitrix ID.'
        Assert-Equal $tasks[0].Title 'AE26081801PK-000001/Client/Board - 5шт.' 'Unexpected task title.'
        Assert-Equal $tasks[0].Status 3 'Unexpected normalized status.'
        Assert-Equal $tasks[0].Deadline '2026-08-21' 'Unexpected normalized deadline.'
        Assert-Equal $tasks[0].GroupId 27 'Unexpected normalized group id.'
        Assert-True ($null -ne $tasks[0].Raw) 'Expected the normalized task to keep the raw payload.'
        Assert-True ($calls.Contains('tasks.task.list')) 'Expected task retrieval to call tasks.task.list.'
        Assert-True ($calls.Contains('task.checklistitem.getlist')) 'Expected task retrieval to inspect checklist completion.'

        Remove-Item Function:\Get-BitrixBlockedTaskIds -ErrorAction SilentlyContinue
    }
}

function Run-RfqTests {
    Invoke-TestCase 'returns null when the exact RFQ checklist item is missing' {
        $item = Get-BitrixChecklistItem -TaskId 900 -Title (Get-BitrixRfqChecklistTitle) -ApiInvoker {
            param($Method, $Params, $TimeoutSec)
            return [pscustomobject]@{
                result = @(
                    [pscustomobject]@{
                        ID = '11'
                        TITLE = 'Other checklist item'
                    }
                )
            }
        }

        Assert-True ($null -eq $item) 'Expected no checklist item match.'
    }

    Invoke-TestCase 'returns Waiting when the RFQ checklist item has no files yet' {
        Reset-RfqTestState
        $task = [pscustomobject]@{ Id = 901; Title = 'AE26081801PK-000001/Client/Board - 5шт.' }
        $procurement = [pscustomobject]@{ id = 17; deal_number = 'AE26081801PK-000001'; bitrix_deal_id = 880 }

        $result = Sync-BitrixRfq -Task $task -Procurement $procurement -ApiInvoker {
            param($Method, $Params, $TimeoutSec)
            return [pscustomobject]@{
                result = @(
                    [pscustomobject]@{
                        ID = '55'
                        TITLE = Get-BitrixRfqChecklistTitle
                    }
                )
            }
        }

        Assert-Equal $result.Status 'Waiting' 'Expected missing-file checklist to stay waiting.'
        Assert-True ($result.Message -match 'file') 'Expected waiting message to mention the missing file.'
    }

    Invoke-TestCase 'selects the latest checklist file using Drive metadata when attachment timestamps are incomplete' {
        $calls = New-Object System.Collections.Generic.List[string]
        $item = [pscustomobject]@{
            ID = '77'
            TITLE = Get-BitrixRfqChecklistTitle
            ATTACHMENTS = [ordered]@{
                '501' = [pscustomobject]@{
                    ATTACHMENT_ID = '501'
                    FILE_ID = '901'
                    NAME = 'RFQ_SC26081202VNV-082372_older.pdf'
                    DOWNLOAD_URL = 'https://example.test/older'
                    UPDATE_TIME = '2026-08-17T10:00:00+03:00'
                }
                '502' = [pscustomobject]@{
                    ATTACHMENT_ID = '502'
                    FILE_ID = '902'
                    NAME = 'Monthly export SC26081202VNV-082372'
                    DOWNLOAD_URL = 'https://example.test/newer'
                }
            }
        }

        $selected = Select-LatestBitrixChecklistFile -ChecklistItem $item -ApiInvoker {
            param($Method, $Params, $TimeoutSec)
            [void]$calls.Add($Method)
            if ($Method -ne 'disk.file.get') {
                throw "Unexpected method: $Method"
            }

            Assert-Equal $Params.id '902' 'Expected Drive lookup for the timestamp-less attachment.'
            return [pscustomobject]@{
                result = [pscustomobject]@{
                    ID = '902'
                    NAME = 'Monthly export SC26081202VNV-082372'
                    DOWNLOAD_URL = 'https://example.test/newer-from-drive'
                    UPDATE_TIME = '2026-08-17T15:45:00+03:00'
                    CREATE_TIME = '2026-08-17T15:40:00+03:00'
                }
            }
        }

        Assert-Equal $selected.FileId '902' 'Expected the latest file to come from Drive metadata.'
        Assert-Equal $selected.Name 'Monthly export SC26081202VNV-082372' 'Unexpected selected file name.'
        Assert-Equal $selected.DownloadUrl 'https://example.test/newer-from-drive' 'Expected Drive metadata download URL to win.'
        Assert-Equal $selected.ModifiedAt '2026-08-17T15:45:00+03:00' 'Unexpected selected modified time.'
        Assert-True ($calls.Contains('disk.file.get')) 'Expected latest-file selection to query Drive metadata.'
    }

    Invoke-TestCase 'uploads the latest matching RFQ and records the processed state' {
        Reset-RfqTestState
        $script:RfqDocumentsRoot = New-TestTempDirectory
        Set-BitrixConfig -BaseUrl 'unirec.bitrix24.ru/' -WebhookUserId 165 -WebhookToken 'secret' -ProjectId 27 -PollIntervalMinutes 10 -RfqFileField 'UF_CRM_RFQ_FILE'

        $task = [pscustomobject]@{ Id = 902; Title = 'SC26081202VNV-082372/Client/Board - 7шт.' }
        $procurement = [pscustomobject]@{ id = 18; deal_number = 'SC26081202VNV-082372'; bitrix_deal_id = 881 }
        $uploadCalls = New-Object System.Collections.Generic.List[object]

        try {
            $result = Sync-BitrixRfq -Task $task -Procurement $procurement -ApiInvoker {
                param($Method, $Params, $TimeoutSec)
                switch ($Method) {
                    'task.checklistitem.getlist' {
                        return [pscustomobject]@{
                            result = @(
                                [pscustomobject]@{
                                    ID = '88'
                                    TITLE = Get-BitrixRfqChecklistTitle
                                    ATTACHMENTS = [ordered]@{
                                        '601' = [pscustomobject]@{
                                            ATTACHMENT_ID = '601'
                                            FILE_ID = '9901'
                                            NAME = 'Monthly export RFQ_SC26081202VNV-082372.final'
                                            DOWNLOAD_URL = 'https://example.test/rfq'
                                            UPDATE_TIME = '2026-08-18T10:15:00+03:00'
                                        }
                                    }
                                }
                            )
                        }
                    }
                    default {
                        throw "Unexpected method: $Method"
                    }
                }
            } -Downloader {
                param($File, $TargetPath)
                [IO.File]::WriteAllBytes($TargetPath, [byte[]](65, 66, 67, 68))
            } -UploadTransport {
                param($Method, $Params, $TimeoutSec)
                [void]$uploadCalls.Add([pscustomobject]@{
                    Method = $Method
                    Params = $Params
                    TimeoutSec = $TimeoutSec
                })
                return [pscustomobject]@{
                    result = [pscustomobject]@{
                        item = [pscustomobject]@{ id = 881 }
                    }
                }
            }

            Assert-Equal $result.Status 'Uploaded' 'Expected matching RFQ to upload.'
            Assert-Equal $uploadCalls.Count 1 'Expected exactly one upload call.'
            Assert-Equal $uploadCalls[0].Method 'crm.item.update' 'Expected CRM item update for upload.'
            Assert-Equal $uploadCalls[0].Params.fields['UF_CRM_RFQ_FILE'][0] 'Monthly export RFQ_SC26081202VNV-082372.final' 'Expected upload to preserve the original RFQ file name without assuming .xlsx.'

            $saved = Get-ProcessedBitrixFile -BitrixTaskId 902 -FileKey 'rfq:88:9901'
            Assert-True ($null -ne $saved) 'Expected uploaded RFQ state to persist.'
            Assert-Equal $saved.upload_state 'Uploaded' 'Expected uploaded RFQ state to be recorded.'
            Assert-Equal $saved.file_name 'Monthly export RFQ_SC26081202VNV-082372.final' 'Unexpected saved file name.'
            Assert-True (-not [string]::IsNullOrWhiteSpace($saved.file_hash)) 'Expected uploaded RFQ hash to be recorded.'
            Assert-True ([IO.File]::Exists($saved.local_path)) 'Expected downloaded RFQ file to exist locally.'
        } finally {
            Remove-Item -LiteralPath $script:RfqDocumentsRoot -Recurse -Force -ErrorAction SilentlyContinue
            $script:RfqDocumentsRoot = ''
        }
    }

    Invoke-TestCase 'blocks mismatched RFQ upload until confirmation' {
        Reset-RfqTestState
        $script:RfqDocumentsRoot = New-TestTempDirectory
        $task = [pscustomobject]@{ Id = 903; Title = 'SC26081202VNV-082372/Client/Board - 7шт.' }
        $procurement = [pscustomobject]@{ id = 19; deal_number = 'SC26081202VNV-082372'; bitrix_deal_id = 882 }
        $uploadCalls = New-Object System.Collections.Generic.List[object]

        try {
            $result = Sync-BitrixRfq -Task $task -Procurement $procurement -ApiInvoker {
                param($Method, $Params, $TimeoutSec)
                return [pscustomobject]@{
                    result = @(
                        [pscustomobject]@{
                            ID = '89'
                            TITLE = Get-BitrixRfqChecklistTitle
                            ATTACHMENTS = [ordered]@{
                                '602' = [pscustomobject]@{
                                    ATTACHMENT_ID = '602'
                                    FILE_ID = '9902'
                                    NAME = 'Monthly export RFQ_AE26081703PK-001949.final'
                                    DOWNLOAD_URL = 'https://example.test/mismatch'
                                    UPDATE_TIME = '2026-08-18T10:20:00+03:00'
                                }
                            }
                        }
                    )
                }
            } -Downloader {
                param($File, $TargetPath)
                [IO.File]::WriteAllBytes($TargetPath, [byte[]](88, 89, 90))
            } -UploadTransport {
                param($Method, $Params, $TimeoutSec)
                $uploadCalled = $true
                throw 'Upload should not be called for a mismatch without confirmation.'
            }

            Assert-Equal $result.Status 'Waiting' 'Expected mismatch to stay waiting for confirmation.'
            Assert-True ($result.Message -match 'SC26081202VNV-082372') 'Expected mismatch message to mention the deal number.'
            Assert-True ($result.Message -match 'AE26081703PK-001949') 'Expected mismatch message to mention the actual file name.'
            Assert-True (-not $uploadCalled) 'Expected mismatch to block upload.'

            $saved = Get-ProcessedBitrixFile -BitrixTaskId 903 -FileKey 'rfq:89:9902'
            Assert-Equal $saved.upload_state 'Waiting' 'Expected mismatch state to be recorded as waiting.'
        } finally {
            Remove-Item -LiteralPath $script:RfqDocumentsRoot -Recurse -Force -ErrorAction SilentlyContinue
            $script:RfqDocumentsRoot = ''
        }
    }

    Invoke-TestCase 'allows a mismatched RFQ upload when confirmation is provided' {
        Reset-RfqTestState
        $script:RfqDocumentsRoot = New-TestTempDirectory
        Set-BitrixConfig -BaseUrl 'unirec.bitrix24.ru/' -WebhookUserId 165 -WebhookToken 'secret' -ProjectId 27 -PollIntervalMinutes 10 -RfqFileField 'UF_CRM_RFQ_FILE'

        $task = [pscustomobject]@{ Id = 904; Title = 'SC26081202VNV-082372/Client/Board - 7шт.' }
        $procurement = [pscustomobject]@{ id = 20; deal_number = 'SC26081202VNV-082372'; bitrix_deal_id = 883 }
        $uploadCalls = New-Object System.Collections.Generic.List[object]

        try {
            $result = Sync-BitrixRfq -Task $task -Procurement $procurement -ConfirmMismatch -ApiInvoker {
                param($Method, $Params, $TimeoutSec)
                return [pscustomobject]@{
                    result = @(
                        [pscustomobject]@{
                            ID = '90'
                            TITLE = Get-BitrixRfqChecklistTitle
                            ATTACHMENTS = [ordered]@{
                                '603' = [pscustomobject]@{
                                    ATTACHMENT_ID = '603'
                                    FILE_ID = '9903'
                                    NAME = 'Monthly export RFQ_AE26081703PK-001949.final'
                                    DOWNLOAD_URL = 'https://example.test/override'
                                    UPDATE_TIME = '2026-08-18T10:25:00+03:00'
                                }
                            }
                        }
                    )
                }
            } -Downloader {
                param($File, $TargetPath)
                [IO.File]::WriteAllBytes($TargetPath, [byte[]](70, 71, 72))
            } -UploadTransport {
                param($Method, $Params, $TimeoutSec)
                [void]$uploadCalls.Add([pscustomobject]@{
                    Method = $Method
                    Params = $Params
                    TimeoutSec = $TimeoutSec
                })
                return [pscustomobject]@{
                    result = [pscustomobject]@{
                        item = [pscustomobject]@{ id = 883 }
                    }
                }
            }

            Assert-Equal $result.Status 'Uploaded' 'Expected mismatch confirmation to allow upload.'
            Assert-Equal $uploadCalls.Count 1 'Expected exactly one upload after mismatch confirmation.'
            Assert-Equal $uploadCalls[0].Method 'crm.item.update' 'Expected mismatch confirmation to use CRM item update.'

            $saved = Get-ProcessedBitrixFile -BitrixTaskId 904 -FileKey 'rfq:90:9903'
            Assert-True ($saved.error_text -match 'override') 'Expected mismatch override note in processed state.'
        } finally {
            Remove-Item -LiteralPath $script:RfqDocumentsRoot -Recurse -Force -ErrorAction SilentlyContinue
            $script:RfqDocumentsRoot = ''
        }
    }

    Invoke-TestCase 'skips a previously uploaded RFQ file idempotently' {
        Reset-RfqTestState
        Save-ProcessedBitrixFile -BitrixTaskId 905 -FileKey 'rfq:91:9904' -BitrixFileId '9904' -FileName 'RFQ_SC26081202VNV-082372.final' -FileHash 'existing-hash' -LocalPath 'C:\temp\RFQ_SC26081202VNV-082372.final' -UploadState 'Uploaded' -ErrorText ''

        $task = [pscustomobject]@{ Id = 905; Title = 'SC26081202VNV-082372/Client/Board - 7шт.' }
        $procurement = [pscustomobject]@{ id = 21; deal_number = 'SC26081202VNV-082372'; bitrix_deal_id = 884 }
        $downloadCalled = $false
        $uploadCalled = $false

        $result = Sync-BitrixRfq -Task $task -Procurement $procurement -ApiInvoker {
            param($Method, $Params, $TimeoutSec)
            return [pscustomobject]@{
                result = @(
                    [pscustomobject]@{
                        ID = '91'
                        TITLE = Get-BitrixRfqChecklistTitle
                        ATTACHMENTS = [ordered]@{
                            '604' = [pscustomobject]@{
                                ATTACHMENT_ID = '604'
                                FILE_ID = '9904'
                                NAME = 'RFQ_SC26081202VNV-082372.final'
                                DOWNLOAD_URL = 'https://example.test/already-uploaded'
                                UPDATE_TIME = '2026-08-18T10:30:00+03:00'
                            }
                        }
                    }
                )
            }
        } -Downloader {
            param($File, $TargetPath)
            $downloadCalled = $true
            throw 'Download should not run for an already processed RFQ.'
        } -UploadTransport {
            param($Method, $Params, $TimeoutSec)
            $uploadCalled = $true
            throw 'Upload should not run for an already processed RFQ.'
        }

        Assert-Equal $result.Status 'Skipped' 'Expected already processed RFQ to skip.'
        Assert-True (-not $downloadCalled) 'Expected already processed RFQ to skip download.'
        Assert-True (-not $uploadCalled) 'Expected already processed RFQ to skip upload.'
    }
}

if ($All -or $TestName -eq 'All' -or $TestName -eq 'Configuration') {
    Run-ConfigurationTests
}

if ($All -or $TestName -eq 'All' -or $TestName -eq 'Parsing') {
    Run-ParsingTests
}

if ($All -or $TestName -eq 'All' -or $TestName -eq 'TaskRetrieval') {
    Run-TaskRetrievalTests
}

if ($All -or $TestName -eq 'All' -or $TestName -eq 'Rfq') {
    Run-RfqTests
}

Write-Host ("Passed: {0}  Failed: {1}" -f $script:TestResults.Passed, $script:TestResults.Failed)
if ($script:TestResults.Failed -gt 0) {
    exit 1
}
