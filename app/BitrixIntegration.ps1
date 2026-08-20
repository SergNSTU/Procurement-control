function Get-BitrixLocalSetting {
    param([string]$Key)

    if (Get-Command -Name Get-PurchaseSetting -ErrorAction SilentlyContinue) {
        return [string](Get-PurchaseSetting $Key)
    }

    if (-not $script:BitrixFallbackSettings) {
        $script:BitrixFallbackSettings = @{}
    }

    if ($script:BitrixFallbackSettings.ContainsKey($Key)) {
        return [string]$script:BitrixFallbackSettings[$Key]
    }

    return $null
}

function Set-BitrixLocalSetting {
    param(
        [string]$Key,
        [string]$Value
    )

    if (Get-Command -Name Set-PurchaseSetting -ErrorAction SilentlyContinue) {
        Set-PurchaseSetting $Key $Value
        return
    }

    if (-not $script:BitrixFallbackSettings) {
        $script:BitrixFallbackSettings = @{}
    }

    $script:BitrixFallbackSettings[$Key] = $Value
}

function ConvertTo-BitrixInt {
    param(
        [object]$Value,
        [int]$Default = 0
    )

    if ($null -eq $Value) {
        return $Default
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $Default
    }

    $parsed = 0
    if ([int]::TryParse($text, [ref]$parsed)) {
        return $parsed
    }

    return $Default
}

function Normalize-BitrixBaseUrl {
    param([string]$BaseUrl)

    $text = ([string]$BaseUrl).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    if ($text -notmatch '^[a-z][a-z0-9+\.-]*://') {
        $text = 'https://' + $text
    }

    $uri = $null
    if (-not [Uri]::TryCreate($text, [UriKind]::Absolute, [ref]$uri)) {
        throw 'Bitrix BaseUrl must be a valid absolute URL.'
    }

    if ($uri.Scheme -ne 'https') {
        $builder = [UriBuilder]::new($uri)
        $builder.Scheme = 'https'
        $builder.Port = -1
        $uri = $builder.Uri
    }

    return $uri.AbsoluteUri.TrimEnd('/')
}

function Get-BitrixConfig {
    # Bitrix credentials are stored only in the local SQLite settings store.
    $baseUrlSetting = Get-BitrixLocalSetting 'bitrix_base_url'
    $userIdSetting = Get-BitrixLocalSetting 'bitrix_webhook_user_id'
    $tokenSetting = Get-BitrixLocalSetting 'bitrix_webhook_token'
    $baseUrl = if ([string]::IsNullOrWhiteSpace([string]$baseUrlSetting)) { 'https://unirec.bitrix24.ru' } else { $baseUrlSetting }
    $webhookUserId = ConvertTo-BitrixInt $userIdSetting 165
    $webhookToken = if ([string]::IsNullOrWhiteSpace([string]$tokenSetting)) { '' } else { $tokenSetting }
    $projectId = ConvertTo-BitrixInt (Get-BitrixLocalSetting 'bitrix_project_id') 27
    $pollIntervalMinutes = ConvertTo-BitrixInt (Get-BitrixLocalSetting 'bitrix_poll_interval_minutes') 10
    $crmEntityTypeId = ConvertTo-BitrixInt (Get-BitrixLocalSetting 'bitrix_crm_entity_type_id') 2
    $dealCategoryId = ConvertTo-BitrixInt (Get-BitrixLocalSetting 'bitrix_deal_category_id') 0
    $dealStageId = [string](Get-BitrixLocalSetting 'bitrix_deal_stage_id')
    $rfqFileField = [string](Get-BitrixLocalSetting 'bitrix_deal_rfq_file_field')

    [pscustomobject]@{
        BaseUrl = [string]$baseUrl
        WebhookUserId = $webhookUserId
        WebhookToken = [string]$webhookToken
        ProjectId = $projectId
        PollIntervalMinutes = $pollIntervalMinutes
        CrmEntityTypeId = $crmEntityTypeId
        DealCategoryId = $dealCategoryId
        DealStageId = $dealStageId
        RfqFileField = $rfqFileField
    }
}

function Set-BitrixConfig {
    param(
        [string]$BaseUrl,
        [object]$WebhookUserId,
        [string]$WebhookToken,
        [object]$ProjectId,
        [object]$PollIntervalMinutes,
        [object]$CrmEntityTypeId,
        [object]$DealCategoryId,
        [string]$DealStageId,
        [string]$RfqFileField
    )

    $current = Get-BitrixConfig
    if ($PSBoundParameters.ContainsKey('BaseUrl')) {
        $current.BaseUrl = Normalize-BitrixBaseUrl $BaseUrl
    }
    if ($PSBoundParameters.ContainsKey('WebhookUserId')) {
        $current.WebhookUserId = ConvertTo-BitrixInt $WebhookUserId 0
    }
    if ($PSBoundParameters.ContainsKey('WebhookToken')) {
        $current.WebhookToken = [string]$WebhookToken
    }
    if ($PSBoundParameters.ContainsKey('ProjectId')) {
        $current.ProjectId = ConvertTo-BitrixInt $ProjectId 27
    }
    if ($PSBoundParameters.ContainsKey('PollIntervalMinutes')) {
        $current.PollIntervalMinutes = ConvertTo-BitrixInt $PollIntervalMinutes 10
    }
    if ($PSBoundParameters.ContainsKey('CrmEntityTypeId')) {
        $current.CrmEntityTypeId = ConvertTo-BitrixInt $CrmEntityTypeId 2
    }
    if ($PSBoundParameters.ContainsKey('DealCategoryId')) {
        $current.DealCategoryId = ConvertTo-BitrixInt $DealCategoryId 0
    }
    if ($PSBoundParameters.ContainsKey('DealStageId')) {
        $current.DealStageId = [string]$DealStageId
    }
    if ($PSBoundParameters.ContainsKey('RfqFileField')) {
        $current.RfqFileField = [string]$RfqFileField
    }

    Set-BitrixLocalSetting 'bitrix_base_url' $current.BaseUrl
    Set-BitrixLocalSetting 'bitrix_webhook_user_id' ([string]$current.WebhookUserId)
    Set-BitrixLocalSetting 'bitrix_webhook_token' $current.WebhookToken
    Set-BitrixLocalSetting 'bitrix_project_id' ([string]$current.ProjectId)
    Set-BitrixLocalSetting 'bitrix_poll_interval_minutes' ([string]$current.PollIntervalMinutes)
    Set-BitrixLocalSetting 'bitrix_crm_entity_type_id' ([string]$current.CrmEntityTypeId)
    Set-BitrixLocalSetting 'bitrix_deal_category_id' ([string]$current.DealCategoryId)
    Set-BitrixLocalSetting 'bitrix_deal_stage_id' $current.DealStageId
    Set-BitrixLocalSetting 'bitrix_deal_rfq_file_field' $current.RfqFileField

    return Get-BitrixConfig
}

function New-BitrixMethodUri {
    param(
        [string]$BaseUrl,
        [int]$UserId,
        [string]$Token,
        [string]$Method
    )

    $normalizedBaseUrl = Normalize-BitrixBaseUrl $BaseUrl
    if ([string]::IsNullOrWhiteSpace($normalizedBaseUrl)) {
        throw 'Bitrix BaseUrl is required.'
    }
    if ($UserId -le 0) {
        throw 'Bitrix WebhookUserId is required.'
    }
    if ([string]::IsNullOrWhiteSpace($Token)) {
        throw 'Bitrix WebhookToken is required.'
    }
    $methodName = ([string]$Method).Trim().TrimStart('/')
    if ([string]::IsNullOrWhiteSpace($methodName)) {
        throw 'Bitrix Method is required.'
    }

    return ('{0}/rest/{1}/{2}/{3}' -f $normalizedBaseUrl, $UserId, $Token.Trim(), $methodName)
}

function Invoke-BitrixMethod {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Method,
        [hashtable]$Params = @{},
        [int]$TimeoutSec = 30,
        [scriptblock]$Transport = $null
    )

    $config = Get-BitrixConfig
    $uri = New-BitrixMethodUri -BaseUrl $config.BaseUrl -UserId $config.WebhookUserId -Token $config.WebhookToken -Method $Method
    $payload = '{}'
    if ($null -ne $Params -and $Params.Count -gt 0) {
        $payload = $Params | ConvertTo-Json -Depth 20 -Compress
    }

    if ($null -eq $Transport) {
        $Transport = {
            param($RequestUri, $JsonBody, $RequestTimeoutSec)
            Invoke-RestMethod -Uri $RequestUri -Method Post -Body $JsonBody -ContentType 'application/json; charset=utf-8' -TimeoutSec $RequestTimeoutSec
        }
    }

    $response = & $Transport $uri $payload $TimeoutSec
    if ($null -ne $response -and $null -ne $response.error) {
        $errorCode = [string]$response.error
        $errorDescription = [string]$response.error_description
        if ([string]::IsNullOrWhiteSpace($errorDescription)) {
            $errorDescription = 'Bitrix returned an API error.'
        }
        throw ('Bitrix API error {0}: {1}' -f $errorCode, $errorDescription)
    }

    return $response
}

function Convert-BitrixTaskTitleToProcurement {
    param([string]$Title)

    $originalTitle = [string]$Title
    $text = $originalTitle.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw 'Bitrix task title is required.'
    }

    $firstSlash = $text.IndexOf('/')
    if ($firstSlash -lt 1) {
        throw 'Bitrix task title must start with a request number before the first slash.'
    }

    $lastSlash = $text.LastIndexOf('/')
    if ($lastSlash -le $firstSlash) {
        throw 'Bitrix task title must contain client and item description segments.'
    }

    $requestNumber = $text.Substring(0, $firstSlash).Trim()
    $client = $text.Substring($firstSlash + 1, $lastSlash - $firstSlash - 1).Trim()
    $itemText = $text.Substring($lastSlash + 1).Trim()

    if ([string]::IsNullOrWhiteSpace($requestNumber)) {
        throw 'Bitrix task title must start with a request number before the first slash.'
    }
    if ([string]::IsNullOrWhiteSpace($client)) {
        throw 'Bitrix task title must include a client between the first and last slash.'
    }
    if ([string]::IsNullOrWhiteSpace($itemText)) {
        throw 'Bitrix task title must include an item description.'
    }

    $itemDescription = $itemText
    $quantity = $null
    $quantityMatch = [regex]::Match($itemText, '^(?<description>.*?)(?:\s*-\s*(?<quantity>\d+)\s*(?:шт\.?)?\s*)$')
    if ($quantityMatch.Success) {
        $itemDescription = $quantityMatch.Groups['description'].Value.Trim()
        $quantity = [int]$quantityMatch.Groups['quantity'].Value
    }

    if ([string]::IsNullOrWhiteSpace($itemDescription)) {
        throw 'Bitrix task title must include an item description.'
    }
    if ($null -eq $quantity) {
        throw 'Bitrix task title must end with a quantity.'
    }

    return [pscustomobject]@{
        RequestNumber = $requestNumber
        Client = $client
        ItemDescription = $itemDescription
        Quantity = $quantity
        OriginalTitle = $originalTitle
    }
}

function Test-RfqFilenameMatchesDeal {
    param(
        [string]$DealNumber,
        [string]$FileName
    )

    $normalizedDealNumber = ([string]$DealNumber).Trim().ToUpperInvariant()
    $fileNameText = [string]$FileName

    if ([string]::IsNullOrWhiteSpace($normalizedDealNumber)) {
        return [pscustomobject]@{
            Matches = $false
            NormalizedDealNumber = $normalizedDealNumber
            FileName = $fileNameText
            Reason = 'Deal number is required.'
        }
    }

    if ([string]::IsNullOrWhiteSpace($fileNameText)) {
        return [pscustomobject]@{
            Matches = $false
            NormalizedDealNumber = $normalizedDealNumber
            FileName = $fileNameText
            Reason = 'File name is required.'
        }
    }

    $normalizedFileName = $fileNameText.Trim().ToUpperInvariant()
    $matches = $normalizedFileName.Contains($normalizedDealNumber)
    $reason = if ($matches) {
        'Deal number matched.'
    } else {
        "Deal number '$normalizedDealNumber' was not found in '$fileNameText'."
    }

    return [pscustomobject]@{
        Matches = $matches
        NormalizedDealNumber = $normalizedDealNumber
        FileName = $fileNameText
        Reason = $reason
    }
}

function Get-BitrixTaskLink {
    param([int]$BitrixTaskId)

    if ($BitrixTaskId -le 0) {
        return $null
    }

    $table = Invoke-PurchaseQuery @'
SELECT
    bitrix_task_id,
    deal_id,
    bitrix_deal_id,
    bitrix_status,
    is_manual_deleted,
    created_at,
    updated_at,
    linked_at,
    last_seen_at,
    deleted_at
FROM bitrix_task_links
WHERE bitrix_task_id = @task_id
'@ @{
        '@task_id' = $BitrixTaskId
    }

    if ($null -eq $table -or $table.Rows.Count -eq 0) {
        return $null
    }

    return $table.Rows[0]
}

function Save-BitrixTaskLink {
    param(
        [int]$BitrixTaskId,
        [AllowNull()][Nullable[int]]$DealId = $null,
        [AllowNull()][Nullable[int]]$BitrixDealId = $null,
        [AllowNull()][Nullable[int]]$BitrixStatus = $null,
        [bool]$IsManualDeleted = $false
    )

    if ($BitrixTaskId -le 0) {
        throw 'BitrixTaskId is required.'
    }

    $hasDealId = $PSBoundParameters.ContainsKey('DealId')
    $hasBitrixDealId = $PSBoundParameters.ContainsKey('BitrixDealId')
    $hasBitrixStatus = $PSBoundParameters.ContainsKey('BitrixStatus')
    $hasIsManualDeleted = $PSBoundParameters.ContainsKey('IsManualDeleted')
    $now = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO bitrix_task_links(
    bitrix_task_id,
    deal_id,
    bitrix_deal_id,
    bitrix_status,
    is_manual_deleted,
    created_at,
    updated_at,
    linked_at,
    last_seen_at,
    deleted_at
)
VALUES(
    @bitrix_task_id,
    @deal_id,
    @bitrix_deal_id,
    @bitrix_status,
    @is_manual_deleted,
    @created_at,
    @updated_at,
    CASE WHEN @deal_id IS NULL THEN NULL ELSE @linked_at END,
    @last_seen_at,
    CASE WHEN @is_manual_deleted = 1 THEN @deleted_at ELSE NULL END
)
ON CONFLICT(bitrix_task_id) DO UPDATE SET
    deal_id = CASE WHEN @has_deal_id = 1 THEN excluded.deal_id ELSE bitrix_task_links.deal_id END,
    bitrix_deal_id = CASE WHEN @has_bitrix_deal_id = 1 THEN excluded.bitrix_deal_id ELSE bitrix_task_links.bitrix_deal_id END,
    bitrix_status = CASE WHEN @has_bitrix_status = 1 THEN excluded.bitrix_status ELSE bitrix_task_links.bitrix_status END,
    is_manual_deleted = CASE WHEN @has_is_manual_deleted = 1 THEN excluded.is_manual_deleted ELSE bitrix_task_links.is_manual_deleted END,
    updated_at = excluded.updated_at,
    linked_at = CASE
        WHEN @has_deal_id = 0 THEN bitrix_task_links.linked_at
        WHEN excluded.deal_id IS NULL THEN bitrix_task_links.linked_at
        WHEN bitrix_task_links.linked_at IS NULL THEN excluded.linked_at
        ELSE bitrix_task_links.linked_at
    END,
    last_seen_at = excluded.last_seen_at,
    deleted_at = CASE
        WHEN @has_is_manual_deleted = 0 THEN bitrix_task_links.deleted_at
        WHEN excluded.is_manual_deleted = 1 THEN excluded.deleted_at
        ELSE NULL
    END
'@ @{
        '@bitrix_task_id' = $BitrixTaskId
        '@deal_id' = if ($null -eq $DealId -or $DealId -le 0) { $null } else { $DealId }
        '@bitrix_deal_id' = if ($null -eq $BitrixDealId -or $BitrixDealId -le 0) { $null } else { $BitrixDealId }
        '@bitrix_status' = if ($null -eq $BitrixStatus) { $null } else { $BitrixStatus }
        '@is_manual_deleted' = if ($IsManualDeleted) { 1 } else { 0 }
        '@has_deal_id' = if ($hasDealId) { 1 } else { 0 }
        '@has_bitrix_deal_id' = if ($hasBitrixDealId) { 1 } else { 0 }
        '@has_bitrix_status' = if ($hasBitrixStatus) { 1 } else { 0 }
        '@has_is_manual_deleted' = if ($hasIsManualDeleted) { 1 } else { 0 }
        '@created_at' = $now
        '@updated_at' = $now
        '@linked_at' = $now
        '@last_seen_at' = $now
        '@deleted_at' = $now
    }

    return Get-BitrixTaskLink -BitrixTaskId $BitrixTaskId
}

function Get-ProcessedBitrixFile {
    param(
        [int]$BitrixTaskId,
        [string]$FileKey
    )

    if ($BitrixTaskId -le 0 -or [string]::IsNullOrWhiteSpace($FileKey)) {
        return $null
    }

    $table = Invoke-PurchaseQuery @'
SELECT
    id,
    bitrix_task_id,
    file_key,
    bitrix_file_id,
    file_name,
    file_hash,
    local_path,
    upload_state,
    error_text,
    created_at,
    updated_at
FROM bitrix_rfq_files
WHERE bitrix_task_id = @task_id AND file_key = @file_key
'@ @{
        '@task_id' = $BitrixTaskId
        '@file_key' = $FileKey.Trim()
    }

    if ($null -eq $table -or $table.Rows.Count -eq 0) {
        return $null
    }

    return $table.Rows[0]
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

    if ($BitrixTaskId -le 0) {
        throw 'BitrixTaskId is required.'
    }
    if ([string]::IsNullOrWhiteSpace($FileKey)) {
        throw 'FileKey is required.'
    }

    $hasBitrixFileId = $PSBoundParameters.ContainsKey('BitrixFileId')
    $hasFileName = $PSBoundParameters.ContainsKey('FileName')
    $hasFileHash = $PSBoundParameters.ContainsKey('FileHash')
    $hasLocalPath = $PSBoundParameters.ContainsKey('LocalPath')
    $hasUploadState = $PSBoundParameters.ContainsKey('UploadState')
    $hasErrorText = $PSBoundParameters.ContainsKey('ErrorText')
    $now = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO bitrix_rfq_files(
    bitrix_task_id,
    file_key,
    bitrix_file_id,
    file_name,
    file_hash,
    local_path,
    upload_state,
    error_text,
    created_at,
    updated_at
)
VALUES(
    @bitrix_task_id,
    @file_key,
    @bitrix_file_id,
    @file_name,
    @file_hash,
    @local_path,
    @upload_state,
    @error_text,
    @created_at,
    @updated_at
)
ON CONFLICT(bitrix_task_id, file_key) DO UPDATE SET
    bitrix_file_id = CASE WHEN @has_bitrix_file_id = 1 THEN excluded.bitrix_file_id ELSE bitrix_rfq_files.bitrix_file_id END,
    file_name = CASE WHEN @has_file_name = 1 THEN excluded.file_name ELSE bitrix_rfq_files.file_name END,
    file_hash = CASE WHEN @has_file_hash = 1 THEN excluded.file_hash ELSE bitrix_rfq_files.file_hash END,
    local_path = CASE WHEN @has_local_path = 1 THEN excluded.local_path ELSE bitrix_rfq_files.local_path END,
    upload_state = CASE WHEN @has_upload_state = 1 THEN excluded.upload_state ELSE bitrix_rfq_files.upload_state END,
    error_text = CASE WHEN @has_error_text = 1 THEN excluded.error_text ELSE bitrix_rfq_files.error_text END,
    updated_at = excluded.updated_at
'@ @{
        '@bitrix_task_id' = $BitrixTaskId
        '@file_key' = $FileKey.Trim()
        '@bitrix_file_id' = if ($hasBitrixFileId) { [string]$BitrixFileId } else { $null }
        '@file_name' = if ($hasFileName) { [string]$FileName } else { $null }
        '@file_hash' = if ($hasFileHash) { [string]$FileHash } else { $null }
        '@local_path' = if ($hasLocalPath) { [string]$LocalPath } else { $null }
        '@upload_state' = if ($hasUploadState) { [string]$UploadState } else { $null }
        '@error_text' = if ($hasErrorText) { [string]$ErrorText } else { $null }
        '@has_bitrix_file_id' = if ($hasBitrixFileId) { 1 } else { 0 }
        '@has_file_name' = if ($hasFileName) { 1 } else { 0 }
        '@has_file_hash' = if ($hasFileHash) { 1 } else { 0 }
        '@has_local_path' = if ($hasLocalPath) { 1 } else { 0 }
        '@has_upload_state' = if ($hasUploadState) { 1 } else { 0 }
        '@has_error_text' = if ($hasErrorText) { 1 } else { 0 }
        '@created_at' = $now
        '@updated_at' = $now
    }

    return Get-ProcessedBitrixFile -BitrixTaskId $BitrixTaskId -FileKey $FileKey
}

function Invoke-BitrixApiInvoker {
    param(
        [string]$Method,
        [hashtable]$Params = @{},
        [int]$TimeoutSec = 30,
        [scriptblock]$ApiInvoker = $null
    )

    if ($null -ne $ApiInvoker) {
        return & $ApiInvoker $Method $Params $TimeoutSec
    }

    return Invoke-BitrixMethod -Method $Method -Params $Params -TimeoutSec $TimeoutSec
}

function ConvertTo-BitrixTaskValue {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [DBNull]) {
        return ''
    }

    return [string]$Value
}

function ConvertTo-BitrixNullableInt {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [DBNull]) {
        return $null
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $parsed = 0
    if ([int]::TryParse($text, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

function Get-BitrixRfqChecklistTitle {
    return 'Заполнение сводного RFQ и выделение компонентов на расчет (Сенько Е., Сидоров М.)'
}

function Get-BitrixRfqChecklistPrefix {
    $title = Get-BitrixRfqChecklistTitle
    $markerIndex = $title.IndexOf('(')
    if ($markerIndex -lt 0) {
        return $title
    }

    return $title.Substring(0, $markerIndex).Trim()
}

function Get-BitrixChecklistItems {
    param(
        [int]$TaskId,
        [scriptblock]$ApiInvoker = $null
    )

    if ($TaskId -le 0) {
        return @()
    }

    $response = Invoke-BitrixApiInvoker -Method 'task.checklistitem.getlist' -Params @{ TASKID = $TaskId } -TimeoutSec 20 -ApiInvoker $ApiInvoker
    if ($null -eq $response -or $null -eq $response.result) {
        return @()
    }

    return @($response.result)
}

function Get-BitrixRecordValue {
    param(
        [object]$Record,
        [string[]]$Names
    )

    if ($null -eq $Record -or $null -eq $Names) {
        return $null
    }

    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if ($Record -is [System.Collections.IDictionary]) {
            if ($Record.Contains($name)) {
                return $Record[$name]
            }
            foreach ($key in $Record.Keys) {
                if ([string]::Equals([string]$key, $name, [StringComparison]::OrdinalIgnoreCase)) {
                    return $Record[$key]
                }
            }
            continue
        }

        $property = $Record.PSObject.Properties | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
        if ($null -ne $property) {
            return $property.Value
        }
    }

    return $null
}

function Normalize-BitrixChecklistTitle {
    param([string]$Title)

    return ((([string]$Title) -replace '\s+', ' ').Trim())
}

function Get-BitrixChecklistItem {
    param(
        [int]$TaskId,
        [string]$Title,
        [scriptblock]$ApiInvoker = $null
    )

    if ($TaskId -le 0 -or [string]::IsNullOrWhiteSpace($Title)) {
        return $null
    }

    $expectedTitle = Normalize-BitrixChecklistTitle $Title
    foreach ($item in @(Get-BitrixChecklistItems -TaskId $TaskId -ApiInvoker $ApiInvoker)) {
        $itemTitle = Normalize-BitrixChecklistTitle (ConvertTo-BitrixTaskValue (Get-BitrixRecordValue $item @('TITLE', 'title')))
        if ($itemTitle -ceq $expectedTitle) {
            return $item
        }
    }

    return $null
}

function ConvertTo-BitrixDateTimeOffset {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [DBNull]) {
        return $null
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $parsed = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse($text, [ref]$parsed)) {
        return $parsed
    }

    return $null
}

function Get-BitrixAttachmentTimestampText {
    param([object]$Attachment)

    foreach ($fieldName in @('UPDATE_TIME', 'UPDATED_AT', 'MODIFIED_AT', 'MODIFY_TIME', 'TIMESTAMP_X', 'CREATE_TIME', 'CREATED_AT')) {
        $value = Get-BitrixRecordValue $Attachment @($fieldName)
        if ($null -ne (ConvertTo-BitrixDateTimeOffset $value)) {
            return [string]$value
        }
    }

    return ''
}

function Get-BitrixDriveFileMetadata {
    param(
        [string]$FileId,
        [scriptblock]$ApiInvoker = $null
    )

    if ([string]::IsNullOrWhiteSpace($FileId) -or $null -eq $ApiInvoker) {
        return $null
    }

    $response = Invoke-BitrixApiInvoker -Method 'disk.file.get' -Params @{ id = $FileId } -TimeoutSec 20 -ApiInvoker $ApiInvoker
    if ($null -eq $response -or $null -eq $response.result) {
        return $null
    }

    return $response.result
}

function Get-BitrixChecklistAttachments {
    param([object]$ChecklistItem)

    $attachments = Get-BitrixRecordValue $ChecklistItem @('ATTACHMENTS', 'attachments')
    if ($null -eq $attachments) {
        return @()
    }

    if ($attachments -is [System.Collections.IDictionary]) {
        return @($attachments.GetEnumerator() | ForEach-Object { $_.Value })
    }

    if ($attachments -is [System.Array] -or $attachments -is [System.Collections.IEnumerable]) {
        return @($attachments)
    }

    $properties = @($attachments.PSObject.Properties)
    if ($properties.Count -gt 0) {
        return @($properties | ForEach-Object { $_.Value })
    }

    return @()
}

function Select-LatestBitrixChecklistFile {
    param(
        $ChecklistItem,
        [scriptblock]$ApiInvoker = $null
    )

    if ($null -eq $ChecklistItem) {
        return $null
    }

    $candidates = New-Object System.Collections.Generic.List[object]
    foreach ($attachment in @(Get-BitrixChecklistAttachments -ChecklistItem $ChecklistItem)) {
        $fileId = [string](Get-BitrixRecordValue $attachment @('FILE_ID', 'fileId', 'file_id'))
        $attachmentId = [string](Get-BitrixRecordValue $attachment @('ATTACHMENT_ID', 'ID', 'attachmentId', 'id'))
        $fileName = [string](Get-BitrixRecordValue $attachment @('NAME', 'name'))
        $downloadUrl = [string](Get-BitrixRecordValue $attachment @('DOWNLOAD_URL', 'downloadUrl'))
        $modifiedAt = Get-BitrixAttachmentTimestampText $attachment
        $metadata = $null

        if ([string]::IsNullOrWhiteSpace($modifiedAt) -and -not [string]::IsNullOrWhiteSpace($fileId)) {
            $metadata = Get-BitrixDriveFileMetadata -FileId $fileId -ApiInvoker $ApiInvoker
            if ($null -ne $metadata) {
                $metadataTimestamp = Get-BitrixAttachmentTimestampText $metadata
                if (-not [string]::IsNullOrWhiteSpace($metadataTimestamp)) {
                    $modifiedAt = $metadataTimestamp
                }
                $metadataName = [string](Get-BitrixRecordValue $metadata @('NAME', 'name'))
                if (-not [string]::IsNullOrWhiteSpace($metadataName)) {
                    $fileName = $metadataName
                }
                $metadataDownloadUrl = [string](Get-BitrixRecordValue $metadata @('DOWNLOAD_URL', 'downloadUrl'))
                if (-not [string]::IsNullOrWhiteSpace($metadataDownloadUrl)) {
                    $downloadUrl = $metadataDownloadUrl
                }
            }
        }

        if ([string]::IsNullOrWhiteSpace($fileId) -and [string]::IsNullOrWhiteSpace($attachmentId)) {
            continue
        }

        $sortTimestamp = ConvertTo-BitrixDateTimeOffset $modifiedAt
        if ($null -eq $sortTimestamp -and $null -ne $metadata) {
            $sortTimestamp = ConvertTo-BitrixDateTimeOffset (Get-BitrixAttachmentTimestampText $metadata)
        }

        [void]$candidates.Add([pscustomobject]@{
            AttachmentId = $attachmentId
            FileId = if (-not [string]::IsNullOrWhiteSpace($fileId)) { $fileId } else { $attachmentId }
            Name = $fileName
            DownloadUrl = $downloadUrl
            ModifiedAt = $modifiedAt
            SortTimestamp = $sortTimestamp
            SortToken = ('{0:D20}:{1:D20}' -f (ConvertTo-BitrixInt $attachmentId 0), (ConvertTo-BitrixInt $fileId 0))
        })
    }

    if ($candidates.Count -eq 0) {
        return $null
    }

    return $candidates |
        Sort-Object @{ Expression = { if ($null -eq $_.SortTimestamp) { [DateTimeOffset]::MinValue } else { $_.SortTimestamp } }; Descending = $true },
                    @{ Expression = { $_.SortToken }; Descending = $true } |
        Select-Object -First 1
}

function Get-BitrixSafePathPart {
    param([string]$Value)

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return 'item'
    }

    $invalidChars = [IO.Path]::GetInvalidFileNameChars()
    foreach ($char in $invalidChars) {
        $text = $text.Replace([string]$char, '_')
    }

    $text = $text.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return 'item'
    }

    return $text
}

function Get-BitrixSafeFileName {
    param([string]$FileName)

    $safeName = Get-BitrixSafePathPart $FileName
    if ($safeName -eq 'item') {
        return 'bitrix-file'
    }

    return $safeName
}

function Download-BitrixFile {
    param(
        $File,
        [string]$TargetPath,
        [scriptblock]$Downloader = $null
    )

    if ($null -eq $File) {
        throw 'File is required.'
    }
    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        throw 'TargetPath is required.'
    }

    $fileName = [string](Get-BitrixRecordValue $File @('Name', 'NAME', 'name'))
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = [IO.Path]::GetFileName($TargetPath)
    }
    $finalPath = $TargetPath
    if ([string]::IsNullOrWhiteSpace([IO.Path]::GetExtension($TargetPath)) -and -not [string]::IsNullOrWhiteSpace($fileName) -and (Test-Path -LiteralPath $TargetPath -PathType Container)) {
        $finalPath = Join-Path $TargetPath (Get-BitrixSafeFileName $fileName)
    }

    $targetDirectory = Split-Path -Parent $finalPath
    if ([string]::IsNullOrWhiteSpace($targetDirectory)) {
        throw 'TargetPath must include a parent directory.'
    }

    New-Item -ItemType Directory -Force -Path $targetDirectory | Out-Null
    $tempPath = Join-Path $targetDirectory (([IO.Path]::GetRandomFileName()) + '.part')

    try {
        if ($null -eq $Downloader) {
            $downloadUrl = [string](Get-BitrixRecordValue $File @('DownloadUrl', 'DOWNLOAD_URL', 'downloadUrl'))
            if ([string]::IsNullOrWhiteSpace($downloadUrl)) {
                throw 'Bitrix file download URL is required.'
            }
            Invoke-WebRequest -Uri $downloadUrl -OutFile $tempPath -UseBasicParsing | Out-Null
        } else {
            & $Downloader $File $tempPath | Out-Null
        }

        if (-not (Test-Path -LiteralPath $tempPath -PathType Leaf)) {
            throw 'Bitrix download did not create a temporary file.'
        }

        $hash = [string](Get-FileHash -LiteralPath $tempPath -Algorithm SHA256).Hash
        Move-Item -LiteralPath $tempPath -Destination $finalPath -Force

        return [pscustomobject]@{
            Path = $finalPath
            Hash = $hash
        }
    } finally {
        if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Upload-RfqToBitrixDeal {
    param(
        [int]$BitrixDealId,
        [string]$FilePath,
        [scriptblock]$UploadTransport = $null
    )

    if ($BitrixDealId -le 0) {
        throw 'BitrixDealId is required.'
    }
    if ([string]::IsNullOrWhiteSpace($FilePath) -or -not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
        throw 'FilePath must point to an existing file.'
    }

    $config = Get-BitrixConfig
    $fieldName = ([string]$config.RfqFileField).Trim()
    if ([string]::IsNullOrWhiteSpace($fieldName)) {
        throw 'Bitrix RFQ file field is not configured.'
    }

    $fileName = [IO.Path]::GetFileName($FilePath)
    $fileBytes = [IO.File]::ReadAllBytes($FilePath)
    $base64Content = [Convert]::ToBase64String($fileBytes)
    $params = @{
        entityTypeId = $config.CrmEntityTypeId
        id = $BitrixDealId
        fields = @{
            $fieldName = @(@($fileName, $base64Content))
        }
    }

    $response = Invoke-BitrixApiInvoker -Method 'crm.item.update' -Params $params -TimeoutSec 60 -ApiInvoker $UploadTransport
    return [pscustomobject]@{
        Method = 'crm.item.update'
        FieldName = $fieldName
        BitrixDealId = $BitrixDealId
        FileName = $fileName
        Response = $response
    }
}

function Get-BitrixRfqFileKey {
    param(
        $ChecklistItem,
        $File
    )

    $checklistId = [string](Get-BitrixRecordValue $ChecklistItem @('ID', 'id'))
    $fileId = [string](Get-BitrixRecordValue $File @('FileId', 'FILE_ID', 'fileId', 'ID', 'id'))
    if ([string]::IsNullOrWhiteSpace($fileId)) {
        $fileId = [string](Get-BitrixRecordValue $File @('AttachmentId', 'ATTACHMENT_ID', 'attachmentId'))
    }

    return ('rfq:{0}:{1}' -f $checklistId, $fileId)
}

function Get-BitrixDocumentsRoot {
    if (Get-Command -Name Get-PurchaseDocumentsRoot -ErrorAction SilentlyContinue) {
        return [string](Get-PurchaseDocumentsRoot)
    }

    $fallbackRoot = Join-Path ([IO.Path]::GetTempPath()) 'procurement-control-bitrix'
    New-Item -ItemType Directory -Force -Path $fallbackRoot | Out-Null
    return $fallbackRoot
}

function New-BitrixRfqResult {
    param(
        [string]$Status,
        [string]$Message,
        [AllowNull()]$File = $null
    )

    return [pscustomobject]@{
        Status = $Status
        Message = $Message
        File = $File
    }
}

function Sync-BitrixRfq {
    param(
        $Task,
        $Procurement,
        [switch]$ConfirmMismatch,
        [scriptblock]$ApiInvoker = $null,
        [scriptblock]$Downloader = $null,
        [scriptblock]$UploadTransport = $null
    )

    if ($null -eq $Task) {
        throw 'Task is required.'
    }
    if ($null -eq $Procurement) {
        throw 'Procurement is required.'
    }

    $taskId = ConvertTo-BitrixInt (Get-BitrixRecordValue $Task @('Id', 'id')) 0
    if ($taskId -le 0) {
        throw 'Bitrix task ID is required.'
    }

    $dealId = ConvertTo-BitrixNullableInt (Get-BitrixRecordValue $Procurement @('id', 'Id'))
    $dealNumber = [string](Get-BitrixRecordValue $Procurement @('deal_number', 'DealNumber'))
    $bitrixDealId = ConvertTo-BitrixNullableInt (Get-BitrixRecordValue $Procurement @('bitrix_deal_id', 'BitrixDealId'))
    if ([string]::IsNullOrWhiteSpace($dealNumber)) {
        throw 'Procurement deal number is required.'
    }
    if ($null -eq $bitrixDealId -or $bitrixDealId -le 0) {
        return New-BitrixRfqResult -Status 'Failed' -Message ('Bitrix deal is not linked for deal {0}.' -f $dealNumber)
    }

    $checklistItem = Get-BitrixChecklistItem -TaskId $taskId -Title (Get-BitrixRfqChecklistTitle) -ApiInvoker $ApiInvoker
    if ($null -eq $checklistItem) {
        return New-BitrixRfqResult -Status 'Waiting' -Message ('RFQ checklist item is missing for Bitrix task {0}.' -f $taskId)
    }

    $file = Select-LatestBitrixChecklistFile -ChecklistItem $checklistItem -ApiInvoker $ApiInvoker
    if ($null -eq $file) {
        return New-BitrixRfqResult -Status 'Waiting' -Message ('RFQ file is not attached yet for deal {0}.' -f $dealNumber)
    }

    $fileKey = Get-BitrixRfqFileKey -ChecklistItem $checklistItem -File $file
    $processed = Get-ProcessedBitrixFile -BitrixTaskId $taskId -FileKey $fileKey
    $selectedFileId = [string](Get-BitrixRecordValue $file @('FileId', 'fileId', 'FILE_ID'))
    $selectedFileName = [string](Get-BitrixRecordValue $file @('Name', 'name', 'NAME'))

    if ($null -ne $processed) {
        $savedState = [string](Get-BitrixRecordValue $processed @('upload_state', 'UploadState'))
        $savedFileId = [string](Get-BitrixRecordValue $processed @('bitrix_file_id', 'BitrixFileId'))
        if ($savedState -eq 'Uploaded' -and $savedFileId -eq $selectedFileId) {
            return New-BitrixRfqResult -Status 'Skipped' -Message ('RFQ file {0} is already processed for deal {1}.' -f $selectedFileName, $dealNumber) -File $file
        }
    }

    $match = Test-RfqFilenameMatchesDeal -DealNumber $dealNumber -FileName $selectedFileName
    if (-not $match.Matches -and -not $ConfirmMismatch) {
        $message = ("RFQ filename '{0}' does not contain deal number '{1}'. Waiting for confirmation." -f $selectedFileName, $dealNumber)
        Save-ProcessedBitrixFile -BitrixTaskId $taskId -FileKey $fileKey -BitrixFileId $selectedFileId -FileName $selectedFileName -UploadState 'Waiting' -ErrorText $message | Out-Null
        return New-BitrixRfqResult -Status 'Waiting' -Message $message -File $file
    }

    $documentsRoot = Get-BitrixDocumentsRoot
    $dealFolder = Join-Path $documentsRoot (Get-BitrixSafePathPart $dealNumber)
    $rfqFolder = Join-Path $dealFolder 'RFQ'
    $targetPath = Join-Path $rfqFolder (Get-BitrixSafeFileName $selectedFileName)

    try {
        $downloadResult = Download-BitrixFile -File $file -TargetPath $targetPath -Downloader $Downloader
    } catch {
        $downloadError = $_.Exception.Message
        Save-ProcessedBitrixFile -BitrixTaskId $taskId -FileKey $fileKey -BitrixFileId $selectedFileId -FileName $selectedFileName -UploadState 'Failed' -ErrorText $downloadError | Out-Null
        return New-BitrixRfqResult -Status 'Failed' -Message $downloadError -File $file
    }

    try {
        $uploadResult = Upload-RfqToBitrixDeal -BitrixDealId $bitrixDealId -FilePath $downloadResult.Path -UploadTransport $UploadTransport
        $errorText = if (-not $match.Matches -and $ConfirmMismatch) {
            ('Mismatch override approved for deal {0} and file {1}.' -f $dealNumber, $selectedFileName)
        } else {
            ''
        }
        Save-ProcessedBitrixFile -BitrixTaskId $taskId -FileKey $fileKey -BitrixFileId $selectedFileId -FileName $selectedFileName -FileHash $downloadResult.Hash -LocalPath $downloadResult.Path -UploadState 'Uploaded' -ErrorText $errorText | Out-Null
        return New-BitrixRfqResult -Status 'Uploaded' -Message ('RFQ file {0} uploaded for deal {1}.' -f $selectedFileName, $dealNumber) -File $uploadResult
    } catch {
        $uploadError = $_.Exception.Message
        Save-ProcessedBitrixFile -BitrixTaskId $taskId -FileKey $fileKey -BitrixFileId $selectedFileId -FileName $selectedFileName -FileHash $downloadResult.Hash -LocalPath $downloadResult.Path -UploadState 'Failed' -ErrorText $uploadError | Out-Null
        return New-BitrixRfqResult -Status 'Failed' -Message $uploadError -File $file
    }
}

function Test-BitrixTaskHasCompletedRfq {
    param(
        [int]$TaskId,
        [scriptblock]$ApiInvoker = $null
    )

    try {
        $checklistItems = @(Get-BitrixChecklistItems -TaskId $TaskId -ApiInvoker $ApiInvoker)
    } catch {
        return $false
    }

    if ($checklistItems.Count -eq 0) {
        return $false
    }

    $prefix = Get-BitrixRfqChecklistPrefix
    foreach ($item in $checklistItems) {
        $itemTitle = ((ConvertTo-BitrixTaskValue $item.TITLE) -replace '\s+', ' ').Trim()
        if ($itemTitle -like ($prefix + '*')) {
            return ([string]$item.IS_COMPLETE -eq 'Y' -or $item.IS_COMPLETE -eq $true)
        }
    }

    return $false
}

function Normalize-BitrixTask {
    param(
        $Task,
        [int]$DefaultGroupId = 27
    )

    if ($null -eq $Task) {
        return $null
    }

    $taskId = ConvertTo-BitrixInt $Task.id 0
    if ($taskId -le 0) {
        return $null
    }

    $deadline = ''
    $deadlineValue = ConvertTo-BitrixTaskValue $Task.deadline
    if (-not [string]::IsNullOrWhiteSpace($deadlineValue)) {
        if ($deadlineValue.Length -ge 10) {
            $deadline = $deadlineValue.Substring(0, 10)
        } else {
            $deadline = $deadlineValue
        }
    }

    $groupId = ConvertTo-BitrixInt $Task.groupId 0
    if ($groupId -le 0) {
        $groupId = ConvertTo-BitrixInt $Task.GROUP_ID 0
    }
    if ($groupId -le 0) {
        $groupId = $DefaultGroupId
    }

    return [pscustomobject]@{
        Id = $taskId
        Title = (ConvertTo-BitrixTaskValue $Task.title).Trim()
        Status = ConvertTo-BitrixInt $Task.status 0
        Deadline = $deadline
        GroupId = $groupId
        Raw = $Task
    }
}

function Get-BitrixBlockedTaskIdSet {
    if (-not (Get-Command -Name Get-BitrixBlockedTaskIds -ErrorAction SilentlyContinue)) {
        return @()
    }

    return @(Get-BitrixBlockedTaskIds)
}

function Get-BitrixNewProjectTasks {
    param([scriptblock]$ApiInvoker = $null)

    $bitrixConfig = Get-BitrixConfig
    $tasksResponse = Invoke-BitrixApiInvoker -Method 'tasks.task.list' -Params @{
        filter = @{
            GROUP_ID = $bitrixConfig.ProjectId
            '>=REAL_STATUS' = 1
            '<=REAL_STATUS' = 6
        }
        select = @('id', 'title', 'deadline', 'status', 'groupId')
        order = @{ ID = 'DESC' }
    } -TimeoutSec 20 -ApiInvoker $ApiInvoker

    $rawTasks = @()
    if ($null -ne $tasksResponse -and $null -ne $tasksResponse.result -and $null -ne $tasksResponse.result.tasks) {
        $rawTasks = @($tasksResponse.result.tasks)
    }

    $allowedStatuses = @(1, 2, 3, 6)
    $blockedTaskIds = @(Get-BitrixBlockedTaskIdSet)
    $normalizedTasks = New-Object System.Collections.Generic.List[object]

    foreach ($rawTask in $rawTasks) {
        $task = Normalize-BitrixTask -Task $rawTask -DefaultGroupId $bitrixConfig.ProjectId
        if ($null -eq $task) { continue }
        if ($blockedTaskIds -contains $task.Id) { continue }
        if ($allowedStatuses -notcontains $task.Status) { continue }
        if (Test-BitrixTaskHasCompletedRfq -TaskId $task.Id -ApiInvoker $ApiInvoker) { continue }
        [void]$normalizedTasks.Add($task)
    }

    return $normalizedTasks.ToArray()
}

function Get-PurchaseDealByNumber {
    param([string]$DealNumber)

    if ([string]::IsNullOrWhiteSpace($DealNumber)) {
        return $null
    }

    $table = Invoke-PurchaseQuery @'
SELECT
    id,
    deal_number,
    client,
    status,
    tracking_status,
    IFNULL(priority, '') AS priority,
    IFNULL(title, '') AS title,
    bitrix_deal_id,
    archived
FROM deals
WHERE deal_number = @deal_number
'@ @{
        '@deal_number' = $DealNumber.Trim()
    }

    if ($null -eq $table -or $table.Rows.Count -eq 0) {
        return $null
    }

    return $table.Rows[0]
}

function Get-PurchaseDealById {
    param([int]$DealId)

    if ($DealId -le 0) {
        return $null
    }

    $table = Invoke-PurchaseQuery @'
SELECT
    id,
    deal_number,
    client,
    status,
    tracking_status,
    IFNULL(priority, '') AS priority,
    IFNULL(title, '') AS title,
    bitrix_deal_id,
    archived
FROM deals
WHERE id = @deal_id
'@ @{
        '@deal_id' = $DealId
    }

    if ($null -eq $table -or $table.Rows.Count -eq 0) {
        return $null
    }

    return $table.Rows[0]
}

function Set-PurchaseDealBitrixDealId {
    param(
        [int]$DealId,
        [AllowNull()][Nullable[int]]$BitrixDealId = $null
    )

    if ($DealId -le 0) {
        return
    }

    Invoke-PurchaseNonQuery @'
UPDATE deals
SET bitrix_deal_id = @bitrix_deal_id,
    updated_at = @updated_at
WHERE id = @deal_id
'@ @{
        '@bitrix_deal_id' = if ($null -eq $BitrixDealId -or $BitrixDealId -le 0) { $null } else { $BitrixDealId }
        '@updated_at' = Get-NowText
        '@deal_id' = $DealId
    }
}

function Set-PurchaseDealImportDefaults {
    param(
        [int]$DealId,
        [string]$Client = ''
    )

    if ($DealId -le 0) {
        return
    }

    Invoke-PurchaseNonQuery @'
UPDATE deals
SET client = CASE WHEN trim(IFNULL(client, '')) = '' AND @client <> '' THEN @client ELSE client END,
    status = 'RFQ',
    tracking_status = 'В работе',
    priority = '',
    title = '',
    updated_at = @updated_at
WHERE id = @deal_id
'@ @{
        '@client' = $Client
        '@updated_at' = Get-NowText
        '@deal_id' = $DealId
    }
}

function Find-OrCreateProcurementFromBitrixTask {
    param($Task)

    if ($null -eq $Task) {
        throw 'Task is required.'
    }

    $parsed = Convert-BitrixTaskTitleToProcurement $Task.Title
    $existingDeal = Get-PurchaseDealByNumber -DealNumber $parsed.RequestNumber
    if ($null -ne $existingDeal) {
        return [pscustomobject]@{
            Procurement = $existingDeal
            Created = $false
            Parsed = $parsed
        }
    }

    $dealId = New-PurchaseDeal -DealNumber $parsed.RequestNumber -Client $parsed.Client -Status 'RFQ'
    Set-PurchaseDealImportDefaults -DealId $dealId -Client $parsed.Client
    $createdDeal = Get-PurchaseDealById -DealId $dealId

    return [pscustomobject]@{
        Procurement = $createdDeal
        Created = $true
        Parsed = $parsed
    }
}

function Find-OrCreateBitrixDeal {
    param(
        $Task,
        $Procurement,
        [scriptblock]$ApiInvoker = $null
    )

    if ($null -eq $Task) {
        throw 'Task is required.'
    }
    if ($null -eq $Procurement) {
        throw 'Procurement is required.'
    }

    $localBitrixDealId = ConvertTo-BitrixNullableInt $Procurement.bitrix_deal_id
    if ($null -ne $localBitrixDealId -and $localBitrixDealId -gt 0) {
        return [pscustomobject]@{
            BitrixDealId = $localBitrixDealId
            Created = $false
        }
    }

    $taskLink = Get-BitrixTaskLink -BitrixTaskId ([int]$Task.Id)
    if ($null -ne $taskLink) {
        $linkedBitrixDealId = ConvertTo-BitrixNullableInt $taskLink.bitrix_deal_id
        if ($null -ne $linkedBitrixDealId -and $linkedBitrixDealId -gt 0) {
            return [pscustomobject]@{
                BitrixDealId = $linkedBitrixDealId
                Created = $false
            }
        }
    }

    $bitrixConfig = Get-BitrixConfig
    $listParams = @{
        entityTypeId = $bitrixConfig.CrmEntityTypeId
        filter = @{
            title = $Task.Title
        }
        select = @('id', 'title', 'categoryId', 'stageId')
        order = @{
            id = 'ASC'
        }
    }
    if ($bitrixConfig.DealCategoryId -ge 0) {
        $listParams.filter.categoryId = $bitrixConfig.DealCategoryId
    }

    $existingResponse = Invoke-BitrixApiInvoker -Method 'crm.item.list' -Params $listParams -TimeoutSec 20 -ApiInvoker $ApiInvoker
    $existingItems = @()
    if ($null -ne $existingResponse -and $null -ne $existingResponse.result -and $null -ne $existingResponse.result.items) {
        $existingItems = @($existingResponse.result.items)
    }
    foreach ($item in $existingItems) {
        $existingId = ConvertTo-BitrixInt $item.id 0
        if ($existingId -gt 0) {
            return [pscustomobject]@{
                BitrixDealId = $existingId
                Created = $false
            }
        }
    }

    $fields = @{
        title = $Task.Title
        categoryId = $bitrixConfig.DealCategoryId
    }
    if (-not [string]::IsNullOrWhiteSpace($bitrixConfig.DealStageId)) {
        $fields.stageId = $bitrixConfig.DealStageId
    }

    $createResponse = Invoke-BitrixApiInvoker -Method 'crm.item.add' -Params @{
        entityTypeId = $bitrixConfig.CrmEntityTypeId
        fields = $fields
    } -TimeoutSec 20 -ApiInvoker $ApiInvoker

    $createdId = 0
    if ($null -ne $createResponse -and $null -ne $createResponse.result) {
        if ($null -ne $createResponse.result.item) {
            $createdId = ConvertTo-BitrixInt $createResponse.result.item.id 0
        }
        if ($createdId -le 0) {
            $createdId = ConvertTo-BitrixInt $createResponse.result.id 0
        }
    }
    if ($createdId -le 0) {
        throw 'Bitrix deal creation did not return an ID.'
    }

    return [pscustomobject]@{
        BitrixDealId = $createdId
        Created = $true
    }
}

function Sync-BitrixTasks {
    param([scriptblock]$ApiInvoker = $null)

    $summary = [pscustomobject]@{
        Imported = 0
        Linked = 0
        Skipped = 0
        Failed = 0
        Errors = New-Object System.Collections.Generic.List[object]
    }

    foreach ($task in @(Get-BitrixNewProjectTasks -ApiInvoker $ApiInvoker)) {
        $existingLink = Get-BitrixTaskLink -BitrixTaskId ([int]$task.Id)
        if ($null -ne $existingLink -and [int]$existingLink.is_manual_deleted -eq 1) {
            Save-BitrixTaskLink -BitrixTaskId ([int]$task.Id) -BitrixStatus ([int]$task.Status) -IsManualDeleted:$true
            $summary.Skipped++
            continue
        }

        try {
            $procurementResult = Find-OrCreateProcurementFromBitrixTask -Task $task
            $procurement = $procurementResult.Procurement
            if ($procurementResult.Created) {
                $summary.Imported++
            }

            Save-BitrixTaskLink -BitrixTaskId ([int]$task.Id) -DealId ([int]$procurement.id) -BitrixStatus ([int]$task.Status) -IsManualDeleted:$false | Out-Null

            $dealResult = Find-OrCreateBitrixDeal -Task $task -Procurement $procurement -ApiInvoker $ApiInvoker
            Set-PurchaseDealBitrixDealId -DealId ([int]$procurement.id) -BitrixDealId ([int]$dealResult.BitrixDealId)
            Save-BitrixTaskLink -BitrixTaskId ([int]$task.Id) -DealId ([int]$procurement.id) -BitrixDealId ([int]$dealResult.BitrixDealId) -BitrixStatus ([int]$task.Status) -IsManualDeleted:$false | Out-Null

            $summary.Linked++
        } catch {
            $summary.Failed++
            [void]$summary.Errors.Add([pscustomobject]@{
                TaskId = [int]$task.Id
                Title = $task.Title
                Message = $_.Exception.Message
            })
        }
    }

    return $summary
}
