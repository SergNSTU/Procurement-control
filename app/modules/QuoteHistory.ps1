# Quote history and Globalist import functions.
# Dot-sourced by RRFQComparer.ps1; no application startup code belongs here.

function Remove-DuplicateQuoteHistory {
    Invoke-PurchaseNonQuery @"
DELETE FROM quote_history
WHERE id NOT IN (
    SELECT MIN(id)
    FROM quote_history
    GROUP BY lower(trim(IFNULL(supplier, ''))), IFNULL(unit_price, -999999999), IFNULL(lead_time, '')
);
DELETE FROM quote_batches
WHERE id NOT IN (SELECT DISTINCT batch_id FROM quote_history);
"@ @{}
}

function Test-QuoteHistoryDuplicate {
    param($Quote)
    $price = if ($null -eq $Quote.UnitPrice) { $null } else { [double]$Quote.UnitPrice }
    $count = Invoke-PurchaseScalar @"
SELECT COUNT(*)
FROM quote_history
WHERE lower(trim(IFNULL(supplier, ''))) = lower(trim(@supplier))
  AND IFNULL(unit_price, -999999999) = IFNULL(@unit_price, -999999999)
  AND IFNULL(lead_time, '') = IFNULL(@lead_time, '')
"@ @{
        '@supplier' = [string]$Quote.Supplier
        '@unit_price' = $price
        '@lead_time' = [string]$Quote.LeadTime
    }
    return ([int]$count -gt 0)
}
function Save-QuoteHistory {
    param($Analysis, [string]$Comment = '')

    if ($null -eq $Analysis) { return 0 }
    Remove-DuplicateQuoteHistory
    $now = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO quote_batches(created_at, rfq_path, priority, comment)
VALUES(@created_at, @rfq_path, @priority, @comment)
'@ @{
        '@created_at' = $now
        '@rfq_path' = [string]$Analysis.RfqPath
        '@priority' = [string]$Analysis.Priority
        '@comment' = $Comment
    }
    $batchId = [int](Invoke-PurchaseScalar 'SELECT id FROM quote_batches ORDER BY id DESC LIMIT 1' @{})
    $winnerIds = @{}
    foreach ($decision in @($Analysis.Decisions)) {
        if ($null -ne $decision.Winner) {
            $winnerIds[$decision.Winner.Id] = $decision
        }
    }

    foreach ($quote in @($Analysis.Quotes)) {
        if (Test-QuoteHistoryDuplicate $quote) { continue }
        $decision = $null
        if ($winnerIds.ContainsKey($quote.Id)) { $decision = $winnerIds[$quote.Id] }
        Invoke-PurchaseNonQuery @'
INSERT INTO quote_history(batch_id, quote_date, rfq_value, pn, supplier, unit_price, lead_time, lead_time_total, mfg, is_winner, warning, sheet_name, row_number, match_status, winner_reason)
VALUES(@batch_id, @quote_date, @rfq_value, @pn, @supplier, @unit_price, @lead_time, @lead_time_total, @mfg, @is_winner, @warning, @sheet_name, @row_number, @match_status, @winner_reason)
'@ @{
            '@batch_id' = $batchId
            '@quote_date' = $now
            '@rfq_value' = [string]$quote.Key
            '@pn' = [string]$quote.PN
            '@supplier' = [string]$quote.Supplier
            '@unit_price' = if ($null -eq $quote.UnitPrice) { $null } else { [double]$quote.UnitPrice }
            '@lead_time' = [string]$quote.LeadTime
            '@lead_time_total' = [string]$quote.LeadTimeTotal
            '@mfg' = if ([string]::IsNullOrWhiteSpace([string]$quote.MfgChina)) { [string]$quote.MfgRussia } else { [string]$quote.MfgChina }
            '@is_winner' = Convert-ToDbBool ($null -ne $decision)
            '@warning' = [string]$quote.Warning
            '@sheet_name' = [string]$quote.SheetName
            '@row_number' = if ($quote.Row) { [int]$quote.Row } else { $null }
            '@match_status' = [string]$quote.MatchStatus
            '@winner_reason' = if ($null -eq $decision) { '' } else { [string]$decision.WinnerReason }
        }
    }
    Write-ActivityLog 'rrfq' $batchId 'Сохранена база квот' "Квот: $(@($Analysis.Quotes).Count)" 0 0
    return $batchId
}

function Convert-GlobalistNumber {
    param([string]$Value)

    $text = ([string]$Value).Replace([char]0x00A0, ' ').Replace(' ', '').Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $text = $text.Replace(',', '.')
    $number = 0.0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    return $null
}

function Import-GlobalistWorkbook {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Файл Globalist не найден: $Path" }
    $excel = Open-ExcelApp
    $wb = $null
    $now = Get-NowText
    $records = New-Object System.Collections.ArrayList
    try {
        $wb = $excel.Workbooks.Open($Path, $false, $true)
        foreach ($ws in @($wb.Worksheets)) {
            $used = $ws.UsedRange
            $values = $used.Value2
            $firstRow = [int]$used.Row
            $firstCol = [int]$used.Column
            $rowCount = [int]$used.Rows.Count
            $colCount = [int]$used.Columns.Count
            $isArray = ($null -ne $values -and $values.GetType().IsArray)
            $getCell = {
                param([int]$row, [int]$col)
                if ($row -lt $firstRow -or $row -ge ($firstRow + $rowCount) -or $col -lt $firstCol -or $col -ge ($firstCol + $colCount)) { return '' }
                if (-not $isArray) { return [string]$values }
                return [string]$values.GetValue($values.GetLowerBound(0) + ($row - $firstRow), $values.GetLowerBound(1) + ($col - $firstCol))
            }
            $headerRow = 0
            $columns = @{}
            for ($r = $firstRow; $r -lt [Math]::Min($firstRow + $rowCount, $firstRow + 25); $r++) {
                $candidate = @{}
                for ($c = $firstCol; $c -lt ($firstCol + $colCount); $c++) {
                    $header = Normalize-Key (&$getCell $r $c)
                    if (-not [string]::IsNullOrWhiteSpace($header)) { $candidate[$header] = $c }
                }
                if ($candidate.ContainsKey('PN') -and $candidate.ContainsKey('PRICE')) { $headerRow = $r; $columns = $candidate; break }
            }
            if ($headerRow -eq 0) { continue }
            for ($r = $headerRow + 1; $r -lt ($firstRow + $rowCount); $r++) {
                $pn = (&$getCell $r $columns['PN']).Trim()
                if ([string]::IsNullOrWhiteSpace($pn)) { continue }
                $get = { param([string]$name) if ($columns.ContainsKey($name)) { return (&$getCell $r $columns[$name]).Trim() } return '' }
                [void]$records.Add([pscustomobject]@{
                    ImportedAt = $now; SourceFile = $Path; SheetName = [string]$ws.Name; RowNumber = $r
                    Factory = &$get 'FACTORY'; PN = $pn; Comment = &$get 'COMMENT'; PiNumber = &$get 'PI NUMBER'
                    Replacement = &$get 'REPLACEMENT'; ChineseRemark = &$get 'CHINESE REMARK'; Package = &$get 'PACKAGE'
                    Brand = &$get 'BRAND'; Datacode = &$get 'DATACODE (DC)'; Moq = &$get 'MINIMUM ORDER QUANTITY (MOQ)'
                    Qty = &$get 'QNTY'; Stock = &$get 'QTY ON STOCK'; NeedSpq = &$get 'NEED SPQ'; Spq = &$get 'SPQ'
                    UnitPrice = Convert-GlobalistNumber (&$get 'PRICE'); TotalAmount = Convert-GlobalistNumber (&$get 'TOTAL AMOUNT')
                    LeadTime = &$get 'LEAD TIME (LT), WKS'; Weight = &$get 'WEIGHT, G'; Target = &$get 'TARGET'; SupplierQuoteId = &$get 'SUPPLIERS QUOTE ID'
                })
            }
            Release-ComObject $used
            Release-ComObject $ws
        }
    } finally {
        if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
        Release-ComObject $wb
    }
    Invoke-PurchaseNonQuery 'DELETE FROM globalist_quotes' @{}
    foreach ($record in $records) {
Invoke-PurchaseNonQuery @'
INSERT INTO globalist_quotes(imported_at, source_file, sheet_name, row_number, factory, pn, comment, pi_number, replacement, chinese_remark, package, brand, datacode, moq, qty, stock, need_spq, spq, unit_price, total_amount, lead_time, weight, target, supplier_quote_id)
VALUES(@imported_at, @source_file, @sheet_name, @row_number, @factory, @pn, @comment, @pi_number, @replacement, @chinese_remark, @package, @brand, @datacode, @moq, @qty, @stock, @need_spq, @spq, @unit_price, @total_amount, @lead_time, @weight, @target, @supplier_quote_id)
'@ @{
            '@imported_at'=$record.ImportedAt; '@source_file'=$record.SourceFile; '@sheet_name'=$record.SheetName; '@row_number'=$record.RowNumber; '@factory'=$record.Factory; '@pn'=$record.PN; '@comment'=$record.Comment; '@pi_number'=$record.PiNumber; '@replacement'=$record.Replacement; '@chinese_remark'=$record.ChineseRemark; '@package'=$record.Package; '@brand'=$record.Brand; '@datacode'=$record.Datacode; '@moq'=$record.Moq; '@qty'=$record.Qty; '@stock'=$record.Stock; '@need_spq'=$record.NeedSpq; '@spq'=$record.Spq; '@unit_price'=$record.UnitPrice; '@total_amount'=$record.TotalAmount; '@lead_time'=$record.LeadTime; '@weight'=$record.Weight; '@target'=$record.Target; '@supplier_quote_id'=$record.SupplierQuoteId
        }
    }
    return $records.Count
}

function Get-GlobalistQuotes {
    param([string]$Search = '', [int]$Limit = 10000)
    return Invoke-PurchaseQuery @'
SELECT * FROM globalist_quotes
WHERE (@search = '' OR IFNULL(pn, '') LIKE @needle OR IFNULL(brand, '') LIKE @needle OR IFNULL(factory, '') LIKE @needle OR IFNULL(pi_number, '') LIKE @needle)
ORDER BY id DESC
LIMIT @limit
'@ @{ '@search'=$Search; '@needle'="%$Search%"; '@limit'=[Math]::Max(1,$Limit) }
}
function Get-QuoteHistory {
    param([string]$Search = '', [string]$DateFrom = '', [string]$DateTo = '', [int]$Limit = 1000)

    $where = 'WHERE (@search = '''' OR IFNULL(q.rfq_value, '''') LIKE @needle OR IFNULL(q.pn, '''') LIKE @needle OR IFNULL(q.supplier, '''') LIKE @needle OR IFNULL(q.mfg, '''') LIKE @needle)'
    if (-not [string]::IsNullOrWhiteSpace($DateFrom)) {
        $where += ' AND substr(q.quote_date, 1, 10) >= @date_from'
    }
    if (-not [string]::IsNullOrWhiteSpace($DateTo)) {
        $where += ' AND substr(q.quote_date, 1, 10) <= @date_to'
    }
    return Invoke-PurchaseQuery @"
SELECT q.*, b.rfq_path, b.priority
FROM quote_history q
JOIN quote_batches b ON b.id = q.batch_id
$where
ORDER BY q.quote_date DESC, q.id DESC
LIMIT @limit
"@ @{
        '@search' = $Search
        '@needle' = "%$Search%"
        '@date_from' = $DateFrom
        '@date_to' = $DateTo
        '@limit' = [Math]::Max(1, $Limit)
    }
}

function Remove-QuoteHistoryItems {
    param([int[]]$Ids)

    $validIds = @($Ids | Where-Object { $_ -gt 0 } | Select-Object -Unique)
    if ($validIds.Count -eq 0) { return 0 }
    $quoteRows = New-Object System.Data.DataTable
    $batchRows = New-Object System.Data.DataTable
    foreach ($id in $validIds) {
        $row = Invoke-PurchaseQuery 'SELECT * FROM quote_history WHERE id = @id' @{ '@id' = $id }
        if ($row.Rows.Count -gt 0) { [void]$quoteRows.Merge($row) }
    }
    foreach ($batchId in @($quoteRows.Rows | ForEach-Object { [int]$_.batch_id } | Select-Object -Unique)) {
        $row = Invoke-PurchaseQuery 'SELECT * FROM quote_batches WHERE id = @id' @{ '@id' = $batchId }
        if ($row.Rows.Count -gt 0) { [void]$batchRows.Merge($row) }
    }
    [void](New-PurchaseTrashItem 'quote_history' 0 ('Квоты: ' + $validIds.Count) @{
        quote_batches = Convert-PurchaseRowsToPayload $batchRows
        quote_history = Convert-PurchaseRowsToPayload $quoteRows
    })
    foreach ($id in $validIds) {
        Invoke-PurchaseNonQuery 'DELETE FROM quote_history WHERE id = @id' @{ '@id' = $id }
    }
    Invoke-PurchaseNonQuery @'
DELETE FROM quote_batches
WHERE id NOT IN (SELECT DISTINCT batch_id FROM quote_history)
'@ @{}
    Write-ActivityLog 'quote_history' 0 'Удалены квоты из базы квот' ("Количество: {0}" -f $validIds.Count)
    return $validIds.Count
}

function Clear-QuoteHistory {
    $quoteRows = Invoke-PurchaseQuery 'SELECT * FROM quote_history' @{}
    $batchRows = Invoke-PurchaseQuery 'SELECT * FROM quote_batches' @{}
    [void](New-PurchaseTrashItem 'quote_history' 0 'Очищена база квот' @{
        quote_batches = Convert-PurchaseRowsToPayload $batchRows
        quote_history = Convert-PurchaseRowsToPayload $quoteRows
    })
    Invoke-PurchaseNonQuery 'DELETE FROM quote_history' @{}
    Invoke-PurchaseNonQuery 'DELETE FROM quote_batches' @{}
}

function Clear-ActivityLog {
    $rows = Invoke-PurchaseQuery 'SELECT * FROM activity_log' @{}
    [void](New-PurchaseTrashItem 'activity_log' 0 'Очищен журнал действий' @{
        activity_log = Convert-PurchaseRowsToPayload $rows
    })
    Invoke-PurchaseNonQuery 'DELETE FROM activity_log' @{}
}
