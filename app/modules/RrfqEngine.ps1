# RRFQ analysis engine: Excel, PDF, quote matching, and result workbook generation.
# Dot-sourced by RRFQComparer.ps1; no application startup code belongs here.

function Release-ComObject {
    param($Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
    }
}

function Open-ExcelApp {
    if ($null -ne $script:Excel) {
        return $script:Excel
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 3
    $script:Excel = $excel
    return $excel
}

function Close-ExcelApp {
    if ($null -ne $script:Excel) {
        try {
            $script:Excel.Quit()
        } finally {
            Release-ComObject $script:Excel
            $script:Excel = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}

function Normalize-Header {
    param($Value)
    $text = ([string]$Value).Replace([char]0x00A0, ' ').Trim().ToLowerInvariant()
    $text = $text -replace '[\r\n\t]+', ' '
    $text = $text -replace '\s+', ' '
    return $text
}

function Normalize-Key {
    param($Value)
    $text = ([string]$Value).Replace([char]0x00A0, ' ').Trim().ToUpperInvariant()
    $text = $text -replace '[\r\n\t]+', ' '
    $text = $text -replace '\s+', ' '
    return $text
}

function Get-CellText {
    param($Worksheet, [int]$Row, [int]$Column)

    $cell = $Worksheet.Cells.Item($Row, $Column)
    $value = $cell.Value2
    if ($null -eq $value) {
        return ([string]$cell.Text).Trim()
    }

    return ([string]$cell.Text).Trim()
}

function Get-CellRawText {
    param($Worksheet, [int]$Row, [int]$Column)

    $cell = $Worksheet.Cells.Item($Row, $Column)
    $value = $cell.Value2
    if ($null -eq $value) {
        return ''
    }

    return ([string]$value).Trim()
}

function Get-CellNumber {
    param($Worksheet, [int]$Row, [int]$Column)

    if ($Column -le 0) {
        return $null
    }

    $cell = $Worksheet.Cells.Item($Row, $Column)
    $value = $cell.Value2
    if ($value -is [double] -or $value -is [int] -or $value -is [decimal]) {
        return [double]$value
    }

    $text = ([string]$cell.Text).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $text = $text -replace '[\$€?]', ''
    $text = $text -replace '(?i)usd|eur|rub|rur', ''
    $text = $text.Replace([char]0x00A0, ' ')
    $text = $text -replace '\s+', ''
    if ($text -match ',' -and $text -notmatch '\.') {
        $text = $text -replace ',', '.'
    } elseif ($text -match ',' -and $text -match '\.') {
        $text = $text -replace ',', ''
    }

    $number = 0.0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }

    return $null
}

function Get-UsedBounds {
    param($Worksheet)

    $used = $Worksheet.UsedRange
    [pscustomobject]@{
        FirstRow = [int]$used.Row
        FirstCol = [int]$used.Column
        LastRow  = [int]($used.Row + $used.Rows.Count - 1)
        LastCol  = [int]($used.Column + $used.Columns.Count - 1)
    }
}

function Test-IsSheetData {
    param($Source)
    return ($null -ne $Source -and $null -ne $Source.PSObject.Properties['Values'])
}

function New-SheetData {
    param($Worksheet)

    $used = $Worksheet.UsedRange
    $values = $used.Value2
    $firstRow = [int]$used.Row
    $firstCol = [int]$used.Column
    $rowCount = [int]$used.Rows.Count
    $colCount = [int]$used.Columns.Count
    $lastRow = $firstRow + $rowCount - 1
    $lastCol = $firstCol + $colCount - 1

    $hiddenRows = @{}
    for ($row = $firstRow; $row -le $lastRow; $row++) {
        $hiddenRows[$row] = [bool]$Worksheet.Rows.Item($row).Hidden
    }

    $hiddenCols = @{}
    for ($col = $firstCol; $col -le $lastCol; $col++) {
        $hiddenCols[$col] = [bool]$Worksheet.Columns.Item($col).Hidden
    }

    $isArray = ($null -ne $values -and $values.GetType().IsArray)
    [pscustomobject]@{
        Name = [string]$Worksheet.Name
        Index = [int]$Worksheet.Index
        FirstRow = $firstRow
        FirstCol = $firstCol
        LastRow = $lastRow
        LastCol = $lastCol
        RowCount = $rowCount
        ColCount = $colCount
        Values = $values
        IsArray = $isArray
        LowerRow = if ($isArray) { [int]$values.GetLowerBound(0) } else { 0 }
        LowerCol = if ($isArray) { [int]$values.GetLowerBound(1) } else { 0 }
        HiddenRows = $hiddenRows
        HiddenCols = $hiddenCols
    }
}

function Get-SheetBounds {
    param($Source)

    if (Test-IsSheetData $Source) {
        return [pscustomobject]@{
            FirstRow = [int]$Source.FirstRow
            FirstCol = [int]$Source.FirstCol
            LastRow = [int]$Source.LastRow
            LastCol = [int]$Source.LastCol
        }
    }

    return Get-UsedBounds $Source
}

function Test-SheetRowHidden {
    param($Source, [int]$Row)

    if (Test-IsSheetData $Source) {
        return ($Source.HiddenRows.ContainsKey($Row) -and [bool]$Source.HiddenRows[$Row])
    }

    return [bool]$Source.Rows.Item($Row).Hidden
}

function Test-SheetColumnHidden {
    param($Source, [int]$Column)

    if (Test-IsSheetData $Source) {
        return ($Source.HiddenCols.ContainsKey($Column) -and [bool]$Source.HiddenCols[$Column])
    }

    return [bool]$Source.Columns.Item($Column).Hidden
}

function Get-SheetCellValue {
    param($Source, [int]$Row, [int]$Column)

    if (-not (Test-IsSheetData $Source)) {
        return $Source.Cells.Item($Row, $Column).Value2
    }

    if ($Row -lt $Source.FirstRow -or $Row -gt $Source.LastRow -or $Column -lt $Source.FirstCol -or $Column -gt $Source.LastCol) {
        return $null
    }

    if (-not $Source.IsArray) {
        return $Source.Values
    }

    $rowIndex = [int]($Source.LowerRow + ($Row - $Source.FirstRow))
    $colIndex = [int]($Source.LowerCol + ($Column - $Source.FirstCol))
    return $Source.Values.GetValue($rowIndex, $colIndex)
}

function Get-SheetCellText {
    param($Source, [int]$Row, [int]$Column)

    if (-not (Test-IsSheetData $Source)) {
        return Get-CellText $Source $Row $Column
    }

    $value = Get-SheetCellValue $Source $Row $Column
    if ($null -eq $value) {
        return ''
    }

    return ([string]$value).Trim()
}

function Convert-ToNumber {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) {
        return [double]$Value
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $text = $text -replace '[\$€?]', ''
    $text = $text -replace '(?i)usd|eur|rub|rur', ''
    $text = $text.Replace([char]0x00A0, ' ')
    $text = $text -replace '\s+', ''
    if ($text -match ',' -and $text -notmatch '\.') {
        $text = $text -replace ',', '.'
    } elseif ($text -match ',' -and $text -match '\.') {
        $text = $text -replace ',', ''
    }

    $number = 0.0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }

    return $null
}

function Format-PurchaseAmount {
    param($Value)

    if ($null -eq $Value -or $Value -is [DBNull]) {
        return ''
    }

    try {
        return ([double]$Value).ToString('#,##0.00', [Globalization.CultureInfo]::InvariantCulture)
    } catch {
        return ''
    }
}

function Convert-ToNullableAmount {
    param($Value)

    if ($null -eq $Value -or $Value -is [DBNull]) {
        return $null
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return Convert-ToNumber $text
}

function Convert-ToNullableInt {
    param($Value)

    if ($null -eq $Value -or $Value -is [DBNull]) {
        return $null
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    if ($text -notmatch '^\d+$') {
        throw 'Срок должен быть целым числом недель.'
    }

    return [int]$text
}

function Convert-PurchaseDateText {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $text = $Value.Trim()
    $cultures = @(
        [Globalization.CultureInfo]::GetCultureInfo('ru-RU'),
        [Globalization.CultureInfo]::InvariantCulture
    )
    $formats = @(
        'dd.MM.yyyy',
        'd.M.yyyy',
        'yyyy-MM-dd',
        'yyyy-MM-dd HH:mm:ss',
        'dd/MM/yyyy',
        'd/M/yyyy'
    )

    foreach ($culture in $cultures) {
        foreach ($format in $formats) {
            $dt = [DateTime]::MinValue
            if ([DateTime]::TryParseExact($text, $format, $culture, [Globalization.DateTimeStyles]::None, [ref]$dt)) {
                return $dt.Date
            }
        }
    }

    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParse($text, [Globalization.CultureInfo]::GetCultureInfo('ru-RU'), [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.Date
    }

    return $null
}

function Format-PurchaseDate {
    param($Value)

    if ($null -eq $Value -or $Value -is [DBNull]) {
        return ''
    }

    if ($Value -is [DateTime]) {
        return $Value.ToString('dd.MM.yyyy')
    }

    $parsed = Convert-PurchaseDateText ([string]$Value)
    if ($null -eq $parsed) {
        return ([string]$Value).Trim()
    }

    return $parsed.ToString('dd.MM.yyyy')
}

function Get-AutoReceiptDate {
    param([string]$InvoiceConfirmedDate, $DeliveryWeeks, [string]$Supplier = '')

    $date = Convert-PurchaseDateText $InvoiceConfirmedDate
    $weeks = Convert-ToNullableInt $DeliveryWeeks
    if ($null -eq $date -or $null -eq $weeks) {
        return ''
    }

    if (Test-CompelSupplier $Supplier) {
        return $date.AddDays(([int]$weeks) * 7).ToString('dd.MM.yyyy')
    }

    $extraWeeks = if (Test-BenSupplier $Supplier) { 1 } else { 2 }
    $target = $date.AddDays(([int]$weeks + $extraWeeks) * 7)
    $daysToTuesday = ([int][DayOfWeek]::Tuesday - [int]$target.DayOfWeek + 7) % 7
    $target = $target.AddDays($daysToTuesday)

    return $target.ToString('dd.MM.yyyy')
}

function Test-SkipPiPdfAmountExtraction {
    param([string]$Supplier)

    $norm = Normalize-Key $Supplier
    return ($norm -match 'BEN|БЕН|COMPEL|КОМПЭЛ|КОМПЕЛ|PROM|ПРОМЭЛЕКТРОНИКА|ПРОМЕЛЕКТРОНИКА')
}

function Decode-PdfLiteralString {
    param([string]$Text)

    $result = $Text
    $result = $result -replace '\\n', "`n"
    $result = $result -replace '\\r', "`r"
    $result = $result -replace '\\t', "`t"
    $result = $result -replace '\\\(', '('
    $result = $result -replace '\\\)', ')'
    $result = $result -replace '\\\\', '\'
    return $result
}

function Get-PdfTextFallback {
    param([string]$Path)

    try {
        $bytes = [IO.File]::ReadAllBytes($Path)
        $raw = [Text.Encoding]::GetEncoding(28591).GetString($bytes)
        $parts = New-Object System.Collections.ArrayList
        foreach ($match in [regex]::Matches($raw, '\((?<text>(?:\\.|[^\\)])*)\)\s*Tj')) {
            [void]$parts.Add((Decode-PdfLiteralString $match.Groups['text'].Value))
        }
        foreach ($match in [regex]::Matches($raw, '\[(?<items>.*?)\]\s*TJ', [Text.RegularExpressions.RegexOptions]::Singleline)) {
            foreach ($item in [regex]::Matches($match.Groups['items'].Value, '\((?<text>(?:\\.|[^\\)])*)\)')) {
                [void]$parts.Add((Decode-PdfLiteralString $item.Groups['text'].Value))
            }
        }

        return ([string]::Join("`r`n", @($parts)))
    } catch {
        return ''
    }
}

function Get-PdfText {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return ''
    }

    if ([IO.Path]::GetExtension($Path).ToLowerInvariant() -ne '.pdf') {
        return ''
    }

    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $confirmConversions = $false
        $readOnly = $true
        $addToRecentFiles = $false
        $doc = $word.Documents.Open([ref]$Path, [ref]$confirmConversions, [ref]$readOnly, [ref]$addToRecentFiles)
        return [string]$doc.Content.Text
    } catch {
        return Get-PdfTextFallback $Path
    } finally {
        if ($null -ne $doc) {
            try { $doc.Close($false) } catch {}
            Release-ComObject $doc
        }
        if ($null -ne $word) {
            try { $word.Quit() } catch {}
            Release-ComObject $word
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Select-CurrencyAmountFromText {
    param([string]$Text, [string]$Currency)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    if ($Currency -eq 'USD') {
        $patterns = @(
            '(?i)(?:USD|US\$|\$)\s*(?<amount>[0-9][0-9\s,]*(?:\.[0-9]{1,4})?)',
            '(?i)(?<amount>[0-9][0-9\s,]*(?:\.[0-9]{1,4})?)\s*(?:USD|US\$|\$)'
        )
    } else {
        $patterns = @(
            '(?i)(?:CNY|RMB|CN?|?|?)\s*(?<amount>[0-9][0-9\s,]*(?:\.[0-9]{1,4})?)',
            '(?i)(?<amount>[0-9][0-9\s,]*(?:\.[0-9]{1,4})?)\s*(?:CNY|RMB|CN?|?|?)'
        )
    }

    $contextPattern = '(?i)grand\s*total|invoice\s*total|amount\s*due|total\s*amount|total\s*value|balance\s*due|??|??|???|????|??|??'
    $candidates = New-Object System.Collections.ArrayList
    $lines = $Text -split '\r?\n'
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = ([string]$lines[$i]).Replace([char]0x00A0, ' ')
        foreach ($pattern in $patterns) {
            foreach ($match in [regex]::Matches($line, $pattern)) {
                $amount = Convert-ToNumber $match.Groups['amount'].Value
                if ($null -ne $amount) {
                    [void]$candidates.Add([pscustomobject]@{
                        Amount = [double]$amount
                        HasContext = ($line -match $contextPattern)
                        Line = $i
                    })
                }
            }
        }
    }

    if ($candidates.Count -eq 0) {
        return $null
    }

    $contextCandidates = @($candidates | Where-Object { $_.HasContext })
    if ($contextCandidates.Count -gt 0) {
        return [double](@($contextCandidates | Sort-Object Line -Descending | Select-Object -First 1)[0].Amount)
    }

    return [double](@($candidates | Sort-Object Amount -Descending | Select-Object -First 1)[0].Amount)
}

function Get-PiAmountsFromPdf {
    param([string]$Path)

    $text = Get-PdfText $Path
    if ([string]::IsNullOrWhiteSpace($text)) {
        return [pscustomobject]@{ Usd = $null; Cny = $null; Warning = 'Не удалось прочитать текст PDF' }
    }

    $usd = Select-CurrencyAmountFromText $text 'USD'
    $cny = Select-CurrencyAmountFromText $text 'CNY'
    $warning = ''
    if ($null -eq $usd -and $null -eq $cny) {
        $warning = 'В PDF не найдена сумма PI'
    }

    return [pscustomobject]@{ Usd = $usd; Cny = $cny; Warning = $warning }
}

function Get-SheetCellNumber {
    param($Source, [int]$Row, [int]$Column)

    if ($Column -le 0) {
        return $null
    }

    if (-not (Test-IsSheetData $Source)) {
        return Get-CellNumber $Source $Row $Column
    }

    return Convert-ToNumber (Get-SheetCellValue $Source $Row $Column)
}

function Add-HeaderCandidate {
    param(
        [hashtable]$Candidates,
        [string]$Field,
        [int]$Column,
        [bool]$Visible
    )

    if (-not $Candidates.ContainsKey($Field)) {
        $Candidates[$Field] = New-Object System.Collections.ArrayList
    }

    [void]$Candidates[$Field].Add([pscustomobject]@{
        Column = $Column
        Visible = $Visible
    })
}

function Select-HeaderColumn {
    param([hashtable]$Candidates, [string]$Field)

    if (-not $Candidates.ContainsKey($Field)) {
        return $null
    }

    $items = @($Candidates[$Field])
    $visible = @($items | Where-Object { $_.Visible })
    if ($visible.Count -gt 0) {
        return [int]$visible[0].Column
    }

    return [int]$items[0].Column
}

function Get-HeaderTexts {
    param($Worksheet, [int]$Row, [int]$FirstCol, [int]$LastCol)

    $items = New-Object System.Collections.ArrayList
    for ($col = $FirstCol; $col -le $LastCol; $col++) {
        $text = Normalize-Header (Get-SheetCellText $Worksheet $Row $col)
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            [void]$items.Add([pscustomobject]@{
                Column = $col
                Text = $text
                Visible = (-not (Test-SheetColumnHidden $Worksheet $col))
            })
        }
    }

    return @($items)
}

function Test-HasHeader {
    param([object[]]$Headers, [string]$Pattern)
    return [bool](@($Headers | Where-Object { $_.Text -match $Pattern }).Count)
}

function Build-HeaderMapForRow {
    param($Worksheet, [int]$Row, [int]$FirstCol, [int]$LastCol)

    $headers = Get-HeaderTexts $Worksheet $Row $FirstCol $LastCol
    if ($headers.Count -eq 0) {
        return $null
    }

    $type = $null
    if ((Test-HasHeader $headers '^исходное наименование$') -and (Test-HasHeader $headers '^цена за шт')) {
        $type = 'Compel'
    } elseif ((Test-HasHeader $headers '^original inquiry$') -and ((Test-HasHeader $headers '^buy unit') -or (Test-HasHeader $headers '^l/t$'))) {
        $type = 'Delivery'
    } elseif ((Test-HasHeader $headers '^value$') -and (Test-HasHeader $headers '^pn$') -and (Test-HasHeader $headers '^unit price$')) {
        $type = 'Standard'
    } else {
        return $null
    }

    $candidates = @{}
    foreach ($header in $headers) {
        $h = $header.Text
        $col = [int]$header.Column
        $visible = [bool]$header.Visible

        if ($type -eq 'Standard') {
            switch -Regex ($h) {
                '^value$' { Add-HeaderCandidate $candidates 'Value' $col $visible; continue }
                '^russian remark$' { Add-HeaderCandidate $candidates 'RussianRemark' $col $visible; continue }
                '^(china|chinese) remark$' { Add-HeaderCandidate $candidates 'ChinaRemark' $col $visible; continue }
                '^pn$' { Add-HeaderCandidate $candidates 'PN' $col $visible; continue }
                '^d/c$' { Add-HeaderCandidate $candidates 'DC' $col $visible; continue }
                '^mfg from russia$' { Add-HeaderCandidate $candidates 'MfgRussia' $col $visible; continue }
                '^mfg from china$' { Add-HeaderCandidate $candidates 'MfgChina' $col $visible; continue }
                '^q-?ty in packing$|^qty in packing$' { Add-HeaderCandidate $candidates 'QtyPacking' $col $visible; continue }
                '^qty to buy/pcs$' { Add-HeaderCandidate $candidates 'QtyToBuy' $col $visible; continue }
                '^unit price \(с ндс\)$' { Add-HeaderCandidate $candidates 'UnitPriceVat' $col $visible; continue }
                '^unit price$' { Add-HeaderCandidate $candidates 'UnitPrice' $col $visible; continue }
                '^lead time \(total\)$' { Add-HeaderCandidate $candidates 'LeadTimeTotal' $col $visible; continue }
                '^lead time' { Add-HeaderCandidate $candidates 'LeadTime' $col $visible; continue }
                '^supplier$|^поставщик$|^source$' { Add-HeaderCandidate $candidates 'Supplier' $col $visible; continue }
            }
        } elseif ($type -eq 'Delivery') {
            switch -Regex ($h) {
                '^original inquiry$' { Add-HeaderCandidate $candidates 'Value' $col $visible; continue }
                '^russian remark$' { Add-HeaderCandidate $candidates 'RussianRemark' $col $visible; continue }
                '^(china|chinese) remark$' { Add-HeaderCandidate $candidates 'ChinaRemark' $col $visible; continue }
                '^replacement$' { Add-HeaderCandidate $candidates 'PN' $col $visible; continue }
                '^mfg$' { Add-HeaderCandidate $candidates 'MfgChina' $col $visible; continue }
                '^d/c$' { Add-HeaderCandidate $candidates 'DC' $col $visible; continue }
                '^qty in packing$|^q-?ty in packing$' { Add-HeaderCandidate $candidates 'QtyPacking' $col $visible; continue }
                '^qty to buy/pcs$' { Add-HeaderCandidate $candidates 'QtyToBuy' $col $visible; continue }
                '^buy unit' { Add-HeaderCandidate $candidates 'UnitPrice' $col $visible; continue }
                '^lead time \(total\)$' { Add-HeaderCandidate $candidates 'LeadTimeTotal' $col $visible; continue }
                '^l/t$|^lead time' { Add-HeaderCandidate $candidates 'LeadTime' $col $visible; continue }
                '^supplier$|^поставщик$|^source$' { Add-HeaderCandidate $candidates 'Supplier' $col $visible; continue }
            }
        } elseif ($type -eq 'Compel') {
            switch -Regex ($h) {
                '^исходное наименование$' { Add-HeaderCandidate $candidates 'Value' $col $visible; continue }
                '^подобранный товар$' { Add-HeaderCandidate $candidates 'PN' $col $visible; continue }
                '^производитель$' { Add-HeaderCandidate $candidates 'MfgChina' $col $visible; continue }
                '^срок поставки$' { Add-HeaderCandidate $candidates 'LeadTime' $col $visible; continue }
                '^количество$' { Add-HeaderCandidate $candidates 'QtyToBuy' $col $visible; continue }
                '^цена за шт' { Add-HeaderCandidate $candidates 'UnitPriceVat' $col $visible; continue }
                '^валюта$' { Add-HeaderCandidate $candidates 'Currency' $col $visible; continue }
            }
        }
    }

    $fields = @{}
    foreach ($field in $candidates.Keys) {
        $fields[$field] = Select-HeaderColumn $candidates $field
    }

    $score = $fields.Count
    if ($fields.ContainsKey('Value')) { $score += 5 }
    if ($fields.ContainsKey('UnitPrice') -or $fields.ContainsKey('UnitPriceVat')) { $score += 4 }
    if ($fields.ContainsKey('LeadTime')) { $score += 3 }

    return [pscustomobject]@{
        Type = $type
        HeaderRow = $Row
        Fields = $fields
        Score = $score
    }
}

function Find-TableMap {
    param(
        $Worksheet,
        [string[]]$AllowedTypes = @('Standard', 'Delivery', 'Compel')
    )

    $bounds = Get-SheetBounds $Worksheet
    $best = $null
    $scanLastRow = [Math]::Min($bounds.LastRow, $bounds.FirstRow + 180)
    for ($row = $bounds.FirstRow; $row -le $scanLastRow; $row++) {
        $map = Build-HeaderMapForRow $Worksheet $row $bounds.FirstCol $bounds.LastCol
        if ($null -eq $map) {
            continue
        }

        if ($AllowedTypes -notcontains $map.Type) {
            continue
        }

        if ($null -eq $best -or $map.Score -gt $best.Score) {
            $best = $map
        }
    }

    return $best
}

function Get-FieldText {
    param($Worksheet, [hashtable]$Fields, [string]$Field, [int]$Row)
    if (-not $Fields.ContainsKey($Field)) {
        return ''
    }

    return Get-SheetCellText $Worksheet $Row ([int]$Fields[$Field])
}

function Get-FieldRawText {
    param($Worksheet, [hashtable]$Fields, [string]$Field, [int]$Row)
    if (-not $Fields.ContainsKey($Field)) {
        return ''
    }

    return Get-SheetCellText $Worksheet $Row ([int]$Fields[$Field])
}

function Get-FieldNumber {
    param($Worksheet, [hashtable]$Fields, [string]$Field, [int]$Row)
    if (-not $Fields.ContainsKey($Field)) {
        return $null
    }

    return Get-SheetCellNumber $Worksheet $Row ([int]$Fields[$Field])
}

function Convert-LeadTime {
    param([string]$LeadTime)

    $raw = ([string]$LeadTime).Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return [pscustomobject]@{ Days = $null; Total = ''; Warning = '' }
    }

    $s = $raw.ToLowerInvariant()
    $s = $s.Replace([char]0x00A0, ' ')
    if ($s -match 'stock|склад') {
        return [pscustomobject]@{ Days = 0.0; Total = 'stock'; Warning = '' }
    }

    $matches = [regex]::Matches($s, '\d+(?:[\.,]\d+)?')
    if ($matches.Count -eq 0) {
        return [pscustomobject]@{ Days = $null; Total = ''; Warning = "Не распознан срок: $raw" }
    }

    $numbers = @()
    foreach ($match in $matches) {
        $numbers += [double](($match.Value) -replace ',', '.')
    }

    $min = ($numbers | Measure-Object -Minimum).Minimum
    $max = ($numbers | Measure-Object -Maximum).Maximum

    if ($s -match 'week|weeks|w\b|нед') {
        $days = $max * 7
        $totalMin = [int][Math]::Ceiling($min) + 1
        $totalMax = [int][Math]::Ceiling($max) + 2
        return [pscustomobject]@{
            Days = [double]$days
            Total = Format-Weeks $totalMin $totalMax
            Warning = ''
        }
    }

    if ($s -match '\bd\b|day|days|д\b|дн') {
        $days = [double]$max
        if ($days -le 3) {
            return [pscustomobject]@{ Days = $days; Total = '1 week'; Warning = '' }
        }

        $baseWeeks = [int][Math]::Ceiling($days / 7.0)
        return [pscustomobject]@{
            Days = $days
            Total = Format-Weeks ($baseWeeks + 1) ($baseWeeks + 2)
            Warning = ''
        }
    }

    return [pscustomobject]@{ Days = $null; Total = ''; Warning = "Не распознан срок: $raw" }
}

function Test-BenSupplier {
    param([string]$Supplier)

    return ((Normalize-Key $Supplier) -match '^(BEN|БЕН)(_[0-9]+)?$')
}

function Test-CompelSupplier {
    param([string]$Supplier)

    return ((Normalize-Key $Supplier) -match '^(COMPEL|КОМПЭЛ|КОМПЕЛ)(_[0-9]+)?$')
}

function Get-LeadTimeTotalForSupplier {
    param([string]$Supplier, [string]$LeadTime, $LeadInfo)

    if (Test-CompelSupplier $Supplier) {
        if ($null -eq $LeadInfo -or $null -eq $LeadInfo.Days) {
            return $LeadInfo.Total
        }
        if ([double]$LeadInfo.Days -le 0) {
            return 'stock'
        }
        $weeks = [int][Math]::Ceiling([double]$LeadInfo.Days / 5.0)
        return Format-Weeks $weeks $weeks
    }

    if (Test-BenSupplier $Supplier) {
        return $LeadTime
    }

    return $LeadInfo.Total
}

function Format-Weeks {
    param([int]$MinWeeks, [int]$MaxWeeks)

    if ($MinWeeks -le 1 -and $MaxWeeks -le 1) {
        return '1 week'
    }

    if ($MinWeeks -eq $MaxWeeks) {
        return "$MinWeeks weeks"
    }

    return "$MinWeeks-$MaxWeeks weeks"
}

function New-RfqRow {
    param($Worksheet, [string]$WorkbookPath, [object]$Map, [int]$Row)

    $fields = $Map.Fields
    $key = Get-FieldText $Worksheet $fields 'Value' $Row
    if ([string]::IsNullOrWhiteSpace($key)) {
        return $null
    }

    if ((Normalize-Key $key) -match '^(TOTAL|ИТОГО)$') {
        return $null
    }

    $id = '{0}::{1}::{2}' -f $Worksheet.Name, $Row, (Normalize-Key $key)
    [pscustomobject]@{
        Id = $id
        WorkbookPath = $WorkbookPath
        SheetName = [string]$Worksheet.Name
        SheetIndex = [int]$Worksheet.Index
        Format = [string]$Map.Type
        HeaderRow = [int]$Map.HeaderRow
        Row = [int]$Row
        Key = $key
        KeyNorm = Normalize-Key $key
        RussianRemark = Get-FieldText $Worksheet $fields 'RussianRemark' $Row
        ChinaRemark = Get-FieldText $Worksheet $fields 'ChinaRemark' $Row
        PN = Get-FieldText $Worksheet $fields 'PN' $Row
        DC = Get-FieldText $Worksheet $fields 'DC' $Row
        MfgRussia = Get-FieldText $Worksheet $fields 'MfgRussia' $Row
        MfgChina = Get-FieldText $Worksheet $fields 'MfgChina' $Row
        QtyPacking = Get-FieldText $Worksheet $fields 'QtyPacking' $Row
        QtyToBuy = Get-FieldText $Worksheet $fields 'QtyToBuy' $Row
        Fields = $fields
        ManualWinnerId = ''
    }
}

function New-Quote {
    param(
        $Worksheet,
        [string]$WorkbookPath,
        [string]$Supplier,
        [object]$Map,
        [int]$Row,
        [string]$ForcedKey = ''
    )

    $fields = $Map.Fields
    $sheetSupplier = Get-FieldText $Worksheet $fields 'Supplier' $Row
    $quoteSupplier = if ([string]::IsNullOrWhiteSpace($sheetSupplier)) { $Supplier } else { $sheetSupplier.Trim() }
    $key = if ([string]::IsNullOrWhiteSpace($ForcedKey)) { Get-FieldText $Worksheet $fields 'Value' $Row } else { $ForcedKey }
    $pn = Get-FieldText $Worksheet $fields 'PN' $Row
    if (-not [string]::IsNullOrWhiteSpace($key) -and -not [string]::IsNullOrWhiteSpace($pn) -and [string]::Equals($key.Trim(), $pn.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) {
        $pn = ''
    }
    $lead = Get-FieldText $Worksheet $fields 'LeadTime' $Row
    $currency = Get-FieldText $Worksheet $fields 'Currency' $Row

    $unitPrice = Get-FieldNumber $Worksheet $fields 'UnitPrice' $Row
    $unitPriceVat = Get-FieldNumber $Worksheet $fields 'UnitPriceVat' $Row
    if ($Map.Type -eq 'Compel') {
        $unitPrice = $null
        if ($null -ne $unitPriceVat) {
            $unitPrice = [double]$unitPriceVat / [double]$script:Config.VatDivisor
        }
    }

    $leadInfo = Convert-LeadTime $lead
    $leadTotal = Get-LeadTimeTotalForSupplier $quoteSupplier $lead $leadInfo
    $isCurrencyComparable = $true
    $warning = ''
    if (-not [string]::IsNullOrWhiteSpace($currency)) {
        $currencyNorm = (Normalize-Key $currency)
        if ($currencyNorm -ne 'USD') {
            $isCurrencyComparable = $false
            $warning = "Валюта не USD: $currency"
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($leadInfo.Warning)) {
        if ([string]::IsNullOrWhiteSpace($warning)) {
            $warning = $leadInfo.Warning
        } else {
            $warning = "$warning; $($leadInfo.Warning)"
        }
    }

    $quote = [pscustomobject]@{
        Id = [guid]::NewGuid().ToString('N')
        WorkbookPath = $WorkbookPath
        FileName = [IO.Path]::GetFileName($WorkbookPath)
        SheetName = [string]$Worksheet.Name
        SheetIndex = [int]$Worksheet.Index
        Format = [string]$Map.Type
        Row = [int]$Row
        Supplier = $quoteSupplier
        Key = $key
        KeyNorm = Normalize-Key $key
        RussianRemark = Get-FieldText $Worksheet $fields 'RussianRemark' $Row
        ChinaRemark = Get-FieldText $Worksheet $fields 'ChinaRemark' $Row
        PN = $pn
        DC = Get-FieldText $Worksheet $fields 'DC' $Row
        MfgRussia = Get-FieldText $Worksheet $fields 'MfgRussia' $Row
        MfgChina = Get-FieldText $Worksheet $fields 'MfgChina' $Row
        QtyPacking = Get-FieldText $Worksheet $fields 'QtyPacking' $Row
        QtyToBuy = Get-FieldText $Worksheet $fields 'QtyToBuy' $Row
        UnitPrice = $unitPrice
        UnitPriceVat = $unitPriceVat
        LeadTime = $lead
        LeadDays = $leadInfo.Days
        LeadTimeTotal = $leadTotal
        Currency = $currency
        IsPriceComparable = $isCurrencyComparable
        Warning = $warning
        TargetId = $null
        MatchStatus = 'Unmatched'
        CandidateIds = @()
    }

    return $quote
}

function New-ManualQuote {
    param(
        $Decision,
        [string]$Supplier,
        [string]$ChinaRemark,
        [string]$PN,
        [string]$DC,
        [string]$MfgRussia,
        [string]$MfgChina,
        [string]$QtyPacking,
        [string]$QtyToBuy,
        $UnitPrice,
        [string]$LeadTime
    )

    $priceNumber = $null
    if ($null -ne $UnitPrice -and -not [string]::IsNullOrWhiteSpace([string]$UnitPrice)) {
        $priceNumber = Convert-ToNumber $UnitPrice
        if ($null -eq $priceNumber) {
            throw "Не удалось распознать цену: $UnitPrice"
        }
    }

    $supplierName = if ([string]::IsNullOrWhiteSpace($Supplier)) { 'Др.' } else { $Supplier.Trim() }
    $leadInfo = Convert-LeadTime $LeadTime

    [pscustomobject]@{
        Id = [guid]::NewGuid().ToString('N')
        WorkbookPath = ''
        FileName = 'Ручной ввод'
        SheetName = [string]$Decision.SheetName
        SheetIndex = 0
        Format = 'Manual'
        Row = [int]$Decision.Row
        Supplier = $supplierName
        Key = [string]$Decision.Value
        KeyNorm = Normalize-Key $Decision.Value
        RussianRemark = [string]$Decision.RussianRemark
        ChinaRemark = [string]$ChinaRemark
        PN = [string]$PN
        DC = [string]$DC
        MfgRussia = [string]$MfgRussia
        MfgChina = [string]$MfgChina
        QtyPacking = [string]$QtyPacking
        QtyToBuy = [string]$QtyToBuy
        UnitPrice = $priceNumber
        UnitPriceVat = $null
        LeadTime = [string]$LeadTime
        LeadDays = $leadInfo.Days
        LeadTimeTotal = Get-LeadTimeTotalForSupplier $supplierName $LeadTime $leadInfo
        Currency = 'USD'
        IsPriceComparable = $true
        Warning = if ([string]::IsNullOrWhiteSpace($leadInfo.Warning)) { 'Ручной ввод' } else { "Ручной ввод; $($leadInfo.Warning)" }
        TargetId = [string]$Decision.Id
        MatchStatus = 'ManualQuote'
        CandidateIds = @()
    }
}

function Test-QuoteHasAnyData {
    param($Quote)

    if ($null -ne $Quote.UnitPrice) { return $true }
    if ($null -ne $Quote.UnitPriceVat) { return $true }
    return $false
}

function Add-QuoteWarning {
    param($Quote, [string]$Message)

    if ($null -eq $Quote -or [string]::IsNullOrWhiteSpace($Message)) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($Quote.Warning)) {
        $Quote.Warning = $Message
    } elseif (-not ([string]$Quote.Warning).Contains($Message)) {
        $Quote.Warning = "$($Quote.Warning); $Message"
    }
}

function Add-SuspiciousPriceWarnings {
    param([object[]]$QuotesForRow)

    $priced = @($QuotesForRow | Where-Object { Test-QuoteSelectable $_ -and $null -ne $_.UnitPrice -and [double]$_.UnitPrice -gt 0 })
    if ($priced.Count -lt 2) { return }

    foreach ($quote in $priced) {
        $others = @($priced | Where-Object { $_.Id -ne $quote.Id })
        if ($others.Count -eq 0) { continue }
        $price = [double]$quote.UnitPrice
        $minOther = [double](@($others | Sort-Object { [double]$_.UnitPrice } | Select-Object -First 1)[0].UnitPrice)
        $maxOther = [double](@($others | Sort-Object { [double]$_.UnitPrice } -Descending | Select-Object -First 1)[0].UnitPrice)
        if (($minOther -gt 0 -and $price -gt ($minOther * 10)) -or ($price -gt 0 -and $maxOther -gt ($price * 10))) {
            Add-QuoteWarning $quote 'Подозрительная цена: отличается более чем в 10 раз от других поставщиков'
        }
    }
}

function Get-WinnerReason {
    param($Winner, [object[]]$Selectable, [string]$PriorityMode, [bool]$IsManualWinner)

    if ($null -eq $Winner) { return '' }
    if ($IsManualWinner) { return 'ручной выбор' }
    $items = @($Selectable | Where-Object { Test-QuoteSelectable $_ })
    if ($items.Count -le 1) {
        return $(if ($PriorityMode -eq 'LeadTime') { 'лучший срок' } else { 'минимальная цена' })
    }

    if ($PriorityMode -eq 'LeadTime') {
        $winnerLead = if ($null -eq $Winner.LeadDays) { [double]::PositiveInfinity } else { [double]$Winner.LeadDays }
        $sameLead = @($items | Where-Object {
            $lead = if ($null -eq $_.LeadDays) { [double]::PositiveInfinity } else { [double]$_.LeadDays }
            [Math]::Abs($lead - $winnerLead) -lt 0.0001
        })
        if ($sameLead.Count -gt 1) { return 'срок равен, выбрана меньшая цена' }
        return 'лучший срок'
    }

    $winnerPrice = [double]$Winner.UnitPrice
    $samePrice = @($items | Where-Object { $null -ne $_.UnitPrice -and [Math]::Abs(([double]$_.UnitPrice) - $winnerPrice) -lt 0.0000001 })
    if ($samePrice.Count -gt 1) { return 'цена равна, выбран меньший срок' }
    return 'минимальная цена'
}

function Get-RfqRowsFromWorkbook {
    param([string]$Path)

    $excel = Open-ExcelApp
    $wb = $null
    $rows = New-Object System.Collections.ArrayList
    try {
        $wb = $excel.Workbooks.Open($Path, 0, $true)
        foreach ($ws in $wb.Worksheets) {
            $sheet = New-SheetData $ws
            $map = Find-TableMap $sheet @('Standard', 'Delivery')
            if ($null -eq $map) {
                continue
            }

            $bounds = Get-SheetBounds $sheet
            $startRow = $map.HeaderRow + 1
            if ($map.Type -eq 'Standard' -and $startRow -lt 49) {
                $startRow = 49
            }
            for ($row = $startRow; $row -le $bounds.LastRow; $row++) {
                if (Test-SheetRowHidden $sheet $row) {
                    continue
                }

                $rfqRow = New-RfqRow $sheet $Path $map $row
                if ($null -ne $rfqRow) {
                    [void]$rows.Add($rfqRow)
                }
            }
        }
    } finally {
        if ($null -ne $wb) {
            $wb.Close($false)
            Release-ComObject $wb
        }
    }

    return @($rows)
}

function Get-QuotesFromWorkbook {
    param(
        [string]$Path,
        [string]$Supplier
    )

    $excel = Open-ExcelApp
    $wb = $null
    $quotes = New-Object System.Collections.ArrayList
    try {
        $wb = $excel.Workbooks.Open($Path, 0, $true)
        foreach ($ws in $wb.Worksheets) {
            $sheet = New-SheetData $ws
            $map = Find-TableMap $sheet @('Standard', 'Delivery', 'Compel')
            if ($null -eq $map) {
                continue
            }

            $bounds = Get-SheetBounds $sheet
            $lastKey = ''
            $startRow = $map.HeaderRow + 1
            if ($map.Type -eq 'Standard' -and $startRow -lt 49) {
                $startRow = 49
            }
            for ($row = $startRow; $row -le $bounds.LastRow; $row++) {
                if (Test-SheetRowHidden $sheet $row) {
                    continue
                }

                $rowKey = Get-FieldText $sheet $map.Fields 'Value' $row
                if ((Normalize-Key $rowKey) -match '^(TOTAL|ИТОГО)$') {
                    continue
                }

                $forcedKey = ''
                if ($map.Type -eq 'Delivery' -and [string]::IsNullOrWhiteSpace($rowKey) -and -not [string]::IsNullOrWhiteSpace($lastKey)) {
                    $forcedKey = $lastKey
                }

                $quote = New-Quote $sheet $Path $Supplier $map $row $forcedKey
                if (-not (Test-QuoteHasAnyData $quote)) {
                    continue
                }

                if (-not [string]::IsNullOrWhiteSpace($quote.Key)) {
                    $lastKey = $quote.Key
                }

                if ([string]::IsNullOrWhiteSpace($quote.Key)) {
                    $quote.MatchStatus = 'NoKey'
                }

                [void]$quotes.Add($quote)
            }
        }
    } finally {
        if ($null -ne $wb) {
            $wb.Close($false)
            Release-ComObject $wb
        }
    }

    return @($quotes)
}

function Resolve-QuoteTargets {
    param([object[]]$RfqRows, [object[]]$Quotes)

    $byExactRow = @{}
    $byKey = @{}
    foreach ($row in $RfqRows) {
        $rowKey = '{0}::{1}' -f $row.SheetName, $row.Row
        $byExactRow[$rowKey] = $row

        if (-not [string]::IsNullOrWhiteSpace($row.KeyNorm)) {
            if (-not $byKey.ContainsKey($row.KeyNorm)) {
                $byKey[$row.KeyNorm] = New-Object System.Collections.ArrayList
            }
            [void]$byKey[$row.KeyNorm].Add($row)
        }
    }

    foreach ($quote in $Quotes) {
        $quote.TargetId = $null
        $quote.CandidateIds = @()

        if ($quote.Format -eq 'Standard') {
            $rowKey = '{0}::{1}' -f $quote.SheetName, $quote.Row
            if ($byExactRow.ContainsKey($rowKey)) {
                $rowCandidate = $byExactRow[$rowKey]
                if ([string]::IsNullOrWhiteSpace($quote.KeyNorm) -or $quote.KeyNorm -eq $rowCandidate.KeyNorm) {
                    $quote.TargetId = $rowCandidate.Id
                    $quote.MatchStatus = 'AutoRow'
                    if ([string]::IsNullOrWhiteSpace($quote.KeyNorm)) {
                        Add-QuoteWarning $quote 'Value пустой, сопоставлено только по номеру строки'
                    }
                    continue
                }

                Add-QuoteWarning $quote ("Value не совпадает со строкой RFQ {0}: в квоте '{1}', в RFQ '{2}'. Сопоставление по строке пропущено." -f $quote.Row, $quote.Key, $rowCandidate.Key)
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($quote.KeyNorm) -and $byKey.ContainsKey($quote.KeyNorm)) {
            $candidates = @($byKey[$quote.KeyNorm])
            $quote.CandidateIds = @($candidates | ForEach-Object { $_.Id })
            if ($candidates.Count -eq 1) {
                $quote.TargetId = $candidates[0].Id
                $quote.MatchStatus = 'AutoKey'
            } else {
                $quote.MatchStatus = 'Ambiguous'
            }
        } elseif ($quote.MatchStatus -ne 'NoKey') {
            $quote.MatchStatus = 'Unmatched'
        }
    }
}

function Test-QuoteSelectable {
    param($Quote)

    if ($null -eq $Quote.UnitPrice) {
        return $false
    }

    if (-not $Quote.IsPriceComparable) {
        return $false
    }

    return $true
}

function Compare-QuoteForWinner {
    param($A, $B, [string]$PriorityMode)

    if ($null -eq $A) { return $B }
    if ($null -eq $B) { return $A }

    if ($PriorityMode -eq 'LeadTime') {
        $aLead = if ($null -eq $A.LeadDays) { [double]::PositiveInfinity } else { [double]$A.LeadDays }
        $bLead = if ($null -eq $B.LeadDays) { [double]::PositiveInfinity } else { [double]$B.LeadDays }
        if ($aLead -lt $bLead) { return $A }
        if ($bLead -lt $aLead) { return $B }

        if ([double]$A.UnitPrice -le [double]$B.UnitPrice) { return $A }
        return $B
    }

    if ([double]$A.UnitPrice -lt [double]$B.UnitPrice) { return $A }
    if ([double]$B.UnitPrice -lt [double]$A.UnitPrice) { return $B }

    $aLeadPrice = if ($null -eq $A.LeadDays) { [double]::PositiveInfinity } else { [double]$A.LeadDays }
    $bLeadPrice = if ($null -eq $B.LeadDays) { [double]::PositiveInfinity } else { [double]$B.LeadDays }
    if ($aLeadPrice -le $bLeadPrice) { return $A }
    return $B
}

function Build-Decisions {
    param(
        [object[]]$RfqRows,
        [object[]]$Quotes,
        [string]$PriorityMode
    )

    $quotesByTarget = @{}
    foreach ($quote in $Quotes) {
        if ([string]::IsNullOrWhiteSpace($quote.TargetId)) {
            continue
        }

        if (-not $quotesByTarget.ContainsKey($quote.TargetId)) {
            $quotesByTarget[$quote.TargetId] = New-Object System.Collections.ArrayList
        }
        [void]$quotesByTarget[$quote.TargetId].Add($quote)
    }

    $decisions = New-Object System.Collections.ArrayList
    foreach ($row in $RfqRows) {
        $quotesForRow = @()
        if ($quotesByTarget.ContainsKey($row.Id)) {
            $quotesForRow = @($quotesByTarget[$row.Id])
        }

        Add-SuspiciousPriceWarnings $quotesForRow
        $selectable = @($quotesForRow | Where-Object { Test-QuoteSelectable $_ })
        $winner = $null
        foreach ($quote in $selectable) {
            $winner = Compare-QuoteForWinner $winner $quote $PriorityMode
        }

        $manualWinnerId = ''
        if ($null -ne $row.PSObject.Properties['ManualWinnerId']) {
            $manualWinnerId = [string]$row.ManualWinnerId
        }
        $manualWinner = $null
        if (-not [string]::IsNullOrWhiteSpace($manualWinnerId)) {
            $manualWinner = @($quotesForRow | Where-Object { $_.Id -eq $manualWinnerId } | Select-Object -First 1)[0]
            if ($null -ne $manualWinner -and (Test-QuoteSelectable $manualWinner)) {
                $winner = $manualWinner
            }
        }

        $status = 'Нет квот'
        if ($quotesForRow.Count -gt 0 -and $selectable.Count -eq 0) {
            $status = 'Нет сравнимой цены'
        } elseif ($null -ne $winner) {
            $status = 'OK'
        }

        $warnings = @()
        foreach ($quote in $quotesForRow) {
            if (-not [string]::IsNullOrWhiteSpace($quote.Warning)) {
                $warnings += ('{0}: {1}' -f $quote.Supplier, $quote.Warning)
            }
            if ($null -eq $quote.UnitPrice) {
                $warnings += ('{0}: нет цены' -f $quote.Supplier)
            }
        }
        if ($null -ne $manualWinner -and (Test-QuoteSelectable $manualWinner)) {
            $warnings += ('Ручной выбор победителя: {0}' -f $manualWinner.Supplier)
        } elseif (-not [string]::IsNullOrWhiteSpace($manualWinnerId)) {
            $warnings += 'Ручной победитель не найден или без сравнимой цены'
        }

        $isManualWinner = ($null -ne $manualWinner -and $null -ne $winner -and $winner.Id -eq $manualWinner.Id)
        $winnerReason = Get-WinnerReason $winner $selectable $PriorityMode $isManualWinner

        $summary = @($quotesForRow | ForEach-Object {
            $priceText = if ($null -eq $_.UnitPrice) { '-' } else { ('{0:N6}' -f [double]$_.UnitPrice) }
            '{0}: {1}; {2}' -f $_.Supplier, $priceText, $_.LeadTime
        }) -join ' | '

        $resultMfgChina = ''
        if ($null -ne $winner) {
            if (-not [string]::IsNullOrWhiteSpace($winner.MfgChina)) {
                $resultMfgChina = $winner.MfgChina
            } elseif (-not [string]::IsNullOrWhiteSpace($winner.MfgRussia) -and [string]::IsNullOrWhiteSpace($row.MfgRussia)) {
                $resultMfgChina = $winner.MfgRussia
            }
        }

        [void]$decisions.Add([pscustomobject]@{
            Id = $row.Id
            Include = ($null -ne $winner)
            Status = $status
            SheetName = $row.SheetName
            Row = $row.Row
            Key = $row.Key
            Value = $row.Key
            RfqRow = $row
            Quotes = $quotesForRow
            Winner = $winner
            ManualWinnerId = $manualWinnerId
            IsManualWinner = $isManualWinner
            WinnerSupplier = if ($null -eq $winner) { '' } else { $winner.Supplier }
            WinnerReason = $winnerReason
            WinnerPrice = if ($null -eq $winner -or $null -eq $winner.UnitPrice) { $null } else { [double]$winner.UnitPrice }
            WinnerLead = if ($null -eq $winner) { '' } else { $winner.LeadTime }
            WinnerLeadTotal = if ($null -eq $winner) { '' } else { $winner.LeadTimeTotal }
            WinnerPN = if ($null -eq $winner) { '' } else { $winner.PN }
            RussianRemark = $row.RussianRemark
            ChinaRemark = if ($null -eq $winner -or [string]::IsNullOrWhiteSpace($winner.ChinaRemark)) { $row.ChinaRemark } else { $winner.ChinaRemark }
            DC = if ($null -eq $winner -or [string]::IsNullOrWhiteSpace($winner.DC)) { $row.DC } else { $winner.DC }
            MfgRussia = $row.MfgRussia
            MfgChina = if ([string]::IsNullOrWhiteSpace($resultMfgChina)) { $row.MfgChina } else { $resultMfgChina }
            QtyPacking = if ($null -eq $winner -or [string]::IsNullOrWhiteSpace($winner.QtyPacking)) { $row.QtyPacking } else { $winner.QtyPacking }
            QtyToBuy = if ($null -eq $winner -or [string]::IsNullOrWhiteSpace($winner.QtyToBuy)) { $row.QtyToBuy } else { $winner.QtyToBuy }
            QuotesSummary = $summary
            Warning = ($warnings -join '; ')
        })
    }

    return @($decisions)
}

function Invoke-Analysis {
    param(
        [string]$RfqPath,
        [object[]]$Suppliers,
        [string]$PriorityMode
    )

    if (-not (Test-Path -LiteralPath $RfqPath)) {
        throw "RFQ не найден: $RfqPath"
    }

    $rfqRows = Get-RfqRowsFromWorkbook $RfqPath
    $quotes = New-Object System.Collections.ArrayList

    foreach ($supplierFile in $Suppliers) {
        if (-not (Test-Path -LiteralPath $supplierFile.Path)) {
            continue
        }

        foreach ($quote in (Get-QuotesFromWorkbook $supplierFile.Path $supplierFile.Supplier)) {
            [void]$quotes.Add($quote)
        }
    }

    Resolve-QuoteTargets $rfqRows @($quotes)
    $decisions = Build-Decisions $rfqRows @($quotes) $PriorityMode

    $analysis = [pscustomobject]@{
        RfqPath = $RfqPath
        Suppliers = $Suppliers
        RfqRows = $rfqRows
        Quotes = @($quotes)
        Decisions = $decisions
        Priority = $PriorityMode
    }
    try {
        Initialize-PurchaseStore
        [void](Save-QuoteHistory $analysis 'analysis')
    } catch {
    }
    return $analysis
}

function Get-UnresolvedQuotes {
    param($Analysis)
    @($Analysis.Quotes | Where-Object { [string]::IsNullOrWhiteSpace($_.TargetId) -or $_.MatchStatus -eq 'Ambiguous' })
}

function Get-OutputPath {
    param([string]$RfqPath)

    $dir = [IO.Path]::GetDirectoryName($RfqPath)
    $stem = [IO.Path]::GetFileNameWithoutExtension($RfqPath)
    $candidate = Join-Path $dir ($stem + '_RRFQ_result.xlsx')
    if (-not (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    return (Join-Path $dir ($stem + "_RRFQ_result_$stamp.xlsx"))
}

function Ensure-HeaderColumn {
    param(
        $Worksheet,
        [object]$Map,
        [string]$Field,
        [string]$HeaderText,
        [int]$InsertBeforeColumn = 0,
        [int]$InsertAfterColumn = 0
    )

    if ($Map.Fields.ContainsKey($Field)) {
        return $Map
    }

    if ($InsertBeforeColumn -gt 0) {
        $Worksheet.Columns.Item($InsertBeforeColumn).Insert() | Out-Null
        $newCol = $InsertBeforeColumn
    } elseif ($InsertAfterColumn -gt 0) {
        $Worksheet.Columns.Item($InsertAfterColumn + 1).Insert() | Out-Null
        $newCol = $InsertAfterColumn + 1
    } else {
        $bounds = Get-UsedBounds $Worksheet
        $newCol = $bounds.LastCol + 1
    }

    $Worksheet.Cells.Item($Map.HeaderRow, $newCol).Value2 = $HeaderText
    $Worksheet.Columns.Item($newCol).Hidden = $false
    foreach ($key in @($Map.Fields.Keys)) {
        if ([int]$Map.Fields[$key] -ge $newCol) {
            $Map.Fields[$key] = [int]$Map.Fields[$key] + 1
        }
    }
    $Map.Fields[$Field] = $newCol
    return $Map
}

function Clear-ResultRow {
    param($Worksheet, [object]$Map, [int]$Row)

    foreach ($field in @('ChinaRemark', 'PN', 'DC', 'MfgChina', 'QtyPacking', 'UnitPrice', 'LeadTime', 'LeadTimeTotal', 'Supplier', 'UnitPriceVat')) {
        if ($Map.Fields.ContainsKey($field)) {
            $Worksheet.Cells.Item($Row, [int]$Map.Fields[$field]).ClearContents() | Out-Null
        }
    }
}

function Set-CellIfPresent {
    param($Worksheet, [object]$Map, [string]$Field, [int]$Row, $Value)
    if (-not $Map.Fields.ContainsKey($Field)) {
        return
    }

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return
    }

    $Worksheet.Cells.Item($Row, [int]$Map.Fields[$Field]).Value2 = $Value
}

function Write-DecisionToWorksheet {
    param($Worksheet, [object]$Map, $Decision)

    $row = [int]$Decision.Row
    $quote = $Decision.Winner
    if ($null -eq $quote) {
        return
    }

    Set-CellIfPresent $Worksheet $Map 'ChinaRemark' $row $quote.ChinaRemark
    Set-CellIfPresent $Worksheet $Map 'PN' $row $quote.PN
    Set-CellIfPresent $Worksheet $Map 'DC' $row $quote.DC

    if (-not [string]::IsNullOrWhiteSpace($quote.MfgChina)) {
        Set-CellIfPresent $Worksheet $Map 'MfgChina' $row $quote.MfgChina
    } elseif (-not [string]::IsNullOrWhiteSpace($quote.MfgRussia) -and [string]::IsNullOrWhiteSpace($Decision.RfqRow.MfgRussia)) {
        Set-CellIfPresent $Worksheet $Map 'MfgChina' $row $quote.MfgRussia
    }

    $qtyPacking = $quote.QtyPacking
    if (Test-CompelSupplier $quote.Supplier) {
        # The RFQ QtyToBuy cell is formula-driven and must remain untouched.
        # If Compel requires more pieces than the RFQ, show that supplier
        # quantity in QtyPacking so the packing requirement is visible.
        $rfqQtyToBuy = Convert-ToNumber $Decision.RfqRow.QtyToBuy
        $compelQtyToBuy = Convert-ToNumber $quote.QtyToBuy
        if ($null -ne $rfqQtyToBuy -and $null -ne $compelQtyToBuy -and $compelQtyToBuy -gt $rfqQtyToBuy) {
            $qtyPacking = $quote.QtyToBuy
        }
    }
    Set-CellIfPresent $Worksheet $Map 'QtyPacking' $row $qtyPacking

    if ($quote.Supplier -eq 'Компэл' -and $Map.Fields.ContainsKey('UnitPriceVat') -and $null -ne $quote.UnitPriceVat) {
        $vatCol = [int]$Map.Fields['UnitPriceVat']
        $unitCol = [int]$Map.Fields['UnitPrice']
        try {
            $Worksheet.Cells.Item($row, $vatCol).NumberFormat = $Worksheet.Cells.Item($row, $unitCol).NumberFormat
        } catch {
            $Worksheet.Cells.Item($row, $vatCol).NumberFormat = '0.000000'
        }
        $Worksheet.Cells.Item($row, $vatCol).Value2 = [double]$quote.UnitPriceVat
        $vatAddress = $Worksheet.Cells.Item($row, $vatCol).Address($false, $false)
        $Worksheet.Cells.Item($row, $unitCol).Formula = "=$vatAddress/$($script:Config.VatDivisor)"
    } elseif ($Map.Fields.ContainsKey('UnitPrice') -and $null -ne $quote.UnitPrice) {
        $Worksheet.Cells.Item($row, [int]$Map.Fields['UnitPrice']).Value2 = [double]$quote.UnitPrice
    }

    Set-CellIfPresent $Worksheet $Map 'LeadTime' $row $quote.LeadTime
    Set-CellIfPresent $Worksheet $Map 'LeadTimeTotal' $row $quote.LeadTimeTotal
    Set-CellIfPresent $Worksheet $Map 'Supplier' $row $quote.Supplier
}

function Write-ResultWorkbook {
    param($Analysis)

    $outputPath = Get-OutputPath $Analysis.RfqPath
    Copy-Item -LiteralPath $Analysis.RfqPath -Destination $outputPath

    $excel = Open-ExcelApp
    $wb = $null
    $oldCalculation = $null
    try {
        try {
            $oldCalculation = $excel.Calculation
            $excel.Calculation = -4135
            $excel.ScreenUpdating = $false
            $excel.EnableEvents = $false
        } catch {
            $oldCalculation = $null
        }

        $wb = $excel.Workbooks.Open($outputPath, 0, $false)
        foreach ($ws in $wb.Worksheets) {
            $sheetDecisionsAll = @($Analysis.Decisions | Where-Object { $_.SheetName -eq $ws.Name })
            if ($sheetDecisionsAll.Count -eq 0) {
                continue
            }

            $sheetDecisionsWithWork = @($sheetDecisionsAll | Where-Object { $_.Winner -or @($_.Quotes).Count -gt 0 })
            if ($sheetDecisionsWithWork.Count -eq 0) {
                continue
            }

            $sheetData = New-SheetData $ws
            $map = Find-TableMap $sheetData @('Standard', 'Delivery')
            if ($null -eq $map) {
                continue
            }

            if ($map.Fields.ContainsKey('LeadTime')) {
                $map = Ensure-HeaderColumn $ws $map 'LeadTimeTotal' 'Lead time (total)' 0 ([int]$map.Fields['LeadTime'])
            }

            $hasCompel = [bool](@($sheetDecisionsAll | Where-Object { $_.Include -and $null -ne $_.Winner -and $_.Winner.Supplier -eq 'Компэл' }).Count)
            if ($hasCompel -and $map.Fields.ContainsKey('UnitPrice')) {
                $map = Ensure-HeaderColumn $ws $map 'UnitPriceVat' 'Unit price (с НДС)' ([int]$map.Fields['UnitPrice']) 0
            }

            if (-not $map.Fields.ContainsKey('Supplier')) {
                $map = Ensure-HeaderColumn $ws $map 'Supplier' 'Supplier' 0 0
            }

            foreach ($decision in $sheetDecisionsWithWork) {
                Clear-ResultRow $ws $map ([int]$decision.Row)
            }

            foreach ($decision in @($sheetDecisionsWithWork | Where-Object { $_.Include -and $null -ne $_.Winner })) {
                Write-DecisionToWorksheet $ws $map $decision
            }
        }

        if ($null -ne $oldCalculation) {
            $excel.Calculation = $oldCalculation
        }
        $wb.Save()
    } finally {
        if ($null -ne $oldCalculation) {
            try { $excel.Calculation = $oldCalculation } catch { }
        }
        if ($null -ne $wb) {
            $wb.Close($true)
            Release-ComObject $wb
        }
    }

    return $outputPath
}
