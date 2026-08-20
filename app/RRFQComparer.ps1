#requires -Version 5.1
param(
    [switch]$SelfTest,
    [switch]$UiSmokeTest,
    [switch]$ImportOnly,
    [switch]$WriteResult,
    [string]$Rfq,
    [string[]]$SupplierFiles = @(),
    [ValidateSet('Price', 'LeadTime')]
    [string]$Priority = 'Price'
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
} catch {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}

$script:AppRoot = Split-Path -Parent $PSScriptRoot
$script:Config = $null
$script:Excel = $null
$script:LastAnalysis = $null
$script:LastPreviewGrid = $null
$script:SqliteLoaded = $false
$script:PurchaseDbPath = $null
$script:NotificationsEnabled = $true
$script:BitrixIntegrationEnabled = $false

$bitrixIntegrationPath = Join-Path $PSScriptRoot 'BitrixIntegration.ps1'
if (Test-Path -LiteralPath $bitrixIntegrationPath -PathType Leaf) {
    . $bitrixIntegrationPath
}

function Get-AppConfig {
    $configPath = Join-Path $script:AppRoot 'config\suppliers.json'
    if (Test-Path -LiteralPath $configPath) {
        return Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }

    return [pscustomobject]@{
        Suppliers = @(
            [pscustomobject]@{ Name = 'Ben'; Aliases = @('ben', 'бен') },
            [pscustomobject]@{ Name = 'BMZ'; Aliases = @('bmz', 'бмз') },
            [pscustomobject]@{ Name = 'HNN'; Aliases = @('hnn') },
            [pscustomobject]@{ Name = 'CPR'; Aliases = @('cpr') },
            [pscustomobject]@{ Name = 'Компэл'; Aliases = @('compel', 'компэл', 'компел') },
            [pscustomobject]@{ Name = 'Промэлектроника'; Aliases = @('promelec', 'promelektronika', 'промэлектроника') },
            [pscustomobject]@{ Name = 'Др.'; Aliases = @('other', 'др', 'другой') }
        )
        OtherSupplier = 'Др.'
        VatDivisor = 1.22
    }
}

function Get-SupplierNames {
    @($script:Config.Suppliers | ForEach-Object { $_.Name })
}

function Get-SupplierFromFileName {
    param([string]$Path)

    $name = ([IO.Path]::GetFileNameWithoutExtension($Path)).ToLowerInvariant()
    foreach ($supplier in $script:Config.Suppliers) {
        foreach ($alias in @($supplier.Aliases)) {
            if ($name.Contains(([string]$alias).ToLowerInvariant())) {
                return [string]$supplier.Name
            }
        }
    }

    return [string]$script:Config.OtherSupplier
}

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

    Set-CellIfPresent $Worksheet $Map 'QtyPacking' $row $quote.QtyPacking
    if ($quote.Supplier -eq 'Компэл' -or $quote.Format -eq 'Manual') {
        Set-CellIfPresent $Worksheet $Map 'QtyToBuy' $row $quote.QtyToBuy
    }

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

function Load-SqliteProvider {
    if ($script:SqliteLoaded) {
        return
    }

    $sqliteDir = Join-Path $script:AppRoot 'app\lib\sqlite'
    $sqliteDll = Join-Path $sqliteDir 'System.Data.SQLite.dll'
    if (-not (Test-Path -LiteralPath $sqliteDll)) {
        throw "SQLite-компонент не найден: $sqliteDll"
    }

    $archDir = Join-Path $sqliteDir ($(if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }))
    if (Test-Path -LiteralPath $archDir) {
        $env:PATH = "$archDir;$sqliteDir;$env:PATH"
    }

    try {
        Add-Type -Path $sqliteDll
    } catch {
        throw "Не удалось загрузить встроенный SQLite-компонент. Установите Microsoft Visual C++ Redistributable 2015–2022 (x64 и x86), затем перезапустите приложение. Технические детали: $($_.Exception.Message)"
    }
    $script:SqliteLoaded = $true
}

function Get-PurchaseDataRoot {
    return (Join-Path $script:AppRoot 'data\purchase_control')
}

function Get-DefaultPurchaseDocumentsRoot {
    return (Join-Path (Get-PurchaseDataRoot) 'files')
}

function Get-ComponentDataRoot {
    return (Join-Path $script:AppRoot 'data\components')
}

function Get-DefaultComponentFilesRoot {
    return (Join-Path (Get-ComponentDataRoot) 'files')
}

function Get-NotesDataRoot {
    return (Join-Path $script:AppRoot 'data\notes')
}

function Get-DefaultNotesFilePath {
    return (Join-Path (Get-NotesDataRoot) 'notes.txt')
}

function Resolve-PortablePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $text = ([string]$Path).Trim()
    if ([IO.Path]::IsPathRooted($text)) { return $text }
    try {
        return [IO.Path]::GetFullPath((Join-Path $script:AppRoot $text))
    } catch {
        return (Join-Path $script:AppRoot $text)
    }
}

function Convert-ToPortablePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    try {
        $full = [IO.Path]::GetFullPath((Resolve-PortablePath $Path))
        $root = [IO.Path]::GetFullPath($script:AppRoot)
        if (-not $root.EndsWith('\')) { $root = "$root\" }
        if ($full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
            return $full.Substring($root.Length)
        }
    } catch {
    }
    return [string]$Path
}

function Get-MovedPortableTreeCandidate {
    param([string]$StoredPath, [string]$PortableTree)

    if ([string]::IsNullOrWhiteSpace($StoredPath) -or [string]::IsNullOrWhiteSpace($PortableTree)) { return '' }
    $separator = [string][IO.Path]::DirectorySeparatorChar
    $normalizedPath = ([string]$StoredPath).Replace([string][IO.Path]::AltDirectorySeparatorChar, $separator)
    $normalizedTree = ([string]$PortableTree).Trim([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar).Replace([string][IO.Path]::AltDirectorySeparatorChar, $separator)
    if ([string]::IsNullOrWhiteSpace($normalizedTree)) { return '' }

    $marker = "$normalizedTree$separator"
    $index = $normalizedPath.IndexOf($marker, [StringComparison]::OrdinalIgnoreCase)
    if ($index -ge 0) {
        $relative = $normalizedPath.Substring($index)
        return (Join-Path $script:AppRoot $relative)
    }
    if ($normalizedPath.EndsWith($normalizedTree, [StringComparison]::OrdinalIgnoreCase)) {
        return (Join-Path $script:AppRoot $normalizedTree)
    }
    return ''
}

function Resolve-PurchaseStoredPath {
    param([string]$StoredPath, [switch]$Directory)

    if ([string]::IsNullOrWhiteSpace($StoredPath)) { return '' }
    $pathType = if ($Directory) { 'Container' } else { 'Leaf' }
    $path = Resolve-PortablePath $StoredPath
    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType $pathType)) {
        try { return [IO.Path]::GetFullPath($path) } catch { return $path }
    }

    $candidate = Get-MovedPortableTreeCandidate $StoredPath 'data\purchase_control\files'
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType $pathType)) {
        try { return [IO.Path]::GetFullPath($candidate) } catch { return $candidate }
    }
    return $path
}

function Resolve-ComponentStoredPath {
    param([string]$StoredPath)

    if ([string]::IsNullOrWhiteSpace($StoredPath)) { return '' }
    $path = Resolve-PortablePath $StoredPath
    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType Container)) {
        try { return [IO.Path]::GetFullPath($path) } catch { return $path }
    }

    $candidate = Get-MovedPortableTreeCandidate $StoredPath 'data\components\files'
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType Container)) {
        try { return [IO.Path]::GetFullPath($candidate) } catch { return $candidate }
    }
    return $path
}

function Get-PurchaseDocumentsRoot {
    $root = Get-PurchaseSetting 'documents_root'
    if ([string]::IsNullOrWhiteSpace($root)) {
        $root = Convert-ToPortablePath (Get-DefaultPurchaseDocumentsRoot)
        Set-PurchaseSetting 'documents_root' $root
    }

    $movedRoot = Get-MovedPortableTreeCandidate $root 'data\purchase_control\files'
    $resolved = if (-not [string]::IsNullOrWhiteSpace($movedRoot)) { $movedRoot } else { Resolve-PortablePath $root }
    if ([string]::IsNullOrWhiteSpace($resolved)) { $resolved = Get-DefaultPurchaseDocumentsRoot }
    return $resolved
}

function Set-PurchaseDocumentsRoot {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Documents folder is required.' }
    $resolved = Resolve-PortablePath $Path
    New-Item -ItemType Directory -Force -Path $resolved | Out-Null
    Set-PurchaseSetting 'documents_root' (Convert-ToPortablePath $resolved)
    $script:PurchaseDocumentsRoot = $resolved
    return $resolved
}

function Repair-PortablePurchasePaths {
    try {
        [void](Set-PurchaseDocumentsRoot (Get-PurchaseDocumentsRoot))
    } catch {
    }

    $docs = Invoke-PurchaseQuery "SELECT id, stored_path FROM documents WHERE IFNULL(stored_path, '') <> ''" @{}
    foreach ($row in $docs.Rows) {
        $stored = [string]$row.stored_path
        $resolved = Resolve-PurchaseStoredPath $stored
        if ([string]::IsNullOrWhiteSpace($resolved) -or -not (Test-Path -LiteralPath $resolved -PathType Leaf)) { continue }
        $portable = Convert-ToPortablePath $resolved
        if ($portable -ne $stored) {
            Invoke-PurchaseNonQuery 'UPDATE documents SET stored_path = @stored_path WHERE id = @id' @{
                '@stored_path' = $portable
                '@id' = [int]$row.id
            }
        }
    }
}

function Repair-PortableComponentPaths {
    $rows = Invoke-PurchaseQuery "SELECT id, folder_path FROM component_deals WHERE IFNULL(folder_path, '') <> ''" @{}
    foreach ($row in $rows.Rows) {
        $stored = [string]$row.folder_path
        $resolved = Resolve-ComponentStoredPath $stored
        if ([string]::IsNullOrWhiteSpace($resolved) -or -not (Test-Path -LiteralPath $resolved -PathType Container)) { continue }
        $portable = Convert-ToPortablePath $resolved
        if ($portable -ne $stored) {
            Invoke-PurchaseNonQuery 'UPDATE component_deals SET folder_path = @folder_path WHERE id = @id' @{
                '@folder_path' = $portable
                '@id' = [int]$row.id
            }
        }
    }
}

function Get-NotesFilePath {
    if (-not [string]::IsNullOrWhiteSpace([string]$script:CurrentNotesFilePath)) {
        return [string]$script:CurrentNotesFilePath
    }
    return Get-DefaultNotesFilePath
}

function Get-NoteDisplayName {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    return [IO.Path]::GetFileNameWithoutExtension($Path)
}

function Get-SafeNoteFileName {
    param([string]$Title)

    $name = if ([string]::IsNullOrWhiteSpace($Title)) { 'Новая заметка' } else { $Title.Trim() }
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) {
        $name = $name.Replace([string]$char, '_')
    }
    $name = ($name -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'Новая заметка' }
    if (-not $name.EndsWith('.txt', [StringComparison]::OrdinalIgnoreCase)) { $name = "$name.txt" }
    return $name
}

function Get-ProjectNoteFiles {
    $root = Get-NotesDataRoot
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $files = @(Get-ChildItem -LiteralPath $root -Filter '*.txt' -File -ErrorAction SilentlyContinue | Sort-Object BaseName)
    if ($files.Count -eq 0) {
        $defaultPath = Join-Path $root 'Общие инструкции.txt'
        Set-Content -LiteralPath $defaultPath -Value '' -Encoding UTF8 -NoNewline
        $files = @(Get-ChildItem -LiteralPath $root -Filter '*.txt' -File -ErrorAction SilentlyContinue | Sort-Object BaseName)
    }
    return $files
}

function New-ProjectNoteFile {
    param([string]$Title)

    $root = Get-NotesDataRoot
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $fileName = Get-SafeNoteFileName $Title
    $stem = [IO.Path]::GetFileNameWithoutExtension($fileName)
    $extension = [IO.Path]::GetExtension($fileName)
    $path = Join-Path $root $fileName
    for ($index = 2; (Test-Path -LiteralPath $path); $index++) {
        $path = Join-Path $root ('{0} ({1}){2}' -f $stem, $index, $extension)
    }
    Set-Content -LiteralPath $path -Value '' -Encoding UTF8 -NoNewline
    $script:CurrentNotesFilePath = $path
    return $path
}

function Read-ProjectNotes {
    $path = Get-NotesFilePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return '' }
    return Get-Content -LiteralPath $path -Raw -Encoding UTF8
}

function Save-ProjectNotes {
    param([string]$Text)

    $root = Get-NotesDataRoot
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $path = Get-NotesFilePath
    if ([string]::IsNullOrWhiteSpace($path)) { $path = Get-DefaultNotesFilePath }
    Set-Content -LiteralPath $path -Value $Text -Encoding UTF8 -NoNewline
    $script:CurrentNotesFilePath = $path
}
function Open-PurchaseConnection {
    Load-SqliteProvider
    $dataRoot = Get-PurchaseDataRoot
    New-Item -ItemType Directory -Force -Path $dataRoot | Out-Null
    if ([string]::IsNullOrWhiteSpace($script:PurchaseDbPath)) {
        $script:PurchaseDbPath = Join-Path $dataRoot 'purchase_control.sqlite'
    }

    $connection = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$script:PurchaseDbPath;Version=3;Foreign Keys=True;Pooling=True;Journal Mode=WAL;Synchronous=Normal;")
    $connection.Open()
    return $connection
}

function Add-SqliteParameters {
    param($Command, [hashtable]$Parameters)

    foreach ($key in @($Parameters.Keys)) {
        $parameter = $Command.CreateParameter()
        $parameter.ParameterName = $key
        $value = $Parameters[$key]
        $parameter.Value = if ($null -eq $value) { [DBNull]::Value } else { $value }
        [void]$Command.Parameters.Add($parameter)
    }
}

function Invoke-PurchaseNonQuery {
    param([string]$Sql, [hashtable]$Parameters = @{})

    $connection = Open-PurchaseConnection
    try {
        $command = $connection.CreateCommand()
        $command.CommandText = $Sql
        Add-SqliteParameters $command $Parameters
        [void]$command.ExecuteNonQuery()
    } finally {
        if ($null -ne $connection) { $connection.Dispose() }
    }
}

function Invoke-PurchaseScalar {
    param([string]$Sql, [hashtable]$Parameters = @{})

    $connection = Open-PurchaseConnection
    try {
        $command = $connection.CreateCommand()
        $command.CommandText = $Sql
        Add-SqliteParameters $command $Parameters
        $value = $command.ExecuteScalar()
        if ($value -is [DBNull]) { return $null }
        return $value
    } finally {
        if ($null -ne $connection) { $connection.Dispose() }
    }
}

function Invoke-PurchaseQuery {
    param([string]$Sql, [hashtable]$Parameters = @{})

    $connection = Open-PurchaseConnection
    try {
        $command = $connection.CreateCommand()
        $command.CommandText = $Sql
        Add-SqliteParameters $command $Parameters
        $reader = $command.ExecuteReader()
        $table = New-Object System.Data.DataTable
        $table.Load($reader)
        return ,$table
    } finally {
        if ($null -ne $connection) { $connection.Dispose() }
    }
}

function Test-PurchaseColumnExists {
    param([string]$TableName, [string]$ColumnName)

    if ($TableName -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') {
        throw "Некорректное имя таблицы: $TableName"
    }

    $table = Invoke-PurchaseQuery "PRAGMA table_info($TableName)"
    foreach ($row in $table.Rows) {
        if ([string]$row.name -eq $ColumnName) {
            return $true
        }
    }

    return $false
}

function Ensure-PurchaseColumn {
    param([string]$TableName, [string]$ColumnName, [string]$Definition)

    if (-not (Test-PurchaseColumnExists $TableName $ColumnName)) {
        Invoke-PurchaseNonQuery "ALTER TABLE $TableName ADD COLUMN $Definition"
    }
}

function Initialize-PurchaseStore {
    $dataRoot = Get-PurchaseDataRoot
    New-Item -ItemType Directory -Force -Path $dataRoot | Out-Null
    $script:PurchaseDbPath = Join-Path $dataRoot 'purchase_control.sqlite'

    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS bitrix_blocked_tasks (
    bitrix_task_id INTEGER PRIMARY KEY,
    title TEXT,
    blocked_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS deals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_number TEXT NOT NULL UNIQUE,
    board_count TEXT,
    client TEXT,
    period TEXT,
    executor TEXT,
    assembly_location TEXT,
    status TEXT NOT NULL DEFAULT 'RFQ',
    tracking_status TEXT NOT NULL DEFAULT 'Ожидание',
    title TEXT,
    comment TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS bitrix_task_links (
    bitrix_task_id INTEGER PRIMARY KEY,
    deal_id INTEGER,
    bitrix_deal_id INTEGER,
    bitrix_status INTEGER,
    is_manual_deleted INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    linked_at TEXT,
    last_seen_at TEXT,
    deleted_at TEXT,
    FOREIGN KEY(deal_id) REFERENCES deals(id) ON DELETE SET NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS bitrix_rfq_files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bitrix_task_id INTEGER NOT NULL,
    file_key TEXT NOT NULL,
    bitrix_file_id TEXT,
    file_name TEXT,
    file_hash TEXT,
    local_path TEXT,
    upload_state TEXT,
    error_text TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(bitrix_task_id, file_key),
    FOREIGN KEY(bitrix_task_id) REFERENCES bitrix_task_links(bitrix_task_id) ON DELETE CASCADE
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS deal_suppliers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_id INTEGER NOT NULL,
    supplier TEXT NOT NULL,
    invoice_received INTEGER NOT NULL DEFAULT 0,
    invoice_confirmed INTEGER NOT NULL DEFAULT 0,
    supplier_order_created INTEGER NOT NULL DEFAULT 0,
    erp_supplier_sent INTEGER NOT NULL DEFAULT 0,
    erp_roger_sent INTEGER NOT NULL DEFAULT 0,
    pi_amount_usd REAL,
    pi_amount_cny REAL,
    paid_amount TEXT,
    delivery_weeks INTEGER,
    payment_submitted INTEGER NOT NULL DEFAULT 0,
    paid INTEGER NOT NULL DEFAULT 0,
    invoice_confirmed_date TEXT,
    components_receipt_date TEXT,
    actual_receipt_date TEXT,
    comment TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(deal_id, supplier),
    FOREIGN KEY(deal_id) REFERENCES deals(id) ON DELETE CASCADE
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_id INTEGER NOT NULL,
    supplier_id INTEGER,
    document_type TEXT NOT NULL,
    original_name TEXT NOT NULL,
    stored_path TEXT NOT NULL,
    uploaded_at TEXT NOT NULL,
    FOREIGN KEY(deal_id) REFERENCES deals(id) ON DELETE CASCADE,
    FOREIGN KEY(supplier_id) REFERENCES deal_suppliers(id) ON DELETE SET NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS component_deals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT NOT NULL,
    deal_number TEXT,
    status TEXT NOT NULL DEFAULT 'В работе',
    stage TEXT NOT NULL DEFAULT 'Запросил поставщиков',
    description TEXT,
    next_action TEXT,
    reminder_date TEXT,
    deadline_date TEXT,
    priority TEXT NOT NULL DEFAULT '3',
    period TEXT,
    folder_path TEXT,
    notes TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS activity_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_type TEXT NOT NULL,
    entity_id INTEGER,
    deal_id INTEGER,
    supplier_id INTEGER,
    action TEXT NOT NULL,
    details TEXT,
    created_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS reminders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    deal_id INTEGER,
    supplier_id INTEGER,
    component_id INTEGER,
    title TEXT NOT NULL,
    due_date TEXT,
    status TEXT NOT NULL DEFAULT 'Open',
    source TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS notification_state (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source TEXT NOT NULL,
    source_id INTEGER NOT NULL,
    due_kind TEXT NOT NULL,
    due_date TEXT NOT NULL,
    handled INTEGER NOT NULL DEFAULT 0,
    snooze_until TEXT,
    last_shown_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(source, source_id, due_kind, due_date)
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS workflow_templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    is_default INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS workflow_steps (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    template_id INTEGER NOT NULL,
    step_order INTEGER NOT NULL,
    name TEXT NOT NULL,
    FOREIGN KEY(template_id) REFERENCES workflow_templates(id) ON DELETE CASCADE
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS quote_batches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT NOT NULL,
    rfq_path TEXT,
    priority TEXT,
    comment TEXT
);
'@
    Invoke-PurchaseNonQuery @'
CREATE TABLE IF NOT EXISTS quote_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id INTEGER NOT NULL,
    quote_date TEXT NOT NULL,
    rfq_value TEXT,
    pn TEXT,
    supplier TEXT,
    unit_price REAL,
    lead_time TEXT,
    lead_time_total TEXT,
    mfg TEXT,
    is_winner INTEGER NOT NULL DEFAULT 0,
    warning TEXT,
    sheet_name TEXT,
    row_number INTEGER,
    match_status TEXT,
    winner_reason TEXT,
    FOREIGN KEY(batch_id) REFERENCES quote_batches(id) ON DELETE CASCADE
);
'@
    $globalistSchema = @(
        'CREATE TABLE IF NOT EXISTS globalist_quotes (',
        '    id INTEGER PRIMARY KEY AUTOINCREMENT,',
        '    imported_at TEXT NOT NULL,', '    source_file TEXT,', '    sheet_name TEXT,', '    row_number INTEGER,',
        '    factory TEXT,', '    pn TEXT,', '    comment TEXT,', '    pi_number TEXT,', '    replacement TEXT,',
        '    chinese_remark TEXT,', '    package TEXT,', '    brand TEXT,', '    datacode TEXT,', '    moq TEXT,',
        '    qty TEXT,', '    stock TEXT,', '    need_spq TEXT,', '    spq TEXT,', '    unit_price REAL,',
        '    total_amount REAL,', '    lead_time TEXT,', '    weight TEXT,', '    target TEXT,', '    supplier_quote_id TEXT',
        ');'
    ) -join "`n"
    Invoke-PurchaseNonQuery $globalistSchema @{}

    Ensure-PurchaseColumn 'quote_batches' 'source_type' 'source_type TEXT NOT NULL DEFAULT ''Our'''
    Ensure-PurchaseColumn 'deals' 'client' 'client TEXT'
    Ensure-PurchaseColumn 'deals' 'status' "status TEXT NOT NULL DEFAULT 'RFQ'"
    Ensure-PurchaseColumn 'deals' 'tracking_status' "tracking_status TEXT NOT NULL DEFAULT 'Ожидание'"
    Ensure-PurchaseColumn 'deals' 'board_count' 'board_count TEXT'
    Ensure-PurchaseColumn 'deals' 'period' 'period TEXT'
    Ensure-PurchaseColumn 'deals' 'executor' 'executor TEXT'
    Ensure-PurchaseColumn 'deals' 'assembly_location' 'assembly_location TEXT'
    Ensure-PurchaseColumn 'deals' 'bitrix_deal_id' 'bitrix_deal_id INTEGER'
    Ensure-PurchaseColumn 'deals' 'priority' "priority TEXT NOT NULL DEFAULT '3'"
    Ensure-PurchaseColumn 'deals' 'archived' 'archived INTEGER NOT NULL DEFAULT 0'
    Ensure-PurchaseColumn 'deals' 'workflow_template_id' 'workflow_template_id INTEGER'
    Ensure-PurchaseColumn 'deals' 'pi_sent' 'pi_sent INTEGER NOT NULL DEFAULT 0'
    Ensure-PurchaseColumn 'deals' 'masks' 'masks INTEGER NOT NULL DEFAULT 0'
    Ensure-PurchaseColumn 'documents' 'file_hash' 'file_hash TEXT'
    Ensure-PurchaseColumn 'deal_suppliers' 'pi_amount_usd' 'pi_amount_usd REAL'
    Ensure-PurchaseColumn 'deal_suppliers' 'pi_amount_cny' 'pi_amount_cny REAL'
    Ensure-PurchaseColumn 'deal_suppliers' 'pi_amount_rub' 'pi_amount_rub REAL'
    Ensure-PurchaseColumn 'deal_suppliers' 'paid_amount' 'paid_amount TEXT'
    Ensure-PurchaseColumn 'deal_suppliers' 'delivery_weeks' 'delivery_weeks INTEGER'
    Ensure-PurchaseColumn 'deal_suppliers' 'payment_submitted' 'payment_submitted INTEGER NOT NULL DEFAULT 0'
    Ensure-PurchaseColumn 'deal_suppliers' 'paid' 'paid INTEGER NOT NULL DEFAULT 0'
    Ensure-PurchaseColumn 'deal_suppliers' 'actual_receipt_date' 'actual_receipt_date TEXT'
    Ensure-PurchaseColumn 'component_deals' 'folder_path' 'folder_path TEXT'
    Ensure-PurchaseColumn 'component_deals' 'notes' 'notes TEXT'

    Initialize-DefaultWorkflowTemplates

    $documentsRoot = Get-PurchaseDocumentsRoot
    $script:PurchaseDocumentsRoot = $documentsRoot
    New-Item -ItemType Directory -Force -Path $script:PurchaseDocumentsRoot | Out-Null
    New-Item -ItemType Directory -Force -Path (Get-DefaultComponentFilesRoot) | Out-Null
    New-Item -ItemType Directory -Force -Path (Get-NotesDataRoot) | Out-Null
    Repair-PortablePurchasePaths
    Repair-PortableComponentPaths
}

function Get-PurchaseSetting {
    param([string]$Key)

    return [string](Invoke-PurchaseScalar 'SELECT value FROM settings WHERE key = @key' @{ '@key' = $Key })
}

function Set-PurchaseSetting {
    param([string]$Key, [string]$Value)

    Invoke-PurchaseNonQuery 'INSERT INTO settings(key, value) VALUES(@key, @value) ON CONFLICT(key) DO UPDATE SET value = excluded.value' @{
        '@key' = $Key
        '@value' = $Value
    }
}

function Get-BitrixBlockedTaskIds {
    return @(Invoke-PurchaseQuery 'SELECT bitrix_task_id FROM bitrix_blocked_tasks ORDER BY bitrix_task_id' @{} | ForEach-Object { [int]$_.bitrix_task_id })
}

function Block-BitrixTask {
    param([int]$TaskId, [string]$Title)

    if ($TaskId -le 0) { return }
    Invoke-PurchaseNonQuery 'INSERT OR IGNORE INTO bitrix_blocked_tasks(bitrix_task_id, title, blocked_at) VALUES(@task_id, @title, @blocked_at)' @{
        '@task_id' = $TaskId
        '@title' = $Title
        '@blocked_at' = Get-NowText
    }
}

function Get-NowText {
    return (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}

function Initialize-DefaultWorkflowTemplates {
    $exists = [int](Invoke-PurchaseScalar "SELECT COUNT(*) FROM workflow_templates WHERE name = 'Базовая закупка'" @{})
    if ($exists -gt 0) { return }

    $now = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO workflow_templates(name, is_default, created_at)
VALUES('Базовая закупка', 1, @created_at)
'@ @{ '@created_at' = $now }
    $templateId = [int](Invoke-PurchaseScalar "SELECT id FROM workflow_templates WHERE name = 'Базовая закупка'" @{})
    $order = 1
    foreach ($step in @('RFQ', 'RRFQ', 'PO', 'PI', 'Подано в оплату', 'Оплачено', 'ERP', 'Поступление')) {
        Invoke-PurchaseNonQuery @'
INSERT INTO workflow_steps(template_id, step_order, name)
VALUES(@template_id, @step_order, @name)
'@ @{
            '@template_id' = $templateId
            '@step_order' = $order
            '@name' = $step
        }
        $order++
    }
}

function Write-ActivityLog {
    param(
        [string]$EntityType,
        [int]$EntityId = 0,
        [string]$Action,
        [string]$Details = '',
        [int]$DealId = 0,
        [int]$SupplierId = 0
    )

    if ([string]::IsNullOrWhiteSpace($Action)) { return }
    try {
        Invoke-PurchaseNonQuery @'
INSERT INTO activity_log(entity_type, entity_id, deal_id, supplier_id, action, details, created_at)
VALUES(@entity_type, @entity_id, @deal_id, @supplier_id, @action, @details, @created_at)
'@ @{
            '@entity_type' = if ([string]::IsNullOrWhiteSpace($EntityType)) { 'system' } else { $EntityType }
            '@entity_id' = if ($EntityId -gt 0) { $EntityId } else { $null }
            '@deal_id' = if ($DealId -gt 0) { $DealId } else { $null }
            '@supplier_id' = if ($SupplierId -gt 0) { $SupplierId } else { $null }
            '@action' = $Action
            '@details' = $Details
            '@created_at' = Get-NowText
        }
    } catch {
        # The log should never block the user's main workflow.
    }
}

function Get-ActivityLog {
    param([int]$DealId = 0, [int]$Limit = 200, [string]$Search = '')

    $whereParts = New-Object System.Collections.ArrayList
    $params = @{ '@limit' = [Math]::Max(1, $Limit) }
    if ($DealId -gt 0) {
        [void]$whereParts.Add('deal_id = @deal_id')
        $params['@deal_id'] = $DealId
    }
    if (-not [string]::IsNullOrWhiteSpace($Search)) {
        [void]$whereParts.Add("(IFNULL(action, '') LIKE @needle OR IFNULL(details, '') LIKE @needle OR IFNULL(entity_type, '') LIKE @needle OR CAST(IFNULL(entity_id, '') AS TEXT) LIKE @needle OR CAST(IFNULL(deal_id, '') AS TEXT) LIKE @needle OR CAST(IFNULL(supplier_id, '') AS TEXT) LIKE @needle)")
        $params['@needle'] = '%' + $Search.Trim() + '%'
    }
    $where = if ($whereParts.Count -gt 0) { 'WHERE ' + [string]::Join(' AND ', [string[]]$whereParts.ToArray()) } else { '' }

    return Invoke-PurchaseQuery @"
SELECT id, entity_type, entity_id, deal_id, supplier_id, action, IFNULL(details, '') AS details, created_at
FROM activity_log
$where
ORDER BY id DESC
LIMIT @limit
"@ $params
}

function Get-DefaultWorkflowTemplateId {
    $id = Invoke-PurchaseScalar 'SELECT id FROM workflow_templates WHERE is_default = 1 ORDER BY id LIMIT 1' @{}
    if ($null -eq $id -or $id -is [DBNull]) { return $null }
    return [int]$id
}

function Get-WorkflowTemplates {
    return Invoke-PurchaseQuery 'SELECT id, name, is_default, created_at FROM workflow_templates ORDER BY is_default DESC, name' @{}
}

function Get-WorkflowSteps {
    param([int]$TemplateId = 0)
    if ($TemplateId -le 0) {
        $TemplateId = Get-DefaultWorkflowTemplateId
    }
    if ($null -eq $TemplateId -or $TemplateId -le 0) {
        return New-Object System.Data.DataTable
    }
    return Invoke-PurchaseQuery @'
SELECT id, template_id, step_order, name
FROM workflow_steps
WHERE template_id = @template_id
ORDER BY step_order
'@ @{ '@template_id' = $TemplateId }
}

function Set-PurchaseDealArchived {
    param([int]$DealId, [bool]$Archived)

    if ($DealId -le 0) { throw 'Выберите сделку.' }
    Invoke-PurchaseNonQuery 'UPDATE deals SET archived = @archived, updated_at = @updated_at WHERE id = @id' @{
        '@archived' = Convert-ToDbBool $Archived
        '@updated_at' = Get-NowText
        '@id' = $DealId
    }
    Write-ActivityLog 'deal' $DealId $(if ($Archived) { 'Сделка отправлена в архив' } else { 'Сделка возвращена из архива' }) '' $DealId
}

function Set-PurchaseDealWorkflowTemplate {
    param([int]$DealId, [int]$TemplateId)

    if ($DealId -le 0) { throw 'Выберите сделку.' }
    if ($TemplateId -le 0) { throw 'Выберите шаблон этапов.' }
    $templateName = Invoke-PurchaseScalar 'SELECT name FROM workflow_templates WHERE id = @id' @{ '@id' = $TemplateId }
    if ($null -eq $templateName -or $templateName -is [DBNull]) { throw 'Шаблон этапов не найден.' }
    Invoke-PurchaseNonQuery 'UPDATE deals SET workflow_template_id = @template_id, updated_at = @updated_at WHERE id = @deal_id' @{
        '@template_id' = $TemplateId
        '@updated_at' = Get-NowText
        '@deal_id' = $DealId
    }
    Write-ActivityLog 'deal' $DealId 'Изменен шаблон этапов' ([string]$templateName) $DealId
}

function Save-Reminder {
    param(
        [string]$Title,
        [string]$DueDate = '',
        [int]$DealId = 0,
        [int]$SupplierId = 0,
        [int]$ComponentId = 0,
        [string]$Source = 'manual'
    )

    if ([string]::IsNullOrWhiteSpace($Title)) { throw 'Введите текст напоминания.' }
    $now = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO reminders(deal_id, supplier_id, component_id, title, due_date, status, source, created_at, updated_at)
VALUES(@deal_id, @supplier_id, @component_id, @title, @due_date, 'Open', @source, @created_at, @updated_at)
'@ @{
        '@deal_id' = if ($DealId -gt 0) { $DealId } else { $null }
        '@supplier_id' = if ($SupplierId -gt 0) { $SupplierId } else { $null }
        '@component_id' = if ($ComponentId -gt 0) { $ComponentId } else { $null }
        '@title' = $Title
        '@due_date' = $DueDate
        '@source' = $Source
        '@created_at' = $now
        '@updated_at' = $now
    }
    Write-ActivityLog 'reminder' 0 'Создано напоминание' $Title $DealId $SupplierId
}

function Set-ReminderDone {
    param([int]$ReminderId, [bool]$Done)

    if ($ReminderId -le 0) { throw 'Выберите напоминание.' }
    Invoke-PurchaseNonQuery 'UPDATE reminders SET status = @status, updated_at = @updated_at WHERE id = @id' @{
        '@status' = if ($Done) { 'Done' } else { 'Open' }
        '@updated_at' = Get-NowText
        '@id' = $ReminderId
    }
}

function Get-NotificationState {
    param(
        [string]$Source,
        [int]$SourceId,
        [string]$DueKind,
        [string]$DueDate
    )

    return Invoke-PurchaseQuery @'
SELECT id, handled, IFNULL(snooze_until, '') AS snooze_until
FROM notification_state
WHERE source = @source AND source_id = @source_id
  AND due_kind = @due_kind AND due_date = @due_date
'@ @{
        '@source' = $Source
        '@source_id' = $SourceId
        '@due_kind' = $DueKind
        '@due_date' = $DueDate
    }
}

function Get-PurchaseBooleanSetting {
    param(
        [string]$Key,
        [bool]$DefaultValue
    )

    $rawValue = ([string](Get-PurchaseSetting $Key)).Trim().ToLowerInvariant()
    if ($rawValue -in @('1', 'true', 'yes')) { return $true }
    if ($rawValue -in @('0', 'false', 'no')) { return $false }

    Set-PurchaseSetting $Key ($(if ($DefaultValue) { '1' } else { '0' }))
    return $DefaultValue
}

function Set-NotificationState {
    param(
        [string]$Source,
        [int]$SourceId,
        [string]$DueKind,
        [string]$DueDate,
        [bool]$Handled = $false,
        [string]$SnoozeUntil = '',
        [bool]$Shown = $false
    )

    $now = Get-NowText
    $shownAt = if ($Shown) { $now } else { $null }
    Invoke-PurchaseNonQuery @'
INSERT INTO notification_state(source, source_id, due_kind, due_date, handled, snooze_until, last_shown_at, created_at, updated_at)
VALUES(@source, @source_id, @due_kind, @due_date, @handled, @snooze_until, @last_shown_at, @created_at, @updated_at)
ON CONFLICT(source, source_id, due_kind, due_date) DO UPDATE SET
    handled = @handled,
    snooze_until = @snooze_until,
    last_shown_at = CASE WHEN @last_shown_at IS NULL THEN notification_state.last_shown_at ELSE @last_shown_at END,
    updated_at = @updated_at
'@ @{
        '@source' = $Source
        '@source_id' = $SourceId
        '@due_kind' = $DueKind
        '@due_date' = $DueDate
        '@handled' = Convert-ToDbBool $Handled
        '@snooze_until' = if ([string]::IsNullOrWhiteSpace($SnoozeUntil)) { $null } else { $SnoozeUntil }
        '@last_shown_at' = $shownAt
        '@created_at' = $now
        '@updated_at' = $now
    }
}

function Get-DueNotificationCandidates {
    $today = (Get-Date).Date
    $currentTime = Get-Date
    $items = New-Object System.Collections.ArrayList
    $stateByKey = @{}
    $states = Invoke-PurchaseQuery @'
SELECT source, source_id, due_kind, due_date, handled, IFNULL(snooze_until, '') AS snooze_until
FROM notification_state
'@ @{}
    foreach ($state in $states.Rows) {
        $stateKey = '{0}|{1}|{2}|{3}' -f [string]$state.source, [int]$state.source_id, [string]$state.due_kind, [string]$state.due_date
        $stateByKey[$stateKey] = $state
    }

    $components = Invoke-PurchaseQuery @'
SELECT id, IFNULL(deal_number, '') AS deal_number,
       IFNULL(description, '') AS description,
       IFNULL(next_action, '') AS next_action,
       IFNULL(status, '') AS status,
       IFNULL(stage, '') AS stage,
       IFNULL(reminder_date, '') AS reminder_date,
       IFNULL(deadline_date, '') AS deadline_date
FROM component_deals
WHERE IFNULL(status, '') NOT IN ('Выполнено', 'Не актуально')
  AND (IFNULL(reminder_date, '') <> '' OR IFNULL(deadline_date, '') <> '')
ORDER BY id DESC
'@ @{}

    foreach ($row in $components.Rows) {
        $due = New-Object System.Collections.ArrayList
        foreach ($entry in @(@('reminder', 'Напоминание', [string]$row.reminder_date), @('deadline', 'Дедлайн', [string]$row.deadline_date))) {
            $date = Convert-PurchaseDateText $entry[2]
            if ($null -eq $date -or $date.Date -gt $today) { continue }
            $dateText = $date.ToString('yyyy-MM-dd')
            $stateKey = 'component|{0}|{1}|{2}' -f [int]$row.id, [string]$entry[0], $dateText
            $stateRow = if ($stateByKey.ContainsKey($stateKey)) { $stateByKey[$stateKey] } else { $null }
            $eligible = $null -eq $stateRow
            if ($null -ne $stateRow) {
                $snooze = [DateTime]::MinValue
                $eligible = (-not (Convert-DbBool $stateRow.handled)) -and
                    ([string]::IsNullOrWhiteSpace([string]$stateRow.snooze_until) -or
                     ([DateTime]::TryParse([string]$stateRow.snooze_until, [ref]$snooze) -and $snooze -le $currentTime))
            }
            if ($eligible) {
                [void]$due.Add([pscustomobject]@{ Kind = [string]$entry[0]; Label = [string]$entry[1]; Date = $date; DateText = $dateText })
            }
        }
        if ($due.Count -eq 0) { continue }
        $name = ([string]$row.deal_number).Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { $name = ([string]$row.description).Trim() }
        if ([string]::IsNullOrWhiteSpace($name)) { $name = 'Задача #' + [string]$row.id }
        [void]$items.Add([pscustomobject]@{
            Source = 'component'; SourceId = [int]$row.id; Title = $name
            Detail = ([string]$row.next_action).Trim(); Due = @($due)
            Status = [string]$row.status; Stage = [string]$row.stage
        })
    }

    $manual = Invoke-PurchaseQuery @'
SELECT id, IFNULL(title, '') AS title, IFNULL(due_date, '') AS due_date,
       IFNULL(deal_id, 0) AS deal_id, IFNULL(supplier_id, 0) AS supplier_id
FROM reminders
WHERE status <> 'Done' AND IFNULL(due_date, '') <> ''
ORDER BY due_date, id
'@ @{}
    foreach ($row in $manual.Rows) {
        $date = Convert-PurchaseDateText ([string]$row.due_date)
        if ($null -eq $date -or $date.Date -gt $today) { continue }
        $dateText = $date.ToString('yyyy-MM-dd')
        $stateKey = 'reminder|{0}|manual|{1}' -f [int]$row.id, $dateText
        $stateRow = if ($stateByKey.ContainsKey($stateKey)) { $stateByKey[$stateKey] } else { $null }
        $eligible = $null -eq $stateRow
        if ($null -ne $stateRow) {
            $snooze = [DateTime]::MinValue
            $eligible = (-not (Convert-DbBool $stateRow.handled)) -and
                ([string]::IsNullOrWhiteSpace([string]$stateRow.snooze_until) -or
                 ([DateTime]::TryParse([string]$stateRow.snooze_until, [ref]$snooze) -and $snooze -le $currentTime))
        }
        if ($eligible) {
            [void]$items.Add([pscustomobject]@{
                Source = 'reminder'; SourceId = [int]$row.id; Title = [string]$row.title
                Detail = ''; Due = @([pscustomobject]@{ Kind = 'manual'; Label = 'Напоминание'; Date = $date; DateText = $dateText })
                DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id
            })
        }
    }
    return @($items)
}

function Mark-NotificationHandled {
    param($Notification)
    foreach ($due in @($Notification.Due)) {
        Set-NotificationState $Notification.Source ([int]$Notification.SourceId) ([string]$due.Kind) ([string]$due.DateText) $true '' $true
    }
}

function Snooze-Notification {
    param($Notification, [DateTime]$Until)
    $untilText = $Until.ToString('yyyy-MM-dd HH:mm:ss')
    foreach ($due in @($Notification.Due)) {
        Set-NotificationState $Notification.Source ([int]$Notification.SourceId) ([string]$due.Kind) ([string]$due.DateText) $false $untilText $true
    }
}

function Convert-PurchaseDateForSort {
    param([string]$Value)
    $date = Convert-PurchaseDateText $Value
    if ($null -eq $date) { return '' }
    return $date.ToString('yyyy-MM-dd')
}

function Test-PurchaseSupplierErpNotRequired {
    param([string]$Supplier)
    $name = (([string]$Supplier).Trim().ToLowerInvariant()) -replace '\s+', ''
    if ([string]::IsNullOrWhiteSpace($name)) { return $false }
    return (
        $name -match '^(компэл|compel)(_\d+)?$' -or
        $name -match '^(чид|chid)(_\d+)?$' -or
        $name -match '^(промэлектроника|promelec|promelektronika)(_\d+)?$'
    )
}

function Get-PurchaseActionItems {
    param([int]$DealId = 0)

    $where = 'WHERE IFNULL(d.archived, 0) = 0'
    $params = @{}
    if ($DealId -gt 0) {
        $where += ' AND d.id = @deal_id'
        $params['@deal_id'] = $DealId
    }

    $suppliers = Invoke-PurchaseQuery @"
SELECT
    d.id AS deal_id,
    d.deal_number,
    ds.id AS supplier_id,
    ds.supplier,
    ds.invoice_received,
    ds.invoice_confirmed,
    ds.supplier_order_created,
    ds.erp_supplier_sent,
    ds.erp_roger_sent,
    ds.payment_submitted,
    ds.paid,
    ds.invoice_confirmed_date,
    ds.components_receipt_date,
    IFNULL(ds.actual_receipt_date, '') AS actual_receipt_date,
    ds.delivery_weeks
FROM deal_suppliers ds
JOIN deals d ON d.id = ds.deal_id
$where
ORDER BY d.updated_at DESC, d.deal_number, ds.supplier
"@ $params

    $items = New-Object System.Collections.ArrayList
    $today = (Get-Date).Date
    foreach ($row in $suppliers.Rows) {
        $dealNumber = [string]$row.deal_number
        $supplier = [string]$row.supplier
        $prefix = "$dealNumber / $supplier"
        if (-not (Convert-DbBool $row.invoice_received)) {
            [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Warn'; Title = "${prefix}: нет PI"; DueDate = ''; Source = 'auto' })
            continue
        }
        if (-not (Convert-DbBool $row.payment_submitted)) {
            [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Info'; Title = "${prefix}: PI есть, но не подан в оплату"; DueDate = ''; Source = 'auto' })
        } elseif (-not (Convert-DbBool $row.paid)) {
            [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Attention'; Title = "${prefix}: подан в оплату, но не оплачен"; DueDate = ''; Source = 'auto' })
        }
        if ([string]::IsNullOrWhiteSpace([string]$row.invoice_confirmed_date)) {
            [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Warn'; Title = "${prefix}: нет даты подтверждения инвойса"; DueDate = ''; Source = 'auto' })
        }
        if (-not (Test-PurchaseSupplierErpNotRequired ([string]$row.supplier))) {
            if (-not (Convert-DbBool $row.erp_supplier_sent)) {
                [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Info'; Title = "${prefix}: Не отправлен ERP поставщику"; DueDate = ''; Source = 'auto' })
            }
            if (-not (Convert-DbBool $row.erp_roger_sent)) {
                [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Info'; Title = "${prefix}: Не отправлен ERP заказ"; DueDate = ''; Source = 'auto' })
            }
        }

        $actualReceiptDate = Convert-PurchaseDateText ([string]$row.actual_receipt_date)
        $receiptDate = Convert-PurchaseDateText ([string]$row.components_receipt_date)
        if ($null -eq $actualReceiptDate -and $null -ne $receiptDate) {
            $days = [int]($receiptDate.Date - $today).TotalDays
            if ($days -lt 0) {
                [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Danger'; Title = "${prefix}: поступление просрочено"; DueDate = $receiptDate.ToString('dd.MM.yyyy'); Source = 'auto' })
            } elseif ($days -ge 0 -and $days -le 7) {
                [void]$items.Add([pscustomobject]@{ DealId = [int]$row.deal_id; SupplierId = [int]$row.supplier_id; Severity = 'Attention'; Title = "${prefix}: скоро поступление"; DueDate = $receiptDate.ToString('dd.MM.yyyy'); Source = 'auto' })
            }
        }
    }

    $manual = Invoke-PurchaseQuery @'
SELECT id, deal_id, supplier_id, title, due_date, status, source
FROM reminders
WHERE status <> 'Done'
ORDER BY
    CASE WHEN IFNULL(due_date, '') = '' THEN 1 ELSE 0 END,
    due_date,
    id DESC
'@ @{}
    foreach ($row in $manual.Rows) {
        $due = Format-PurchaseDate $row.due_date
        $severity = 'Info'
        $date = Convert-PurchaseDateText $due
        if ($null -eq $date) {
            $severity = 'Warn'
        } elseif ($date.Date -lt $today) {
            $severity = 'Danger'
        } elseif ([int]($date.Date - $today).TotalDays -le 7) {
            $severity = 'Attention'
        }
        [void]$items.Add([pscustomobject]@{ DealId = if ($null -eq $row.deal_id -or $row.deal_id -is [DBNull]) { 0 } else { [int]$row.deal_id }; SupplierId = if ($null -eq $row.supplier_id -or $row.supplier_id -is [DBNull]) { 0 } else { [int]$row.supplier_id }; ReminderId = [int]$row.id; Severity = $severity; Title = [string]$row.title; DueDate = $due; Source = [string]$row.source })
    }
    return @($items)
}

function Get-PurchaseDashboard {
    $items = @(Get-PurchaseActionItems)
    [pscustomobject]@{
        AttentionCount = $items.Count
        OverdueCount = @($items | Where-Object { $_.Severity -eq 'Danger' }).Count
        TodayCount = @($items | Where-Object { (Convert-PurchaseDateText $_.DueDate) -and (Convert-PurchaseDateText $_.DueDate).Date -eq (Get-Date).Date }).Count
        WeekCount = @($items | Where-Object {
            $date = Convert-PurchaseDateText $_.DueDate
            $null -ne $date -and [int]($date.Date - (Get-Date).Date).TotalDays -ge 0 -and [int]($date.Date - (Get-Date).Date).TotalDays -le 7
        }).Count
        Items = $items
    }
}

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
    Invoke-PurchaseNonQuery 'DELETE FROM quote_history' @{}
    Invoke-PurchaseNonQuery 'DELETE FROM quote_batches' @{}
}

function Clear-ActivityLog {
    Invoke-PurchaseNonQuery 'DELETE FROM activity_log' @{}
}

function Convert-DbBool {
    param($Value)
    if ($null -eq $Value -or $Value -is [DBNull]) { return $false }
    return ([int]$Value -ne 0)
}

function Convert-ToDbBool {
    param($Value)
    if ([bool]$Value) { return 1 }
    return 0
}

function Convert-DealYesNoToDbValue {
    param($Value)

    if ($null -eq $Value -or $Value -is [DBNull]) { return $null }
    if ($Value -is [bool]) { return (Convert-ToDbBool $Value) }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    if ($text -eq 'Да' -or $text -eq 'Yes' -or $text -eq 'True' -or $text -eq '1') { return 1 }
    return 0
}

function Convert-MasksToDbValue {
    param($Value)

    if ($null -eq $Value -or $Value -is [DBNull]) { return 2 }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return 2 }
    if ($text -eq 'Да' -or $text -eq 'Yes' -or $text -eq 'True' -or $text -eq '1') { return 1 }
    return 0
}

function Convert-PurchaseDocumentType {
    param([string]$DocumentType)

    $text = if ([string]::IsNullOrWhiteSpace($DocumentType)) { 'Other' } else { $DocumentType.Trim() }
    if ($text -eq 'ERP supplier' -or $text -eq 'ERP Roger') { return 'ERP' }
    return $text
}

function Test-DealLevelDocumentType {
    param([string]$DocumentType)

    $text = if ([string]::IsNullOrWhiteSpace($DocumentType)) { '' } else { $DocumentType.Trim() }
    return ($text -eq 'Общий PI' -or $text -eq 'RFQ' -or $text -eq 'RRFQ' -or $text -eq 'PO' -or $text -eq 'Other')
}

function Get-UniquePurchaseDocumentPath {
    param([string]$Directory, [string]$FileName)

    $safeName = if ([string]::IsNullOrWhiteSpace($FileName)) { 'document' } else { $FileName.Trim() }
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) {
        $safeName = $safeName.Replace([string]$char, '_')
    }
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = 'document' }
    $candidate = Join-Path $Directory $safeName
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }

    $stem = [IO.Path]::GetFileNameWithoutExtension($safeName)
    $extension = [IO.Path]::GetExtension($safeName)
    for ($index = 2; $index -lt 1000; $index++) {
        $candidate = Join-Path $Directory ('{0} ({1}){2}' -f $stem, $index, $extension)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }

    throw "Не удалось подобрать свободное имя файла: $FileName"
}

function Get-SafePathPart {
    param([string]$Value)

    $text = if ([string]::IsNullOrWhiteSpace($Value)) { 'empty' } else { $Value.Trim() }
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) {
        $text = $text.Replace([string]$char, '_')
    }
    return ($text -replace '\s+', '_')
}

function Get-ComponentDealFolderPath {
    param([int]$Id, [string]$DealNumber = '')

    $root = Get-DefaultComponentFilesRoot
    $name = if ([string]::IsNullOrWhiteSpace($DealNumber)) {
        'component_{0:D5}' -f $Id
    } else {
        '{0:D5}_{1}' -f $Id, (Get-SafePathPart $DealNumber)
    }
    return (Join-Path $root $name)
}

function Ensure-ComponentDealFolder {
    param([int]$Id, [string]$DealNumber = '')

    if ($Id -le 0) { return '' }
    $root = Get-DefaultComponentFilesRoot
    New-Item -ItemType Directory -Force -Path $root | Out-Null
    $target = Get-ComponentDealFolderPath $Id $DealNumber
    $currentStored = [string](Invoke-PurchaseScalar 'SELECT IFNULL(folder_path, '''') FROM component_deals WHERE id = @id' @{ '@id' = $Id })
    $current = Resolve-ComponentStoredPath $currentStored
    $chosen = $target
    $samePath = $false
    if (-not [string]::IsNullOrWhiteSpace($current)) {
        try {
            $samePath = ([IO.Path]::GetFullPath($current).TrimEnd('\') -eq [IO.Path]::GetFullPath($target).TrimEnd('\'))
        } catch {
            $samePath = ($current -eq $target)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($current) -and -not $samePath) {
        try {
            if ((Test-Path -LiteralPath $current -PathType Container) -and -not (Test-Path -LiteralPath $target)) {
                Move-Item -LiteralPath $current -Destination $target -Force
                $chosen = $target
            } elseif (Test-Path -LiteralPath $current -PathType Container) {
                $chosen = $current
            }
        } catch {
            $chosen = $current
        }
    }

    if ([string]::IsNullOrWhiteSpace($chosen)) { $chosen = $target }
    New-Item -ItemType Directory -Force -Path $chosen | Out-Null
    Invoke-PurchaseNonQuery 'UPDATE component_deals SET folder_path = @folder_path WHERE id = @id' @{
        '@folder_path' = Convert-ToPortablePath $chosen
        '@id' = $Id
    }
    return $chosen
}

function Remove-ComponentDealFolder {
    param([int]$Id, [string]$FolderPath = '')

    if ($Id -le 0) { return }
    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        $FolderPath = [string](Invoke-PurchaseScalar 'SELECT IFNULL(folder_path, '''') FROM component_deals WHERE id = @id' @{ '@id' = $Id })
    }
    if ([string]::IsNullOrWhiteSpace($FolderPath)) { return }

    $root = [IO.Path]::GetFullPath((Get-DefaultComponentFilesRoot))
    $resolvedFolderPath = Resolve-ComponentStoredPath $FolderPath
    $target = [IO.Path]::GetFullPath($resolvedFolderPath)
    if (-not $target.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { return }
    if (Test-Path -LiteralPath $target -PathType Container) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
}

function New-PurchaseDeal {
    param(
        [string]$DealNumber,
        [string]$Title = '',
        [string]$Comment = '',
        [string]$Client = '',
        [string]$Status = ''
    )

    if ([string]::IsNullOrWhiteSpace($DealNumber)) {
        throw 'Введите номер сделки.'
    }

    $dealNumberParam = $DealNumber.Trim()
    $now = Get-NowText
    $workflowTemplateId = Get-DefaultWorkflowTemplateId
    Invoke-PurchaseNonQuery @'
INSERT INTO deals(deal_number, client, status, title, comment, workflow_template_id, masks, created_at, updated_at)
VALUES(@deal_number, @client, CASE WHEN @status <> '' THEN @status ELSE 'RFQ' END, @title, @comment, @workflow_template_id, 2, @created_at, @updated_at)
ON CONFLICT(deal_number) DO UPDATE SET
    client = CASE WHEN excluded.client <> '' THEN excluded.client ELSE client END,
    status = CASE WHEN @status <> '' THEN @status ELSE status END,
    title = CASE WHEN excluded.title <> '' THEN excluded.title ELSE title END,
    comment = CASE WHEN excluded.comment <> '' THEN excluded.comment ELSE comment END,
    archived = 0,
    updated_at = excluded.updated_at;
'@ @{
        '@deal_number' = $dealNumberParam
        '@client' = $Client
        '@status' = $Status
        '@title' = $Title
        '@comment' = $Comment
        '@workflow_template_id' = $workflowTemplateId
        '@created_at' = $now
        '@updated_at' = $now
    }
    $id = [int](Invoke-PurchaseScalar 'SELECT id FROM deals WHERE deal_number = @deal_number' @{ '@deal_number' = $dealNumberParam })
    Write-ActivityLog 'deal' $id 'Создана/обновлена сделка' $dealNumberParam $id
    return $id
}

function Update-PurchaseDeal {
    param(
        [int]$DealId,
        [string]$Client,
        [string]$Status,
        [string]$Comment,
        [AllowNull()][string]$BoardCount = $null,
        [AllowNull()][string]$DealNumber = $null,
        [AllowNull()][int]$WorkflowTemplateId = $null,
        [AllowNull()][string]$Period = $null,
        [AllowNull()]$Masks = $null,
        [AllowNull()]$PiSent = $null,
        [AllowNull()][string]$Priority = $null,
        [AllowNull()][string]$TrackingStatus = $null,
        [AllowNull()][string]$Executor = $null,
        [AllowNull()][string]$AssemblyLocation = $null
    )

    if ($DealId -le 0) { throw 'Выберите сделку.' }
    $dealNumberParam = $null
    if (-not [string]::IsNullOrWhiteSpace($DealNumber)) {
        $dealNumberParam = $DealNumber.Trim()
    }
    Invoke-PurchaseNonQuery @'
UPDATE deals SET
    deal_number = CASE WHEN @deal_number IS NULL THEN deal_number ELSE @deal_number END,
    board_count = CASE WHEN @board_count IS NULL THEN board_count ELSE @board_count END,
    workflow_template_id = CASE WHEN @workflow_template_id IS NULL THEN workflow_template_id ELSE @workflow_template_id END,
    period = CASE WHEN @period IS NULL THEN period ELSE @period END,
    priority = CASE WHEN @priority IS NULL THEN priority ELSE @priority END,
    masks = CASE WHEN @masks IS NULL THEN masks ELSE @masks END,
    pi_sent = CASE WHEN @pi_sent IS NULL THEN pi_sent ELSE @pi_sent END,
    client = @client,
    status = @status,
    tracking_status = CASE WHEN @tracking_status IS NULL THEN tracking_status ELSE @tracking_status END,
    executor = CASE WHEN @executor IS NULL OR trim(@executor) = '' THEN executor ELSE @executor END,
    assembly_location = CASE WHEN @assembly_location IS NULL THEN assembly_location ELSE @assembly_location END,
    comment = @comment,
    updated_at = @updated_at
WHERE id = @id
'@ @{
        '@deal_number' = $dealNumberParam
        '@board_count' = $BoardCount
        '@workflow_template_id' = if ($null -eq $WorkflowTemplateId -or $WorkflowTemplateId -le 0) { $null } else { $WorkflowTemplateId }
        '@period' = $Period
        '@priority' = if ($null -eq $Priority) { $null } else { $Priority.Trim() }
        '@masks' = Convert-MasksToDbValue $Masks
        '@pi_sent' = Convert-DealYesNoToDbValue $PiSent
        '@client' = $Client
        '@status' = $Status
        '@tracking_status' = if ([string]::IsNullOrWhiteSpace($TrackingStatus)) { $null } else { $TrackingStatus.Trim() }
        '@executor' = if ($null -eq $Executor) { $null } else { ([string]$Executor).Trim() }
        '@assembly_location' = if ($null -eq $AssemblyLocation) { $null } else { ([string]$AssemblyLocation).Trim() }
        '@comment' = $Comment
        '@updated_at' = Get-NowText
        '@id' = $DealId
    }
    Write-ActivityLog 'deal' $DealId 'Изменена сделка' "Этап: $Status; Статус: $TrackingStatus; Клиент: $Client; Период: $Period; Исполнитель: $Executor; Место сборки: $AssemblyLocation" $DealId
}

function Touch-PurchaseDeal {
    param([int]$DealId)

    if ($DealId -le 0) { return }
    Invoke-PurchaseNonQuery 'UPDATE deals SET updated_at = @updated_at WHERE id = @id' @{
        '@updated_at' = Get-NowText
        '@id' = $DealId
    }
}

function Touch-PurchaseDealBySupplier {
    param([int]$SupplierId)

    if ($SupplierId -le 0) { return }
    Invoke-PurchaseNonQuery @'
UPDATE deals SET updated_at = @updated_at
WHERE id = (SELECT deal_id FROM deal_suppliers WHERE id = @supplier_id)
'@ @{
        '@updated_at' = Get-NowText
        '@supplier_id' = $SupplierId
    }
}

function Add-PurchaseSupplier {
    param([int]$DealId, [string]$Supplier)

    if ($DealId -le 0) { throw 'Выберите сделку.' }
    if ([string]::IsNullOrWhiteSpace($Supplier)) { throw 'Введите поставщика.' }

    $now = Get-NowText
    Invoke-PurchaseNonQuery @'
INSERT INTO deal_suppliers(deal_id, supplier, created_at, updated_at)
VALUES(@deal_id, @supplier, @created_at, @updated_at)
ON CONFLICT(deal_id, supplier) DO UPDATE SET updated_at = excluded.updated_at;
'@ @{
        '@deal_id' = $DealId
        '@supplier' = $Supplier.Trim()
        '@created_at' = $now
        '@updated_at' = $now
    }
    Touch-PurchaseDeal $DealId
    Write-ActivityLog 'supplier' 0 'Добавлен/обновлен поставщик' $Supplier.Trim() $DealId
}

function Delete-PurchaseSupplier {
    param([int]$SupplierId)

    if ($SupplierId -le 0) { throw 'Выберите поставщика.' }
    $table = Invoke-PurchaseQuery 'SELECT documents.id, documents.deal_id, deal_suppliers.supplier, documents.original_name, documents.stored_path FROM documents LEFT JOIN deal_suppliers ON deal_suppliers.id = documents.supplier_id WHERE documents.supplier_id = @supplier_id' @{ '@supplier_id' = $SupplierId }
    $supplier = [string](Invoke-PurchaseScalar 'SELECT supplier FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId })
    $dealId = [int](Invoke-PurchaseScalar 'SELECT deal_id FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId })
    if ([string]::IsNullOrWhiteSpace($supplier) -or $dealId -le 0) { throw 'Поставщик не найден.' }

    foreach ($row in $table.Rows) {
        $storedPath = [string]$row.stored_path
        $pathToDelete = Resolve-PurchaseStoredPath $storedPath
        if (-not [string]::IsNullOrWhiteSpace($pathToDelete) -and (Test-Path -LiteralPath $pathToDelete -PathType Leaf)) {
            Remove-Item -LiteralPath $pathToDelete -Force
        }
    }
    Invoke-PurchaseNonQuery 'DELETE FROM documents WHERE supplier_id = @supplier_id' @{ '@supplier_id' = $SupplierId }
    Invoke-PurchaseNonQuery 'DELETE FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId }
    Touch-PurchaseDeal $dealId
    Write-ActivityLog 'supplier' $SupplierId 'Поставщик удален' $supplier $dealId
}

function Get-PurchaseDeals {
    param([string]$Search = '', [string]$Filter = 'All')

    $where = 'WHERE (@search = '''' OR d.deal_number LIKE @needle OR IFNULL(d.title, '''') LIKE @needle OR IFNULL(d.client, '''') LIKE @needle OR IFNULL(d.status, '''') LIKE @needle OR IFNULL(d.tracking_status, '''') LIKE @needle OR IFNULL(d.period, '''') LIKE @needle OR IFNULL(d.executor, '''') LIKE @needle OR IFNULL(d.assembly_location, '''') LIKE @needle OR IFNULL(d.board_count, '''') LIKE @needle)'
    if ($Filter -eq 'Archived') {
        $where += ' AND IFNULL(d.archived, 0) = 1'
    } elseif ($Filter -eq 'AllRecords') {
    } elseif ($Filter -eq 'China') {
        $where += " AND trim(IFNULL(d.assembly_location, '')) = 'Китай'"
    } else {
        $where += ' AND IFNULL(d.archived, 0) = 0'
        if ($Filter -eq 'Active') {
$where += " AND trim(IFNULL(d.executor, '')) = '' AND (IFNULL(d.status, '') IN ('RFQ', 'PO', 'Закупка', 'В работе') OR (IFNULL(d.status, '') = 'Заказано' AND (IFNULL(stats.supplier_count, 0) = 0 OR IFNULL(stats.invoice_confirmed_count, 0) < IFNULL(stats.supplier_count, 0))))"
        }
    }

    # AllRecords is used when navigating from reminders and must really include
    # every deal, including deals assembled in China. The regular/default views
    # keep their existing non-China restriction.
    if ($Filter -ne 'China' -and $Filter -ne 'Tracking' -and $Filter -ne 'AllRecords') {
        $where += " AND trim(IFNULL(d.assembly_location, '')) <> 'Китай'"
    }

    if ($Filter -eq 'Tracking') {
$where += " AND trim(IFNULL(d.executor, '')) <> ''"
    } elseif ($Filter -eq 'China') {
        $where += " AND trim(IFNULL(d.assembly_location, '')) = 'Китай'"
    } elseif ($Filter -eq 'Purchase') {
        $where += ' AND IFNULL(d.status, '''') = ''Закупка'''
    } elseif ($Filter -eq 'Pause') {
        $where += ' AND IFNULL(d.status, '''') IN (''RRFQ'', ''PI'')'
    } elseif ($Filter -eq 'NoInvoices') {
        $where += ' AND IFNULL(stats.invoice_count, 0) = 0'
    } elseif ($Filter -eq 'WaitingReceipt') {
        $where += ' AND IFNULL(stats.supplier_count, 0) > 0 AND IFNULL(stats.receipt_count, 0) < IFNULL(stats.supplier_count, 0)'
    } elseif ($Filter -eq 'Done') {
        $where += " AND IFNULL(d.status, '') = 'Заказано' AND IFNULL(stats.supplier_count, 0) > 0 AND IFNULL(stats.invoice_confirmed_count, 0) = IFNULL(stats.supplier_count, 0)"
    } elseif ($Filter -eq 'Open') {
        $where += ' AND (IFNULL(stats.supplier_count, 0) = 0 OR IFNULL(stats.done_count, 0) < IFNULL(stats.supplier_count, 0))'
    }

    $sql = @"
SELECT
    d.id,
    d.deal_number,
    IFNULL(d.board_count, '') AS board_count,
    IFNULL(d.client, '') AS client,
    IFNULL(d.period, '') AS period,
    IFNULL(d.executor, '') AS executor,
    IFNULL(d.assembly_location, '') AS assembly_location,
    IFNULL(d.priority, '3') AS priority,
    IFNULL(d.tracking_status, 'Ожидание') AS tracking_status,
    IFNULL(d.status, 'RFQ') AS status,
    IFNULL(d.pi_sent, 0) AS pi_sent,
    IFNULL(d.masks, 0) AS masks,
    IFNULL(d.archived, 0) AS archived,
    d.workflow_template_id,
    d.title,
    d.comment,
    d.updated_at,
    IFNULL(stats.supplier_count, 0) AS supplier_count,
    IFNULL(stats.invoice_count, 0) AS invoice_count,
    IFNULL(stats.payment_submitted_count, 0) AS payment_submitted_count,
    IFNULL(stats.paid_count, 0) AS paid_count,
    IFNULL(stats.done_count, 0) AS done_count,
    IFNULL(stats.invoice_confirmed_count, 0) AS invoice_confirmed_count,
    IFNULL(stats.receipt_count, 0) AS receipt_count,
    IFNULL(stats.max_receipt_date, '') AS completion_receipt_date
FROM deals d
LEFT JOIN (
    SELECT
        deal_id,
        COUNT(*) AS supplier_count,
        SUM(invoice_received) AS invoice_count,
        SUM(payment_submitted) AS payment_submitted_count,
        SUM(paid) AS paid_count,
        SUM(CASE WHEN invoice_received = 1 AND invoice_confirmed = 1 THEN 1 ELSE 0 END) AS invoice_confirmed_count,
        SUM(CASE WHEN invoice_received = 1 AND invoice_confirmed = 1 AND ((supplier LIKE 'Компэл%' OR supplier LIKE 'Компел%' OR supplier LIKE 'ЧиД%' OR supplier LIKE 'ЧИД%' OR supplier LIKE 'чид%' OR supplier LIKE 'Промэлектроника%' OR supplier LIKE 'промэлектроника%' OR supplier LIKE 'Promelec%' OR supplier LIKE 'promelec%') OR (erp_supplier_sent = 1 AND erp_roger_sent = 1)) THEN 1 ELSE 0 END) AS done_count,
        SUM(CASE WHEN IFNULL(actual_receipt_date, '') <> '' THEN 1 ELSE 0 END) AS receipt_count,
        MAX(CASE
            WHEN IFNULL(actual_receipt_date, '') LIKE '__.__.____' THEN substr(actual_receipt_date, 7, 4) || '-' || substr(actual_receipt_date, 4, 2) || '-' || substr(actual_receipt_date, 1, 2)
            WHEN IFNULL(actual_receipt_date, '') LIKE '____-__-__%' THEN substr(actual_receipt_date, 1, 10)
            ELSE ''
        END) AS max_receipt_date
    FROM deal_suppliers
    GROUP BY deal_id
) stats ON stats.deal_id = d.id
$where
ORDER BY d.updated_at DESC, d.id DESC
"@

    return Invoke-PurchaseQuery $sql @{
        '@search' = $Search
        '@needle' = "%$Search%"
    }
}

function Get-PurchaseSuppliers {
    param([int]$DealId)

    if ($DealId -le 0) {
        return New-Object System.Data.DataTable
    }

    return Invoke-PurchaseQuery @'
SELECT
    id,
    deal_id,
    supplier,
    invoice_received,
    invoice_confirmed,
    supplier_order_created,
    erp_supplier_sent,
    erp_roger_sent,
    pi_amount_usd,
    pi_amount_cny,
    pi_amount_rub,
    IFNULL(paid_amount, '') AS paid_amount,
    delivery_weeks,
    payment_submitted,
    paid,
    invoice_confirmed_date,
    components_receipt_date,
    IFNULL(actual_receipt_date, '') AS actual_receipt_date,
    comment
FROM deal_suppliers
WHERE deal_id = @deal_id
ORDER BY supplier
'@ @{ '@deal_id' = $DealId }
}

function Update-PurchaseSupplier {
    param(
        [int]$SupplierId,
        [bool]$InvoiceReceived,
        [bool]$InvoiceConfirmed,
        [bool]$SupplierOrderCreated,
        [bool]$ErpSupplierSent,
        [bool]$ErpRogerSent,
        $PiAmountUsd,
        $PiAmountCny,
        $PiAmountRub,
        $PaidAmount,
        $DeliveryWeeks,
        [bool]$PaymentSubmitted,
        [bool]$Paid,
        [string]$InvoiceConfirmedDate,
        [string]$ComponentsReceiptDate,
        [string]$ActualReceiptDate,
        [string]$Comment
    )

    Invoke-PurchaseNonQuery @'
UPDATE deal_suppliers SET
    invoice_received = @invoice_received,
    invoice_confirmed = @invoice_confirmed,
    supplier_order_created = @supplier_order_created,
    erp_supplier_sent = @erp_supplier_sent,
    erp_roger_sent = @erp_roger_sent,
    pi_amount_usd = @pi_amount_usd,
    pi_amount_cny = @pi_amount_cny,
    pi_amount_rub = @pi_amount_rub,
    paid_amount = @paid_amount,
    delivery_weeks = @delivery_weeks,
    payment_submitted = @payment_submitted,
    paid = @paid,
    invoice_confirmed_date = @invoice_confirmed_date,
    components_receipt_date = @components_receipt_date,
    actual_receipt_date = @actual_receipt_date,
    comment = @comment,
    updated_at = @updated_at
WHERE id = @id
'@ @{
        '@invoice_received' = Convert-ToDbBool $InvoiceReceived
        '@invoice_confirmed' = Convert-ToDbBool $InvoiceConfirmed
        '@supplier_order_created' = Convert-ToDbBool $SupplierOrderCreated
        '@erp_supplier_sent' = Convert-ToDbBool $ErpSupplierSent
        '@erp_roger_sent' = Convert-ToDbBool $ErpRogerSent
        '@pi_amount_usd' = Convert-ToNullableAmount $PiAmountUsd
        '@pi_amount_cny' = Convert-ToNullableAmount $PiAmountCny
        '@pi_amount_rub' = Convert-ToNullableAmount $PiAmountRub
        '@paid_amount' = $PaidAmount
        '@delivery_weeks' = Convert-ToNullableInt $DeliveryWeeks
        '@payment_submitted' = Convert-ToDbBool $PaymentSubmitted
        '@paid' = Convert-ToDbBool $Paid
        '@invoice_confirmed_date' = $InvoiceConfirmedDate
        '@components_receipt_date' = $ComponentsReceiptDate
        '@actual_receipt_date' = $ActualReceiptDate
        '@comment' = $Comment
        '@updated_at' = Get-NowText
        '@id' = $SupplierId
    }
    Touch-PurchaseDealBySupplier $SupplierId
    $dealId = Invoke-PurchaseScalar 'SELECT deal_id FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId }
    Write-ActivityLog 'supplier' $SupplierId 'Изменен поставщик' "Оплата: $Paid; PI: $InvoiceReceived; срок: $DeliveryWeeks" $(if ($null -eq $dealId -or $dealId -is [DBNull]) { 0 } else { [int]$dealId }) $SupplierId
}

function Set-PurchaseSupplierActualReceiptDate {
    param([int]$SupplierId, [string]$ActualReceiptDate)

    if ($SupplierId -le 0) { throw 'Choose supplier.' }
    $actualText = ''
    if (-not [string]::IsNullOrWhiteSpace($ActualReceiptDate)) {
        $date = Convert-PurchaseDateText $ActualReceiptDate
        if ($null -eq $date) { throw 'Фактическая дата поступления должна быть в формате дд.мм.гггг.' }
        $actualText = $date.ToString('dd.MM.yyyy')
    }

    Invoke-PurchaseNonQuery 'UPDATE deal_suppliers SET actual_receipt_date = @actual_receipt_date, updated_at = @updated_at WHERE id = @id' @{
        '@actual_receipt_date' = $actualText
        '@updated_at' = Get-NowText
        '@id' = $SupplierId
    }
    Touch-PurchaseDealBySupplier $SupplierId
    $dealId = Invoke-PurchaseScalar 'SELECT deal_id FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId }
    Write-ActivityLog 'supplier' $SupplierId 'Поступление подтверждено на склад' ('Факт: {0}' -f $actualText) $(if ($null -eq $dealId -or $dealId -is [DBNull]) { 0 } else { [int]$dealId }) $SupplierId
}
function Update-PurchaseSupplierPiAmounts {
    param(
        [int]$SupplierId,
        $PiAmountUsd,
        $PiAmountCny,
        [bool]$MarkInvoiceReceived = $true
    )

    if ($SupplierId -le 0) { return }
    Invoke-PurchaseNonQuery @'
UPDATE deal_suppliers SET
    invoice_received = CASE WHEN @mark_invoice_received = 1 THEN 1 ELSE invoice_received END,
    pi_amount_usd = CASE WHEN @pi_amount_usd IS NULL THEN pi_amount_usd ELSE @pi_amount_usd END,
    pi_amount_cny = CASE WHEN @pi_amount_cny IS NULL THEN pi_amount_cny ELSE @pi_amount_cny END,
    updated_at = @updated_at
WHERE id = @id
'@ @{
        '@mark_invoice_received' = Convert-ToDbBool $MarkInvoiceReceived
        '@pi_amount_usd' = Convert-ToNullableAmount $PiAmountUsd
        '@pi_amount_cny' = Convert-ToNullableAmount $PiAmountCny
        '@updated_at' = Get-NowText
        '@id' = $SupplierId
    }
    Touch-PurchaseDealBySupplier $SupplierId
    $dealId = Invoke-PurchaseScalar 'SELECT deal_id FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId }
    Write-ActivityLog 'supplier' $SupplierId 'Обновлены суммы PI' "USD=$(Format-PurchaseAmount $PiAmountUsd); CNY=$(Format-PurchaseAmount $PiAmountCny)" $(if ($null -eq $dealId -or $dealId -is [DBNull]) { 0 } else { [int]$dealId }) $SupplierId
}

function Get-PurchaseDocuments {
    param([int]$DealId)

    if ($DealId -le 0) {
        return New-Object System.Data.DataTable
    }

    return Invoke-PurchaseQuery @'
SELECT
    doc.id,
    doc.document_type,
    IFNULL(ds.supplier, '') AS supplier,
    doc.original_name,
    doc.stored_path,
    IFNULL(doc.file_hash, '') AS file_hash,
    doc.uploaded_at
FROM documents doc
LEFT JOIN deal_suppliers ds ON ds.id = doc.supplier_id
WHERE doc.deal_id = @deal_id
ORDER BY doc.uploaded_at DESC, doc.id DESC
'@ @{ '@deal_id' = $DealId }
}

function Get-PurchaseDocumentFileHash {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    try {
        return [string](Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    } catch {
        return ''
    }
}

function Update-PurchaseDocumentFileReference {
    param([int]$DocumentId, [string]$Path, [string]$FileHash = '')

    if ($DocumentId -le 0 -or [string]::IsNullOrWhiteSpace($Path)) { return }
    $resolvedPath = Resolve-PurchaseStoredPath $Path
    if ([string]::IsNullOrWhiteSpace($resolvedPath)) { $resolvedPath = $Path }
    $name = [IO.Path]::GetFileName($resolvedPath)
    if ([string]::IsNullOrWhiteSpace($FileHash)) {
        $FileHash = Get-PurchaseDocumentFileHash $resolvedPath
    }

    Invoke-PurchaseNonQuery 'UPDATE documents SET stored_path = @stored_path, original_name = @original_name, file_hash = CASE WHEN @file_hash <> '''' THEN @file_hash ELSE file_hash END WHERE id = @id' @{
        '@stored_path' = Convert-ToPortablePath $resolvedPath
        '@original_name' = $name
        '@file_hash' = $FileHash
        '@id' = $DocumentId
    }
}

function Resolve-PurchaseDocumentFile {
    param(
        [int]$DocumentId,
        [string]$StoredPath,
        [string]$OriginalName = '',
        [string]$FileHash = ''
    )

    $storedPath = [string]$StoredPath
    $path = Resolve-PurchaseStoredPath $storedPath
    $name = if ([string]::IsNullOrWhiteSpace($OriginalName)) { [IO.Path]::GetFileName($path) } else { $OriginalName }
    $hash = [string]$FileHash

    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType Leaf)) {
        $currentName = [IO.Path]::GetFileName($path)
        if ([string]::IsNullOrWhiteSpace($hash)) { $hash = Get-PurchaseDocumentFileHash $path }
        $portablePath = Convert-ToPortablePath $path
        if ($DocumentId -gt 0 -and ($currentName -ne $name -or [string]::IsNullOrWhiteSpace($FileHash) -or $portablePath -ne $storedPath)) {
            Update-PurchaseDocumentFileReference $DocumentId $path $hash
        }
        return [pscustomobject]@{ Id = $DocumentId; Path = $path; Name = $currentName; FileHash = $hash; Found = $true; Changed = ($currentName -ne $name -or $portablePath -ne $storedPath) }
    }

    if ([string]::IsNullOrWhiteSpace($path)) {
        return [pscustomobject]@{ Id = $DocumentId; Path = $path; Name = $name; FileHash = $hash; Found = $false; Changed = $false }
    }

    try {
        $directory = [IO.Path]::GetDirectoryName($path)
    } catch {
        $directory = ''
    }
    if ([string]::IsNullOrWhiteSpace($directory) -or -not (Test-Path -LiteralPath $directory -PathType Container)) {
        return [pscustomobject]@{ Id = $DocumentId; Path = $path; Name = $name; FileHash = $hash; Found = $false; Changed = $false }
    }

    $candidates = @(Get-ChildItem -LiteralPath $directory -File -ErrorAction SilentlyContinue)
    $chosen = $null
    if (-not [string]::IsNullOrWhiteSpace($hash)) {
        $hashMatches = New-Object System.Collections.ArrayList
        foreach ($candidate in $candidates) {
            if ((Get-PurchaseDocumentFileHash $candidate.FullName) -eq $hash) {
                [void]$hashMatches.Add($candidate)
            }
        }
        if ($hashMatches.Count -eq 1) { $chosen = $hashMatches[0] }
    }

    if ($null -eq $chosen) {
        $knownPaths = @{}
        $otherDocs = Invoke-PurchaseQuery 'SELECT stored_path FROM documents WHERE id <> @id' @{ '@id' = $DocumentId }
        foreach ($row in $otherDocs.Rows) {
            $otherPath = Resolve-PurchaseStoredPath ([string]$row.stored_path)
            if ([string]::IsNullOrWhiteSpace($otherPath)) { continue }
            try { $knownPaths[[IO.Path]::GetFullPath($otherPath).ToLowerInvariant()] = $true } catch { }
        }
        $available = @($candidates | Where-Object {
            try { -not $knownPaths.ContainsKey([IO.Path]::GetFullPath($_.FullName).ToLowerInvariant()) } catch { $false }
        })
        $extension = [IO.Path]::GetExtension($path)
        if (-not [string]::IsNullOrWhiteSpace($extension)) {
            $sameExtension = @($available | Where-Object { [IO.Path]::GetExtension($_.FullName).Equals($extension, [StringComparison]::OrdinalIgnoreCase) })
            if ($sameExtension.Count -eq 1) { $chosen = $sameExtension[0] }
        }
        if ($null -eq $chosen -and $available.Count -eq 1) { $chosen = $available[0] }
    }

    if ($null -ne $chosen) {
        $newPath = [string]$chosen.FullName
        $newHash = Get-PurchaseDocumentFileHash $newPath
        Update-PurchaseDocumentFileReference $DocumentId $newPath $newHash
        return [pscustomobject]@{ Id = $DocumentId; Path = $newPath; Name = $chosen.Name; FileHash = $newHash; Found = $true; Changed = $true }
    }

    return [pscustomobject]@{ Id = $DocumentId; Path = $path; Name = $name; FileHash = $hash; Found = $false; Changed = $false }
}
function Get-PurchaseSupplierPiDocuments {
    param([int]$SupplierId)

    if ($SupplierId -le 0) {
        return New-Object System.Data.DataTable
    }

    return Invoke-PurchaseQuery @'
SELECT
    doc.id,
    doc.original_name,
    doc.stored_path,
    IFNULL(doc.file_hash, '') AS file_hash,
    doc.uploaded_at
FROM documents doc
WHERE doc.supplier_id = @supplier_id
  AND doc.document_type = 'PI/Invoice'
ORDER BY doc.uploaded_at DESC, doc.id DESC
'@ @{ '@supplier_id' = $SupplierId }
}

function Get-PurchaseSupplierName {
    param([int]$SupplierId)

    if ($SupplierId -le 0) { return '' }
    return [string](Invoke-PurchaseScalar 'SELECT supplier FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId })
}

function Import-PiAmountsFromPdfForSupplier {
    param([int]$SupplierId, [string]$Supplier, [string]$Path)

    if ($SupplierId -le 0) { throw 'Выберите поставщика.' }
    if ([string]::IsNullOrWhiteSpace($Supplier)) {
        $Supplier = Get-PurchaseSupplierName $SupplierId
    }

    if (Test-SkipPiPdfAmountExtraction $Supplier) {
        Update-PurchaseSupplierPiAmounts $SupplierId $null $null $true
        return [pscustomobject]@{
            Updated = $false
            Skipped = $true
            Message = "Для $Supplier авто-чтение суммы PI отключено."
        }
    }

    if ([IO.Path]::GetExtension($Path).ToLowerInvariant() -ne '.pdf') {
        Update-PurchaseSupplierPiAmounts $SupplierId $null $null $true
        return [pscustomobject]@{
            Updated = $false
            Skipped = $true
            Message = 'Файл не PDF, сумму нужно заполнить вручную.'
        }
    }

    $amounts = Get-PiAmountsFromPdf $Path
    Update-PurchaseSupplierPiAmounts $SupplierId $amounts.Usd $amounts.Cny $true
    $updated = ($null -ne $amounts.Usd -or $null -ne $amounts.Cny)
    $message = if ($updated) {
        "Суммы найдены: `$=$(Format-PurchaseAmount $amounts.Usd), ?=$(Format-PurchaseAmount $amounts.Cny)"
    } else {
        $amounts.Warning
    }

    return [pscustomobject]@{
        Updated = $updated
        Skipped = $false
        Message = $message
    }
}

function Add-PurchaseDocument {
    param(
        [int]$DealId,
        [int]$SupplierId,
        [string]$DocumentType,
        [string]$SourcePath,
        [string]$FolderName = ''
    )

    if ($DealId -le 0) { throw 'Выберите сделку.' }
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { throw "Файл не найден: $SourcePath" }
    if ([string]::IsNullOrWhiteSpace($DocumentType)) { $DocumentType = 'Other' }
    $DocumentType = Convert-PurchaseDocumentType $DocumentType
    if (Test-DealLevelDocumentType $DocumentType) { $SupplierId = 0 }

    $dealNumber = [string](Invoke-PurchaseScalar 'SELECT deal_number FROM deals WHERE id = @id' @{ '@id' = $DealId })
    $supplier = ''
    if ($SupplierId -gt 0) {
        $supplier = [string](Invoke-PurchaseScalar 'SELECT supplier FROM deal_suppliers WHERE id = @id' @{ '@id' = $SupplierId })
    }

    $root = Get-PurchaseDocumentsRoot

    $folderPart = if (-not [string]::IsNullOrWhiteSpace($FolderName)) {
        $FolderName
    } elseif (Test-DealLevelDocumentType $DocumentType) {
        $DocumentType
    } else {
        $supplier
    }
    $targetDir = Join-Path (Join-Path $root (Get-SafePathPart $dealNumber)) (Get-SafePathPart $folderPart)
    New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
    $originalName = [IO.Path]::GetFileName($SourcePath)
    $targetPath = Get-UniquePurchaseDocumentPath $targetDir $originalName
    Copy-Item -LiteralPath $SourcePath -Destination $targetPath

    Invoke-PurchaseNonQuery @'
INSERT INTO documents(deal_id, supplier_id, document_type, original_name, stored_path, file_hash, uploaded_at)
VALUES(@deal_id, @supplier_id, @document_type, @original_name, @stored_path, @file_hash, @uploaded_at)
'@ @{
        '@deal_id' = $DealId
        '@supplier_id' = if ($SupplierId -gt 0) { $SupplierId } else { $null }
        '@document_type' = $DocumentType
        '@original_name' = $originalName
        '@stored_path' = Convert-ToPortablePath $targetPath
        '@file_hash' = Get-PurchaseDocumentFileHash $targetPath
        '@uploaded_at' = Get-NowText
    }
    Touch-PurchaseDeal $DealId
    Write-ActivityLog 'document' 0 'Загружен документ' "${DocumentType}: $originalName" $DealId $SupplierId
    return $targetPath
}

function Delete-PurchaseDocument {
    param([int]$DocumentId)

    if ($DocumentId -le 0) {
        throw 'Выберите документ.'
    }

    $table = Invoke-PurchaseQuery 'SELECT id, deal_id, original_name, stored_path, IFNULL(file_hash, '''') AS file_hash FROM documents WHERE id = @id' @{ '@id' = $DocumentId }
    if ($table.Rows.Count -eq 0) {
        throw 'Документ не найден.'
    }

    $dealId = [int]$table.Rows[0].deal_id
    $resolvedDoc = Resolve-PurchaseDocumentFile $DocumentId ([string]$table.Rows[0].stored_path) ([string]$table.Rows[0].original_name) ([string]$table.Rows[0].file_hash)
    $path = [string]$resolvedDoc.Path
    if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType Leaf)) {
        Remove-Item -LiteralPath $path -Force
    }

    Invoke-PurchaseNonQuery 'DELETE FROM documents WHERE id = @id' @{ '@id' = $DocumentId }
    Touch-PurchaseDeal $dealId
    Write-ActivityLog 'document' $DocumentId 'Удален документ' $path $dealId
}

function Delete-PurchaseDeal {
    param([int]$DealId)

    if ($DealId -le 0) {
        throw 'Выберите сделку.'
    }

    $dealNumber = [string](Invoke-PurchaseScalar 'SELECT deal_number FROM deals WHERE id = @id' @{ '@id' = $DealId })
    $safeDealPart = Get-SafePathPart $dealNumber
    $dealDirs = New-Object System.Collections.ArrayList
    $docs = Invoke-PurchaseQuery 'SELECT stored_path FROM documents WHERE deal_id = @deal_id' @{ '@deal_id' = $DealId }
    foreach ($row in $docs.Rows) {
        $path = Resolve-PurchaseStoredPath ([string]$row.stored_path)
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType Leaf)) {
            try {
                $parent = [IO.Directory]::GetParent([IO.Path]::GetFullPath($path))
                if ($null -ne $parent -and $null -ne $parent.Parent -and $parent.Parent.Name -eq $safeDealPart) {
                    [void]$dealDirs.Add($parent.Parent.FullName)
                }
            } catch {
            }
            Remove-Item -LiteralPath $path -Force
        }
    }

    if (Get-Command -Name Save-BitrixTaskLink -ErrorAction SilentlyContinue) {
        $bitrixTaskLinks = Invoke-PurchaseQuery 'SELECT bitrix_task_id FROM bitrix_task_links WHERE deal_id = @deal_id' @{ '@deal_id' = $DealId }
        foreach ($row in $bitrixTaskLinks.Rows) {
            Save-BitrixTaskLink -BitrixTaskId ([int]$row.bitrix_task_id) -DealId $null -IsManualDeleted:$true
        }
    }

    Invoke-PurchaseNonQuery 'DELETE FROM deals WHERE id = @id' @{ '@id' = $DealId }
    Write-ActivityLog 'deal' $DealId 'Удалена сделка' $dealNumber $DealId

    $currentRoot = Get-PurchaseDocumentsRoot
    if (-not [string]::IsNullOrWhiteSpace($currentRoot)) {
        [void]$dealDirs.Add((Join-Path $currentRoot $safeDealPart))
    }
    [void]$dealDirs.Add((Join-Path (Get-DefaultPurchaseDocumentsRoot) $safeDealPart))

    foreach ($dir in @($dealDirs | Sort-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        try {
            $fullDir = [IO.Path]::GetFullPath($dir)
            if ((Split-Path -Leaf $fullDir) -ne $safeDealPart) { continue }
            if (Test-Path -LiteralPath $fullDir -PathType Container) {
                Remove-Item -LiteralPath $fullDir -Recurse -Force
            }
        } catch {
        }
    }
}

function New-ComponentDeal {
    $now = Get-NowText
    $entryDate = (Get-Date).ToString('dd.MM.yyyy')
    Invoke-PurchaseNonQuery @'
INSERT INTO component_deals(entry_date, status, stage, priority, created_at, updated_at)
VALUES(@entry_date, 'В работе', 'Запросил поставщиков', '3', @created_at, @updated_at)
'@ @{
        '@entry_date' = $entryDate
        '@created_at' = $now
        '@updated_at' = $now
    }
    $id = [int](Invoke-PurchaseScalar 'SELECT id FROM component_deals ORDER BY id DESC LIMIT 1')
    $folder = Ensure-ComponentDealFolder $id ''
    Write-ActivityLog 'component' $id 'Создана задача' $folder 0 0
    return $id
}

function Get-ComponentDeals {
    param([string]$Search = '')

    $needle = if ([string]::IsNullOrWhiteSpace($Search)) { '%' } else { '%' + $Search.Trim() + '%' }
    return Invoke-PurchaseQuery @'
SELECT
    id,
    entry_date,
    IFNULL(deal_number, '') AS deal_number,
    IFNULL(status, 'В работе') AS status,
    IFNULL(stage, 'Запросил поставщиков') AS stage,
    IFNULL(description, '') AS description,
    IFNULL(next_action, '') AS next_action,
    IFNULL(reminder_date, '') AS reminder_date,
    IFNULL(deadline_date, '') AS deadline_date,
    IFNULL(priority, '3') AS priority,
    IFNULL(period, '') AS period,
    IFNULL(folder_path, '') AS folder_path,
    IFNULL(notes, '') AS notes,
    updated_at
FROM component_deals
WHERE IFNULL(deal_number, '') LIKE @needle
   OR IFNULL(description, '') LIKE @needle
ORDER BY id DESC
'@ @{ '@needle' = $needle }
}


function Update-ComponentDeal {
    param(
        [int]$Id,
        [string]$DealNumber,
        [string]$Status,
        [string]$Stage,
        [string]$Description,
        [string]$NextAction,
        [string]$ReminderDate,
        [string]$DeadlineDate,
        [string]$Priority,
        [string]$Period
    )

    if ($Id -le 0) { return }
    $dealNumberParam = if ([string]::IsNullOrWhiteSpace($DealNumber)) { '' } else { $DealNumber.Trim() }
    Invoke-PurchaseNonQuery @'
UPDATE component_deals SET
    deal_number = @deal_number,
    status = @status,
    stage = @stage,
    description = @description,
    next_action = @next_action,
    reminder_date = @reminder_date,
    deadline_date = @deadline_date,
    priority = @priority,
    period = @period,
    updated_at = @updated_at
WHERE id = @id
'@ @{
        '@deal_number' = $dealNumberParam
        '@status' = $Status
        '@stage' = $Stage
        '@description' = $Description
        '@next_action' = $NextAction
        '@reminder_date' = $ReminderDate
        '@deadline_date' = $DeadlineDate
        '@priority' = $Priority
        '@period' = $Period
        '@updated_at' = Get-NowText
        '@id' = $Id
    }
    [void](Ensure-ComponentDealFolder $Id $dealNumberParam)
    Write-ActivityLog 'component' $Id 'Изменена задача' $dealNumberParam 0 0
}

function Get-ComponentDealNotes {
    param([int]$Id)

    if ($Id -le 0) { return '' }
    $value = Invoke-PurchaseScalar 'SELECT IFNULL(notes, '''') FROM component_deals WHERE id = @id' @{ '@id' = $Id }
    if ($null -eq $value -or $value -is [DBNull]) { return '' }
    return [string]$value
}

function Set-ComponentDealNotes {
    param([int]$Id, [string]$Notes)

    if ($Id -le 0) { return }
    $current = Get-ComponentDealNotes $Id
    if ($current -eq $Notes) { return }

    Invoke-PurchaseNonQuery 'UPDATE component_deals SET notes = @notes, updated_at = @updated_at WHERE id = @id' @{
        '@notes' = $Notes
        '@updated_at' = Get-NowText
        '@id' = $Id
    }
    Write-ActivityLog 'component' $Id 'Обновлены записи по задаче' '' 0 0
}
function Delete-ComponentDeal {
    param([int]$Id)
    if ($Id -le 0) { return }
    $folderPath = [string](Invoke-PurchaseScalar 'SELECT IFNULL(folder_path, '''') FROM component_deals WHERE id = @id' @{ '@id' = $Id })
    Invoke-PurchaseNonQuery 'DELETE FROM component_deals WHERE id = @id' @{ '@id' = $Id }
    Remove-ComponentDealFolder $Id $folderPath
    Write-ActivityLog 'component' $Id 'Удалена задача' '' 0 0
}

function Get-SupplierNamesFromRrfqResult {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Файл не найден: $Path"
    }

    $excel = Open-ExcelApp
    $wb = $null
    $names = New-Object System.Collections.ArrayList
    try {
        $wb = $excel.Workbooks.Open($Path, 0, $true)
        foreach ($ws in $wb.Worksheets) {
            $sheet = New-SheetData $ws
            $map = Find-TableMap $sheet @('Standard', 'Delivery')
            if ($null -eq $map -or -not $map.Fields.ContainsKey('Supplier')) {
                continue
            }

            $bounds = Get-SheetBounds $sheet
            $startRow = $map.HeaderRow + 1
            if ($map.Type -eq 'Standard' -and $startRow -lt 49) {
                $startRow = 49
            }

            for ($row = $startRow; $row -le $bounds.LastRow; $row++) {
                if (Test-SheetRowHidden $sheet $row) { continue }
                $supplier = Get-FieldText $sheet $map.Fields 'Supplier' $row
                if (-not [string]::IsNullOrWhiteSpace($supplier)) {
                    [void]$names.Add($supplier.Trim())
                }
            }
        }
    } finally {
        if ($null -ne $wb) {
            $wb.Close($false)
            Release-ComObject $wb
        }
    }

    return @($names | Sort-Object -Unique)
}

function Format-Price {
    param($Price)
    if ($null -eq $Price) { return '' }
    return ('{0:N6}' -f [double]$Price)
}

function Read-CompelHtmlText {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "HTML-файл не найден: $Path"
    }

    $bytes = [IO.File]::ReadAllBytes($Path)
    $utf8 = [Text.UTF8Encoding]::new($false, $false)
    $text = $utf8.GetString($bytes)
    if ($text -match '(?i)charset\s*=\s*"?windows-1251') {
        $text = [Text.Encoding]::GetEncoding(1251).GetString($bytes)
    }
    return $text
}

function Get-CompelRegexFirst {
    param([string]$Pattern, [string]$Text)

    $match = [regex]::Match($Text, $Pattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    if ($match.Success -and $match.Groups.Count -gt 1) {
        return $match.Groups[1].Value
    }
    return ''
}

function ConvertTo-CompelCleanText {
    param($Value)

    if ($null -eq $Value) { return '' }
    $text = [regex]::Replace([string]$Value, '<[^>]+>', ' ')
    $text = [Net.WebUtility]::HtmlDecode($text)
    if ($null -eq $text) { return '' }
    $text = $text.Replace([char]0x00A0, ' ')
    return ([regex]::Replace($text, '\s+', ' ')).Trim()
}

function ConvertTo-CompelInt {
    param($Value)

    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Replace(' ', '').Trim()
    if ([string]::IsNullOrWhiteSpace($text) -or $text -notmatch '^-?\d+$') {
        return $null
    }
    return [int]$text
}

function ConvertTo-CompelFloat {
    param($Value)

    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Replace(' ', '').Replace(',', '.').Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    [double]$number = 0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return $null
}

function Split-CompelMatchedPart {
    param([string]$Value)

    $text = ([string]$Value).Trim()
    $match = [regex]::Match($text, '^(.+?)\s*\(([^()]*)\)$')
    if (-not $match.Success) {
        return [pscustomobject]@{ Part = $text; Manufacturer = '' }
    }

    return [pscustomobject]@{
        Part = $match.Groups[1].Value.Trim()
        Manufacturer = $match.Groups[2].Value.Trim()
    }
}

function Get-CompelLeadDays {
    param([string]$Value)

    $match = [regex]::Match(([string]$Value), '\d+')
    if ($match.Success) { return [int]$match.Value }
    return $null
}

function Get-CompelMaxLeadTime {
    param([object[]]$Rows)

    $maxDays = $null
    $maxText = ''
    foreach ($row in @($Rows)) {
        $days = Get-CompelLeadDays $row.LeadTime
        if ($null -ne $days -and ($null -eq $maxDays -or $days -gt $maxDays)) {
            $maxDays = $days
            $maxText = [string]$row.LeadTime
        }
    }
    return $maxText
}

function New-CompelSummary {
    param([string]$HtmlText, [object[]]$Rows)

    $rowsList = @($Rows)
    $missingRows = New-Object System.Collections.ArrayList
    [double]$selectedTotalUsd = 0
    [int]$pricedCount = 0
    $longestLeadDays = $null
    $longestLeadTime = ''

    foreach ($row in $rowsList) {
        if ($null -ne $row.Total) {
            $pricedCount++
            $selectedTotalUsd += [double]$row.Total
        } elseif ($null -ne $row.Number) {
            [void]$missingRows.Add([int]$row.Number)
        }

        $days = Get-CompelLeadDays $row.LeadTime
        if ($null -ne $days -and ($null -eq $longestLeadDays -or $days -gt $longestLeadDays)) {
            $longestLeadDays = $days
            $longestLeadTime = [string]$row.LeadTime
        }
    }

    $cheapTotalUsd = $null
    $cheapLeadTime = ''
    $cheapMatch = [regex]::Match(
        $HtmlText,
        '<div data-set="cheap"[^>]*>.*?<price2\s+val="([^"]+)"\s+cur="USD".*?<span class="bom-settings-offer-type-dlv">([^<]+)</span>',
        [Text.RegularExpressions.RegexOptions]::Singleline
    )
    if ($cheapMatch.Success) {
        $cheapTotalUsd = ConvertTo-CompelFloat $cheapMatch.Groups[1].Value
        $cheapLeadTime = ConvertTo-CompelCleanText $cheapMatch.Groups[2].Value
    }

    $optimalTotalUsd = $null
    $optimalLeadTime = ''
    $optimalMatch = [regex]::Match(
        $HtmlText,
        '<div data-set="optimal"[^>]*>.*?<price2\s+val="([^"]+)"\s+cur="USD".*?<span class="bom-settings-offer-type-dlv">([^<]+)</span>',
        [Text.RegularExpressions.RegexOptions]::Singleline
    )
    if ($optimalMatch.Success) {
        $optimalTotalUsd = ConvertTo-CompelFloat $optimalMatch.Groups[1].Value
        $optimalLeadTime = ConvertTo-CompelCleanText $optimalMatch.Groups[2].Value
    }

    return [pscustomobject]@{
        TotalRows = $rowsList.Count
        PricedRows = $pricedCount
        MissingRows = @($missingRows)
        SelectedTotalUsd = $selectedTotalUsd
        LongestLeadDays = $longestLeadDays
        LongestLeadTime = $longestLeadTime
        CheapTotalUsd = $cheapTotalUsd
        CheapLeadTime = $cheapLeadTime
        OptimalTotalUsd = $optimalTotalUsd
        OptimalLeadTime = $optimalLeadTime
    }
}

function Parse-CompelHtml {
    param([string]$Path)

    $text = Read-CompelHtmlText $Path
    $chunks = [regex]::Split($text, '(?=<tr data-is="ac-row" class="ac-row)')
    $rows = New-Object System.Collections.ArrayList

    foreach ($chunk in $chunks) {
        if (-not ([string]$chunk).StartsWith('<tr data-is="ac-row"')) {
            continue
        }

        $number = ConvertTo-CompelInt (ConvertTo-CompelCleanText (Get-CompelRegexFirst '<div class="ac-row-num">(.*?)</div>' $chunk))
        $sourceName = ConvertTo-CompelCleanText (Get-CompelRegexFirst '<span[^>]*class="ac-row-name[^>]*>(.*?)</span>' $chunk)
        $matchedText = ConvertTo-CompelCleanText (Get-CompelRegexFirst '<a[^>]*class="ac-item-pn-link"[^>]*>(.*?)</a>' $chunk)
        $matched = Split-CompelMatchedPart $matchedText
        $package = ConvertTo-CompelCleanText (Get-CompelRegexFirst '<span[^>]*class="ac-part-mpq"[^>]*>(.*?)</span>' $chunk)
        $stockCount = ConvertTo-CompelInt (ConvertTo-CompelCleanText (Get-CompelRegexFirst '<span[^>]*class="ac-item-stocks-key"[^>]*>(.*?)</span>' $chunk))
        $leadTime = ConvertTo-CompelCleanText (Get-CompelRegexFirst '<span[^>]*class="ac-item-offer-dlv"[^>]*>(.*?)</span>' $chunk)
        $quantityText = ConvertTo-CompelCleanText (Get-CompelRegexFirst '<span[^>]*class="ac-item-offer-qty"[^>]*>(.*?)</span>' $chunk)
        $quantity = ConvertTo-CompelInt ($quantityText -replace '\s*x\s*$', '')

        $unitPrice = $null
        $currency = ''
        $unitMatch = [regex]::Match(
            $chunk,
            '<span class="ac-item-offer-price">.*?<price2\s+val="([^"]+)"\s+cur="([^"]+)"',
            [Text.RegularExpressions.RegexOptions]::Singleline
        )
        if ($unitMatch.Success) {
            $unitPrice = ConvertTo-CompelFloat $unitMatch.Groups[1].Value
            $currency = $unitMatch.Groups[2].Value
        }

        $total = $null
        $totalCurrency = ''
        $totalBlock = Get-CompelRegexFirst '<div class="ac-row-total-sum[^"]*"[^>]*>(.*?)</div>' $chunk
        if (-not [string]::IsNullOrWhiteSpace($totalBlock)) {
            $totalMatch = [regex]::Match($totalBlock, '<price2\s+val="([^"]+)"\s+cur="([^"]+)"', [Text.RegularExpressions.RegexOptions]::Singleline)
            if ($totalMatch.Success) {
                $total = ConvertTo-CompelFloat $totalMatch.Groups[1].Value
                $totalCurrency = $totalMatch.Groups[2].Value
            }
        }

        [void]$rows.Add([pscustomobject]@{
            Number = $number
            SourceName = $sourceName
            MatchedPart = $matched.Part
            Manufacturer = $matched.Manufacturer
            Package = $package
            StockCount = $stockCount
            LeadTime = $leadTime
            Quantity = $quantity
            UnitPrice = $unitPrice
            Currency = $currency
            Total = $total
            TotalCurrency = $totalCurrency
        })
    }

    $rowsArray = @($rows)
    return [pscustomobject]@{
        Rows = $rowsArray
        Summary = New-CompelSummary $text $rowsArray
    }
}

function Format-CompelNumber {
    param($Value, [int]$Digits = 2)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return '' }
    try {
        return (("{0:N$Digits}" -f [double]$Value))
    } catch {
        return [string]$Value
    }
}

function Set-CompelCellValue {
    param($Worksheet, [int]$Row, [int]$Column, $Value)

    if ($null -eq $Value) { return }
    $cell = $Worksheet.Cells.Item($Row, $Column)
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int] -or $Value -is [long] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        [void]($cell.Value2 = [double]$Value)
    } else {
        [void]($cell.Value2 = [string]$Value)
    }
}

function Set-CompelNumberFormat {
    param($Range, [string]$Format, [string]$LocalFormat = '')

    if ($null -eq $Range -or [string]::IsNullOrWhiteSpace($Format)) { return }
    try {
        $Range.NumberFormat = $Format
    } catch {
        if (-not [string]::IsNullOrWhiteSpace($LocalFormat)) {
            try { $Range.NumberFormatLocal = $LocalFormat } catch { }
        }
    }
}

function Write-CompelRrfqWorkbook {
    param([string]$Path, [object[]]$Rows)

    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Не задан файл выгрузки.' }
    $fullPath = [IO.Path]::GetFullPath($Path)
    $dir = Split-Path -Parent $fullPath
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }

    $excel = Open-ExcelApp
    $wb = $null
    $ws = $null
    try {
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        [void]($ws.Name = 'RRFQ')
        $headers = @('Value', 'Russian remark', 'China remark', 'PN', 'D/C', 'Mfg from Russia', 'Mfg from China', 'Q-ty in packing', 'Q-ty to Buy/pcs', 'Unit price', 'Currency', 'Lead time', 'Supplier')
        for ($i = 0; $i -lt $headers.Count; $i++) { Set-CompelCellValue $ws 1 ($i + 1) $headers[$i] }
        $rowIndex = 2
        foreach ($row in @($Rows)) {
            $source = [string]$row.SourceName
            $matched = [string]$row.MatchedPart
            $pn = $matched
            if (-not [string]::IsNullOrWhiteSpace($source) -and [string]::Equals($source.Trim(), $matched.Trim(), [System.StringComparison]::OrdinalIgnoreCase)) { $pn = '' }
            $values = @($source, '', '', $pn, '', '', [string]$row.Manufacturer, [string]$row.Package, [string]$row.Quantity, $row.UnitPrice, [string]$row.Currency, [string]$row.LeadTime, 'Компэл')
            for ($i = 0; $i -lt $values.Count; $i++) { Set-CompelCellValue $ws $rowIndex ($i + 1) $values[$i] }
            $rowIndex++
        }
        [void]($ws.Range('A1:M1').Font.Bold = $true)
        [void]($ws.Range('A1:M1').Interior.ColorIndex = 15)
        $ws.Range('A1:M' + [Math]::Max(1, $rowIndex - 1)).AutoFilter() | Out-Null
        [void]$ws.Columns.AutoFit()
        [void]$wb.SaveAs($fullPath, 51)
        return $fullPath
    } finally {
        if ($null -ne $wb) { try { [void]$wb.Close($false) } catch { } }
        Release-ComObject $ws
        Release-ComObject $wb
    }
}

function Write-CompelWorkbook {
    param([string]$Path, [object[]]$Rows, $Summary)

    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Documents folder is required.' }
    if ($null -eq $Summary) { throw 'Сначала разберите HTML-файл Компэла.' }

    $rowsList = @($Rows)
    $fullPath = [IO.Path]::GetFullPath($Path)
    $dir = Split-Path -Parent $fullPath
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        [void](New-Item -ItemType Directory -Path $dir -Force)
    }

    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    $excel = Open-ExcelApp
    $wb = $null
    $summaryWs = $null
    $detailWs = $null
    $titleColor = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(31, 78, 121))
    $subColor = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(217, 234, 247))
    $whiteColor = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::White)

    try {
        $wb = $excel.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) {
            $sheet = $wb.Worksheets.Item($wb.Worksheets.Count)
            $sheet.Delete()
            Release-ComObject $sheet
        }

        $summaryWs = $wb.Worksheets.Item(1)
        $summaryWs.Name = 'Сводка'
        $detailWs = $wb.Worksheets.Add([Type]::Missing, $summaryWs)
        $detailWs.Name = 'Позиции'

        Set-CompelCellValue $summaryWs 1 1 'Расчет SDS Compel'
        [void]$summaryWs.Range('A1:D1').Merge()
        $summaryWs.Range('A1:D1').Interior.Color = $titleColor
        $summaryWs.Range('A1:D1').Font.Bold = $true
        $summaryWs.Range('A1:D1').Font.Color = $whiteColor
        $summaryWs.Range('A1:D1').Font.Size = 14

        Set-CompelCellValue $summaryWs 3 1 'Показатель'
        Set-CompelCellValue $summaryWs 3 2 'USD'
        Set-CompelCellValue $summaryWs 3 3 'Комментарий'
        $summaryWs.Range('A3:D3').Interior.Color = $subColor
        $summaryWs.Range('A3:D3').Font.Bold = $true

        $missingText = [string]::Join(', ', @($Summary.MissingRows | ForEach-Object { [string]$_ }))
        Set-CompelCellValue $summaryWs 4 1 'Итого строк с подобранной ценой'
        Set-CompelCellValue $summaryWs 4 2 ([double]$Summary.SelectedTotalUsd)
        Set-CompelCellValue $summaryWs 4 3 'Сумма по листу Позиции'
        Set-CompelCellValue $summaryWs 5 1 'Итог сайта: оптимизировано по цене'
        Set-CompelCellValue $summaryWs 5 2 $Summary.CheapTotalUsd
        Set-CompelCellValue $summaryWs 5 3 $Summary.CheapLeadTime
        Set-CompelCellValue $summaryWs 6 1 'Итог сайта: оптимизировано по сроку'
        Set-CompelCellValue $summaryWs 6 2 $Summary.OptimalTotalUsd
        Set-CompelCellValue $summaryWs 6 3 $Summary.OptimalLeadTime
        Set-CompelCellValue $summaryWs 7 1 'Всего строк'
        Set-CompelCellValue $summaryWs 7 2 ([int]$Summary.TotalRows)
        Set-CompelCellValue $summaryWs 8 1 'Строк с ценой'
        Set-CompelCellValue $summaryWs 8 2 ([int]$Summary.PricedRows)
        Set-CompelCellValue $summaryWs 9 1 'Строк без подобранного предложения'
        Set-CompelCellValue $summaryWs 9 2 ([int]$Summary.MissingRows.Count)
        Set-CompelCellValue $summaryWs 9 3 $missingText
        Set-CompelCellValue $summaryWs 10 1 'Самый долгий срок по позициям'
        Set-CompelCellValue $summaryWs 10 2 $Summary.LongestLeadDays
        Set-CompelCellValue $summaryWs 10 3 $Summary.LongestLeadTime

        foreach ($row in 4..10) {
            $summaryWs.Cells.Item($row, 1).Font.Bold = $true
            Set-CompelNumberFormat ($summaryWs.Cells.Item($row, 2)) '#,##0.00' '# ##0,00'
        }
        $summaryWs.Columns.Item(1).ColumnWidth = 34
        $summaryWs.Columns.Item(2).ColumnWidth = 16
        $summaryWs.Columns.Item(3).ColumnWidth = 38
        $summaryWs.Columns.Item(4).ColumnWidth = 8
        $summaryWs.Range('A1:D10').WrapText = $true
        $summaryWs.Range('A1:D10').VerticalAlignment = -4160

        $headers = @('N', 'Исходное наименование', 'Подобранный товар', 'Производитель', 'Срок поставки', 'Количество', 'Цена за шт', 'Валюта', 'Сумма', 'Валюта суммы')
        for ($i = 0; $i -lt $headers.Count; $i++) {
            Set-CompelCellValue $detailWs 1 ($i + 1) $headers[$i]
        }

        $rowIndex = 2
        foreach ($row in $rowsList) {
            Set-CompelCellValue $detailWs $rowIndex 1 $row.Number
            Set-CompelCellValue $detailWs $rowIndex 2 $row.SourceName
            Set-CompelCellValue $detailWs $rowIndex 3 $row.MatchedPart
            Set-CompelCellValue $detailWs $rowIndex 4 $row.Manufacturer
            Set-CompelCellValue $detailWs $rowIndex 5 $row.LeadTime
            Set-CompelCellValue $detailWs $rowIndex 6 $row.Quantity
            Set-CompelCellValue $detailWs $rowIndex 7 $row.UnitPrice
            Set-CompelCellValue $detailWs $rowIndex 8 $row.Currency
            Set-CompelCellValue $detailWs $rowIndex 9 $row.Total
            Set-CompelCellValue $detailWs $rowIndex 10 $row.TotalCurrency
            $rowIndex++
        }

        $totalRow = $rowsList.Count + 2
        Set-CompelCellValue $detailWs $totalRow 2 'Итого'
        Set-CompelCellValue $detailWs $totalRow 5 'Самый долгий срок'
        Set-CompelCellValue $detailWs $totalRow 6 (Get-CompelMaxLeadTime $rowsList)
        Set-CompelCellValue $detailWs $totalRow 9 ([double]$Summary.SelectedTotalUsd)
        Set-CompelCellValue $detailWs $totalRow 10 'USD'

        $detailWs.Range('A1:J1').Interior.Color = $titleColor
        $detailWs.Range('A1:J1').Font.Bold = $true
        $detailWs.Range('A1:J1').Font.Color = $whiteColor
        $detailWs.Range("A$totalRow:J$totalRow").Interior.Color = $subColor
        $detailWs.Range("A$totalRow:J$totalRow").Font.Bold = $true

        $detailWs.Columns.Item(1).ColumnWidth = 6
        $detailWs.Columns.Item(2).ColumnWidth = 28
        $detailWs.Columns.Item(3).ColumnWidth = 32
        $detailWs.Columns.Item(4).ColumnWidth = 18
        $detailWs.Columns.Item(5).ColumnWidth = 13
        $detailWs.Columns.Item(6).ColumnWidth = 13
        $detailWs.Columns.Item(7).ColumnWidth = 13
        $detailWs.Columns.Item(8).ColumnWidth = 10
        $detailWs.Columns.Item(9).ColumnWidth = 14
        $detailWs.Columns.Item(10).ColumnWidth = 13
        Set-CompelNumberFormat ($detailWs.Columns.Item(1)) '0' '0'
        Set-CompelNumberFormat ($detailWs.Columns.Item(6)) '0' '0'
        Set-CompelNumberFormat ($detailWs.Columns.Item(7)) '0.00000' '0,00000'
        Set-CompelNumberFormat ($detailWs.Columns.Item(9)) '#,##0.00' '# ##0,00'
        $detailWs.Range("A1:J$totalRow").WrapText = $true
        $detailWs.Range("A1:J$totalRow").VerticalAlignment = -4160

        $filterLastRow = [Math]::Max(1, $rowsList.Count + 1)
        [void]$detailWs.Range("A1:J$filterLastRow").AutoFilter()
        [void]$detailWs.Activate()
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.FreezePanes = $true
        [void]$summaryWs.Activate()

        $wb.SaveAs($fullPath, 51)
        return $fullPath
    } finally {
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch { }
        }
        Release-ComObject $detailWs
        Release-ComObject $summaryWs
        Release-ComObject $wb
    }
}
function Add-TableRowStyle {
    param($Table, [string]$SizeType, [double]$Value)
    $style = New-Object Windows.Forms.RowStyle
    $style.SizeType = [Enum]::Parse([Windows.Forms.SizeType], $SizeType)
    if ($SizeType -eq 'Absolute') {
        $style.Height = [single]$Value
    } elseif ($SizeType -eq 'Percent') {
        $style.Height = [single]$Value
    }
    [void]$Table.RowStyles.Add($style)
}

function Add-TableColumnStyle {
    param($Table, [string]$SizeType, [double]$Value)
    $style = New-Object Windows.Forms.ColumnStyle
    $style.SizeType = [Enum]::Parse([Windows.Forms.SizeType], $SizeType)
    if ($SizeType -eq 'Absolute') {
        $style.Width = [single]$Value
    } elseif ($SizeType -eq 'Percent') {
        $style.Width = [single]$Value
    }
    [void]$Table.ColumnStyles.Add($style)
}

function Get-UiColor {
    param([string]$Name)
    switch ($Name) {
        'Canvas'          { return [Drawing.Color]::FromArgb(248, 250, 252) }
        'Surface'         { return [Drawing.Color]::FromArgb(255, 255, 255) }
        'SurfaceAlt'      { return [Drawing.Color]::FromArgb(241, 245, 249) }
        'Line'            { return [Drawing.Color]::FromArgb(226, 232, 240) }
        'Border'          { return [Drawing.Color]::FromArgb(203, 213, 225) }
        'Header'          { return [Drawing.Color]::FromArgb(241, 245, 249) }
        'Text'            { return [Drawing.Color]::FromArgb(15, 23, 42) }
        'Muted'           { return [Drawing.Color]::FromArgb(100, 116, 139) }
        'Primary'         { return [Drawing.Color]::FromArgb(59, 130, 246) }
        'PrimaryHover'    { return [Drawing.Color]::FromArgb(37, 99, 235) }
        'PrimaryPress'    { return [Drawing.Color]::FromArgb(29, 78, 216) }
        'PrimaryText'     { return [Drawing.Color]::FromArgb(255, 255, 255) }
        'Secondary'       { return [Drawing.Color]::FromArgb(241, 245, 249) }
        'SecondaryText'   { return [Drawing.Color]::FromArgb(51, 65, 85) }
        'SecondaryBorder' { return [Drawing.Color]::FromArgb(203, 213, 225) }
        'Nav'             { return [Drawing.Color]::FromArgb(15, 23, 42) }
        'NavActive'       { return [Drawing.Color]::FromArgb(30, 41, 59) }
        'NavHover'        { return [Drawing.Color]::FromArgb(30, 41, 59) }
        'NavText'         { return [Drawing.Color]::FromArgb(226, 232, 240) }
        'NavMuted'        { return [Drawing.Color]::FromArgb(148, 163, 184) }
        'NavDivider'      { return [Drawing.Color]::FromArgb(51, 65, 85) }
        'NavAccent'       { return [Drawing.Color]::FromArgb(59, 130, 246) }
        'Success'         { return [Drawing.Color]::FromArgb(220, 252, 231) }
        'SuccessText'     { return [Drawing.Color]::FromArgb(22, 101, 52) }
        'Warn'            { return [Drawing.Color]::FromArgb(254, 249, 195) }
        'WarnText'        { return [Drawing.Color]::FromArgb(133, 77, 14) }
        'Info'            { return [Drawing.Color]::FromArgb(219, 234, 254) }
        'Attention'       { return [Drawing.Color]::FromArgb(243, 232, 255) }
        'Danger'          { return [Drawing.Color]::FromArgb(254, 226, 226) }
        'DangerText'      { return [Drawing.Color]::FromArgb(153, 27, 27) }
        'Shadow'          { return [Drawing.Color]::FromArgb(15, 0, 0, 0) }
        'Overlay'         { return [Drawing.Color]::FromArgb(80, 0, 0, 0) }
        default { return [Drawing.Color]::FromArgb(255, 255, 255) }
    }
}


function New-CockpitMetricCard {
    param([string]$Caption, [string]$Value, [string]$Hint, [string]$Tone = 'Info')
    $panel = New-Object Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $panel.Margin = New-Object Windows.Forms.Padding(0, 0, 10, 0)
    $panel.Padding = New-Object Windows.Forms.Padding(16, 10, 12, 8)
    $panel.BackColor = Get-UiColor $Tone

    $accent = New-Object Windows.Forms.Panel
    $accent.Dock = 'Left'
    $accent.Width = 5
    $accent.BackColor = switch ($Tone) {
        'Danger' { Get-UiColor 'DangerText' }
        'Warn' { Get-UiColor 'WarnText' }
        'Attention' { [Drawing.Color]::FromArgb(126, 34, 206) }
        default { Get-UiColor 'Primary' }
    }
    $panel.Controls.Add($accent)

    $captionLabel = New-Object Windows.Forms.Label
    $captionLabel.Text = $Caption
    $captionLabel.Dock = 'Top'
    $captionLabel.Height = 23
    $captionLabel.ForeColor = Get-UiColor 'Muted'
    $captionLabel.Font = New-Object Drawing.Font('Segoe UI Semibold', 9)
    $panel.Controls.Add($captionLabel)
    $valueLabel = New-Object Windows.Forms.Label
    $valueLabel.Text = $Value
    $valueLabel.Dock = 'Top'
    $valueLabel.Height = 36
    $valueLabel.ForeColor = Get-UiColor 'Text'
    $valueLabel.Font = New-Object Drawing.Font('Segoe UI', 19, [Drawing.FontStyle]::Bold)
    $panel.Controls.Add($valueLabel)
    $hintLabel = New-Object Windows.Forms.Label
    $hintLabel.Text = $Hint
    $hintLabel.Dock = 'Fill'
    $hintLabel.ForeColor = Get-UiColor 'Muted'
    $hintLabel.Font = New-Object Drawing.Font('Segoe UI', 8.5)
    $panel.Controls.Add($hintLabel)

    $baseColor = $panel.BackColor
    $hoverColor = [Drawing.Color]::FromArgb(
        [Math]::Min(255, $baseColor.R + 7),
        [Math]::Min(255, $baseColor.G + 7),
        [Math]::Min(255, $baseColor.B + 7))
    $panel.Add_MouseEnter({ $this.BackColor = $this.Tag.HoverColor })
    $panel.Add_MouseLeave({ $this.BackColor = $this.Tag.BaseColor })
    foreach ($control in @($accent, $captionLabel, $valueLabel, $hintLabel)) {
        $control.Add_MouseEnter({ $this.Parent.BackColor = $this.Parent.Tag.HoverColor })
        $control.Add_MouseLeave({ $this.Parent.BackColor = $this.Parent.Tag.BaseColor })
    }
    $panel.Tag = [pscustomobject]@{ Value = $valueLabel; Hint = $hintLabel; BaseColor = $baseColor; HoverColor = $hoverColor }
    return $panel
}

function Set-CockpitMetricCard {
    param($Card, [string]$Value, [string]$Hint)
    if ($null -eq $Card -or $null -eq $Card.Tag) { return }
    $Card.Tag.Value.Text = $Value
    $Card.Tag.Hint.Text = $Hint
}

function Add-DebouncedTextChanged {
    param($Control, [scriptblock]$Action, [int]$Delay = 280)

    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = $Delay
    $timer.Tag = $Action
    $timer.Add_Tick({
        $this.Stop()
        & $this.Tag
    })
    $Control.Tag = $timer
    $Control.Add_TextChanged({
        $this.Tag.Stop()
        $this.Tag.Start()
    })
}
function Enable-GridDoubleBuffer {
    param($Grid)

    try {
        $property = [Windows.Forms.DataGridView].GetProperty('DoubleBuffered', [Reflection.BindingFlags]'Instance, NonPublic')
        $property.SetValue($Grid, $true, $null)
    } catch {
    }
}

function Set-GridLook {
    param($Grid)
    Enable-GridDoubleBuffer $Grid
    $Grid.BackgroundColor = Get-UiColor 'Surface'
    $Grid.BorderStyle = [Windows.Forms.BorderStyle]::None
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.ColumnHeadersBorderStyle = [Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $Grid.ColumnHeadersHeight = 38
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = Get-UiColor 'Header'
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = Get-UiColor 'Muted'
    $Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Get-UiColor 'Header'
    $Grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Get-UiColor 'Muted'
    $Grid.ColumnHeadersDefaultCellStyle.Font = New-Object Drawing.Font('Segoe UI', 8.5, [Drawing.FontStyle]::Bold)
    $Grid.ColumnHeadersDefaultCellStyle.Padding = New-Object Windows.Forms.Padding(8, 0, 8, 0)
    $Grid.DefaultCellStyle.BackColor = Get-UiColor 'Surface'
    $Grid.DefaultCellStyle.ForeColor = Get-UiColor 'Text'
    $Grid.DefaultCellStyle.Font = New-Object Drawing.Font('Segoe UI', 9)
    $Grid.DefaultCellStyle.SelectionBackColor = [Drawing.Color]::FromArgb(219, 234, 254)
    $Grid.DefaultCellStyle.SelectionForeColor = Get-UiColor 'Text'
    $Grid.DefaultCellStyle.Padding = New-Object Windows.Forms.Padding(8, 0, 8, 0)
    $Grid.AlternatingRowsDefaultCellStyle.BackColor = Get-UiColor 'SurfaceAlt'
    $Grid.GridColor = Get-UiColor 'Line'
    $Grid.CellBorderStyle = [Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $Grid.RowTemplate.Height = 34
    $Grid.AutoSizeRowsMode = [Windows.Forms.DataGridViewAutoSizeRowsMode]::None
    if ($null -eq $script:GridHoverStates) { $script:GridHoverStates = @{} }
    $gridKey = [Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Grid)
    $script:GridHoverStates[$gridKey] = @{}
    $Grid.Add_CellMouseEnter({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $key = [Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($s)
        if (-not $script:GridHoverStates.ContainsKey($key)) { $script:GridHoverStates[$key] = @{} }
        $state = $script:GridHoverStates[$key]
        $row = $s.Rows[$e.RowIndex]
        if ($row.Selected) { return }
        if ($state.ContainsKey($e.RowIndex)) { return }
        $darken = 18
        $cellHoverColors = @{}
        $cellBackColors = @{}
        foreach ($cell in $row.Cells) {
            $cellBackColors[$cell.ColumnIndex] = $cell.Style.BackColor
            $baseColor = $cell.InheritedStyle.BackColor
            $cellHoverColors[$cell.ColumnIndex] = [Drawing.Color]::FromArgb(
                [Math]::Max(0, $baseColor.R - $darken),
                [Math]::Max(0, $baseColor.G - $darken),
                [Math]::Max(0, $baseColor.B - $darken)
            )
        }
        $state[$e.RowIndex] = [pscustomobject]@{
            RowBackColor = $row.DefaultCellStyle.BackColor
            CellBackColors = $cellBackColors
        }
        $row.DefaultCellStyle.BackColor = $cellHoverColors[0]
        foreach ($cell in $row.Cells) {
            $cell.Style.BackColor = $cellHoverColors[$cell.ColumnIndex]
        }
    })
    $Grid.Add_CellMouseLeave({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $key = [Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($s)
        if ($null -eq $script:GridHoverStates -or -not $script:GridHoverStates.ContainsKey($key)) { return }
        $state = $script:GridHoverStates[$key]
        if ($state.ContainsKey($e.RowIndex)) {
            $row = $s.Rows[$e.RowIndex]
            $rowState = $state[$e.RowIndex]
            $row.DefaultCellStyle.BackColor = $rowState.RowBackColor
            foreach ($cell in $row.Cells) {
                $cell.Style.BackColor = $rowState.CellBackColors[$cell.ColumnIndex]
            }
            $state.Remove($e.RowIndex)
        }
    })
    $Grid.Add_Disposed({
        param($s, $e)
        if ($null -eq $script:GridHoverStates) { return }
        $key = [Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($s)
        [void]$script:GridHoverStates.Remove($key)
    })
    $emptyText = -join @([char]0x41D, [char]0x435, [char]0x442, [char]0x20, [char]0x434, [char]0x430, [char]0x43D, [char]0x43D, [char]0x44B, [char]0x445)
$emptyLabel = New-Object Windows.Forms.Label
$emptyLabel.Name = 'EmptyStateLabel'
$emptyLabel.Text = $emptyText
$emptyLabel.ForeColor = Get-UiColor 'Muted'
$emptyLabel.Font = New-Object Drawing.Font('Segoe UI', 10)
$emptyLabel.BackColor = Get-UiColor 'Surface'
$emptyLabel.TextAlign = [Drawing.ContentAlignment]::MiddleCenter
$emptyLabel.Visible = ($Grid.Rows.Count -eq 0)
$Grid.Controls.Add($emptyLabel)
$emptyLabel.SetBounds(0, $Grid.ColumnHeadersHeight, $Grid.ClientSize.Width, [Math]::Max(0, $Grid.ClientSize.Height - $Grid.ColumnHeadersHeight))
$Grid.Add_SizeChanged({
    param($sender, $eventArgs)
    $label = $sender.Controls['EmptyStateLabel']
    if ($null -eq $label) { return }
    $headerHeight = $sender.ColumnHeadersHeight
    $label.SetBounds(0, $headerHeight, $sender.ClientSize.Width, [Math]::Max(0, $sender.ClientSize.Height - $headerHeight))
})
$Grid.Add_RowsAdded({
    param($sender, $eventArgs)
    $label = $sender.Controls['EmptyStateLabel']
    if ($null -ne $label) { $label.Visible = ($sender.Rows.Count -eq 0) }
})
$Grid.Add_RowsRemoved({
    param($sender, $eventArgs)
    $label = $sender.Controls['EmptyStateLabel']
    if ($null -ne $label) { $label.Visible = ($sender.Rows.Count -eq 0) }
})
}

function Set-PrimaryButtonLook {
    param($Button)
    $Button.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $Button.BackColor = Get-UiColor 'Primary'
    $Button.ForeColor = Get-UiColor 'PrimaryText'
    $Button.FlatAppearance.BorderSize = 0
    $Button.FlatAppearance.MouseOverBackColor = Get-UiColor 'PrimaryHover'
    $Button.FlatAppearance.MouseDownBackColor = Get-UiColor 'PrimaryPress'
    $Button.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
    $Button.Cursor = [Windows.Forms.Cursors]::Hand
    $applyRoundedRegion = {
        param($b)
        try {
            if ($b.Width -gt 2 -and $b.Height -gt 2) {
                $r = New-Object Drawing.Rectangle(0, 0, $b.Width, $b.Height)
                $b.Region = New-RoundedRectPath $r 6
            }
        } catch { }
    }
    & $applyRoundedRegion $Button
    $Button.Add_Resize({
        param($s, $e2)
        try {
            if ($s.Width -gt 2 -and $s.Height -gt 2) {
                $r = New-Object Drawing.Rectangle(0, 0, $s.Width, $s.Height)
                $s.Region = New-RoundedRectPath $r 6
            }
        } catch { }
    })
}

function Set-SecondaryButtonLook {
    param($Button)
    $Button.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $Button.BackColor = Get-UiColor 'Secondary'
    $Button.ForeColor = Get-UiColor 'SecondaryText'
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = Get-UiColor 'SecondaryBorder'
    $Button.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(226, 232, 240)
    $Button.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb(203, 213, 225)
    $Button.Font = New-Object Drawing.Font('Segoe UI', 9)
    $Button.Cursor = [Windows.Forms.Cursors]::Hand
}

function Set-NavButtonLook {
    param($Button, [bool]$Active = $false)
    $Button.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.FlatAppearance.MouseOverBackColor = Get-UiColor 'SurfaceAlt'
    $Button.FlatAppearance.MouseDownBackColor = Get-UiColor 'Info'
    $Button.BackColor = $(if ($Active) { Get-UiColor 'SurfaceAlt' } else { Get-UiColor 'Surface' })
    $Button.ForeColor = Get-UiColor 'Text'
    $fs = if ($Active) { [Drawing.FontStyle]::Bold } else { [Drawing.FontStyle]::Regular }
    $Button.Font = New-Object Drawing.Font('Segoe UI', 9, $fs)
    $Button.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $Button.Padding = New-Object Windows.Forms.Padding(16, 0, 0, 0)
    $Button.Cursor = [Windows.Forms.Cursors]::Hand
    if ($null -ne $Button.Image) {
        $Button.ImageAlign = [Drawing.ContentAlignment]::MiddleLeft
        $Button.TextImageRelation = [Windows.Forms.TextImageRelation]::ImageBeforeText
        $Button.Padding = New-Object Windows.Forms.Padding(12, 0, 0, 0)
    }
    $alreadyStyled = ($Button.Tag -is [hashtable] -and $Button.Tag.ContainsKey('NavStyled'))
    $Button.Tag = @{ Active = [bool]$Active; NavStyled = $true }
    if (-not $alreadyStyled) {
        $Button.Add_Paint({
            param($s, $e)
            $t = $s.Tag
            if ($t -is [hashtable] -and $t.Active) {
                $e.Graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $barBrush = New-Object Drawing.SolidBrush(Get-UiColor 'Primary')
                $e.Graphics.FillRectangle($barBrush, 0, 8, 3, $s.Height - 16)
                $barBrush.Dispose()
            }
        })
    }
}

function Set-InputLook {
    param($Control)

    $Control.BackColor = Get-UiColor 'Surface'
    $Control.ForeColor = Get-UiColor 'Text'
    $Control.Font = New-Object Drawing.Font('Segoe UI', 9)
    if ($Control -is [Windows.Forms.TextBox]) {
        $Control.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
        if ($Control.ReadOnly) {
            $Control.BackColor = Get-UiColor 'SurfaceAlt'
        }
    } elseif ($Control -is [Windows.Forms.ComboBox]) {
        $Control.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    }
}

function Set-GroupLook {
    param($Group)

    $Group.BackColor = Get-UiColor 'Surface'
    $Group.ForeColor = Get-UiColor 'Text'
    $Group.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
}

function Apply-ControlTreeLook {
    param($Control)

    foreach ($child in $Control.Controls) {
        if ($child -is [Windows.Forms.TextBox] -or $child -is [Windows.Forms.ComboBox]) {
            Set-InputLook $child
        } elseif ($child -is [Windows.Forms.Label]) {
            $child.ForeColor = Get-UiColor 'Text'
            $child.Font = New-Object Drawing.Font('Segoe UI', 9)
        } elseif ($child -is [Windows.Forms.GroupBox]) {
            Set-GroupLook $child
        } elseif ($child -is [Windows.Forms.Panel] -or $child -is [Windows.Forms.TableLayoutPanel] -or $child -is [Windows.Forms.FlowLayoutPanel]) {
            if ($child.BackColor -eq [Drawing.Color]::Transparent -or $child.BackColor.ToArgb() -eq [Drawing.SystemColors]::Control.ToArgb()) {
                $child.BackColor = Get-UiColor 'Canvas'
            }
        }

        if ($child.Controls.Count -gt 0) {
            Apply-ControlTreeLook $child
        }
    }
}

function Set-TextBoxPlaceholder {
    param($TextBox, [string]$Text)
    try {
        $TextBox.PlaceholderText = $Text
    } catch {
        $TextBox.Tag = $Text
    }
}

function Get-AppIconPath {
    return (Join-Path $script:AppRoot 'app\assets\procurement-control.ico')
}


function Set-TaskbarWindowIcon {
    param($Form)

    try {
        if (-not ('ProcurementControl.TaskbarNative' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace ProcurementControl {
    public static class TaskbarNative {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern int SetCurrentProcessExplicitAppUserModelID(string appID);

        [DllImport("user32.dll", EntryPoint = "SetClassLongPtrW", SetLastError = true)]
        public static extern IntPtr SetClassLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
    }
}

function Remove-Reminder {
    param([int]$ReminderId)

    if ($ReminderId -le 0) { throw 'Автоматическое напоминание нельзя удалить вручную.' }
    $reminder = Invoke-PurchaseQuery 'SELECT id, title, deal_id, supplier_id FROM reminders WHERE id = @id' @{ '@id' = $ReminderId }
    if ($reminder.Rows.Count -eq 0) { throw 'Напоминание не найдено.' }
    $row = $reminder.Rows[0]
    Invoke-PurchaseNonQuery 'DELETE FROM reminders WHERE id = @id' @{ '@id' = $ReminderId }
    $dealId = if ($null -eq $row.deal_id -or $row.deal_id -is [DBNull]) { 0 } else { [int]$row.deal_id }
    $supplierId = if ($null -eq $row.supplier_id -or $row.supplier_id -is [DBNull]) { 0 } else { [int]$row.supplier_id }
    Write-ActivityLog 'reminder' $ReminderId 'Удалено напоминание' ([string]$row.title) $dealId $supplierId
}
'@
        }

        [void][ProcurementControl.TaskbarNative]::SetCurrentProcessExplicitAppUserModelID('ProcurementControl.Desktop')
        if ($null -ne $Form.Icon) {
            $handle = $Form.Handle
            [void][ProcurementControl.TaskbarNative]::SetClassLongPtr($handle, -14, $Form.Icon.Handle)
            [void][ProcurementControl.TaskbarNative]::SetClassLongPtr($handle, -34, $Form.Icon.Handle)
        }
    } catch {
        # Taskbar decoration must never prevent the application from opening.
    }
}

function Set-AppBranding {
    param($Form)

    $Form.Text = 'Procurement control'
    $Form.BackColor = Get-UiColor 'Canvas'
    $iconPath = Get-AppIconPath
    if (Test-Path -LiteralPath $iconPath -PathType Leaf) {
        try {
            $Form.Icon = New-Object Drawing.Icon($iconPath)
            Set-TaskbarWindowIcon $Form
        } catch {
            # A broken icon file should not block the app from opening.
        }
    }
}


function New-RoundedRectPath {
    param([Drawing.Rectangle]$Rect, [int]$Radius)
    $path = New-Object Drawing.Drawing2D.GraphicsPath
    $d = [Math]::Max(0, [Math]::Min($Radius, [Math]::Min($Rect.Width, $Rect.Height) / 2))
    if ($d -le 0) { $path.AddRectangle($Rect); return $path }
    $path.AddArc($Rect.X, $Rect.Y, $d * 2, $d * 2, 180, 90)
    $path.AddArc($Rect.Right - $d * 2, $Rect.Y, $d * 2, $d * 2, 270, 90)
    $path.AddArc($Rect.Right - $d * 2, $Rect.Bottom - $d * 2, $d * 2, $d * 2, 0, 90)
    $path.AddArc($Rect.X, $Rect.Bottom - $d * 2, $d * 2, $d * 2, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-NavIcon {
    param([string]$Type)
    $bmp = New-Object Drawing.Bitmap(20, 20)
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([Drawing.Color]::Transparent)
    $pen = New-Object Drawing.Pen([Drawing.Color]::Black, 1.8)
    $pen.StartCap = [Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [Drawing.Drawing2D.LineCap]::Round
    $pen.LineJoin = [Drawing.Drawing2D.LineJoin]::Round
    $fb = New-Object Drawing.SolidBrush([Drawing.Color]::Black)
    switch ($Type) {
        'RRFQ' {
            $g.DrawRectangle($pen, 3, 2, 13, 16)
            $g.DrawLine($pen, 11, 2, 16, 7)
            $g.DrawLine($pen, 11, 2, 11, 7)
            $g.DrawLine($pen, 11, 7, 16, 7)
            $g.DrawLine($pen, 6, 10, 13, 10)
            $g.DrawLine($pen, 6, 13, 13, 13)
        }
        'Bitrix' {
            $g.DrawArc($pen, 3, 5, 8, 6, 180, 180)
            $g.DrawArc($pen, 9, 9, 8, 6, 0, 180)
            $g.DrawLine($pen, 3, 8, 3, 12)
            $g.DrawLine($pen, 17, 8, 17, 12)
        }
        'ExcelCompare' {
            $g.DrawRectangle($pen, 2, 4, 10, 13)
            $g.DrawRectangle($pen, 7, 2, 10, 13)
            $g.DrawLine($pen, 5, 8, 10, 8); $g.DrawLine($pen, 5, 11, 10, 11)
            $g.DrawLine($pen, 10, 6, 15, 6); $g.DrawLine($pen, 10, 9, 15, 9)
        }
        'Purchase' {
            $g.DrawRectangle($pen, 3, 4, 14, 14)
            $g.DrawRectangle($pen, 7, 2, 6, 4)
            $g.DrawLine($pen, 6, 9, 8, 11); $g.DrawLine($pen, 8, 11, 12, 7)
            $g.DrawLine($pen, 6, 14, 14, 14)
        }
        'Components' {
            [Drawing.PointF[]]$pts = @((New-Object Drawing.PointF(10,2)),(New-Object Drawing.PointF(17,6)),
                (New-Object Drawing.PointF(17,14)),(New-Object Drawing.PointF(10,18)),
                (New-Object Drawing.PointF(3,14)),(New-Object Drawing.PointF(3,6)))
            $g.DrawPolygon($pen, $pts)
            $g.FillEllipse($fb, 8.5, 8.5, 3, 3)
        }
        'Reminders' {
            $g.DrawArc($pen, 4, 3, 12, 10, 180, 180)
            $g.DrawLine($pen, 4, 8, 4, 14); $g.DrawLine($pen, 16, 8, 16, 14)
            $g.DrawLine($pen, 3, 14, 17, 14); $g.DrawLine($pen, 8, 17, 12, 17)
            $g.DrawLine($pen, 10, 1, 10, 3)
        }
        'QuoteBase' {
            $g.DrawEllipse($pen, 3, 2, 14, 5)
            $g.DrawLine($pen, 3, 4.5, 3, 15.5); $g.DrawLine($pen, 17, 4.5, 17, 15.5)
            $g.DrawArc($pen, 3, 13, 14, 5, 0, 180)
            $g.DrawArc($pen, 3, 8, 14, 4, 0, 180)
        }
        'History' {
            $g.DrawEllipse($pen, 3, 3, 14, 14)
            $g.DrawLine($pen, 10, 6, 10, 10); $g.DrawLine($pen, 10, 10, 14, 12)
            $g.DrawLine($pen, 1, 5, 3, 8); $g.DrawLine($pen, 3, 8, 6, 6)
        }
        'CompelParser' {
            $g.DrawEllipse($pen, 6, 6, 8, 8)
            $g.DrawLine($pen, 10, 1, 10, 5); $g.DrawLine($pen, 10, 15, 10, 19)
            $g.DrawLine($pen, 1, 10, 5, 10); $g.DrawLine($pen, 15, 10, 19, 10)
            $g.DrawLine($pen, 4, 4, 7, 7); $g.DrawLine($pen, 13, 13, 16, 16)
            $g.DrawLine($pen, 16, 4, 13, 7); $g.DrawLine($pen, 7, 13, 4, 16)
        }
        'PriceSearch' {
            $g.DrawEllipse($pen, 3, 3, 10, 10)
            $g.DrawLine($pen, 11, 11, 17, 17)
        }
        'Notes' {
            $g.DrawRectangle($pen, 3, 2, 14, 16)
            $g.DrawLine($pen, 3, 6, 17, 6)
            $g.DrawLine($pen, 6, 9, 14, 9); $g.DrawLine($pen, 6, 12, 14, 12)
            $g.DrawLine($pen, 6, 15, 11, 15)
        }
        'Instructions' {
            $g.DrawLine($pen, 10, 4, 10, 17)
            $g.DrawBezier($pen, 10, 4, 7, 3, 4, 3, 2, 5)
            $g.DrawLine($pen, 2, 5, 2, 16)
            $g.DrawBezier($pen, 2, 16, 4, 14, 7, 14, 10, 17)
            $g.DrawBezier($pen, 10, 4, 13, 3, 16, 3, 18, 5)
            $g.DrawLine($pen, 18, 5, 18, 16)
            $g.DrawBezier($pen, 18, 16, 16, 14, 13, 14, 10, 17)
        }
        'Settings' {
            $g.DrawLine($pen, 3, 6, 17, 6); $g.DrawLine($pen, 3, 10, 17, 10)
            $g.DrawLine($pen, 3, 14, 17, 14)
            $g.FillEllipse($fb, 6, 4, 4, 4)
            $g.FillEllipse($fb, 12, 8, 4, 4)
            $g.FillEllipse($fb, 8, 12, 4, 4)
        }
    }
    $fb.Dispose(); $pen.Dispose(); $g.Dispose()
    return $bmp
}

function New-CardPanel {
    param([int]$Radius = 8, [switch]$Shadow)
    $p = New-Object Windows.Forms.Panel
    try { [void]$p.GetType().InvokeMember('DoubleBuffered', ([System.Reflection.BindingFlags]::SetProperty -bor [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic), $null, $p, @($true)) } catch { }
    $p.Tag = @{ Radius = $Radius; HasShadow = [bool]$Shadow }
    $p.BackColor = [Drawing.Color]::Transparent
    $p.Add_Paint({
        param($s, $e)
        $c = $s; $g = $e.Graphics
        $g.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $t = $c.Tag; $r = if ($t -is [hashtable]) { $t.Radius } else { 8 }
        $hs = if ($t -is [hashtable]) { $t.HasShadow } else { $false }
        $ox = 0; $oy = 0; if ($hs) { $ox = 2; $oy = 2 }
        $bw = $c.ClientSize.Width - $ox - 1
        $bh = $c.ClientSize.Height - $oy - 1
        $b = New-Object Drawing.Rectangle($ox, $oy, $bw, $bh)
        $path = New-RoundedRectPath $b $r
        if ($hs) {
            $sr = New-Object Drawing.Rectangle(2, 2, $c.ClientSize.Width - 3, $c.ClientSize.Height - 3)
            $sp = New-RoundedRectPath $sr $r
            $g.FillPath((New-Object Drawing.SolidBrush([Drawing.Color]::FromArgb(25,0,0,0))), $sp)
            $sp.Dispose()
        }
        $g.FillPath((New-Object Drawing.SolidBrush([Drawing.Color]::White)), $path)
        $bp = New-Object Drawing.Pen((Get-UiColor 'Line'), 1)
        $g.DrawPath($bp, $path); $bp.Dispose(); $path.Dispose()
    })
    return $p
}


function Show-Toast {
    param([string]$Text, [string]$Kind = 'Info')
    $toast = New-Object Windows.Forms.Form
    $toast.FormBorderStyle = [Windows.Forms.FormBorderStyle]::None
    $toast.ShowInTaskbar = $false
    $toast.TopMost = $true
    $toast.StartPosition = [Windows.Forms.FormStartPosition]::Manual
    $toast.BackColor = [Drawing.Color]::FromArgb(15, 23, 42)
    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.ForeColor = [Drawing.Color]::White
    $lbl.Font = New-Object Drawing.Font('Segoe UI', 9)
    $lbl.AutoSize = $true
    $lbl.MaximumSize = New-Object Drawing.Size(460, 0)
    $lbl.Location = New-Object Drawing.Point(20, 12)
    $toast.Controls.Add($lbl)
    $bar = New-Object Windows.Forms.Panel
    $bar.Dock = [Windows.Forms.DockStyle]::Left
    $bar.Width = 4
    $barColor = switch ($Kind) {
        'Success' { [Drawing.Color]::FromArgb(74, 222, 128) }
        'Warn'    { [Drawing.Color]::FromArgb(250, 204, 21) }
        'Danger'  { [Drawing.Color]::FromArgb(248, 113, 113) }
        default   { [Drawing.Color]::FromArgb(96, 165, 250) }
    }
    $bar.BackColor = $barColor
    $toast.Controls.Add($bar)
    $toast.Add_Shown({
        $t = $toast
        $ps = $lbl.GetPreferredSize((New-Object Drawing.Size(460, 0)))
        $t.ClientSize = New-Object Drawing.Size([int]($ps.Width + 34), [int]($ps.Height + 24))
        $wa = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $t.Location = New-Object Drawing.Point([int]($wa.Right - $t.Width - 24), [int]($wa.Bottom - $t.Height - 56))
        $tm = New-Object Windows.Forms.Timer
        $tm.Interval = 3500
        $tm.Add_Tick({ $tm.Stop(); $t.Close(); $t.Dispose() })
        $tm.Start()
    })
    $toast.Show()
}

function Refresh-PreviewGrid {
    param($Analysis, $Grid, $LogBox, [string]$Filter = 'All', [string]$Search = '')

    $Grid.Rows.Clear()
    foreach ($decision in (Get-FilteredDecisions $Analysis $Filter $Search)) {
        $rowIndex = $Grid.Rows.Add(
            [bool]$decision.Include,
            [string]$decision.Status,
            [string]$decision.Row,
            [string]$decision.Value,
            [string]$decision.WinnerPN,
            [string]$decision.RussianRemark,
            [string]$decision.ChinaRemark,
            [string]$decision.DC,
            [string]$decision.MfgRussia,
            [string]$decision.MfgChina,
            [string]$decision.QtyPacking,
            [string]$decision.QtyToBuy,
            [string]$decision.WinnerSupplier,
            [string]$decision.WinnerReason,
            (Format-Price $decision.WinnerPrice),
            [string]$decision.WinnerLead,
            [string]$decision.WinnerLeadTotal,
            [string]$decision.QuotesSummary,
            [string]$decision.Warning
        )
        $Grid.Rows[$rowIndex].Tag = $decision.Id

        if ($decision.Status -eq 'OK') {
            $Grid.Rows[$rowIndex].DefaultCellStyle.BackColor = Get-UiColor 'Success'
        } elseif ($decision.Status -eq 'Нет квот') {
            $Grid.Rows[$rowIndex].DefaultCellStyle.BackColor = Get-UiColor 'SurfaceAlt'
            $Grid.Rows[$rowIndex].DefaultCellStyle.ForeColor = Get-UiColor 'Muted'
        } else {
            $Grid.Rows[$rowIndex].DefaultCellStyle.BackColor = Get-UiColor 'Warn'
        }

        if (-not [string]::IsNullOrWhiteSpace($decision.Warning)) {
            $Grid.Rows[$rowIndex].Cells['Warnings'].Style.BackColor = Get-UiColor 'Danger'
        }
    }

    $unresolved = @(Get-UnresolvedQuotes $Analysis)
    $okCount = @($Analysis.Decisions | Where-Object { $_.Winner }).Count
    $noQuoteCount = @($Analysis.Decisions | Where-Object { $_.Status -eq 'Нет квот' }).Count
    $LogBox.Text = "RFQ позиций: $($Analysis.RfqRows.Count)`r`nКвот прочитано: $($Analysis.Quotes.Count)`r`nПобедителей: $okCount`r`nБез квот: $noQuoteCount`r`nТребуют ручной проверки: $($unresolved.Count)"
}

function Get-FilteredDecisions {
    param($Analysis, [string]$Filter = 'All', [string]$Search = '')

    if ($null -eq $Analysis) {
        return @()
    }

    $items = @($Analysis.Decisions)
    switch ($Filter) {
        'Winners' { $items = @($items | Where-Object { $_.Winner }) }
        'NoQuotes' { $items = @($items | Where-Object { $_.Status -eq 'Нет квот' }) }
        'Warnings' { $items = @($items | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Warning) -or $_.Status -eq 'Нет сравнимой цены' }) }
        'Manual' {
            $manualIds = @{}
            foreach ($quote in (Get-UnresolvedQuotes $Analysis)) {
                foreach ($id in @($quote.CandidateIds)) {
                    $manualIds[$id] = $true
                }
                if (-not [string]::IsNullOrWhiteSpace($quote.TargetId)) {
                    $manualIds[$quote.TargetId] = $true
                }
            }
            foreach ($decision in @($Analysis.Decisions)) {
                if ($decision.IsManualWinner) {
                    $manualIds[$decision.Id] = $true
                    continue
                }
                if (@($decision.Quotes | Where-Object { ([string]$_.MatchStatus).StartsWith('Manual') }).Count -gt 0) {
                    $manualIds[$decision.Id] = $true
                }
            }
            $items = @($items | Where-Object { $manualIds.ContainsKey($_.Id) })
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Search)) {
        $needle = $Search.Trim().ToLowerInvariant()
        $items = @($items | Where-Object {
            ([string]$_.Value).ToLowerInvariant().Contains($needle) -or
            ([string]$_.WinnerPN).ToLowerInvariant().Contains($needle) -or
            ([string]$_.RussianRemark).ToLowerInvariant().Contains($needle) -or
            ([string]$_.ChinaRemark).ToLowerInvariant().Contains($needle) -or
            ([string]$_.WinnerSupplier).ToLowerInvariant().Contains($needle) -or
            ([string]$_.QuotesSummary).ToLowerInvariant().Contains($needle)
        })
    }

    return @($items)
}

function Get-DecisionById {
    param($Analysis, [string]$Id)
    if ($null -eq $Analysis -or [string]::IsNullOrWhiteSpace($Id)) {
        return $null
    }

    return @($Analysis.Decisions | Where-Object { $_.Id -eq $Id } | Select-Object -First 1)[0]
}

function Get-QuoteById {
    param($Analysis, [string]$Id)
    if ($null -eq $Analysis -or [string]::IsNullOrWhiteSpace($Id)) {
        return $null
    }

    return @($Analysis.Quotes | Where-Object { $_.Id -eq $Id } | Select-Object -First 1)[0]
}

function Set-RfqRowManualWinner {
    param($Analysis, [string]$DecisionId, [string]$QuoteId)

    $decision = Get-DecisionById $Analysis $DecisionId
    if ($null -eq $decision) {
        throw 'Строка preview не найдена.'
    }

    $quote = Get-QuoteById $Analysis $QuoteId
    if ($null -eq $quote -or $quote.TargetId -ne $decision.Id) {
        throw 'Выбранная квота не относится к этой строке RFQ.'
    }

    if (-not (Test-QuoteSelectable $quote)) {
        throw 'У выбранной квоты нет сравнимой цены.'
    }

    $decision.RfqRow.ManualWinnerId = $quote.Id
    $quote.MatchStatus = 'ManualWinner'
    Rebuild-DecisionsAfterManualMatch $Analysis
}

function Add-ManualQuoteToDecision {
    param(
        $Analysis,
        [string]$DecisionId,
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

    $decision = Get-DecisionById $Analysis $DecisionId
    if ($null -eq $decision) {
        throw 'Строка preview не найдена.'
    }

    $quote = New-ManualQuote $decision $Supplier $ChinaRemark $PN $DC $MfgRussia $MfgChina $QtyPacking $QtyToBuy $UnitPrice $LeadTime
    $Analysis.Quotes += $quote
    $decision.RfqRow.ManualWinnerId = $quote.Id
    Rebuild-DecisionsAfterManualMatch $Analysis
}

function Refresh-DecisionDetails {
    param($Analysis, $PreviewGrid, $SummaryBox, $QuotesGrid)

    $QuotesGrid.Rows.Clear()
    if ($null -eq $Analysis -or $PreviewGrid.SelectedRows.Count -eq 0) {
        $SummaryBox.Text = ''
        return
    }

    $decision = Get-DecisionById $Analysis ([string]$PreviewGrid.SelectedRows[0].Tag)
    if ($null -eq $decision) {
        $SummaryBox.Text = ''
        return
    }

    $SummaryBox.Text = "Лист: $($decision.SheetName)    Строка: $($decision.Row)`r`nValue: $($decision.Value)`r`nPN победителя: $($decision.WinnerPN)`r`nСтатус: $($decision.Status)    Победитель: $($decision.WinnerSupplier)    Почему выбран: $($decision.WinnerReason)`r`nЦена: $(Format-Price $decision.WinnerPrice)    Срок: $($decision.WinnerLead) / $($decision.WinnerLeadTotal)`r`nПредупреждения: $($decision.Warning)"

    foreach ($quote in @($decision.Quotes)) {
        $isWinner = ($null -ne $decision.Winner -and $quote.Id -eq $decision.Winner.Id)
        $rowIndex = $QuotesGrid.Rows.Add(
            $(if ($isWinner) { 'Да' } else { '' }),
            [string]$quote.Supplier,
            (Format-Price $quote.UnitPrice),
            [string]$quote.LeadTime,
            [string]$quote.LeadTimeTotal,
            [string]$quote.PN,
            [string]$quote.MfgChina,
            [string]$quote.QtyPacking,
            [string]$quote.QtyToBuy,
            [string]$quote.MatchStatus,
            [string]$quote.Warning
        )
        $QuotesGrid.Rows[$rowIndex].Tag = $quote.Id
        if ($isWinner) {
            $QuotesGrid.Rows[$rowIndex].DefaultCellStyle.BackColor = Get-UiColor 'Success'
        }
    }
}

function Update-DecisionsFromGrid {
    param($Analysis, $Grid)

    $byId = @{}
    foreach ($decision in $Analysis.Decisions) {
        $byId[$decision.Id] = $decision
    }

    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow -or $null -eq $row.Tag) {
            continue
        }

        if ($byId.ContainsKey([string]$row.Tag)) {
            $byId[[string]$row.Tag].Include = [bool]$row.Cells[0].Value
        }
    }
}

function Rebuild-DecisionsAfterManualMatch {
    param($Analysis)

    $oldIncludes = @{}
    foreach ($decision in $Analysis.Decisions) {
        $oldIncludes[$decision.Id] = [bool]$decision.Include
    }

    $Analysis.Decisions = Build-Decisions $Analysis.RfqRows $Analysis.Quotes $Analysis.Priority
    foreach ($decision in $Analysis.Decisions) {
        if ($oldIncludes.ContainsKey($decision.Id)) {
            $decision.Include = [bool]$oldIncludes[$decision.Id]
        }
    }
}

function Show-ManualQuoteDialog {
    param($Analysis, $Decision, $Owner)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    if ($null -eq $Analysis -or $null -eq $Decision) {
        return $false
    }

    $form = New-Object Windows.Forms.Form
    $form.Text = 'Ручной ввод квоты'
    $form.Width = 560
    $form.Height = 520
    $form.MinimumSize = New-Object Drawing.Size(520, 480)
    $form.StartPosition = 'CenterParent'
    $form.Font = New-Object Drawing.Font('Segoe UI', 9)
    $form.BackColor = Get-UiColor 'Canvas'
    $form.AutoScaleMode = [Windows.Forms.AutoScaleMode]::Dpi
    $form.KeyPreview = $true

    $layout = New-Object Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = New-Object Windows.Forms.Padding(12)
    $layout.ColumnCount = 2
    $layout.RowCount = 12
    Add-TableColumnStyle $layout 'Absolute' 150
    Add-TableColumnStyle $layout 'Percent' 100
    for ($i = 0; $i -lt 10; $i++) {
        Add-TableRowStyle $layout 'Absolute' 34
    }
    Add-TableRowStyle $layout 'Percent' 100
    Add-TableRowStyle $layout 'Absolute' 44
    $form.Controls.Add($layout)

    function Add-ManualQuoteLabel {
        param([string]$Text, [int]$Row)
        $label = New-Object Windows.Forms.Label
        $label.Text = $Text
        $label.Dock = 'Fill'
        $label.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
        $layout.Controls.Add($label, 0, $Row)
    }

    function Add-ManualQuoteTextBox {
        param([int]$Row, [string]$Value = '')
        $box = New-Object Windows.Forms.TextBox
        $box.Dock = 'Fill'
        $box.Margin = New-Object Windows.Forms.Padding(0, 3, 0, 3)
        $box.Text = $Value
        $layout.Controls.Add($box, 1, $Row)
        return $box
    }

    Add-ManualQuoteLabel 'Поставщик' 0
    $cmbSupplier = New-Object Windows.Forms.ComboBox
    $cmbSupplier.Dock = 'Fill'
    $cmbSupplier.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDown
    $cmbSupplier.Margin = New-Object Windows.Forms.Padding(0, 3, 0, 3)
    foreach ($name in (Get-SupplierNames)) {
        [void]$cmbSupplier.Items.Add($name)
    }
    $cmbSupplier.Text = 'Др.'
    $layout.Controls.Add($cmbSupplier, 1, 0)

    Add-ManualQuoteLabel 'Цена Unit price' 1
    $txtPrice = Add-ManualQuoteTextBox 1 ''

    Add-ManualQuoteLabel 'Lead Time' 2
    $txtLead = Add-ManualQuoteTextBox 2 ''

    Add-ManualQuoteLabel 'PN' 3
    $txtPN = Add-ManualQuoteTextBox 3 ''

    Add-ManualQuoteLabel 'D/C' 4
    $txtDC = Add-ManualQuoteTextBox 4 ''

    Add-ManualQuoteLabel 'MFG from Russia' 5
    $txtMfgRussia = Add-ManualQuoteTextBox 5 ([string]$Decision.MfgRussia)

    Add-ManualQuoteLabel 'MFG from China' 6
    $txtMfgChina = Add-ManualQuoteTextBox 6 ([string]$Decision.MfgChina)

    Add-ManualQuoteLabel 'Q-ty in packing' 7
    $txtQtyPacking = Add-ManualQuoteTextBox 7 ([string]$Decision.QtyPacking)

    Add-ManualQuoteLabel 'Q-ty to Buy/pcs' 8
    $txtQtyToBuy = Add-ManualQuoteTextBox 8 ([string]$Decision.QtyToBuy)

    Add-ManualQuoteLabel 'China remark' 9
    $txtChinaRemark = Add-ManualQuoteTextBox 9 ([string]$Decision.ChinaRemark)

    $info = New-Object Windows.Forms.TextBox
    $info.Multiline = $true
    $info.ReadOnly = $true
    $info.Dock = 'Fill'
    $info.BackColor = Get-UiColor 'SurfaceAlt'
    $info.Text = "Строка: $($Decision.Row)`r`nValue: $($Decision.Value)`r`nПосле сохранения эта квота станет победителем для выбранной строки."
    $layout.Controls.Add($info, 0, 10)
    $layout.SetColumnSpan($info, 2)

    $buttonPanel = New-Object Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = 'Fill'
    $buttonPanel.FlowDirection = 'RightToLeft'
    $buttonPanel.WrapContents = $false
    $layout.Controls.Add($buttonPanel, 0, 11)
    $layout.SetColumnSpan($buttonPanel, 2)

    $btnOk = New-Object Windows.Forms.Button
    $btnOk.Text = 'Сохранить'
    $btnOk.Width = 120
    $btnOk.Height = 30
    Set-PrimaryButtonLook $btnOk
    $buttonPanel.Controls.Add($btnOk)

    $btnCancel = New-Object Windows.Forms.Button
    $btnCancel.Text = 'Отмена'
    $btnCancel.Width = 100
    $btnCancel.Height = 30
    Set-SecondaryButtonLook $btnCancel
    $buttonPanel.Controls.Add($btnCancel)

    $form.Tag = $false
    $btnOk.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($txtPrice.Text)) {
                throw 'Введите цену.'
            }

            Add-ManualQuoteToDecision `
                $Analysis `
                $Decision.Id `
                $cmbSupplier.Text `
                $txtChinaRemark.Text `
                $txtPN.Text `
                $txtDC.Text `
                $txtMfgRussia.Text `
                $txtMfgChina.Text `
                $txtQtyPacking.Text `
                $txtQtyToBuy.Text `
                $txtPrice.Text `
                $txtLead.Text
            $form.Tag = $true
            $form.Close()
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Ошибка ручной квоты') | Out-Null
        }
    })

    $btnCancel.Add_Click({ $form.Close() })
    Apply-ControlTreeLook $form
    [void]$form.ShowDialog($Owner)
    return [bool]$form.Tag
}

function Show-ManualMatchDialog {
    param($Analysis, $PreviewGrid, $LogBox)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object Windows.Forms.Form
    $form.Text = 'Ручное сопоставление'
    $form.Width = 1220
    $form.Height = 760
    $form.Font = New-Object Drawing.Font('Segoe UI', 9)
    $form.BackColor = Get-UiColor 'Canvas'
    $form.StartPosition = 'CenterParent'

    $split = New-Object Windows.Forms.SplitContainer
    $split.Dock = 'Fill'
    $split.Orientation = 'Vertical'
    $split.SplitterDistance = 580
    $form.Controls.Add($split)

    $leftPanel = New-Object Windows.Forms.TableLayoutPanel
    $leftPanel.Dock = 'Fill'
    $leftPanel.RowCount = 2
    $leftPanel.ColumnCount = 1
    Add-TableRowStyle $leftPanel 'Absolute' 36
    Add-TableRowStyle $leftPanel 'Percent' 100
    $split.Panel1.Controls.Add($leftPanel)

    $txtQuoteSearch = New-Object Windows.Forms.TextBox
    $txtQuoteSearch.Dock = 'Fill'
    $txtQuoteSearch.Margin = New-Object Windows.Forms.Padding(4, 6, 4, 4)
    Set-TextBoxPlaceholder $txtQuoteSearch 'Поиск по несопоставленным квотам'
    $leftPanel.Controls.Add($txtQuoteSearch, 0, 0)

    $unmatchedGrid = New-Object Windows.Forms.DataGridView
    $unmatchedGrid.Dock = 'Fill'
    $unmatchedGrid.ReadOnly = $true
    $unmatchedGrid.SelectionMode = 'FullRowSelect'
    $unmatchedGrid.MultiSelect = $false
    $unmatchedGrid.AllowUserToAddRows = $false
    $unmatchedGrid.RowHeadersVisible = $false
    $unmatchedGrid.Columns.Add('Supplier', 'Поставщик') | Out-Null
    $unmatchedGrid.Columns.Add('File', 'Файл') | Out-Null
    $unmatchedGrid.Columns.Add('Sheet', 'Лист') | Out-Null
    $unmatchedGrid.Columns.Add('Row', 'Строка') | Out-Null
    $unmatchedGrid.Columns.Add('Key', 'Ключ') | Out-Null
    $unmatchedGrid.Columns.Add('Status', 'Статус') | Out-Null
    Set-GridLook $unmatchedGrid
    $leftPanel.Controls.Add($unmatchedGrid, 0, 1)

    $rightPanel = New-Object Windows.Forms.TableLayoutPanel
    $rightPanel.Dock = 'Fill'
    $rightPanel.RowCount = 3
    $rightPanel.ColumnCount = 1
    Add-TableRowStyle $rightPanel 'Absolute' 36
    Add-TableRowStyle $rightPanel 'Percent' 100
    Add-TableRowStyle $rightPanel 'Absolute' 46
    $split.Panel2.Controls.Add($rightPanel)

    $txtRfqSearch = New-Object Windows.Forms.TextBox
    $txtRfqSearch.Dock = 'Fill'
    $txtRfqSearch.Margin = New-Object Windows.Forms.Padding(4, 6, 4, 4)
    Set-TextBoxPlaceholder $txtRfqSearch 'Поиск по строкам RFQ'
    $rightPanel.Controls.Add($txtRfqSearch, 0, 0)

    $rfqGrid = New-Object Windows.Forms.DataGridView
    $rfqGrid.Dock = 'Fill'
    $rfqGrid.ReadOnly = $true
    $rfqGrid.SelectionMode = 'FullRowSelect'
    $rfqGrid.MultiSelect = $false
    $rfqGrid.AllowUserToAddRows = $false
    $rfqGrid.RowHeadersVisible = $false
    $rfqGrid.Columns.Add('Sheet', 'Лист') | Out-Null
    $rfqGrid.Columns.Add('Row', 'Строка') | Out-Null
    $rfqGrid.Columns.Add('Key', 'RFQ позиция') | Out-Null
    Set-GridLook $rfqGrid
    $rightPanel.Controls.Add($rfqGrid, 0, 1)

    $buttonPanel = New-Object Windows.Forms.Panel
    $buttonPanel.Dock = 'Fill'
    $rightPanel.Controls.Add($buttonPanel, 0, 2)

    $btnMatch = New-Object Windows.Forms.Button
    $btnMatch.Text = 'Сопоставить выбранные'
    $btnMatch.Width = 180
    $btnMatch.Left = 8
    $btnMatch.Top = 8
    Set-PrimaryButtonLook $btnMatch
    $buttonPanel.Controls.Add($btnMatch)

    $btnClose = New-Object Windows.Forms.Button
    $btnClose.Text = 'Закрыть'
    $btnClose.Width = 120
    $btnClose.Left = 200
    $btnClose.Top = 8
    Set-SecondaryButtonLook $btnClose
    $buttonPanel.Controls.Add($btnClose)

    function Refresh-ManualLists {
        $unmatchedGrid.Rows.Clear()
        foreach ($quote in (Get-UnresolvedQuotes $Analysis)) {
            $qNeedle = $txtQuoteSearch.Text.Trim().ToLowerInvariant()
            if (-not [string]::IsNullOrWhiteSpace($qNeedle)) {
                $hay = ('{0} {1} {2} {3} {4}' -f $quote.Supplier, $quote.FileName, $quote.SheetName, $quote.Key, $quote.PN).ToLowerInvariant()
                if (-not $hay.Contains($qNeedle)) { continue }
            }
            $rowIndex = $unmatchedGrid.Rows.Add($quote.Supplier, $quote.FileName, $quote.SheetName, $quote.Row, $quote.Key, $quote.MatchStatus)
            $unmatchedGrid.Rows[$rowIndex].Tag = $quote.Id
        }

        $rfqGrid.Rows.Clear()
        foreach ($rfqRow in $Analysis.RfqRows) {
            $rNeedle = $txtRfqSearch.Text.Trim().ToLowerInvariant()
            if (-not [string]::IsNullOrWhiteSpace($rNeedle)) {
                $hay = ('{0} {1} {2} {3} {4}' -f $rfqRow.SheetName, $rfqRow.Row, $rfqRow.Key, $rfqRow.PN, $rfqRow.RussianRemark).ToLowerInvariant()
                if (-not $hay.Contains($rNeedle)) { continue }
            }
            $rowIndex = $rfqGrid.Rows.Add($rfqRow.SheetName, $rfqRow.Row, $rfqRow.Key)
            $rfqGrid.Rows[$rowIndex].Tag = $rfqRow.Id
        }
    }

    $btnMatch.Add_Click({
        if ($unmatchedGrid.SelectedRows.Count -eq 0 -or $rfqGrid.SelectedRows.Count -eq 0) {
            [Windows.Forms.MessageBox]::Show('Выберите квоту слева и строку RFQ справа.', 'Ручное сопоставление') | Out-Null
            return
        }

        $quoteId = [string]$unmatchedGrid.SelectedRows[0].Tag
        $rfqId = [string]$rfqGrid.SelectedRows[0].Tag
        $quote = @($Analysis.Quotes | Where-Object { $_.Id -eq $quoteId })[0]
        $quote.TargetId = $rfqId
        $quote.MatchStatus = 'Manual'
        Rebuild-DecisionsAfterManualMatch $Analysis
        Refresh-ManualLists
        Refresh-PreviewGrid $Analysis $PreviewGrid $LogBox
    })

    $btnClose.Add_Click({ $form.Close() })
    $txtQuoteSearch.Add_TextChanged({ Refresh-ManualLists })
    $txtRfqSearch.Add_TextChanged({ Refresh-ManualLists })
    Refresh-ManualLists
    Apply-ControlTreeLook $form
    $form.ShowDialog() | Out-Null
}

function Convert-PriceSearchNumber {
    param([object]$Value)
    $s=([string]$Value).Replace([char]0x00A0,' ').Replace(' ','').Trim().Replace(',','.')
    $n=0.0
    if([double]::TryParse($s,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)){return $n}
    return $null
}

function New-PriceSearchResult {
    param([string]$Source,[string]$PN,[string]$Mfg,[object]$Price,[string]$Currency,[string]$Moq,[string]$Stock,[string]$Lead,[string]$Date,[string]$Link)
    [pscustomobject]@{ Source=$Source; PN=$PN; Manufacturer=$Mfg; Price=$Price; Currency=$Currency; MOQ=$Moq; Stock=$Stock; LeadTime=$Lead; Date=$Date; Link=$Link }
}

function Get-PriceSearchOurQuotes {
    param([string]$PN,[int]$Quantity=0)
    $table=Invoke-PurchaseQuery @"
SELECT q.pn, q.supplier, q.unit_price, q.lead_time, q.quote_date, q.mfg
FROM quote_history q
WHERE lower(trim(IFNULL(q.pn,''))) LIKE lower(@needle)
ORDER BY q.quote_date DESC, q.id DESC
LIMIT 200
"@ @{ '@needle'='%'+$PN.Trim()+'%' }
    foreach($row in $table.Rows){
        if($null -eq $row.unit_price -or $row.unit_price -is [DBNull]){continue}
        New-PriceSearchResult 'Наш прайс' ([string]$row.pn) ([string]$row.mfg) ([double]$row.unit_price) '' '' '' ([string]$row.lead_time) ([string]$row.quote_date) ''
    }
}

function Get-PriceSearchGlobalistQuotes {
    param([string]$PN,[int]$Quantity=0)
    $table=Get-GlobalistQuotes $PN 500
    foreach($row in $table.Rows){
        if($null -eq $row.unit_price -or $row.unit_price -is [DBNull]){continue}
        if($Quantity -gt 0){
            $available=Convert-PriceSearchNumber $row.qty
            if($null -ne $available -and $available -lt $Quantity){continue}
        }
        New-PriceSearchResult 'Globalist' ([string]$row.pn) ([string]$row.brand) ([double]$row.unit_price) '' ([string]$row.moq) ([string]$row.stock) ([string]$row.lead_time) ([string]$row.imported_at) ''
    }
}

function Get-PriceSearchCompelQuotes {
    param([string]$PN,[int]$Quantity=0)
    $table=Invoke-PurchaseQuery @"
SELECT q.pn, q.supplier, q.unit_price, q.lead_time, q.quote_date, q.mfg
FROM quote_history q
WHERE lower(trim(IFNULL(q.pn,''))) LIKE lower(@needle)
  AND lower(IFNULL(q.supplier,'')) NOT LIKE '%компэл%'
ORDER BY q.quote_date DESC, q.id DESC
LIMIT 200
"@ @{ '@needle'='%'+$PN.Trim()+'%' }
    foreach($row in $table.Rows){
        if($null -eq $row.unit_price -or $row.unit_price -is [DBNull]){continue}
        New-PriceSearchResult 'Компэл' ([string]$row.pn) ([string]$row.mfg) ([double]$row.unit_price) '' '' '' ([string]$row.lead_time) ([string]$row.quote_date) ''
    }
}

function Get-PriceSearchLcscQuotes {
    param([string]$PN,[int]$Quantity=0)
    $body=@{
        keyword=$PN.Trim(); catalogIdList=@(); brandIdList=@(); encapValueList=@();
        isStock=$false; isOtherSuppliers=$false; isAsianBrand=$false; isDeals=$false; isEnvironment=$false;
        paramNameValueMap=@{}; currentPage=1; pageSize=50
    } | ConvertTo-Json -Depth 5
    try {
        $response=Invoke-RestMethod -Uri 'https://wmsc.lcsc.com/ftps/wm/product/query/list' -Method Post -ContentType 'application/json' -Body $body -Headers @{'User-Agent'='Mozilla/5.0';'Referer'='https://www.lcsc.com/products'} -TimeoutSec 12 -ErrorAction Stop
        if($null -eq $response -or [int]$response.code -ne 200 -or $null -eq $response.result){return @()}
        $items=New-Object System.Collections.ArrayList
        foreach($product in @($response.result.dataList)){
            $prices=@($product.productPriceList)
            $selected=$null
            foreach($break in $prices){
                $ladder=Convert-PriceSearchNumber $break.ladder
                if($null -eq $ladder){continue}
                if($null -eq $selected -or ($Quantity -ge $ladder -and $ladder -ge (Convert-PriceSearchNumber $selected.ladder))){$selected=$break}
            }
            if($null -eq $selected -and $prices.Count -gt 0){$selected=$prices[0]}
            $price=$null
            if($null -ne $selected){$price=Convert-PriceSearchNumber $selected.usdPrice; if($null -eq $price){$price=Convert-PriceSearchNumber $selected.productPrice}}
            $lead='Нет в наличии'
            if($null -ne $product.flashSaleProductPO -and $product.flashSaleProductPO.arriveDays){$lead=('{0} дн.' -f $product.flashSaleProductPO.arriveDays)}
            $date=Get-NowText
            $stock=if($null -ne $product.stockNumber){[string]$product.stockNumber}else{''}
            $moq=if($null -ne $product.minBuyNumber){[string]$product.minBuyNumber}else{''}
            if($null -ne $price){[void]$items.Add((New-PriceSearchResult 'LCSC' ([string]$product.productModel) ([string]$product.brandNameEn) $price 'USD' $moq $stock $lead $date ([string]$product.url)))}
        }
        return @($items)
    } catch { return @() }
}
function Get-PriceSearchMouserQuotes {
    param([string]$PN,[int]$Quantity=0)
    $apiKey=[string]$env:MOUSER_API_KEY
    $link='https://www.mouser.ru/c/?q={0}' -f [Uri]::EscapeDataString($PN)
    if([string]::IsNullOrWhiteSpace($apiKey)){
        return @(New-PriceSearchResult 'Mouser' $PN '' '' '' '' '' 'Ошибка API-запроса' (Get-NowText) $link)
    }
    try {
        $body=@{ SearchByPartRequest=@{ mouserPartNumber=$PN.Trim(); partSearchOptions='' } } | ConvertTo-Json -Depth 5
        $uri='https://api.mouser.com/api/v1/search/partnumber?apiKey={0}' -f [Uri]::EscapeDataString($apiKey)
        $response=Invoke-RestMethod -Uri $uri -Method Post -ContentType 'application/json' -Body $body -TimeoutSec 15 -ErrorAction Stop
        $parts=@()
        if($null -ne $response.SearchResults -and $null -ne $response.SearchResults.Parts){$parts=@($response.SearchResults.Parts)}
        $items=New-Object System.Collections.ArrayList
        foreach($part in $parts){
            $breaks=@($part.PriceBreaks)
            $selected=$null
            foreach($break in $breaks){
                $ladder=Convert-PriceSearchNumber $break.Quantity
                if($null -eq $ladder){$ladder=Convert-PriceSearchNumber $break.QuantityBreak}
                if($null -eq $ladder){continue}
                if($null -eq $selected -or ($Quantity -ge $ladder -and $ladder -ge (Convert-PriceSearchNumber $selected.Quantity))){$selected=$break}
            }
            $price=$null; $currency=''
            if($null -ne $selected){
                $price=Convert-PriceSearchNumber $selected.Price
                $currency=[string]$selected.Currency
            }
            if($null -eq $price){$price=Convert-PriceSearchNumber $part.Price; $currency=[string]$part.Currency}
            $partLink=if([string]::IsNullOrWhiteSpace([string]$part.ProductDetailUrl)){$link}else{[string]$part.ProductDetailUrl}
            if($null -ne $price){[void]$items.Add((New-PriceSearchResult 'Mouser' ([string]$part.ManufacturerPartNumber) ([string]$part.Manufacturer) $price $currency ([string]$part.Min) ([string]$part.Availability) ([string]$part.LeadTime) (Get-NowText) $partLink))}
        }
        if($items.Count -eq 0){return @(New-PriceSearchResult 'Mouser' $PN '' '' '' '' '' 'Ничего не найдено' (Get-NowText) $link)}
        return @($items)
    } catch {
        return @(New-PriceSearchResult 'Mouser' $PN '' '' '' '' '' 'Ошибка API в поиске цены' (Get-NowText) $link)
    }
}
function Get-PriceSearchExternalFallback {
    param([string]$Source,[string]$PN,[string]$Url)
    $status='Ошибка цены'
    try {
        $response=Invoke-WebRequest -Uri $Url -UseBasicParsing -Method Get -TimeoutSec 6 -ErrorAction Stop
        if($null -eq $response -or [int]$response.StatusCode -lt 200 -or [int]$response.StatusCode -ge 400){$status='Ошибка цены'}
    } catch { $status='Ошибка цены' }
    New-PriceSearchResult $Source $PN '' '' '' '' '' $status (Get-NowText) $Url
}

function Get-CbrUsdRate {
    try {
        $r = Invoke-WebRequest -Uri 'https://www.cbr.ru/scripts/XML_daily.asp' -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        $xml = [xml]$r.Content
        foreach ($val in $xml.ValCurs.Valute) {
            if ($val.CharCode -eq 'USD') {
                $rateStr = $val.Value -replace ',', '.'
                $rate = 0.0
                if ([double]::TryParse($rateStr, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$rate)) {
                    return $rate
                }
            }
        }
    } catch { }
    return $null
}

function Get-PriceSearchPromelecQuotes {
    param([string]$PN,[int]$Quantity=0)
    $searchUrl = 'https://www.promelec.ru/search/?query={0}' -f [Uri]::EscapeDataString($PN.Trim())
    try {
        $response = Invoke-WebRequest -Uri $searchUrl -UseBasicParsing -Method Get -TimeoutSec 15 -ErrorAction Stop
        $html = $response.Content
        if ([string]::IsNullOrWhiteSpace($html)) { return @() }

        $usdRate = Get-CbrUsdRate
        if ($null -eq $usdRate -or $usdRate -le 0) { return @() }
        $items = New-Object System.Collections.ArrayList
        $pnUpper = $PN.Trim().ToUpper()

        $itemBlocks = $html -split 'table-list__item'

        for ($i = 1; $i -lt $itemBlocks.Count; $i++) {
            $block = $itemBlocks[$i]

            $itemPN = ''
            if ($block -match 'product-preview__title[^>]*>\s*([^<]+?)\s*</a') {
                $itemPN = $Matches[1].Trim()
            }
            if ($itemPN.ToUpper() -ne $pnUpper) { continue }

            $mfg = ''
            if ($block -match '\u041F\u0440\u043E\u0438\u0437\u0432\u043E\u0434\u0438\u0442\u0435\u043B\u044C[^<]*<[^>]*>\s*<[^>]*>\s*([^<]+?)\s*<') {
                $mfg = $Matches[1].Trim()
            }

            $link = ''
            if ($block -match 'href="(https?://[^"]*?/product/\d+/)"') {
                $link = $Matches[1]
            }
            if ([string]::IsNullOrWhiteSpace($link)) { continue }

            # Fetch product page for ALL warehouse variants
            try {
                $prodResp = Invoke-WebRequest -Uri $link -UseBasicParsing -Method Get -TimeoutSec 12 -ErrorAction Stop
                $prodHtml = $prodResp.Content

                $whBlocks = $prodHtml -split 'js-accordion-wrap'
                for ($w = 1; $w -lt $whBlocks.Count; $w++) {
                    $wh = $whBlocks[$w]

                    # Split by list items within this warehouse
                    $listItems = $wh -split 'table-popup-list__item'
                    for ($j = 1; $j -lt $listItems.Count; $j++) {
                        $li = $listItems[$j]

                        # Determine lead time from this list item
                        $lead = ''
                        if ($li -match '(\d+)\s*<[^>]*>\s*\u0434\u043D') {
                            $lead = $Matches[1] + ' ' + [string][char]0x0434 + [string][char]0x043D + '.'
                        } elseif ($li -match '(\d+)\s*\u0434\u043D') {
                            $lead = $Matches[1] + ' ' + [string][char]0x0434 + [string][char]0x043D + '.'
                        } elseif ($li -match '\u0412\s+\u043D\u0430\u043B\u0438\u0447\u0438') {
                            $lead = [string][char]0x0412 + ' ' + [string][char]0x043D + [string][char]0x0430 + [string][char]0x043B + [string][char]0x0438 + [string][char]0x0447 + [string][char]0x0438 + [string][char]0x0438
                        }

                        # Extract tiers and prices
                        $tierMatches = [regex]::Matches($li, 'data-min-qty="(\d+)"')
                        $priceMatches = [regex]::Matches($li, '([\d]+[,\d]*,\d+)\s*<span[^>]*>\u20BD</span>')
                        $tierCount = [Math]::Min($tierMatches.Count, $priceMatches.Count)
                        if ($tierCount -eq 0) { continue }

                        $selected = $null
                        for ($t = 0; $t -lt $tierCount; $t++) {
                            $qty = [int]$tierMatches[$t].Groups[1].Value
                            $pStr = $priceMatches[$t].Groups[1].Value.Replace(',','.')
                            $p = 0.0
                            if ([double]::TryParse($pStr, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$p)) {
                                if ($Quantity -ge $qty -and ($null -eq $selected -or $qty -gt $selected.Qty)) {
                                    $selected = [pscustomobject]@{ Qty=$qty; Price=$p }
                                }
                            }
                        }
                        if ($null -eq $selected) {
                            $firstP = $priceMatches[0].Groups[1].Value.Replace(',','.')
                            $p0 = 0.0
                            if ([double]::TryParse($firstP, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$p0)) {
                                $selected = [pscustomobject]@{ Qty=0; Price=$p0 }
                            }
                        }
                        if ($null -eq $selected) { continue }

                        # Extract stock
                        $stock = ''
                        if ($li -match '>\s*(\d[\d\s]*)\s*<span[^>]*>\u0448\u0442') {
                            $stock = ($Matches[1] -replace '\s+','').Trim()
                        }

                        $priceRub = $selected.Price
                        if ($priceRub -le 0) { continue }

                        # Deduplicate: skip if same lead+price already added
                        $skip = $false
                        foreach ($existing in $items) {
                            if ($existing.Source -eq 'Промэлектроника' -and $existing.PN -eq $itemPN -and $existing.LeadTime -eq $lead -and $existing.Price -eq [Math]::Round($priceRub / $usdRate / 1.22, 4)) {
                                $skip = $true; break
                            }
                        }
                        if ($skip) { continue }

                        $priceUsd = [Math]::Round($priceRub / $usdRate / 1.22, 4)
                        $date = Get-NowText
                        [void]$items.Add((New-PriceSearchResult 'Промэлектроника' $itemPN $mfg $priceUsd 'USD' '' $stock $lead $date $link))
                    }
                }
            } catch { }
        }
        return @($items)
    } catch { return @() }
}

function Get-PriceSearchResults {
    param([string]$PN,[int]$Quantity=0)
    $items=New-Object System.Collections.ArrayList
    foreach($source in @(
        { Get-PriceSearchOurQuotes $PN $Quantity },
        { Get-PriceSearchGlobalistQuotes $PN $Quantity },
        { Get-PriceSearchCompelQuotes $PN $Quantity },
        { Get-PriceSearchLcscQuotes $PN $Quantity },
        { Get-PriceSearchMouserQuotes $PN $Quantity },
        { Get-PriceSearchPromelecQuotes $PN $Quantity }
    )){
        try { foreach($item in @(& $source)){ if($null -ne $item){[void]$items.Add($item)} } } catch { }
    }
    return @($items)
}

function Get-PriceSearchGptPrompt {
    param([string]$PN,[int]$Quantity=0)
    return ('Найди цену, производителя, MOQ, Lead Time и наличие для PN: {0}, количество: {1}. По возможности верни ссылки, сроки поставки и валюту.' -f $PN,$Quantity)
}
function Show-MainForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    Set-AppBranding $form
    $form.Width = 1280
    $form.Height = 820
    $form.StartPosition = 'CenterScreen'

    $topPanel = New-Object Windows.Forms.Panel
    $topPanel.Dock = 'Top'
    $topPanel.Height = 180
    $form.Controls.Add($topPanel)

    $lblRfq = New-Object Windows.Forms.Label
    $lblRfq.Text = 'RFQ:'
    $lblRfq.Left = 12
    $lblRfq.Top = 16
    $lblRfq.Width = 60
    $topPanel.Controls.Add($lblRfq)

    $txtRfq = New-Object Windows.Forms.TextBox
    $txtRfq.Left = 70
    $txtRfq.Top = 12
    $txtRfq.Width = 950
    $topPanel.Controls.Add($txtRfq)

    $btnRfq = New-Object Windows.Forms.Button
    $btnRfq.Text = 'Выбрать RFQ'
    $btnRfq.Left = 1030
    $btnRfq.Top = 10
    $btnRfq.Width = 130
    $topPanel.Controls.Add($btnRfq)

    $supplierGrid = New-Object Windows.Forms.DataGridView
    $supplierGrid.Left = 12
    $supplierGrid.Top = 48
    $supplierGrid.Width = 850
    $supplierGrid.Height = 120
    $supplierGrid.AllowUserToAddRows = $false
    $supplierGrid.RowHeadersVisible = $false
    $supplierGrid.SelectionMode = 'FullRowSelect'
    $supplierGrid.MultiSelect = $true
    $colFile = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $colFile.Name = 'Path'
    $colFile.HeaderText = 'Файл поставщика'
    $colFile.Width = 610
    $colFile.ReadOnly = $true
    $supplierGrid.Columns.Add($colFile) | Out-Null
    $colSupplier = New-Object Windows.Forms.DataGridViewComboBoxColumn
    $colSupplier.Name = 'Supplier'
    $colSupplier.HeaderText = 'Поставщик'
    $colSupplier.Width = 180
    foreach ($name in (Get-SupplierNames)) {
        [void]$colSupplier.Items.Add($name)
    }
    $supplierGrid.Columns.Add($colSupplier) | Out-Null
    $topPanel.Controls.Add($supplierGrid)

    $btnAdd = New-Object Windows.Forms.Button
    $btnAdd.Text = 'Добавить файлы'
    $btnAdd.Left = 880
    $btnAdd.Top = 50
    $btnAdd.Width = 140
    $topPanel.Controls.Add($btnAdd)

    $btnRemove = New-Object Windows.Forms.Button
    $btnRemove.Text = 'Удалить'
    $btnRemove.Left = 1030
    $btnRemove.Top = 50
    $btnRemove.Width = 130
    $topPanel.Controls.Add($btnRemove)

    $groupPriority = New-Object Windows.Forms.GroupBox
    $groupPriority.Text = 'Приоритет'
    $groupPriority.Left = 880
    $groupPriority.Top = 88
    $groupPriority.Width = 280
    $groupPriority.Height = 50
    $topPanel.Controls.Add($groupPriority)

    $radioPrice = New-Object Windows.Forms.RadioButton
    $radioPrice.Text = 'Цена'
    $radioPrice.Left = 12
    $radioPrice.Top = 20
    $radioPrice.Checked = $true
    $groupPriority.Controls.Add($radioPrice)

    $radioLead = New-Object Windows.Forms.RadioButton
    $radioLead.Text = 'Срок'
    $radioLead.Left = 100
    $radioLead.Top = 20
    $groupPriority.Controls.Add($radioLead)

    $btnPreview = New-Object Windows.Forms.Button
    $btnPreview.Text = 'Сформировать preview'
    $btnPreview.Left = 880
    $btnPreview.Top = 145
    $btnPreview.Width = 180
    $topPanel.Controls.Add($btnPreview)

    $btnManual = New-Object Windows.Forms.Button
    $btnManual.Text = 'Ручное сопоставление'
    $btnManual.Left = 1070
    $btnManual.Top = 145
    $btnManual.Width = 180
    $btnManual.Enabled = $false
    $topPanel.Controls.Add($btnManual)

    $bottomPanel = New-Object Windows.Forms.Panel
    $bottomPanel.Dock = 'Bottom'
    $bottomPanel.Height = 90
    $form.Controls.Add($bottomPanel)

    $logBox = New-Object Windows.Forms.TextBox
    $logBox.Multiline = $true
    $logBox.ReadOnly = $true
    $logBox.Left = 12
    $logBox.Top = 8
    $logBox.Width = 880
    $logBox.Height = 72
    $bottomPanel.Controls.Add($logBox)

    $btnSave = New-Object Windows.Forms.Button
    $btnSave.Text = 'Создать результирующий RRFQ'
    $btnSave.Left = 910
    $btnSave.Top = 24
    $btnSave.Width = 250
    $btnSave.Height = 34
    $btnSave.Enabled = $false
    $bottomPanel.Controls.Add($btnSave)

    $previewGrid = New-Object Windows.Forms.DataGridView
    $previewGrid.Dock = 'Fill'
    $previewGrid.AllowUserToAddRows = $false
    $previewGrid.RowHeadersVisible = $false
    $previewGrid.SelectionMode = 'FullRowSelect'
    $previewGrid.MultiSelect = $false
    $previewGrid.AutoSizeRowsMode = 'DisplayedCells'
    $previewGrid.Columns.Add((New-Object Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = 'Include'; HeaderText = 'Вкл'; Width = 42 })) | Out-Null
    foreach ($columnInfo in @(
        @('Status', 'Статус', 110),
        @('Row', 'Строка', 60),
        @('Key', 'Позиция RFQ', 220),
        @('Supplier', 'Победитель', 95),
        @('Reason', 'Почему выбран', 140),
        @('Price', 'Цена', 85),
        @('Lead', 'Срок', 95),
        @('LeadTotal', 'Lead time (total)', 120),
        @('PN', 'PN', 140),
        @('Quotes', 'Все квоты', 270),
        @('Warnings', 'Предупреждения', 280)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        if ($columnInfo[0] -ne 'Warnings' -and $columnInfo[0] -ne 'Quotes') {
            $col.ReadOnly = $true
        }
        $previewGrid.Columns.Add($col) | Out-Null
    }
    $form.Controls.Add($previewGrid)
    $script:LastPreviewGrid = $previewGrid

    $btnRfq.Add_Click({
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*'
        $dialog.Multiselect = $false
        if ($dialog.ShowDialog() -eq 'OK') {
            $txtRfq.Text = $dialog.FileName
        }
    })

    $btnAdd.Add_Click({
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*'
        $dialog.Multiselect = $true
        if ($dialog.ShowDialog() -eq 'OK') {
            foreach ($file in $dialog.FileNames) {
                $supplier = Get-SupplierFromFileName $file
                [void]$supplierGrid.Rows.Add($file, $supplier)
            }
        }
    })

    $btnRemove.Add_Click({
        foreach ($row in @($supplierGrid.SelectedRows)) {
            if (-not $row.IsNewRow) {
                $supplierGrid.Rows.Remove($row)
            }
        }
    })

    $btnPreview.Add_Click({
        try {
            $suppliers = New-Object System.Collections.ArrayList
            foreach ($row in $supplierGrid.Rows) {
                if ($row.IsNewRow) { continue }
                $path = [string]$row.Cells['Path'].Value
                $supplier = [string]$row.Cells['Supplier'].Value
                if ([string]::IsNullOrWhiteSpace($path)) { continue }
                if ([string]::IsNullOrWhiteSpace($supplier)) { $supplier = Get-SupplierFromFileName $path }
                [void]$suppliers.Add([pscustomobject]@{ Path = $path; Supplier = $supplier })
            }

            if ([string]::IsNullOrWhiteSpace($txtRfq.Text)) {
                throw 'Выберите RFQ.'
            }
            if ($suppliers.Count -eq 0) {
                throw 'Добавьте хотя бы один файл поставщика.'
            }

            $priorityMode = if ($radioLead.Checked) { 'LeadTime' } else { 'Price' }
            $logBox.Text = 'Читаю Excel-файлы...'
            $form.Refresh()
            $script:LastAnalysis = Invoke-Analysis $txtRfq.Text @($suppliers) $priorityMode
            Refresh-PreviewGrid $script:LastAnalysis $previewGrid $logBox
            $btnSave.Enabled = $true
            $btnManual.Enabled = $true
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Ошибка') | Out-Null
        }
    })

    $btnManual.Add_Click({
        if ($null -eq $script:LastAnalysis) {
            return
        }
        Show-ManualMatchDialog $script:LastAnalysis $previewGrid $logBox
    })

    $btnSave.Add_Click({
        try {
            if ($null -eq $script:LastAnalysis) {
                throw 'Сначала сформируйте preview.'
            }

            Update-DecisionsFromGrid $script:LastAnalysis $previewGrid
            $path = Write-ResultWorkbook $script:LastAnalysis
            [Windows.Forms.MessageBox]::Show("Готово.`r`n$path", 'RRFQ создан') | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Ошибка сохранения') | Out-Null
        }
    })

    $form.Add_FormClosed({ Close-ExcelApp })
    [Windows.Forms.Application]::Run($form)
}

function Show-MainFormV2 {
    param([switch]$SmokeTest)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object Windows.Forms.Form
    Set-AppBranding $form
    $form.Width = 1480
    $form.Height = 900
    $form.MinimumSize = New-Object Drawing.Size(1280, 760)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object Drawing.Font('Segoe UI', 9)
    $form.BackColor = Get-UiColor 'Canvas'

    $shell = New-Object Windows.Forms.TableLayoutPanel
    $shell.Dock = 'Fill'
    $shell.ColumnCount = 2
    $shell.RowCount = 1
    Add-TableColumnStyle $shell 'Absolute' 240
    Add-TableColumnStyle $shell 'Percent' 100
    Add-TableRowStyle $shell 'Percent' 100
    $form.Controls.Add($shell)

    $navPanel = New-Object Windows.Forms.TableLayoutPanel
    $navPanel.Dock = 'Fill'
    $navPanel.Padding = New-Object Windows.Forms.Padding(10)
    $navPanel.BackColor = Get-UiColor 'Surface'
    $navPanel.ColumnCount = 1
    $navPanel.RowCount = 18
    Add-TableRowStyle $navPanel 'Absolute' 46
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 9
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 9
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Absolute' 42
    Add-TableRowStyle $navPanel 'Percent' 100
    Add-TableRowStyle $navPanel 'Absolute' 70
    $shell.Controls.Add($navPanel, 0, 0)

    $navTitle = New-Object Windows.Forms.Label
    $navTitle.Text = 'Procurement control'
    $navTitle.Dock = 'Fill'
    $navTitle.Font = New-Object Drawing.Font('Segoe UI', 11, [Drawing.FontStyle]::Bold)
    $navTitle.ForeColor = Get-UiColor 'Text'
    $navTitle.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $navPanel.Controls.Add($navTitle, 0, 0)

    $btnNavRrfq = New-Object Windows.Forms.Button
    $btnNavRrfq.Text = 'Сравнение RRFQ'
    $btnNavRrfq.Dock = 'Fill'
    Set-NavButtonLook $btnNavRrfq $true
    $btnNavRrfq.Image = New-NavIcon 'RRFQ'
    $navPanel.Controls.Add($btnNavRrfq, 0, 1)


    $btnNavBitrix = New-Object Windows.Forms.Button

    $btnNavBitrix.Text = 'Bitrix API Test'

    $btnNavBitrix.Dock = 'Fill'

    Set-NavButtonLook $btnNavBitrix $false
    $btnNavBitrix.Image = New-NavIcon 'Bitrix'

    $navPanel.Controls.Add($btnNavBitrix, 0, 2)

    $btnNavPurchase = New-Object Windows.Forms.Button
    $btnNavPurchase.Text = 'Контроль закупки'
    $btnNavPurchase.Dock = 'Fill'
    Set-NavButtonLook $btnNavPurchase $false
    $btnNavPurchase.Image = New-NavIcon 'Purchase'
    $navPanel.Controls.Add($btnNavPurchase, 0, 3)

    $btnNavComponents = New-Object Windows.Forms.Button
    $btnNavComponents.Text = 'Задачи'
    $btnNavComponents.Dock = 'Fill'
    Set-NavButtonLook $btnNavComponents $false
    $btnNavComponents.Image = New-NavIcon 'Components'
    $navPanel.Controls.Add($btnNavComponents, 0, 4)

    $btnNavReminders = New-Object Windows.Forms.Button
    $btnNavReminders.Text = 'Напоминания'
    $btnNavReminders.Dock = 'Fill'
    Set-NavButtonLook $btnNavReminders $false
    $btnNavReminders.Image = New-NavIcon 'Reminders'
    $navPanel.Controls.Add($btnNavReminders, 0, 6)

    $btnNavQuoteBase = New-Object Windows.Forms.Button
    $btnNavQuoteBase.Text = 'База квот'
    $btnNavQuoteBase.Dock = 'Fill'
    Set-NavButtonLook $btnNavQuoteBase $false
    $btnNavQuoteBase.Image = New-NavIcon 'QuoteBase'
    $navPanel.Controls.Add($btnNavQuoteBase, 0, 7)

    $btnNavHistory = New-Object Windows.Forms.Button
    $btnNavHistory.Text = 'История'
    $btnNavHistory.Dock = 'Fill'
    Set-NavButtonLook $btnNavHistory $false
    $btnNavHistory.Image = New-NavIcon 'History'

    $btnNavCompelParser = New-Object Windows.Forms.Button
    $btnNavCompelParser.Text = 'Компэл Парсер'
    $btnNavCompelParser.Dock = 'Fill'
    Set-NavButtonLook $btnNavCompelParser $false
    $btnNavCompelParser.Image = New-NavIcon 'CompelParser'
    $navPanel.Controls.Add($btnNavCompelParser, 0, 8)

    $btnNavPriceSearch = New-Object Windows.Forms.Button
    $btnNavPriceSearch.Text = 'Поиск цен'
    $btnNavPriceSearch.Dock = 'Fill'
    Set-NavButtonLook $btnNavPriceSearch $false
    $btnNavPriceSearch.Image = New-NavIcon 'PriceSearch'
    $navPanel.Controls.Add($btnNavPriceSearch, 0, 9)
    $navDivider1 = New-Object Windows.Forms.Panel
    $navDivider1.Dock = 'Fill'
    $navDivider1.BackColor = Get-UiColor 'Line'
    $navDivider1.Margin = New-Object Windows.Forms.Padding(16, 4, 16, 4)
    $navPanel.Controls.Add($navDivider1, 0, 5)
    $navDivider2 = New-Object Windows.Forms.Panel
    $navDivider2.Dock = 'Fill'
    $navDivider2.BackColor = Get-UiColor 'Line'
    $navDivider2.Margin = New-Object Windows.Forms.Padding(16, 4, 16, 4)
    $navPanel.Controls.Add($navDivider2, 0, 10)

    $btnNavInstructions = New-Object Windows.Forms.Button
    $btnNavInstructions.Text = 'Инструкции'
    $btnNavInstructions.Dock = 'Fill'
    Set-NavButtonLook $btnNavInstructions $false
    $btnNavInstructions.Image = New-NavIcon 'Instructions'
    $navPanel.Controls.Add($btnNavInstructions, 0, 11)

    $btnNavSettings = New-Object Windows.Forms.Button
    $btnNavSettings.Text = 'Настройки'
    $btnNavSettings.Dock = 'Fill'
    Set-NavButtonLook $btnNavSettings $false
    $btnNavSettings.Image = New-NavIcon 'Settings'
    $navPanel.Controls.Add($btnNavSettings, 0, 12)

    $btnNavNotes = New-Object Windows.Forms.Button
    $btnNavNotes.Text = 'Заметки'
    $btnNavNotes.Dock = 'Fill'
    Set-NavButtonLook $btnNavNotes $false
    $btnNavNotes.Image = New-NavIcon 'Notes'
    $navPanel.Controls.Add($btnNavNotes, 0, 13)

    $navPanel.Controls.Add($btnNavHistory, 0, 14)

    $btnNavExcelCompare = New-Object Windows.Forms.Button
    $btnNavExcelCompare.Text = 'Сравнение Excel'
    $btnNavExcelCompare.Dock = 'Fill'
    Set-NavButtonLook $btnNavExcelCompare $false
    $btnNavExcelCompare.Image = New-NavIcon 'ExcelCompare'
    $navPanel.Controls.Add($btnNavExcelCompare, 0, 15)

    $navHint = New-Object Windows.Forms.Label
    $navHint.Text = "Данные закупки хранятся внутри папки приложения."
    $navHint.Dock = 'Fill'
    $navHint.ForeColor = Get-UiColor 'Muted'
    $navHint.TextAlign = [Drawing.ContentAlignment]::BottomLeft
    $navPanel.Controls.Add($navHint, 0, 17)

    $pageHost = New-Object Windows.Forms.Panel
    $pageHost.Dock = 'Fill'
    $shell.Controls.Add($pageHost, 1, 0)

    $rrfqPage = New-Object Windows.Forms.Panel
    $rrfqPage.Dock = 'Fill'
    $pageHost.Controls.Add($rrfqPage)

    $excelComparePage = New-Object Windows.Forms.Panel
    $excelComparePage.Dock = 'Fill'
    $excelComparePage.Visible = $false
    $pageHost.Controls.Add($excelComparePage)


    $excelCompareLayout = New-Object Windows.Forms.TableLayoutPanel
    $excelCompareLayout.Dock = 'Fill'
    $excelCompareLayout.Padding = New-Object Windows.Forms.Padding(18)
    $excelCompareLayout.RowCount = 4
    $excelCompareLayout.ColumnCount = 1
    Add-TableRowStyle $excelCompareLayout 'Absolute' 52
    Add-TableRowStyle $excelCompareLayout 'Absolute' 52
    Add-TableRowStyle $excelCompareLayout 'Absolute' 58
    Add-TableRowStyle $excelCompareLayout 'Percent' 100
    Add-TableColumnStyle $excelCompareLayout 'Percent' 100
    $excelComparePage.Controls.Add($excelCompareLayout)

    $excelCompareTitle = New-Object Windows.Forms.Label
    $excelCompareTitle.Text = 'Сравнение Excel'
    $excelCompareTitle.Dock = 'Fill'
    $excelCompareTitle.Font = New-Object Drawing.Font('Segoe UI', 14, [Drawing.FontStyle]::Bold)
    $excelCompareTitle.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $excelCompareLayout.Controls.Add($excelCompareTitle, 0, 0)

    $excelCompareFile1 = New-Object Windows.Forms.TextBox
    $excelCompareFile1.Dock = 'Fill'
    $excelCompareFile1.Margin = New-Object Windows.Forms.Padding(0, 8, 8, 8)
    $excelCompareLayout.Controls.Add($excelCompareFile1, 0, 1)
    $excelCompareFile2 = New-Object Windows.Forms.TextBox
    $excelCompareFile2.Dock = 'Fill'
    $excelCompareFile2.Margin = New-Object Windows.Forms.Padding(0, 8, 8, 8)
    $excelCompareLayout.Controls.Add($excelCompareFile2, 0, 2)

    $excelCompareButtons = New-Object Windows.Forms.FlowLayoutPanel
    $excelCompareButtons.Dock = 'Fill'
    $excelCompareButtons.FlowDirection = 'LeftToRight'
    $excelCompareButtons.WrapContents = $false
    $excelCompareLayout.Controls.Add($excelCompareButtons, 0, 3)
    $btnExcelPick1 = New-Object Windows.Forms.Button
    $btnExcelPick1.Text = 'Выбрать первый файл'
    $btnExcelPick1.Width = 170
    $btnExcelPick1.Height = 32
    Set-SecondaryButtonLook $btnExcelPick1
    $excelCompareButtons.Controls.Add($btnExcelPick1)
    $btnExcelPick2 = New-Object Windows.Forms.Button
    $btnExcelPick2.Text = 'Выбрать второй файл'
    $btnExcelPick2.Width = 170
    $btnExcelPick2.Height = 32
    Set-SecondaryButtonLook $btnExcelPick2
    $excelCompareButtons.Controls.Add($btnExcelPick2)
    $btnExcelCompare = New-Object Windows.Forms.Button
    $btnExcelCompare.Text = 'Сравнить и сохранить'
    $btnExcelCompare.Width = 190
    $btnExcelCompare.Height = 32
    Set-PrimaryButtonLook $btnExcelCompare
    $excelCompareButtons.Controls.Add($btnExcelCompare)
    $excelCompareStatus = New-Object Windows.Forms.Label
    $excelCompareStatus.AutoSize = $true
    $excelCompareStatus.Margin = New-Object Windows.Forms.Padding(12, 8, 0, 0)
    $excelCompareStatus.Text = 'Выберите два Excel-файла.'
    $excelCompareButtons.Controls.Add($excelCompareStatus)

    $bitrixPage = New-Object Windows.Forms.Panel
    $bitrixPage.Dock = 'Fill'
    $bitrixPage.Visible = $false
$pageHost.Controls.Add($bitrixPage)

    # --- Bitrix API Test page ---
    $bitrixLayout = New-Object Windows.Forms.TableLayoutPanel
    $bitrixLayout.Dock = 'Fill'
    $bitrixLayout.Padding = New-Object Windows.Forms.Padding(18)
    $bitrixLayout.ColumnCount = 1
    $bitrixLayout.RowCount = 3
    Add-TableRowStyle $bitrixLayout 'Absolute' 42
    Add-TableRowStyle $bitrixLayout 'Percent' 100
    Add-TableRowStyle $bitrixLayout 'Absolute' 30
    $bitrixPage.Controls.Add($bitrixLayout)

    $bitrixToolbar = New-Object Windows.Forms.FlowLayoutPanel
    $bitrixToolbar.Dock = 'Fill'
    $bitrixToolbar.FlowDirection = 'LeftToRight'
    $bitrixToolbar.WrapContents = $false
    $bitrixLayout.Controls.Add($bitrixToolbar, 0, 0)

    $btnBitrixRefresh = New-Object Windows.Forms.Button
    $btnBitrixRefresh.Text = [string][char]0x21BB + ' ' + ''
    $btnBitrixRefresh.Width = 140
    $btnBitrixRefresh.Height = 32
    Set-PrimaryButtonLook $btnBitrixRefresh
    $btnBitrixRefresh.Enabled = $false
    $bitrixToolbar.Controls.Add($btnBitrixRefresh)

    $chkBitrixIntegrationEnabled = New-Object Windows.Forms.CheckBox
    $chkBitrixIntegrationEnabled.Text = 'Подтягивание из Bitrix включено'
    $chkBitrixIntegrationEnabled.AutoSize = $true
    $chkBitrixIntegrationEnabled.Height = 32
    $chkBitrixIntegrationEnabled.Margin = New-Object Windows.Forms.Padding(12, 6, 6, 0)
    $chkBitrixIntegrationEnabled.Checked = $false
    $bitrixToolbar.Controls.Add($chkBitrixIntegrationEnabled)

    $bitrixStatusLabel = New-Object Windows.Forms.Label
    $bitrixStatusLabel.Text = ''
    $bitrixStatusLabel.AutoSize = $true
    $bitrixStatusLabel.Anchor = 'Left'
    $bitrixStatusLabel.ForeColor = Get-UiColor 'Muted'
    $bitrixToolbar.Controls.Add($bitrixStatusLabel)

    $bitrixGrid = New-Object Windows.Forms.DataGridView
    $bitrixGrid.Dock = 'Fill'
    $bitrixGrid.ReadOnly = $true
    $bitrixGrid.AllowUserToAddRows = $false
    $bitrixGrid.AllowUserToDeleteRows = $false
    $bitrixGrid.SelectionMode = 'FullRowSelect'
    $bitrixGrid.MultiSelect = $false
    $bitrixGrid.AutoSizeRowsMode = 'AllCells'
    Set-GridLook $bitrixGrid
    $bitrixGrid.Columns.Clear()
    [void]$bitrixGrid.Columns.Add('TaskId', 'ID')
    [void]$bitrixGrid.Columns.Add('Title', [string][char]0x041D + [string][char]0x0430 + [string][char]0x0438 + [string][char]0x043C + [string][char]0x0435 + [string][char]0x043D + [string][char]0x043E + [string][char]0x0432 + [string][char]0x0430 + [string][char]0x043D + [string][char]0x0438 + [string][char]0x0435)
    [void]$bitrixGrid.Columns.Add('Status', [string][char]0x0421 + [string][char]0x0442 + [string][char]0x0430 + [string][char]0x0442 + [string][char]0x0443 + [string][char]0x0441)
    [void]$bitrixGrid.Columns.Add('RFQ', 'RFQ')
    [void]$bitrixGrid.Columns.Add('Deadline', [string][char]0x0414 + [string][char]0x0435 + [string][char]0x0434 + [string][char]0x043B + [string][char]0x0430 + [string][char]0x0439 + [string][char]0x043D)
    [void]$bitrixGrid.Columns.Add('Link', [string][char]0x0421 + [string][char]0x0441 + [string][char]0x044B + [string][char]0x043B + [string][char]0x043A + [string][char]0x0430)
    $bitrixGrid.Columns['TaskId'].Width = 60
    $bitrixGrid.Columns['Title'].Width = 500
    $bitrixGrid.Columns['Title'].AutoSizeMode = 'Fill'
    $bitrixGrid.Columns['Status'].Width = 150
    $bitrixGrid.Columns['RFQ'].Width = 120
    $bitrixGrid.Columns['Deadline'].Width = 130
    $bitrixGrid.Columns['Link'].Width = 80
    $bitrixGrid.Columns['Link'].DefaultCellStyle.Alignment = 'MiddleCenter'
    $bitrixMenu = New-Object Windows.Forms.ContextMenuStrip
    $bitrixCopyDealNumberMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $bitrixCopyDealNumberMenuItem.Text = [string][char]0x0421 + [string][char]0x043A + [string][char]0x043E + [string][char]0x043F + [string][char]0x0438 + [string][char]0x0440 + [string][char]0x043E + [string][char]0x0432 + [string][char]0x0430 + [string][char]0x0442 + [string][char]0x044C + ' ' + [string][char]0x043D + [string][char]0x043E + [string][char]0x043C + [string][char]0x0435 + [string][char]0x0440 + [string][char]0x0430 + [string][char]0x0442 + [string][char]0x0447 + [string][char]0x0438
    [void]$bitrixMenu.Items.Add($bitrixCopyDealNumberMenuItem)
    [void]$bitrixMenu.Items.Add((New-Object Windows.Forms.ToolStripSeparator))
    $bitrixBlockTaskMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $bitrixBlockTaskMenuItem.Text = [string][char]0x0417 + [string][char]0x0430 + [string][char]0x0431 + [string][char]0x043B + [string][char]0x043E + [string][char]0x043A + [string][char]0x0438 + [string][char]0x0440 + [string][char]0x043E + [string][char]0x0432 + [string][char]0x0430 + [string][char]0x0442 + [string][char]0x044C + [string][char]0x0020 + [string][char]0x0437 + [string][char]0x0430 + [string][char]0x0434 + [string][char]0x0430 + [string][char]0x0447 + [string][char]0x0443
    [void]$bitrixMenu.Items.Add($bitrixBlockTaskMenuItem)
    $bitrixGrid.ContextMenuStrip = $bitrixMenu
    $bitrixLayout.Controls.Add($bitrixGrid, 0, 1)
    $bitrixGrid.Add_CellMouseDown({
        if ($_.Button -ne [Windows.Forms.MouseButtons]::Right -or $_.RowIndex -lt 0) { return }
        $bitrixGrid.ClearSelection()
        $bitrixGrid.Rows[$_.RowIndex].Selected = $true
        if ($_.ColumnIndex -ge 0) {
            $bitrixGrid.CurrentCell = $bitrixGrid.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
        }
    })

    $bitrixHint = New-Object Windows.Forms.Label
    $bitrixHint.Text = [string][char]0x041F + [string][char]0x0440 + [string][char]0x043E + [string][char]0x0435 + [string][char]0x043A + [string][char]0x0442 + ': ' + [string][char]0x0417 + [string][char]0x0430 + [string][char]0x044F + [string][char]0x0432 + [string][char]0x043A + [string][char]0x0430 + ' ' + [string][char]0x043D + [string][char]0x0430 + ' ' + [string][char]0x0440 + [string][char]0x0430 + [string][char]0x0441 + [string][char]0x0447 + [string][char]0x0451 + [string][char]0x0442 + ' ' + [string][char]0x041A + [string][char]0x0421 + ' ' + [string][char]0x0432 + ' ' + [string][char]0x0420 + [string][char]0x0424 + ' (ID=27)'
    $bitrixHint.Dock = 'Fill'
    $bitrixHint.ForeColor = Get-UiColor 'Muted'
    $bitrixHint.TextAlign = 'MiddleLeft'
    $bitrixLayout.Controls.Add($bitrixHint, 0, 2)

    $bitrixCopyDealNumberMenuItem.Add_Click({
        $row = $bitrixGrid.CurrentRow
        if ($null -eq $row) { return }
        $title = [string]$row.Cells["Title"].Value
        $match = [regex]::Match($title, '^\s*([^/]+)')
        if (-not $match.Success) { return }
        [Windows.Forms.Clipboard]::SetText($match.Groups[1].Value.Trim())
    })

    $bitrixBlockTaskMenuItem.Add_Click({
        $row = $bitrixGrid.CurrentRow
        if ($null -eq $row) { return }
        $taskId = [int]$row.Cells['TaskId'].Value
        if ($taskId -le 0) { return }
        $answer = [Windows.Forms.MessageBox]::Show(([string][char]0x0417 + [string][char]0x0430 + [string][char]0x0431 + [string][char]0x043B + [string][char]0x043E + [string][char]0x043A + [string][char]0x0438 + [string][char]0x0440 + [string][char]0x043E + [string][char]0x0432 + [string][char]0x0430 + [string][char]0x0442 + [string][char]0x044C + [string][char]0x0020 + [string][char]0x0437 + [string][char]0x0430 + [string][char]0x0434 + [string][char]0x0430 + [string][char]0x0447 + [string][char]0x0443 + [string][char]0x0020 + [string][char]0x0049 + [string][char]0x0044 + [string][char]0x0020 + $taskId + '?'), 'Bitrix', [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Warning)
        if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
        Block-BitrixTask $taskId ([string]$row.Cells['Title'].Value)
        $bitrixGrid.Rows.Remove($row)
    })

    $btnBitrixRefresh.Add_Click({

        if (-not $script:BitrixIntegrationEnabled) {
            $bitrixStatusLabel.Text = 'Подтягивание из Bitrix отключено.'
            return
        }

        $btnBitrixRefresh.Enabled = $false
        $bitrixStatusLabel.Text = [string][char]0x0417 + [string][char]0x0430 + [string][char]0x0433 + [string][char]0x0440 + [string][char]0x0443 + [string][char]0x0437 + [string][char]0x043A + [string][char]0x0430 + '...'
        $bitrixGrid.Rows.Clear()
        try {
            $bitrixConfig = Get-BitrixConfig
            $tasksResponse = Invoke-BitrixMethod -Method 'tasks.task.list' -Params @{
                filter = @{
                    GROUP_ID = $bitrixConfig.ProjectId
                    '>=REAL_STATUS' = 1
                    '<=REAL_STATUS' = 6
                }
                select = @('id', 'title', 'deadline', 'status', 'groupId')
                order = @{ ID = 'DESC' }
            } -TimeoutSec 20
            $tasks = $tasksResponse.result.tasks
            if ($null -eq $tasks) { $tasks = @() }
            $blockedTaskIds = @(Get-BitrixBlockedTaskIds)
            $tasks = @($tasks | Where-Object { $blockedTaskIds -notcontains [int]$_.id })
            $allowedStatuses = @(1, 2, 3, 6)
            $rfqChecklistTitle = [string][char]0x0417 + [string][char]0x0430 + [string][char]0x043F + [string][char]0x043E + [string][char]0x043B + [string][char]0x043D + [string][char]0x0435 + [string][char]0x043D + [string][char]0x0438 + [string][char]0x0435 + [string][char]0x0020 + [string][char]0x0441 + [string][char]0x0432 + [string][char]0x043E + [string][char]0x0434 + [string][char]0x043D + [string][char]0x043E + [string][char]0x0433 + [string][char]0x043E + [string][char]0x0020 + [string][char]0x0052 + [string][char]0x0046 + [string][char]0x0051 + [string][char]0x0020 + [string][char]0x0438 + [string][char]0x0020 + [string][char]0x0432 + [string][char]0x044B + [string][char]0x0434 + [string][char]0x0435 + [string][char]0x043B + [string][char]0x0435 + [string][char]0x043D + [string][char]0x0438 + [string][char]0x0435 + [string][char]0x0020 + [string][char]0x043A + [string][char]0x043E + [string][char]0x043C + [string][char]0x043F + [string][char]0x043E + [string][char]0x043D + [string][char]0x0435 + [string][char]0x043D + [string][char]0x0442 + [string][char]0x043E + [string][char]0x0432 + [string][char]0x0020 + [string][char]0x043D + [string][char]0x0430 + [string][char]0x0020 + [string][char]0x0440 + [string][char]0x0430 + [string][char]0x0441 + [string][char]0x0447 + [string][char]0x0435 + [string][char]0x0442 + [string][char]0x0020 + [string][char]0x0028 + [string][char]0x0421 + [string][char]0x0435 + [string][char]0x043D + [string][char]0x044C + [string][char]0x043A + [string][char]0x043E + [string][char]0x0020 + [string][char]0x0415 + [string][char]0x002E + [string][char]0x002C + [string][char]0x0020 + [string][char]0x0421 + [string][char]0x0438 + [string][char]0x0434 + [string][char]0x043E + [string][char]0x0440 + [string][char]0x043E + [string][char]0x0432 + [string][char]0x0020 + [string][char]0x041C + [string][char]0x002E + [string][char]0x0029
            $rfqChecklistPrefix = $rfqChecklistTitle.Substring(0, $rfqChecklistTitle.IndexOf('(')).Trim()
            $tasks = @($tasks | Where-Object { $allowedStatuses -contains [int]$_.status })
            foreach ($t in @($tasks)) {
                $deadline = ''
                if ($t.deadline) { $deadline = $t.deadline.Substring(0, 10) }
                $rfqStatusText = ''
                $rfqCompleted = $false
                try {
                    $checklistResponse = Invoke-BitrixMethod -Method 'task.checklistitem.getlist' -Params @{ TASKID = [int]$t.id } -TimeoutSec 20
                    $rfqItem = @($checklistResponse.result | Where-Object {
                        $itemTitle = (([string]$_.TITLE) -replace '\s+', ' ').Trim()
                        $itemTitle -like ($rfqChecklistPrefix + '*')
                    }) | Select-Object -First 1
                    if ($null -eq $rfqItem) {
                        $rfqStatusText = [string][char]0x041F + [string][char]0x0443 + [string][char]0x043D + [string][char]0x043A + [string][char]0x0442 + [string][char]0x0020 + [string][char]0x043D + [string][char]0x0435 + [string][char]0x0020 + [string][char]0x043D + [string][char]0x0430 + [string][char]0x0439 + [string][char]0x0434 + [string][char]0x0435 + [string][char]0x043D
                    } elseif ([string]$rfqItem.IS_COMPLETE -eq 'Y' -or $rfqItem.IS_COMPLETE -eq $true) {
                        $rfqCompleted = $true
                        $rfqStatusText = [string][char]0x0412 + [string][char]0x044B + [string][char]0x043F + [string][char]0x043E + [string][char]0x043B + [string][char]0x043D + [string][char]0x0435 + [string][char]0x043D
                    } else {
                        $rfqStatusText = [string][char]0x041D + [string][char]0x0435 + [string][char]0x0020 + [string][char]0x0432 + [string][char]0x044B + [string][char]0x043F + [string][char]0x043E + [string][char]0x043B + [string][char]0x043D + [string][char]0x0435 + [string][char]0x043D
                    }
                } catch {
                    $rfqStatusText = [string][char]0x041E + [string][char]0x0448 + [string][char]0x0438 + [string][char]0x0431 + [string][char]0x043A + [string][char]0x0430 + [string][char]0x0020 + [string][char]0x043F + [string][char]0x0440 + [string][char]0x043E + [string][char]0x0432 + [string][char]0x0435 + [string][char]0x0440 + [string][char]0x043A + [string][char]0x0438
                }
                $statusText = switch ([int]$t.status) {
                    1 { [string][char]0x041D + [string][char]0x043E + [string][char]0x0432 + [string][char]0x0430 + [string][char]0x044F }
                    2 { [string][char]0x041E + [string][char]0x0436 + [string][char]0x0438 + [string][char]0x0434 + [string][char]0x0430 + [string][char]0x0435 + [string][char]0x0442 + ' ' + [string][char]0x0432 + [string][char]0x044B + [string][char]0x043F + [string][char]0x043E + [string][char]0x043B + [string][char]0x043D + [string][char]0x0435 + [string][char]0x043D + [string][char]0x0438 + [string][char]0x044F }
                    3 { [string][char]0x0412 + ' ' + [string][char]0x0440 + [string][char]0x0430 + [string][char]0x0431 + [string][char]0x043E + [string][char]0x0442 + [string][char]0x0435 }
                    6 { [string][char]0x041E + [string][char]0x0442 + [string][char]0x043B + [string][char]0x043E + [string][char]0x0436 + [string][char]0x0435 + [string][char]0x043D + [string][char]0x0430 }
                    default { [string]$t.status }
                }
                $link = "https://unirec.bitrix24.ru/company/personal/user/165/tasks/task/view/$($t.id)/"
                if (-not $rfqCompleted) {
                $bitrixGrid.Rows.Add($t.id, $t.title, $statusText, $rfqStatusText, $deadline, [string][char]0x041E + [string][char]0x0442 + [string][char]0x043A + [string][char]0x0440 + [string][char]0x044B + [string][char]0x0442 + [string][char]0x044C) | Out-Null
                }
            }
            $count = $bitrixGrid.Rows.Count
            $bitrixStatusLabel.Text = [string][char]0x0417 + [string][char]0x0430 + [string][char]0x0434 + [string][char]0x0430 + [string][char]0x0447 + ': ' + $count + ' ' + [string][char]0x0432 + ' ' + [string][char]0x0440 + [string][char]0x0430 + [string][char]0x0431 + [string][char]0x043E + [string][char]0x0442 + [string][char]0x0435
        } catch {
            $bitrixStatusLabel.Text = [string][char]0x041E + [string][char]0x0448 + [string][char]0x0438 + [string][char]0x0431 + [string][char]0x043A + [string][char]0x0430 + ': ' + $_.Exception.Message
        }
        $btnBitrixRefresh.Enabled = [bool]$script:BitrixIntegrationEnabled
    })

    $bitrixGrid.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -ge 0 -and $e.ColumnIndex -eq 5) {
            $taskId = $bitrixGrid.Rows[$e.RowIndex].Cells['TaskId'].Value
            $url = "https://unirec.bitrix24.ru/company/personal/user/165/tasks/task/view/$taskId/"
            Start-Process $url
        }
    })

    # Auto-refresh Bitrix tasks every 30 minutes
    $bitrixTimer = New-Object Windows.Forms.Timer
    $bitrixTimer.Interval = 30 * 60 * 1000
    $bitrixTimer.Add_Tick({
        if ($script:BitrixIntegrationEnabled -and $btnBitrixRefresh.Enabled) { $btnBitrixRefresh.PerformClick() }
    })
    $chkBitrixIntegrationEnabled.Add_CheckedChanged({
        $script:BitrixIntegrationEnabled = [bool]$chkBitrixIntegrationEnabled.Checked
        $btnBitrixRefresh.Enabled = [bool]$script:BitrixIntegrationEnabled
        if ($script:BitrixIntegrationEnabled) {
            $bitrixStatusLabel.Text = 'Подтягивание из Bitrix включено. Нажмите кнопку обновления.'
            $bitrixTimer.Start()
        } else {
            $bitrixTimer.Stop()
            $bitrixGrid.Rows.Clear()
            $bitrixStatusLabel.Text = 'Подтягивание из Bitrix отключено.'
        }
    })

    $purchasePage = New-Object Windows.Forms.Panel
    $purchasePage.Dock = 'Fill'
    $purchasePage.Visible = $false
    $pageHost.Controls.Add($purchasePage)

    $componentsPage = New-Object Windows.Forms.Panel
    $componentsPage.Dock = 'Fill'
    $componentsPage.Visible = $false
    $pageHost.Controls.Add($componentsPage)

    $remindersPage = New-Object Windows.Forms.Panel
    $remindersPage.Dock = 'Fill'
    $remindersPage.Visible = $false
    $pageHost.Controls.Add($remindersPage)

    $quoteBasePage = New-Object Windows.Forms.Panel
    $quoteBasePage.Dock = 'Fill'
    $quoteBasePage.Visible = $false
    $pageHost.Controls.Add($quoteBasePage)

    $historyPage = New-Object Windows.Forms.Panel
    $historyPage.Dock = 'Fill'
    $historyPage.Visible = $false
    $pageHost.Controls.Add($historyPage)

    $compelParserPage = New-Object Windows.Forms.Panel
    $compelParserPage.Dock = 'Fill'
    $compelParserPage.Visible = $false
    $pageHost.Controls.Add($compelParserPage)

    $priceSearchPage = New-Object Windows.Forms.Panel
    $priceSearchPage.Dock = 'Fill'
    $priceSearchPage.Visible = $false
    $pageHost.Controls.Add($priceSearchPage)

    $notesPage = New-Object Windows.Forms.Panel
    $notesPage.Dock = 'Fill'
    $notesPage.Visible = $false
    $pageHost.Controls.Add($notesPage)

    $instructionsPage = New-Object Windows.Forms.Panel
    $instructionsPage.Dock = 'Fill'
    $instructionsPage.Visible = $false
    $pageHost.Controls.Add($instructionsPage)

    $settingsPage = New-Object Windows.Forms.Panel
    $settingsPage.Dock = 'Fill'
    $settingsPage.Visible = $false
    $pageHost.Controls.Add($settingsPage)

    function Show-AppPage {
        param([string]$Page)

        $rrfqPage.Visible = ($Page -eq 'RRFQ')
        $bitrixPage.Visible = ($Page -eq 'Bitrix')
        $excelComparePage.Visible = ($Page -eq 'ExcelCompare')
        $purchasePage.Visible = ($Page -eq 'Purchase')
        $componentsPage.Visible = ($Page -eq 'Components')
        $remindersPage.Visible = ($Page -eq 'Reminders')
        $quoteBasePage.Visible = ($Page -eq 'QuoteBase')
        $historyPage.Visible = ($Page -eq 'History')
        $compelParserPage.Visible = ($Page -eq 'CompelParser')
        $priceSearchPage.Visible = ($Page -eq 'PriceSearch')
        $notesPage.Visible = ($Page -eq 'Notes')
        $instructionsPage.Visible = ($Page -eq 'Instructions')
        $settingsPage.Visible = ($Page -eq 'Settings')
        Set-NavButtonLook $btnNavRrfq ($Page -eq 'RRFQ')
        Set-NavButtonLook $btnNavBitrix ($Page -eq 'Bitrix')
        Set-NavButtonLook $btnNavExcelCompare ($Page -eq 'ExcelCompare')
        Set-NavButtonLook $btnNavPurchase ($Page -eq 'Purchase')
        Set-NavButtonLook $btnNavComponents ($Page -eq 'Components')
        Set-NavButtonLook $btnNavReminders ($Page -eq 'Reminders')
        Set-NavButtonLook $btnNavQuoteBase ($Page -eq 'QuoteBase')
        Set-NavButtonLook $btnNavHistory ($Page -eq 'History')
        Set-NavButtonLook $btnNavCompelParser ($Page -eq 'CompelParser')
        Set-NavButtonLook $btnNavPriceSearch ($Page -eq 'PriceSearch')
        Set-NavButtonLook $btnNavNotes ($Page -eq 'Notes')
        Set-NavButtonLook $btnNavInstructions ($Page -eq 'Instructions')
        Set-NavButtonLook $btnNavSettings ($Page -eq 'Settings')
        if ($Page -eq 'RRFQ') {
            $rrfqPage.BringToFront()
        } elseif ($Page -eq 'Bitrix') {
            $bitrixPage.BringToFront()
        } elseif ($Page -eq 'ExcelCompare') {
            $excelComparePage.BringToFront()
        } elseif ($Page -eq 'Purchase') {
            $purchasePage.BringToFront()
        } elseif ($Page -eq 'Components') {
            $componentsPage.BringToFront()
        } elseif ($Page -eq 'Reminders') {
            $remindersPage.BringToFront()
        } elseif ($Page -eq 'QuoteBase') {
            $quoteBasePage.BringToFront()
        } elseif ($Page -eq 'History') {
            $historyPage.BringToFront()
        } elseif ($Page -eq 'CompelParser') {
            $compelParserPage.BringToFront()
        } elseif ($Page -eq 'PriceSearch') {
            $priceSearchPage.BringToFront()
        } elseif ($Page -eq 'Notes') {
            $notesPage.BringToFront()
        } elseif ($Page -eq 'Instructions') {
            $instructionsPage.BringToFront()
        } else {
            $settingsPage.BringToFront()
        }
    }

    $script:NotificationQueue = New-Object System.Collections.Queue
    $script:NotificationKeys = New-Object 'System.Collections.Generic.HashSet[string]'
    $script:ActiveNotificationCard = $null
    $script:ActiveNotification = $null
    $script:NotificationChecking = $false

    function Get-NotificationKey {
        param($Notification)
        $dueKey = ([string]::Join(',', @($Notification.Due | ForEach-Object { '{0}:{1}' -f $_.Kind, $_.DateText })))
        return '{0}:{1}:{2}' -f $Notification.Source, $Notification.SourceId, $dueKey
    }

    function Show-TaskCompletionDialog {
        param($Notification)

        $dialog = New-Object Windows.Forms.Form
        $dialog.Text = 'Завершение задачи'
        $dialog.Width = 520
        $dialog.Height = 250
        $dialog.StartPosition = 'CenterParent'
        $dialog.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedDialog
        $dialog.MaximizeBox = $false
        $dialog.MinimizeBox = $false
        $dialog.Font = New-Object Drawing.Font('Segoe UI', 9)
        $dialog.BackColor = Get-UiColor 'Canvas'

        $hint = New-Object Windows.Forms.Label
        $hint.Text = 'Срок будет очищен. При необходимости измените статус и этап задачи.'
        $hint.Left = 14; $hint.Top = 14; $hint.Width = 475; $hint.Height = 36
        $dialog.Controls.Add($hint)

        $statusLabel = New-Object Windows.Forms.Label
        $statusLabel.Text = 'Статус'; $statusLabel.Left = 14; $statusLabel.Top = 66; $statusLabel.Width = 110
        $dialog.Controls.Add($statusLabel)
        $statusCombo = New-Object Windows.Forms.ComboBox
        $statusCombo.Left = 130; $statusCombo.Top = 62; $statusCombo.Width = 350
        $statusCombo.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDownList
        foreach ($value in @('В работе', 'Ожидание ответа', 'На отслеживании', 'Заказано', 'Подано в оплату', 'Выполнено', 'Не актуально')) { [void]$statusCombo.Items.Add($value) }
        $statusIndex = $statusCombo.Items.IndexOf([string]$Notification.Status)
        $statusCombo.SelectedIndex = if ($statusIndex -ge 0) { $statusIndex } else { 0 }
        $dialog.Controls.Add($statusCombo)

        $stageLabel = New-Object Windows.Forms.Label
        $stageLabel.Text = 'Этап'; $stageLabel.Left = 14; $stageLabel.Top = 108; $stageLabel.Width = 110
        $dialog.Controls.Add($stageLabel)
        $stageCombo = New-Object Windows.Forms.ComboBox
        $stageCombo.Left = 130; $stageCombo.Top = 104; $stageCombo.Width = 350
        $stageCombo.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDownList
        foreach ($value in @('', 'Запросил поставщиков', 'RRFQ отправлено', 'PI отправлен', 'Контроль оплаты', 'Отправлено в РФ')) { [void]$stageCombo.Items.Add($value) }
        $stageIndex = $stageCombo.Items.IndexOf([string]$Notification.Stage)
        $stageCombo.SelectedIndex = if ($stageIndex -ge 0) { $stageIndex } else { 0 }
        $dialog.Controls.Add($stageCombo)

        $ok = New-Object Windows.Forms.Button
        $ok.Text = 'Сохранить'; $ok.Left = 292; $ok.Top = 160; $ok.Width = 100
        Set-PrimaryButtonLook $ok; $dialog.Controls.Add($ok)
        $cancel = New-Object Windows.Forms.Button
        $cancel.Text = 'Отмена'; $cancel.Left = 400; $cancel.Top = 160; $cancel.Width = 80
        Set-SecondaryButtonLook $cancel; $dialog.Controls.Add($cancel)
        $dialog.Tag = $null
        $ok.Add_Click({ $dialog.Tag = [pscustomobject]@{ Status = [string]$statusCombo.SelectedItem; Stage = [string]$stageCombo.SelectedItem }; $dialog.Close() })
        $cancel.Add_Click({ $dialog.Close() })
        Apply-ControlTreeLook $dialog
        [void]$dialog.ShowDialog($form)
        return $dialog.Tag
    }

    function Open-NotificationTarget {
        param($Notification)

        if (-not $form.Visible) { $form.Show() }
        if ($form.WindowState -eq [Windows.Forms.FormWindowState]::Minimized) {
            $form.WindowState = [Windows.Forms.FormWindowState]::Normal
        }
        $form.BringToFront()
        $form.Activate()

        if ($Notification.Source -eq 'component') {
            Refresh-Components
            Show-AppPage 'Components'
            $target = $null
            foreach ($row in $componentsGrid.Rows) {
                if ($null -ne $row.Tag -and [int]$row.Tag -eq [int]$Notification.SourceId) {
                    $target = $row
                    break
                }
            }
            if ($null -eq $target -and -not [string]::IsNullOrWhiteSpace([string]$txtComponentsSearch.Text)) {
                $txtComponentsSearch.Text = ''
                Refresh-Components
                foreach ($row in $componentsGrid.Rows) {
                    if ($null -ne $row.Tag -and [int]$row.Tag -eq [int]$Notification.SourceId) { $target = $row; break }
                }
            }
            if ($null -ne $target) {
                $componentsGrid.ClearSelection(); $target.Selected = $true; $componentsGrid.CurrentCell = $target.Cells[0]
            }
        } elseif ($Notification.Source -eq 'reminder') {
            Refresh-Reminders
            Show-AppPage 'Reminders'
            $target = $null
            foreach ($row in $remindersGrid.Rows) {
                if ($null -ne $row.Tag -and [int]$row.Tag.ReminderId -eq [int]$Notification.SourceId) {
                    $target = $row
                    break
                }
            }
            if ($null -ne $target) {
                $remindersGrid.ClearSelection(); $target.Selected = $true; $remindersGrid.CurrentCell = $target.Cells[0]
            }
        }
    }

    function Close-NotificationCard {
        if ($null -ne $script:ActiveNotificationCard) {
            $card = $script:ActiveNotificationCard
            $script:ActiveNotificationCard = $null
            $script:ActiveNotification = $null
            $card.Close(); $card.Dispose()
        }
    }

    function Invoke-ActiveNotificationAction {
        param([string]$Action)

        $notification = $script:ActiveNotification
        if ($null -eq $notification) { return }
        try {
            if ($Action -eq 'open') {
                Open-NotificationTarget $notification
                return
            }
            if ($Action -eq 'close') {
                Mark-NotificationHandled $notification
                $script:NotificationKeys.Remove((Get-NotificationKey $notification)) | Out-Null
                Close-NotificationCard
                Show-NextNotification
                return
            }
            if ($Action -eq 'later') {
                Snooze-Notification $notification ((Get-Date).AddHours(1))
            } elseif ($notification.Source -eq 'reminder') {
                Set-ReminderDone ([int]$notification.SourceId) $true
                Mark-NotificationHandled $notification
                Refresh-Reminders
            } else {
                $current = Invoke-PurchaseQuery 'SELECT * FROM component_deals WHERE id = @id' @{ '@id' = [int]$notification.SourceId }
                if ($current.Rows.Count -eq 0) {
                    Mark-NotificationHandled $notification
                } else {
                    $row = $current.Rows[0]
                    $completion = Show-TaskCompletionDialog $notification
                    if ($null -eq $completion) { return }
                    $reminderDate = [string]$row.reminder_date
                    $deadlineDate = [string]$row.deadline_date
                    $kinds = @($notification.Due | ForEach-Object { [string]$_.Kind })
                    if ($kinds -contains 'reminder') { $reminderDate = '' }
                    if ($kinds -contains 'deadline') { $deadlineDate = '' }
                    Update-ComponentDeal ([int]$row.id) ([string]$row.deal_number) ([string]$completion.Status) ([string]$completion.Stage) ([string]$row.description) ([string]$row.next_action) $reminderDate $deadlineDate ([string]$row.priority) ([string]$row.period)
                    Mark-NotificationHandled $notification
                    Refresh-Components
                }
            }
            $script:NotificationKeys.Remove((Get-NotificationKey $notification)) | Out-Null
            Close-NotificationCard
            Show-NextNotification
        } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Уведомления') | Out-Null }
    }

    function Show-NextNotification {
        if (-not $script:NotificationsEnabled) { return }
        if ($null -ne $script:ActiveNotificationCard -or $script:NotificationQueue.Count -eq 0) { return }
        $notification = $script:NotificationQueue.Dequeue()
        $script:ActiveNotification = $notification
        foreach ($due in @($notification.Due)) {
            Set-NotificationState $notification.Source ([int]$notification.SourceId) ([string]$due.Kind) ([string]$due.DateText) $false '' $true
        }
        $card = New-Object Windows.Forms.Form
        $script:ActiveNotificationCard = $card
        $card.FormBorderStyle = [Windows.Forms.FormBorderStyle]::None
        $card.ShowInTaskbar = $false; $card.TopMost = $true
        $card.StartPosition = [Windows.Forms.FormStartPosition]::Manual
        $card.BackColor = Get-UiColor 'Surface'
        $card.Width = 440; $card.Height = 190
        $card.Padding = New-Object Windows.Forms.Padding(1)
        $card.Region = New-RoundedRectPath (New-Object Drawing.Rectangle(0, 0, $card.Width, $card.Height)) 14

        $icon = New-Object Windows.Forms.Panel
        $icon.Left = 20; $icon.Top = 18; $icon.Width = 54; $icon.Height = 54
        $icon.BackColor = [Drawing.Color]::FromArgb(239, 246, 255)
        $icon.Add_Paint({
            param($s, $e)
            $g = $e.Graphics
            $g.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $pen = New-Object Drawing.Pen((Get-UiColor 'Primary'), 2)
            $pen.StartCap = [Drawing.Drawing2D.LineCap]::Round
            $pen.EndCap = [Drawing.Drawing2D.LineCap]::Round
            $calendarPath = New-RoundedRectPath (New-Object Drawing.Rectangle(15, 14, 24, 24)) 4
            $g.DrawPath($pen, $calendarPath)
            $g.DrawLine($pen, 15, 21, 39, 21)
            $g.DrawLine($pen, 21, 10, 21, 17)
            $g.DrawLine($pen, 33, 10, 33, 17)
            $calendarPath.Dispose(); $pen.Dispose()
        })
        $card.Controls.Add($icon)
        $title = New-Object Windows.Forms.Label
        $title.Text = 'Напоминание'; $title.Left = 92; $title.Top = 21; $title.Width = 300; $title.Height = 25
        $title.Font = New-Object Drawing.Font('Segoe UI Semibold', 10, [Drawing.FontStyle]::Bold); $title.ForeColor = Get-UiColor 'Text'
        $card.Controls.Add($title)
        $close = New-Object Windows.Forms.Button
        $close.Text = '×'; $close.Left = 394; $close.Top = 15; $close.Width = 28; $close.Height = 28
        $close.FlatStyle = [Windows.Forms.FlatStyle]::Flat; $close.FlatAppearance.BorderSize = 0
        $close.BackColor = Get-UiColor 'Surface'; $close.ForeColor = Get-UiColor 'Muted'
        $close.Font = New-Object Drawing.Font('Segoe UI', 12)
        $close.Cursor = [Windows.Forms.Cursors]::Hand
        $card.Controls.Add($close)
        $divider = New-Object Windows.Forms.Panel
        $divider.Left = 20; $divider.Top = 70; $divider.Width = 400; $divider.Height = 1
        $divider.BackColor = Get-UiColor 'Line'; $card.Controls.Add($divider)

        $body = New-Object Windows.Forms.Label
        $body.Text = [string]$notification.Title
        $body.Left = 92; $body.Top = 82; $body.Width = 320; $body.Height = 24
        $body.Font = New-Object Drawing.Font('Segoe UI Semibold', 9.5, [Drawing.FontStyle]::Bold)
        $body.ForeColor = Get-UiColor 'Text'; $body.AutoEllipsis = $true
        $card.Controls.Add($body)

        $dueIcon = New-Object Windows.Forms.Panel
        $dueIcon.Left = 92; $dueIcon.Top = 113; $dueIcon.Width = 18; $dueIcon.Height = 18
        $dueIcon.BackColor = Get-UiColor 'Surface'
        $dueIcon.Add_Paint({
            param($s, $e)
            $pen = New-Object Drawing.Pen((Get-UiColor 'Muted'), 1.3)
            $e.Graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $e.Graphics.DrawRectangle($pen, 2, 4, 13, 12)
            $e.Graphics.DrawLine($pen, 2, 8, 15, 8)
            $e.Graphics.DrawLine($pen, 5, 2, 5, 6)
            $e.Graphics.DrawLine($pen, 12, 2, 12, 6)
            $pen.Dispose()
        })
        $card.Controls.Add($dueIcon)
        $dueDate = ([string]::Join(', ', @($notification.Due | ForEach-Object { $_.Date.ToString('dd.MM.yyyy') })))
        $dueLabel = New-Object Windows.Forms.Label
        $dueLabel.Text = 'Срок:'; $dueLabel.Left = 118; $dueLabel.Top = 111; $dueLabel.Width = 42; $dueLabel.Height = 22
        $dueLabel.ForeColor = Get-UiColor 'Muted'; $card.Controls.Add($dueLabel)
        $dueValue = New-Object Windows.Forms.Label
        $dueValue.Text = $dueDate; $dueValue.Left = 160; $dueValue.Top = 111; $dueValue.Width = 210; $dueValue.Height = 22
        $dueValue.ForeColor = Get-UiColor 'Primary'; $dueValue.Font = New-Object Drawing.Font('Segoe UI Semibold', 9)
        $card.Controls.Add($dueValue)
        $open = New-Object Windows.Forms.Button
        $open.Text = 'Открыть'; $open.Left = 16; $open.Top = 148; $open.Width = 114; $open.Height = 30
        Set-SecondaryButtonLook $open; $card.Controls.Add($open)
        $later = New-Object Windows.Forms.Button
        $later.Text = 'Напомнить позже'; $later.Left = 138; $later.Top = 148; $later.Width = 138; $later.Height = 30
        Set-SecondaryButtonLook $later; $card.Controls.Add($later)
        $done = New-Object Windows.Forms.Button
        $done.Text = 'Выполнено'; $done.Left = 284; $done.Top = 148; $done.Width = 122; $done.Height = 30
        Set-PrimaryButtonLook $done; $card.Controls.Add($done)

        $close.Add_Click({ Invoke-ActiveNotificationAction 'close' })
        $open.Add_Click({ Invoke-ActiveNotificationAction 'open' })
        $later.Add_Click({ Invoke-ActiveNotificationAction 'later' })
        $done.Add_Click({ Invoke-ActiveNotificationAction 'done' })
        $card.Add_Shown({
            $activeCard = $script:ActiveNotificationCard
            if ($null -eq $activeCard) { return }
            $activeCard.WindowState = [Windows.Forms.FormWindowState]::Normal
            $activeCard.TopMost = $true
            $screen = [Windows.Forms.Screen]::PrimaryScreen
            if ($null -ne $form -and $form.IsHandleCreated) {
                $formScreen = [Windows.Forms.Screen]::FromControl($form)
                if ($null -ne $formScreen) { $screen = $formScreen }
            }
            $wa = $screen.WorkingArea
            $activeCard.Location = New-Object Drawing.Point([int]($wa.Left + 24), [int]($wa.Bottom - $activeCard.Height - 56))
        })
        # Do not assign the main form as owner: an owned window is minimized
        # together with its owner and would hide notifications from the user.
        $card.Show()
        $card.BringToFront()
        $card.Activate()
    }

    function Check-Notifications {
        if (-not $script:NotificationsEnabled) { return }
        if ($script:NotificationChecking) { return }
        if ($null -ne $script:ActiveNotificationCard -or $script:NotificationQueue.Count -gt 0) {
            Show-NextNotification
            return
        }
        $script:NotificationChecking = $true
        try {
            foreach ($item in @(Get-DueNotificationCandidates)) {
                $key = Get-NotificationKey $item
                if ($script:NotificationKeys.Add($key)) { [void]$script:NotificationQueue.Enqueue($item) }
            }
            Show-NextNotification
        } finally { $script:NotificationChecking = $false }
    }

    $root = New-Object Windows.Forms.TableLayoutPanel
    $root.Dock = 'Fill'
    $root.Padding = New-Object Windows.Forms.Padding(16)
    $root.RowCount = 2
    $root.ColumnCount = 1
    Add-TableRowStyle $root 'Absolute' 318
    Add-TableRowStyle $root 'Percent' 100
    $rrfqPage.Controls.Add($root)

    $topLayout = New-Object Windows.Forms.TableLayoutPanel
    $topLayout.Dock = 'Fill'
    $topLayout.ColumnCount = 2
    $topLayout.RowCount = 1
    Add-TableColumnStyle $topLayout 'Percent' 66
    Add-TableColumnStyle $topLayout 'Percent' 34
    Add-TableRowStyle $topLayout 'Percent' 100
    $root.Controls.Add($topLayout, 0, 0)

    $filesGroup = New-CardPanel
    $filesGroupTitle = New-Object Windows.Forms.Label
    $filesGroupTitle.Text = '1. Выбор файла'
    $filesGroupTitle.Dock = 'Top'
    $filesGroupTitle.Height = 30
    $filesGroupTitle.Font = New-Object Drawing.Font('Segoe UI', 9.5, [Drawing.FontStyle]::Bold)
    $filesGroupTitle.ForeColor = Get-UiColor 'Text'
    $filesGroupTitle.Padding = New-Object Windows.Forms.Padding(4, 8, 0, 2)
    $filesGroup.Dock = 'Fill'
    $filesGroup.Padding = New-Object Windows.Forms.Padding(12)
    $topLayout.Controls.Add($filesGroup, 0, 0)

    $filesLayout = New-Object Windows.Forms.TableLayoutPanel
    $filesLayout.Dock = 'Fill'
    $filesLayout.ColumnCount = 3
    $filesLayout.RowCount = 3
    Add-TableColumnStyle $filesLayout 'Absolute' 46
    Add-TableColumnStyle $filesLayout 'Percent' 100
    Add-TableColumnStyle $filesLayout 'Absolute' 142
    Add-TableRowStyle $filesLayout 'Absolute' 34
    Add-TableRowStyle $filesLayout 'Percent' 100
    Add-TableRowStyle $filesLayout 'Absolute' 40
    $filesGroup.Controls.Add($filesLayout)
    $filesGroup.Controls.Add($filesGroupTitle)

    $lblRfq = New-Object Windows.Forms.Label
    $lblRfq.Text = 'RFQ'
    $lblRfq.Dock = 'Fill'
    $lblRfq.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $filesLayout.Controls.Add($lblRfq, 0, 0)

    $txtRfq = New-Object Windows.Forms.TextBox
    $txtRfq.Dock = 'Fill'
    $txtRfq.Margin = New-Object Windows.Forms.Padding(0, 4, 8, 0)
    $filesLayout.Controls.Add($txtRfq, 1, 0)

    $btnRfq = New-Object Windows.Forms.Button
    $btnRfq.Text = 'Выбрать...'
    $btnRfq.Dock = 'Fill'
    $btnRfq.Margin = New-Object Windows.Forms.Padding(0, 2, 0, 0)
    Set-SecondaryButtonLook $btnRfq
    $filesLayout.Controls.Add($btnRfq, 2, 0)

    $supplierGrid = New-Object Windows.Forms.DataGridView
    $supplierGrid.Dock = 'Fill'
    $supplierGrid.AllowUserToAddRows = $false
    $supplierGrid.AllowUserToResizeRows = $false
    $supplierGrid.RowHeadersVisible = $false
    $supplierGrid.SelectionMode = 'FullRowSelect'
    $supplierGrid.MultiSelect = $true
    $supplierGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    Set-GridLook $supplierGrid
    $colFile = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $colFile.Name = 'Path'
    $colFile.HeaderText = 'Файл поставщика'
    $colFile.FillWeight = 72
    $colFile.ReadOnly = $true
    $supplierGrid.Columns.Add($colFile) | Out-Null
    $colSupplier = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $colSupplier.Name = 'Supplier'
    $colSupplier.HeaderText = 'Поставщик (можно вписать)'
    $colSupplier.FillWeight = 28
    $supplierGrid.Columns.Add($colSupplier) | Out-Null
    $filesLayout.Controls.Add($supplierGrid, 0, 1)
    $filesLayout.SetColumnSpan($supplierGrid, 3)

    $fileButtons = New-Object Windows.Forms.FlowLayoutPanel
    $fileButtons.Dock = 'Fill'
    $fileButtons.FlowDirection = 'LeftToRight'
    $fileButtons.WrapContents = $false
    $fileButtons.Padding = New-Object Windows.Forms.Padding(0, 6, 0, 0)
    $filesLayout.Controls.Add($fileButtons, 0, 2)
    $filesLayout.SetColumnSpan($fileButtons, 3)

    $btnAdd = New-Object Windows.Forms.Button
    $btnAdd.Text = 'Добавить поставщиков'
    $btnAdd.Width = 180
    $btnAdd.Height = 30
    Set-SecondaryButtonLook $btnAdd
    $fileButtons.Controls.Add($btnAdd)

    $btnRemove = New-Object Windows.Forms.Button
    $btnRemove.Text = 'Удалить выбранные'
    $btnRemove.Width = 160
    $btnRemove.Height = 30
    Set-SecondaryButtonLook $btnRemove
    $fileButtons.Controls.Add($btnRemove)

    $actionsGroup = New-CardPanel
    $actionsGroupTitle = New-Object Windows.Forms.Label
    $actionsGroupTitle.Text = '2. Параметры и запуск'
    $actionsGroupTitle.Dock = 'Top'
    $actionsGroupTitle.Height = 30
    $actionsGroupTitle.Font = New-Object Drawing.Font('Segoe UI', 9.5, [Drawing.FontStyle]::Bold)
    $actionsGroupTitle.ForeColor = Get-UiColor 'Text'
    $actionsGroupTitle.Padding = New-Object Windows.Forms.Padding(4, 8, 0, 2)
    $actionsGroup.Dock = 'Fill'
    $actionsGroup.Padding = New-Object Windows.Forms.Padding(12)
    $topLayout.Controls.Add($actionsGroup, 1, 0)

    $actionsLayout = New-Object Windows.Forms.TableLayoutPanel
    $actionsLayout.Dock = 'Fill'
    $actionsLayout.ColumnCount = 1
    $actionsLayout.RowCount = 6
    Add-TableRowStyle $actionsLayout 'Absolute' 54
    Add-TableRowStyle $actionsLayout 'Absolute' 42
    Add-TableRowStyle $actionsLayout 'Absolute' 42
    Add-TableRowStyle $actionsLayout 'Absolute' 42
    Add-TableRowStyle $actionsLayout 'Absolute' 10
    Add-TableRowStyle $actionsLayout 'Absolute' 106
    $actionsGroup.Controls.Add($actionsLayout)
    $actionsGroup.Controls.Add($actionsGroupTitle)

    $groupPriority = New-CardPanel
    $groupPriorityTitle = New-Object Windows.Forms.Label
    $groupPriorityTitle.Text = 'Приоритет'
    $groupPriorityTitle.AutoSize = $true
    $groupPriorityTitle.Location = New-Object Drawing.Point(14, 5)
    $groupPriorityTitle.Font = New-Object Drawing.Font('Segoe UI', 8.5, [Drawing.FontStyle]::Bold)
    $groupPriorityTitle.ForeColor = Get-UiColor 'Muted'
    $groupPriority.Controls.Add($groupPriorityTitle)
    $groupPriority.Dock = 'Fill'
    $actionsLayout.Controls.Add($groupPriority, 0, 0)

    $radioPrice = New-Object Windows.Forms.RadioButton
    $radioPrice.Text = 'Цена'
    $radioPrice.Left = 14
    $radioPrice.Top = 28
    $radioPrice.Width = 90
    $radioPrice.Checked = $true
    $groupPriority.Controls.Add($radioPrice)

    $radioLead = New-Object Windows.Forms.RadioButton
    $radioLead.Text = 'Срок'
    $radioLead.Left = 118
    $radioLead.Top = 28
    $radioLead.Width = 90
    $groupPriority.Controls.Add($radioLead)

    $btnPreview = New-Object Windows.Forms.Button
    $btnPreview.Text = 'Сформировать preview'
    $btnPreview.Dock = 'Fill'
    Set-PrimaryButtonLook $btnPreview
    $actionsLayout.Controls.Add($btnPreview, 0, 1)

    $btnManual = New-Object Windows.Forms.Button
    $btnManual.Text = 'Ручное сопоставление'
    $btnManual.Dock = 'Fill'
    $btnManual.Enabled = $false
    Set-SecondaryButtonLook $btnManual
    $actionsLayout.Controls.Add($btnManual, 0, 2)

    $btnSave = New-Object Windows.Forms.Button
    $btnSave.Text = 'Создать результирующий RRFQ'
    $btnSave.Dock = 'Fill'
    $btnSave.Enabled = $false
    Set-PrimaryButtonLook $btnSave
    $actionsLayout.Controls.Add($btnSave, 0, 3)

    $logBox = New-Object Windows.Forms.TextBox
    $logBox.Multiline = $true
    $logBox.ReadOnly = $true
    $logBox.Dock = 'Fill'
    $logBox.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
    $logBox.BackColor = Get-UiColor 'SurfaceAlt'
    $logBox.Text = "Выберите RFQ и добавьте файлы поставщиков.`r`nПосле preview здесь появится сводка."
    $actionsLayout.Controls.Add($logBox, 0, 5)

    $previewGroup = New-CardPanel
    $previewGroupTitle = New-Object Windows.Forms.Label
    $previewGroupTitle.Text = '3. Preview и ручная корректировка'
    $previewGroupTitle.Dock = 'Top'
    $previewGroupTitle.Height = 30
    $previewGroupTitle.Font = New-Object Drawing.Font('Segoe UI', 9.5, [Drawing.FontStyle]::Bold)
    $previewGroupTitle.ForeColor = Get-UiColor 'Text'
    $previewGroupTitle.Padding = New-Object Windows.Forms.Padding(4, 8, 0, 2)
    $previewGroup.Dock = 'Fill'
    $previewGroup.Padding = New-Object Windows.Forms.Padding(10)
    $root.Controls.Add($previewGroup, 0, 1)

    $previewLayout = New-Object Windows.Forms.TableLayoutPanel
    $previewLayout.Dock = 'Fill'
    $previewLayout.RowCount = 3
    $previewLayout.ColumnCount = 1
    Add-TableRowStyle $previewLayout 'Absolute' 38
    Add-TableRowStyle $previewLayout 'Percent' 55
    Add-TableRowStyle $previewLayout 'Percent' 45
    $previewGroup.Controls.Add($previewLayout)
    $previewGroup.Controls.Add($previewGroupTitle)

    $filterPanel = New-Object Windows.Forms.FlowLayoutPanel
    $filterPanel.Dock = 'Fill'
    $filterPanel.FlowDirection = 'LeftToRight'
    $filterPanel.WrapContents = $false
    $filterPanel.Padding = New-Object Windows.Forms.Padding(0, 3, 0, 0)
    $previewLayout.Controls.Add($filterPanel, 0, 0)

    $lblFilter = New-Object Windows.Forms.Label
    $lblFilter.Text = 'Фильтр:'
    $lblFilter.Width = 56
    $lblFilter.Height = 26
    $lblFilter.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $filterPanel.Controls.Add($lblFilter)

    $cmbFilter = New-Object Windows.Forms.ComboBox
    $cmbFilter.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbFilter.Width = 190
    [void]$cmbFilter.Items.Add('Все строки')
    [void]$cmbFilter.Items.Add('Только победители')
    [void]$cmbFilter.Items.Add('Без квот')
    [void]$cmbFilter.Items.Add('Предупреждения')
    [void]$cmbFilter.Items.Add('Ручная проверка')
    $cmbFilter.SelectedIndex = 0
    $filterPanel.Controls.Add($cmbFilter)

    $lblSearch = New-Object Windows.Forms.Label
    $lblSearch.Text = 'Поиск:'
    $lblSearch.Width = 52
    $lblSearch.Height = 26
    $lblSearch.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $filterPanel.Controls.Add($lblSearch)

    $txtSearch = New-Object Windows.Forms.TextBox
    $txtSearch.Width = 240
    $filterPanel.Controls.Add($txtSearch)

    $btnIncludeSelected = New-Object Windows.Forms.Button
    $btnIncludeSelected.Text = 'Включить'
    $btnIncludeSelected.Width = 92
    $btnIncludeSelected.Height = 28
    Set-SecondaryButtonLook $btnIncludeSelected
    $filterPanel.Controls.Add($btnIncludeSelected)

    $btnExcludeSelected = New-Object Windows.Forms.Button
    $btnExcludeSelected.Text = 'Исключить'
    $btnExcludeSelected.Width = 96
    $btnExcludeSelected.Height = 28
    Set-SecondaryButtonLook $btnExcludeSelected
    $filterPanel.Controls.Add($btnExcludeSelected)

    $previewGrid = New-Object Windows.Forms.DataGridView
    $previewGrid.Dock = 'Fill'
    $previewGrid.AllowUserToAddRows = $false
    $previewGrid.AllowUserToResizeRows = $false
    $previewGrid.RowHeadersVisible = $false
    $previewGrid.SelectionMode = 'FullRowSelect'
    $previewGrid.MultiSelect = $false
    $previewGrid.AutoSizeRowsMode = [Windows.Forms.DataGridViewAutoSizeRowsMode]::DisplayedCells
    $previewGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $previewGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both
    Set-GridLook $previewGrid

    $includeCol = New-Object Windows.Forms.DataGridViewCheckBoxColumn
    $includeCol.Name = 'Include'
    $includeCol.HeaderText = 'Вкл'
    $includeCol.Width = 44
    $includeCol.Frozen = $true
    $previewGrid.Columns.Add($includeCol) | Out-Null

    foreach ($columnInfo in @(
        @('Status', 'Статус', 115, 'None', $true),
        @('Row', '№', 54, 'None', $false),
        @('Value', 'Value', 230, 'None', $false),
        @('PN', 'PN', 170, 'None', $false),
        @('RussianRemark', 'Russian remark', 210, 'None', $false),
        @('ChinaRemark', 'China remark', 210, 'None', $false),
        @('DC', 'D/C', 90, 'None', $false),
        @('MfgRussia', 'MFG from Russia', 150, 'None', $false),
        @('MfgChina', 'MFG from China', 150, 'None', $false),
        @('QtyPacking', 'Q-ty in packing', 130, 'None', $false),
        @('QtyToBuy', 'Q-ty to Buy/pcs', 130, 'None', $false),
        @('Supplier', 'Победитель', 110, 'None', $false),
        @('Reason', 'Почему выбран', 170, 'None', $false),
        @('Price', 'Цена', 92, 'None', $false),
        @('Lead', 'Срок', 105, 'None', $false),
        @('LeadTotal', 'Lead time total', 128, 'None', $false),
        @('Quotes', 'Все квоты', 330, 'None', $false),
        @('Warnings', 'Предупреждения', 330, 'None', $false)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        $col.AutoSizeMode = [Enum]::Parse([Windows.Forms.DataGridViewAutoSizeColumnMode], [string]$columnInfo[3])
        $col.Frozen = [bool]$columnInfo[4]
        $col.ReadOnly = $true
        $previewGrid.Columns.Add($col) | Out-Null
    }
    $previewLayout.Controls.Add($previewGrid, 0, 1)
    $script:LastPreviewGrid = $previewGrid

    $detailsSplit = New-Object Windows.Forms.SplitContainer
    $detailsSplit.Dock = 'Fill'
    $detailsSplit.Orientation = 'Vertical'
    $detailsSplit.SplitterDistance = 520
    $previewLayout.Controls.Add($detailsSplit, 0, 2)

    $detailsBox = New-Object Windows.Forms.TextBox
    $detailsBox.Multiline = $true
    $detailsBox.ReadOnly = $true
    $detailsBox.Dock = 'Fill'
    $detailsBox.ScrollBars = [Windows.Forms.ScrollBars]::Vertical
    $detailsBox.BackColor = Get-UiColor 'SurfaceAlt'
    $detailsBox.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
    $detailsSplit.Panel1.Controls.Add($detailsBox)

    $quotePanel = New-Object Windows.Forms.TableLayoutPanel
    $quotePanel.Dock = 'Fill'
    $quotePanel.RowCount = 2
    $quotePanel.ColumnCount = 1
    Add-TableRowStyle $quotePanel 'Absolute' 36
    Add-TableRowStyle $quotePanel 'Percent' 100
    $detailsSplit.Panel2.Controls.Add($quotePanel)

    $quoteButtons = New-Object Windows.Forms.FlowLayoutPanel
    $quoteButtons.Dock = 'Fill'
    $quoteButtons.FlowDirection = 'LeftToRight'
    $quoteButtons.WrapContents = $false
    $quoteButtons.Padding = New-Object Windows.Forms.Padding(0, 3, 0, 0)
    $quotePanel.Controls.Add($quoteButtons, 0, 0)

    $btnChooseQuote = New-Object Windows.Forms.Button
    $btnChooseQuote.Text = 'Сделать победителем'
    $btnChooseQuote.Width = 160
    $btnChooseQuote.Height = 28
    Set-SecondaryButtonLook $btnChooseQuote
    $quoteButtons.Controls.Add($btnChooseQuote)

    $btnAddManualQuote = New-Object Windows.Forms.Button
    $btnAddManualQuote.Text = 'Добавить квоту вручную'
    $btnAddManualQuote.Width = 180
    $btnAddManualQuote.Height = 28
    Set-SecondaryButtonLook $btnAddManualQuote
    $quoteButtons.Controls.Add($btnAddManualQuote)

    $quoteDetailsGrid = New-Object Windows.Forms.DataGridView
    $quoteDetailsGrid.Dock = 'Fill'
    $quoteDetailsGrid.AllowUserToAddRows = $false
    $quoteDetailsGrid.AllowUserToResizeRows = $false
    $quoteDetailsGrid.RowHeadersVisible = $false
    $quoteDetailsGrid.ReadOnly = $true
    $quoteDetailsGrid.SelectionMode = 'FullRowSelect'
    $quoteDetailsGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $quoteDetailsGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both
    Set-GridLook $quoteDetailsGrid
    foreach ($columnInfo in @(
        @('Winner', 'Побед.', 62),
        @('Supplier', 'Поставщик', 100),
        @('Price', 'Цена', 86),
        @('Lead', 'Срок', 96),
        @('LeadTotal', 'Total', 96),
        @('PN', 'PN', 150),
        @('MfgChina', 'MFG China', 120),
        @('QtyPacking', 'Pack', 80),
        @('QtyToBuy', 'Buy qty', 80),
        @('Match', 'Матчинг', 92),
        @('Warning', 'Предупреждение', 220)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        $quoteDetailsGrid.Columns.Add($col) | Out-Null
    }
    $quotePanel.Controls.Add($quoteDetailsGrid, 0, 1)

    Initialize-PurchaseStore
    $script:NotificationsEnabled = Get-PurchaseBooleanSetting 'notifications.enabled' $true
    # Bitrix access is deliberately disabled on every fresh application start.
    $script:BitrixIntegrationEnabled = $false
    $chkBitrixIntegrationEnabled.Checked = $false
    $bitrixStatusLabel.Text = 'Подтягивание из Bitrix отключено.'

    function Save-GridColumnWidths {
        param($Grid, [string]$SettingKey)

        if ($null -eq $Grid -or [string]::IsNullOrWhiteSpace($SettingKey)) { return }
        $parts = New-Object System.Collections.ArrayList
        foreach ($column in $Grid.Columns) {
            if (-not [string]::IsNullOrWhiteSpace([string]$column.Name)) {
                [void]$parts.Add(('{0}={1}' -f $column.Name, [int]$column.Width))
            }
        }
        Set-PurchaseSetting $SettingKey ([string]::Join(';', @($parts)))
    }

    function Restore-GridColumnWidths {
        param($Grid, [string]$SettingKey)

        if ($null -eq $Grid -or [string]::IsNullOrWhiteSpace($SettingKey)) { return }
        $text = Get-PurchaseSetting $SettingKey
        if ([string]::IsNullOrWhiteSpace($text)) { return }

        foreach ($part in $text.Split(';')) {
            if ([string]::IsNullOrWhiteSpace($part) -or $part -notmatch '=') { continue }
            $items = $part.Split('=', 2)
            $name = $items[0]
            $width = 0
            if ([string]::IsNullOrWhiteSpace($name) -or -not [int]::TryParse($items[1], [ref]$width)) { continue }
            if ($width -lt 30) { continue }
            if ($Grid.Columns.Contains($name)) {
                $Grid.Columns[$name].Width = $width
            }
        }
    }

    function Show-SimpleInputDialog {
        param([string]$Title, [string]$Label, [string]$DefaultValue = '')

        $dialog = New-Object Windows.Forms.Form
        $dialog.Text = $Title
        $dialog.Width = 420
        $dialog.Height = 160
        $dialog.StartPosition = 'CenterParent'
        $dialog.Font = New-Object Drawing.Font('Segoe UI', 9)
        $dialog.BackColor = Get-UiColor 'Canvas'

        $labelControl = New-Object Windows.Forms.Label
        $labelControl.Text = $Label
        $labelControl.Left = 12
        $labelControl.Top = 14
        $labelControl.Width = 370
        $dialog.Controls.Add($labelControl)

        $textBox = New-Object Windows.Forms.TextBox
        $textBox.Left = 12
        $textBox.Top = 42
        $textBox.Width = 380
        $textBox.Text = $DefaultValue
        $dialog.Controls.Add($textBox)

        $btnOk = New-Object Windows.Forms.Button
        $btnOk.Text = 'OK'
        $btnOk.Left = 206
        $btnOk.Top = 80
        $btnOk.Width = 88
        Set-PrimaryButtonLook $btnOk
        $dialog.Controls.Add($btnOk)

        $btnCancel = New-Object Windows.Forms.Button
        $btnCancel.Text = 'Отмена'
        $btnCancel.Left = 304
        $btnCancel.Top = 80
        $btnCancel.Width = 88
        Set-SecondaryButtonLook $btnCancel
        $dialog.Controls.Add($btnCancel)

        $dialog.Tag = $null
        $btnOk.Add_Click({
            $dialog.Tag = $textBox.Text
            $dialog.Close()
        })
        $btnCancel.Add_Click({ $dialog.Close() })
        Apply-ControlTreeLook $dialog
        [void]$dialog.ShowDialog($form)
        return [string]$dialog.Tag
    }

    function Show-DocumentTypeDialog {
        $dialog = New-Object Windows.Forms.Form
        $dialog.Text = 'Тип документа'
        $dialog.Width = 360
        $dialog.Height = 150
        $dialog.StartPosition = 'CenterParent'
        $dialog.Font = New-Object Drawing.Font('Segoe UI', 9)
        $dialog.BackColor = Get-UiColor 'Canvas'

        $labelControl = New-Object Windows.Forms.Label
        $labelControl.Text = 'Выберите тип для загружаемых файлов'
        $labelControl.Left = 12
        $labelControl.Top = 14
        $labelControl.Width = 320
        $dialog.Controls.Add($labelControl)

        $combo = New-Object Windows.Forms.ComboBox
        $combo.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDownList
        $combo.Left = 12
        $combo.Top = 42
        $combo.Width = 320
        foreach ($item in @('PI/Invoice', 'Общий PI', 'RFQ', 'RRFQ', 'PO', 'ERP', 'Other')) {
            [void]$combo.Items.Add($item)
        }
        $combo.SelectedIndex = 0
        $dialog.Controls.Add($combo)

        $btnOk = New-Object Windows.Forms.Button
        $btnOk.Text = 'OK'
        $btnOk.Left = 244
        $btnOk.Top = 78
        $btnOk.Width = 88
        Set-PrimaryButtonLook $btnOk
        $dialog.Controls.Add($btnOk)

        $dialog.Tag = $null
        $btnOk.Add_Click({
            $dialog.Tag = [string]$combo.SelectedItem
            $dialog.Close()
        })
        Apply-ControlTreeLook $dialog
        [void]$dialog.ShowDialog($form)
        return [string]$dialog.Tag
    }

    function Show-DatePickerDialog {
        param([string]$CurrentValue)

        $dialog = New-Object Windows.Forms.Form
        $dialog.Text = 'Выберите дату'
        $dialog.Width = 320
        $dialog.Height = 280
        $dialog.StartPosition = 'CenterParent'
        $dialog.FormBorderStyle = [Windows.Forms.FormBorderStyle]::FixedDialog
        $dialog.MaximizeBox = $false
        $dialog.MinimizeBox = $false
        $dialog.Font = New-Object Drawing.Font('Segoe UI', 9)
        $dialog.BackColor = Get-UiColor 'Canvas'

        $calendar = New-Object Windows.Forms.MonthCalendar
        $calendar.Left = 12
        $calendar.Top = 10
        $calendar.MaxSelectionCount = 1
        $date = Convert-PurchaseDateText $CurrentValue
        if ($null -ne $date) {
            $calendar.SetDate($date)
        } else {
            $calendar.SetDate((Get-Date).Date)
        }
        $dialog.Controls.Add($calendar)

        $btnOk = New-Object Windows.Forms.Button
        $btnOk.Text = 'OK'
        $btnOk.Left = 62
        $btnOk.Top = 205
        $btnOk.Width = 58
        Set-PrimaryButtonLook $btnOk
        $dialog.Controls.Add($btnOk)

        $btnClear = New-Object Windows.Forms.Button
        $btnClear.Text = 'Очистить'
        $btnClear.Left = 126
        $btnClear.Top = 205
        $btnClear.Width = 78
        Set-SecondaryButtonLook $btnClear
        $dialog.Controls.Add($btnClear)

        $btnCancel = New-Object Windows.Forms.Button
        $btnCancel.Text = 'Отмена'
        $btnCancel.Left = 210
        $btnCancel.Top = 205
        $btnCancel.Width = 78
        Set-SecondaryButtonLook $btnCancel
        $dialog.Controls.Add($btnCancel)

        $dialog.Tag = $null
        $btnOk.Add_Click({
            $dialog.Tag = $calendar.SelectionStart.ToString('dd.MM.yyyy')
            $dialog.Close()
        })
        $calendar.Add_DoubleClick({
            $dialog.Tag = $calendar.SelectionStart.ToString('dd.MM.yyyy')
            $dialog.Close()
        })
        $btnClear.Add_Click({
            $dialog.Tag = ''
            $dialog.Close()
        })
        $btnCancel.Add_Click({ $dialog.Close() })

        Apply-ControlTreeLook $dialog
        [void]$dialog.ShowDialog($form)
        return $dialog.Tag
    }

    function New-PlaceholderPage {
        param($Page, [string]$Title, [string]$Message)

        $layout = New-Object Windows.Forms.TableLayoutPanel
        $layout.Dock = 'Fill'
        $layout.Padding = New-Object Windows.Forms.Padding(24)
        $layout.RowCount = 3
        $layout.ColumnCount = 1
        Add-TableRowStyle $layout 'Absolute' 44
        Add-TableRowStyle $layout 'Absolute' 70
        Add-TableRowStyle $layout 'Percent' 100
        $Page.Controls.Add($layout)

        $titleLabel = New-Object Windows.Forms.Label
        $titleLabel.Text = $Title
        $titleLabel.Dock = 'Fill'
        $titleLabel.Font = New-Object Drawing.Font('Segoe UI', 15, [Drawing.FontStyle]::Bold)
        $titleLabel.ForeColor = Get-UiColor 'Text'
        $layout.Controls.Add($titleLabel, 0, 0)

        $messageLabel = New-Object Windows.Forms.Label
        $messageLabel.Text = $Message
        $messageLabel.Dock = 'Fill'
        $messageLabel.ForeColor = Get-UiColor 'Muted'
        $messageLabel.TextAlign = [Drawing.ContentAlignment]::TopLeft
        $layout.Controls.Add($messageLabel, 0, 1)
    }

    $purchaseLayout = New-Object Windows.Forms.TableLayoutPanel
    $purchaseLayout.Dock = 'Fill'
    $purchaseLayout.Padding = New-Object Windows.Forms.Padding(12)
    $purchaseLayout.RowCount = 3
    $purchaseLayout.ColumnCount = 1
    Add-TableRowStyle $purchaseLayout 'Absolute' 86
    Add-TableRowStyle $purchaseLayout 'Absolute' 46
    Add-TableRowStyle $purchaseLayout 'Percent' 100
    $purchasePage.Controls.Add($purchaseLayout)

    $purchaseMetricsPanel = New-Object Windows.Forms.TableLayoutPanel
    $purchaseMetricsPanel.Dock = 'Fill'
    $purchaseMetricsPanel.ColumnCount = 4
    $purchaseMetricsPanel.RowCount = 1
    $purchaseMetricsPanel.Padding = New-Object Windows.Forms.Padding(0, 0, 0, 8)
    foreach ($i in 1..4) { Add-TableColumnStyle $purchaseMetricsPanel 'Percent' 25 }
    Add-TableRowStyle $purchaseMetricsPanel 'Percent' 100
    $purchaseLayout.Controls.Add($purchaseMetricsPanel, 0, 0)
    $metricActive = New-CockpitMetricCard 'Активные сделки' '0' 'В текущей очереди' 'Info'
    $metricOverdue = New-CockpitMetricCard 'Просрочено' '0' 'Нужна реакция' 'Danger'
    $metricToday = New-CockpitMetricCard 'На сегодня' '0' 'Ожидается поступление' 'Warn'
    $metricAttention = New-CockpitMetricCard 'Требуют внимания' '0' 'Нужна ручная проверка' 'Attention'
    $purchaseMetricsPanel.Controls.Add($metricActive, 0, 0)
    $purchaseMetricsPanel.Controls.Add($metricOverdue, 1, 0)
    $purchaseMetricsPanel.Controls.Add($metricToday, 2, 0)
    $purchaseMetricsPanel.Controls.Add($metricAttention, 3, 0)


    $purchaseToolbar = New-Object Windows.Forms.TableLayoutPanel
    $purchaseToolbar.Dock = 'Fill'
    $purchaseToolbar.ColumnCount = 11
    $purchaseToolbar.RowCount = 1
    $purchaseToolbar.Padding = New-Object Windows.Forms.Padding(0, 5, 0, 5)
    Add-TableColumnStyle $purchaseToolbar 'Absolute' 64
    Add-TableColumnStyle $purchaseToolbar 'Absolute' 170
    Add-TableColumnStyle $purchaseToolbar 'Absolute' 68
    Add-TableColumnStyle $purchaseToolbar 'Absolute' 160
    Add-TableColumnStyle $purchaseToolbar 'Percent' 100
    Add-TableColumnStyle $purchaseToolbar 'Absolute' 112
    Add-TableColumnStyle $purchaseToolbar 'Absolute' 154
    Add-TableColumnStyle $purchaseToolbar 'Absolute' 132
    Add-TableColumnStyle $purchaseToolbar 'Absolute' 92
    Add-TableColumnStyle $purchaseToolbar 'Absolute' 124
    Add-TableColumnStyle $purchaseToolbar 'Absolute' 96
    Add-TableRowStyle $purchaseToolbar 'Percent' 100
    $purchaseLayout.Controls.Add($purchaseToolbar, 0, 1)

    $purchaseToolbar.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'Поиск:'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }), 0, 0)
    $txtDealSearch = New-Object Windows.Forms.TextBox
    $txtDealSearch.Dock = 'Fill'
    $txtDealSearch.Margin = New-Object Windows.Forms.Padding(0, 2, 10, 2)
    $purchaseToolbar.Controls.Add($txtDealSearch, 1, 0)

    $purchaseToolbar.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'Фильтр:'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }), 2, 0)
    $cmbDealFilter = New-Object Windows.Forms.ComboBox
    $cmbDealFilter.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbDealFilter.Dock = 'Fill'
    $cmbDealFilter.Margin = New-Object Windows.Forms.Padding(0, 2, 10, 2)
    foreach ($item in @('Активные', 'На отслеживании', 'Китай', 'Закупка', 'Пауза', 'Ожидается поступление', 'Готово', 'Архив', 'Все включая архив')) {
        [void]$cmbDealFilter.Items.Add($item)
    }
    $cmbDealFilter.SelectedIndex = 0
    $purchaseToolbar.Controls.Add($cmbDealFilter, 3, 0)

    $btnNewDeal = New-Object Windows.Forms.Button
    $btnNewDeal.Text = 'Новая сделка'
    $btnNewDeal.Dock = 'Fill'
    $btnNewDeal.Margin = New-Object Windows.Forms.Padding(8, 0, 4, 0)
    Set-PrimaryButtonLook $btnNewDeal
    $purchaseToolbar.Controls.Add($btnNewDeal, 5, 0)

    $btnAddPurchaseSupplier = New-Object Windows.Forms.Button
    $btnAddPurchaseSupplier.Text = 'Добавить поставщика'
    $btnAddPurchaseSupplier.Dock = 'Fill'
    $btnAddPurchaseSupplier.Margin = New-Object Windows.Forms.Padding(4, 0, 4, 0)
    Set-SecondaryButtonLook $btnAddPurchaseSupplier
    $purchaseToolbar.Controls.Add($btnAddPurchaseSupplier, 6, 0)

    $btnImportRrfq = New-Object Windows.Forms.Button
    $btnImportRrfq.Text = 'Импорт из RRFQ'
    $btnImportRrfq.Dock = 'Fill'
    $btnImportRrfq.Margin = New-Object Windows.Forms.Padding(4, 0, 4, 0)
    Set-SecondaryButtonLook $btnImportRrfq
    $purchaseToolbar.Controls.Add($btnImportRrfq, 7, 0)

    $btnArchiveDeal = New-Object Windows.Forms.Button
    $btnArchiveDeal.Text = 'В архив'
    $btnArchiveDeal.Dock = 'Fill'
    $btnArchiveDeal.Margin = New-Object Windows.Forms.Padding(4, 0, 4, 0)
    Set-SecondaryButtonLook $btnArchiveDeal
    $purchaseToolbar.Controls.Add($btnArchiveDeal, 8, 0)

    $btnDeleteDeal = New-Object Windows.Forms.Button
    $btnDeleteDeal.Text = 'Удалить сделку'
    $btnDeleteDeal.Dock = 'Fill'
    $btnDeleteDeal.Margin = New-Object Windows.Forms.Padding(4, 0, 4, 0)
    Set-SecondaryButtonLook $btnDeleteDeal
    $purchaseToolbar.Controls.Add($btnDeleteDeal, 9, 0)

    $btnRefreshPurchase = New-Object Windows.Forms.Button
    $btnRefreshPurchase.Text = 'Обновить'
    $btnRefreshPurchase.Dock = 'Fill'
    $btnRefreshPurchase.Margin = New-Object Windows.Forms.Padding(4, 0, 0, 0)
    Set-SecondaryButtonLook $btnRefreshPurchase
    $purchaseToolbar.Controls.Add($btnRefreshPurchase, 10, 0)

    $purchaseContent = New-Object Windows.Forms.Panel
    $purchaseContent.Dock = 'Fill'
    $purchaseLayout.Controls.Add($purchaseContent, 0, 2)

    $purchaseMainSplit = New-Object Windows.Forms.SplitContainer
    $purchaseMainSplit.Dock = 'Fill'
    $purchaseMainSplit.Orientation = [Windows.Forms.Orientation]::Horizontal
    $purchaseMainSplit.SplitterWidth = 7
    $purchaseMainSplit.Panel1MinSize = 25
    $purchaseMainSplit.Panel2MinSize = 25
    $purchaseContent.Controls.Add($purchaseMainSplit)

    $purchaseRestLayout = New-Object Windows.Forms.TableLayoutPanel
    $purchaseRestLayout.Dock = 'Fill'
    $purchaseRestLayout.RowCount = 2
    $purchaseRestLayout.ColumnCount = 1
    Add-TableRowStyle $purchaseRestLayout 'Absolute' 58
    Add-TableRowStyle $purchaseRestLayout 'Percent' 100
    Add-TableColumnStyle $purchaseRestLayout 'Percent' 100
    $purchaseMainSplit.Panel2.Controls.Add($purchaseRestLayout)

    $purchaseLowerSplit = New-Object Windows.Forms.SplitContainer
    $purchaseLowerSplit.Dock = 'Fill'
    $purchaseLowerSplit.Orientation = [Windows.Forms.Orientation]::Horizontal
    $purchaseLowerSplit.SplitterWidth = 7
    $purchaseLowerSplit.Panel1MinSize = 25
    $purchaseLowerSplit.Panel2MinSize = 25
    $purchaseRestLayout.Controls.Add($purchaseLowerSplit, 0, 1)

    $dealsGrid = New-Object Windows.Forms.DataGridView
    $dealsGrid.Dock = 'Fill'
    $dealsGrid.AllowUserToAddRows = $false
    $dealsGrid.AllowUserToResizeRows = $false
    $dealsGrid.RowHeadersVisible = $false
    $dealsGrid.SelectionMode = 'FullRowSelect'
    $dealsGrid.MultiSelect = $false
    $dealsGrid.ReadOnly = $false
    $dealsGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $dealsGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both
    $dealsGrid.AllowUserToResizeColumns = $true
    $dealsGrid.EditMode = [Windows.Forms.DataGridViewEditMode]::EditOnKeystrokeOrF2
    Set-GridLook $dealsGrid
    $dealsGrid.DefaultCellStyle.SelectionBackColor = Get-UiColor 'SurfaceAlt'
    $dealsGrid.DefaultCellStyle.SelectionForeColor = Get-UiColor 'Text'
    $dealTrackingStatusValues = @('На отслеживании', 'Ожидание', 'Готово')
    foreach ($columnInfo in @(
        @('DealNumber', 'Сделка', 190),
        @('BoardCount', 'Кол-во плат', 92),
        @('Client', 'Клиент', 150),
        @('Priority', 'Приоритет', 92),
        @('Stage', 'Этап', 86),
        @('Suppliers', 'Поставщики', 92),
        @('Invoices', 'PI', 58),
        @('Payment', 'Оплачено', 86),
        @('Done', 'Готово', 74),
        @('Masks', 'Маски', 82),
        @('Comment', 'Комментарий', 130),
        @('CompletionReceipt', 'Поступление комплектации', 180),
        @('Updated', 'Обновлено', 142),
        @('Period', 'Период', 110),
        @('Executor', 'Исполнитель', 110),
        @('AssemblyLocation', 'Место сборки', 120)
    )) {
        if ($columnInfo[0] -eq 'DealStatus') {
            $col = New-Object Windows.Forms.DataGridViewComboBoxColumn
            $col.FlatStyle = [Windows.Forms.FlatStyle]::Flat
            foreach ($item in $dealTrackingStatusValues) { [void]$col.Items.Add($item) }
        } elseif ($columnInfo[0] -eq 'Masks') {
            $col = New-Object Windows.Forms.DataGridViewComboBoxColumn
            $col.FlatStyle = [Windows.Forms.FlatStyle]::Flat
            [void]$col.Items.Add('')
            [void]$col.Items.Add('Нет')
            # Convert-MasksToDbValue and Refresh-PurchaseDeals use Да/Нет for this boolean field.
            # Keeping the same display value in the ComboBox prevents archived rows with masks=1
            # from raising DataGridViewComboBoxCell's "invalid value" error.
            [void]$col.Items.Add('Да')
        } elseif ($columnInfo[0] -eq 'Priority') {
            $col = New-Object Windows.Forms.DataGridViewComboBoxColumn
            $col.FlatStyle = [Windows.Forms.FlatStyle]::Flat
            [void]$col.Items.Add('')
            foreach ($item in @('Срочно', '1', '2', '3', '4', '5')) { [void]$col.Items.Add($item) }
            $col.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
        } elseif ($columnInfo[0] -eq 'Executor') {
            $col = New-Object Windows.Forms.DataGridViewComboBoxColumn
            $col.FlatStyle = [Windows.Forms.FlatStyle]::Flat
            [void]$col.Items.Add('')
            [void]$col.Items.Add('Евгений')
            $col.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
        } elseif ($columnInfo[0] -eq 'AssemblyLocation') {
            $col = New-Object Windows.Forms.DataGridViewComboBoxColumn
            $col.FlatStyle = [Windows.Forms.FlatStyle]::Flat
            [void]$col.Items.Add('')
            [void]$col.Items.Add('Китай')
            [void]$col.Items.Add('РФ')
            $col.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
        } else {
            $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        }
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        if ($columnInfo[0] -ne 'DealNumber' -and $columnInfo[0] -ne 'BoardCount' -and $columnInfo[0] -ne 'Client' -and $columnInfo[0] -ne 'Priority' -and $columnInfo[0] -ne 'Period' -and $columnInfo[0] -ne 'Executor' -and $columnInfo[0] -ne 'AssemblyLocation' -and $columnInfo[0] -ne 'Masks') {
            $col.ReadOnly = $true
        }
        $dealsGrid.Columns.Add($col) | Out-Null
    }
    Restore-GridColumnWidths $dealsGrid 'ui.purchase.deals.column_widths'

    $purchaseMainSplit.Panel1.Controls.Add($dealsGrid)

    $dealsMenu = New-Object Windows.Forms.ContextMenuStrip
    $dealsCopyNumberMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $dealsCopyNumberMenuItem.Text = 'Копировать номер сделки'
    [void]$dealsMenu.Items.Add($dealsCopyNumberMenuItem)
    [void]$dealsMenu.Items.Add((New-Object Windows.Forms.ToolStripSeparator))
    $dealsArchiveMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $dealsArchiveMenuItem.Text = 'В архив / вернуть'
    [void]$dealsMenu.Items.Add($dealsArchiveMenuItem)
    $dealsDeleteMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $dealsDeleteMenuItem.Text = 'Удалить сделку'
    [void]$dealsMenu.Items.Add($dealsDeleteMenuItem)
    $dealsGrid.ContextMenuStrip = $dealsMenu

    $dealEditPanel = New-Object Windows.Forms.TableLayoutPanel
    $dealEditPanel.Dock = 'Fill'
    $dealEditPanel.Padding = New-Object Windows.Forms.Padding(0, 6, 0, 6)
    $dealEditPanel.ColumnCount = 10
    $dealEditPanel.RowCount = 1
    Add-TableColumnStyle $dealEditPanel 'Absolute' 64
    Add-TableColumnStyle $dealEditPanel 'Absolute' 180
    Add-TableColumnStyle $dealEditPanel 'Absolute' 58
    Add-TableColumnStyle $dealEditPanel 'Absolute' 120
    Add-TableColumnStyle $dealEditPanel 'Absolute' 62
    Add-TableColumnStyle $dealEditPanel 'Absolute' 120
    Add-TableColumnStyle $dealEditPanel 'Absolute' 95
    Add-TableColumnStyle $dealEditPanel 'Percent' 100
    Add-TableColumnStyle $dealEditPanel 'Absolute' 116
    Add-TableColumnStyle $dealEditPanel 'Absolute' 116
    Add-TableRowStyle $dealEditPanel 'Percent' 100
    $purchaseRestLayout.Controls.Add($dealEditPanel, 0, 0)

    $dealEditPanel.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'Клиент'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }), 0, 0)
    $txtDealClient = New-Object Windows.Forms.TextBox
    $txtDealClient.Dock = 'Fill'
    $txtDealClient.Margin = New-Object Windows.Forms.Padding(0, 9, 8, 9)
    $dealEditPanel.Controls.Add($txtDealClient, 1, 0)

    $dealEditPanel.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'Этап'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }), 2, 0)
    $cmbDealStatusEdit = New-Object Windows.Forms.ComboBox
    $cmbDealStatusEdit.Dock = 'Fill'
    $cmbDealStatusEdit.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbDealStatusEdit.Margin = New-Object Windows.Forms.Padding(0, 9, 8, 9)
    foreach ($status in @('RFQ', 'RRFQ', 'PO', 'PI', 'Закупка', 'Заказано', 'В работе', 'Ожидание')) {
        [void]$cmbDealStatusEdit.Items.Add($status)
    }
    $cmbDealStatusEdit.SelectedIndex = 0
    $dealEditPanel.Controls.Add($cmbDealStatusEdit, 3, 0)

    $dealEditPanel.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'Период'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }), 4, 0)
    $cmbDealPeriodEdit = New-Object Windows.Forms.ComboBox
    $cmbDealPeriodEdit.Dock = 'Fill'
    $cmbDealPeriodEdit.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDown
    $cmbDealPeriodEdit.Margin = New-Object Windows.Forms.Padding(0, 9, 8, 9)
    foreach ($period in @('', 'I квартал', 'II квартал', 'III квартал', 'IV квартал')) {
        [void]$cmbDealPeriodEdit.Items.Add($period)
    }
    $dealEditPanel.Controls.Add($cmbDealPeriodEdit, 5, 0)

    $dealEditPanel.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'Комментарий'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }), 6, 0)
    $txtDealComment = New-Object Windows.Forms.TextBox
    $txtDealComment.Dock = 'Fill'
    $txtDealComment.Margin = New-Object Windows.Forms.Padding(0, 9, 8, 9)
    $dealEditPanel.Controls.Add($txtDealComment, 7, 0)

    $btnSaveDealInfo = New-Object Windows.Forms.Button
    $btnSaveDealInfo.Text = 'Сохранить'
    $btnSaveDealInfo.Dock = 'Fill'
    $btnSaveDealInfo.Margin = New-Object Windows.Forms.Padding(0, 8, 8, 8)
    $btnSaveDealInfo.CausesValidation = $false
    Set-PrimaryButtonLook $btnSaveDealInfo
    $dealEditPanel.Controls.Add($btnSaveDealInfo, 8, 0)

    $btnExtractPiAmounts = New-Object Windows.Forms.Button
    $btnExtractPiAmounts.Text = 'Считать PI'
    $btnExtractPiAmounts.Dock = 'Fill'
    $btnExtractPiAmounts.Margin = New-Object Windows.Forms.Padding(0, 8, 0, 8)
    Set-SecondaryButtonLook $btnExtractPiAmounts
    $dealEditPanel.Controls.Add($btnExtractPiAmounts, 9, 0)

    $suppliersPanel = New-Object Windows.Forms.Panel
    $suppliersPanel.Dock = 'Fill'
    $purchaseLowerSplit.Panel1.Controls.Add($suppliersPanel)

    $suppliersGrid = New-Object Windows.Forms.DataGridView
    $suppliersGrid.Dock = 'Fill'
    $suppliersGrid.AllowUserToAddRows = $false
    $suppliersGrid.AllowUserToResizeRows = $false
    $suppliersGrid.RowHeadersVisible = $false
    $suppliersGrid.SelectionMode = 'FullRowSelect'
    $suppliersGrid.MultiSelect = $false
    $suppliersGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    Set-GridLook $suppliersGrid
    foreach ($columnInfo in @(
        @('Supplier', 'Поставщик', 130, 'Text'),
        @('PiUsd', 'Сумма PI $', 92, 'Text'),
        @('PiCny', 'Сумма PI ?', 92, 'Text'),
        @('PiRub', 'Сумма RUB', 92, 'Text'),
        @('PaidAmount', 'Оплачено', 92, 'Text'),
        @('DeliveryWeeks', 'Срок', 58, 'Text'),
        @('PaymentSubmitted', 'Подано в оплату', 95, 'Check'),
        @('Paid', 'Оплачено?', 76, 'Check'),
        @('InvoiceReceived', 'Инвойс', 62, 'Check'),
        @('InvoiceConfirmed', 'Инвойс подтв.', 98, 'Check'),
        @('ErpSupplier', 'ERP пост.', 76, 'Check'),
        @('ErpRoger', 'ERP Roger', 80, 'Check'),
        @('InvoiceDate', 'Дата подтв.', 90, 'Text'),
        @('ReceiptDate', 'План поступл.', 100, 'Text'),
        @('ActualReceiptDate', 'Факт поступл.', 100, 'Text'),
        @('Comment', 'Комментарий', 220, 'Text')
    )) {
        if ($columnInfo[3] -eq 'Check') {
            $col = New-Object Windows.Forms.DataGridViewCheckBoxColumn
        } else {
            $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        }
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        if ($columnInfo[0] -eq 'Supplier') { $col.ReadOnly = $true }
        if ($columnInfo[0] -eq 'InvoiceDate' -or $columnInfo[0] -eq 'ReceiptDate' -or $columnInfo[0] -eq 'ActualReceiptDate') { $col.ReadOnly = $true }
        $suppliersGrid.Columns.Add($col) | Out-Null
    }
    $suppliersPanel.Controls.Add($suppliersGrid)

    $suppliersMenu = New-Object Windows.Forms.ContextMenuStrip
    $suppliersDeleteMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $suppliersDeleteMenuItem.Text = 'Удалить поставщика'
    [void]$suppliersMenu.Items.Add($suppliersDeleteMenuItem)
    $suppliersGrid.ContextMenuStrip = $suppliersMenu

    $docsPanel = New-Object Windows.Forms.TableLayoutPanel
    $docsPanel.Dock = 'Fill'
    $docsPanel.RowCount = 2
    $docsPanel.ColumnCount = 1
    $docsPanel.AllowDrop = $true
    $docsPanel.Padding = New-Object Windows.Forms.Padding(8)
    $docsPanel.BackColor = Get-UiColor 'Surface'
    $docsPanel.CellBorderStyle = [Windows.Forms.TableLayoutPanelCellBorderStyle]::Single
    Add-TableRowStyle $docsPanel 'Absolute' 34
    Add-TableRowStyle $docsPanel 'Percent' 100
    $purchaseLowerSplit.Panel2.Controls.Add($docsPanel)

    $docsHeaderPanel = New-Object Windows.Forms.TableLayoutPanel
    $docsHeaderPanel.Dock = 'Fill'
    $docsHeaderPanel.RowCount = 1
    $docsHeaderPanel.ColumnCount = 2
    $docsHeaderPanel.AllowDrop = $true
    $docsHeaderPanel.BackColor = Get-UiColor 'Surface'
    Add-TableColumnStyle $docsHeaderPanel 'Percent' 100
    Add-TableColumnStyle $docsHeaderPanel 'Absolute' 136
    Add-TableRowStyle $docsHeaderPanel 'Percent' 100
    $docsPanel.Controls.Add($docsHeaderPanel, 0, 0)

    $docsLabel = New-Object Windows.Forms.Label
    $docsLabel.Text = 'Документы выбранной сделки/поставщика: перетащите или выберите файлы'
    $docsLabel.Dock = 'Fill'
    $docsLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $docsLabel.AllowDrop = $true
    $docsLabel.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
    $docsLabel.ForeColor = Get-UiColor 'Text'
    $docsLabel.BackColor = Get-UiColor 'Surface'
    $docsHeaderPanel.Controls.Add($docsLabel, 0, 0)

    $btnPickPurchaseDocuments = New-Object Windows.Forms.Button
    $btnPickPurchaseDocuments.Text = 'Выбрать файлы'
    $btnPickPurchaseDocuments.Dock = 'Fill'
    $btnPickPurchaseDocuments.Margin = New-Object Windows.Forms.Padding(4, 2, 0, 2)
    Set-SecondaryButtonLook $btnPickPurchaseDocuments
    $docsHeaderPanel.Controls.Add($btnPickPurchaseDocuments, 1, 0)

    $docsGrid = New-Object Windows.Forms.DataGridView
    $docsGrid.Dock = 'Fill'
    $docsGrid.AllowUserToAddRows = $false
    $docsGrid.AllowUserToResizeRows = $false
    $docsGrid.RowHeadersVisible = $false
    $docsGrid.SelectionMode = 'FullRowSelect'
    $docsGrid.MultiSelect = $false
    $docsGrid.ReadOnly = $true
    $docsGrid.AllowDrop = $true
    Set-GridLook $docsGrid
    foreach ($columnInfo in @(
        @('Type', 'Тип', 110),
        @('Supplier', 'Поставщик', 120),
        @('File', 'Файл', 260),
        @('Uploaded', 'Загружен', 140),
        @('Path', 'Путь', 360)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        $docsGrid.Columns.Add($col) | Out-Null
    }
    $docsPanel.Controls.Add($docsGrid, 0, 1)

    $docsMenu = New-Object Windows.Forms.ContextMenuStrip
    $docsOpenExplorerMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $docsOpenExplorerMenuItem.Text = 'Открыть в проводнике'
    [void]$docsMenu.Items.Add($docsOpenExplorerMenuItem)
    $docsLoadQuotesMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $docsLoadQuotesMenuItem.Text = 'Загрузить квоты'
    [void]$docsMenu.Items.Add($docsLoadQuotesMenuItem)
    [void]$docsMenu.Items.Add((New-Object Windows.Forms.ToolStripSeparator))
    $docsDeleteMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $docsDeleteMenuItem.Text = 'Удалить'
    [void]$docsMenu.Items.Add($docsDeleteMenuItem)
    $docsGrid.ContextMenuStrip = $docsMenu

    $componentsLayout = New-Object Windows.Forms.TableLayoutPanel
    $componentsLayout.Dock = 'Fill'
    $componentsLayout.Padding = New-Object Windows.Forms.Padding(12)
    $componentsLayout.RowCount = 2
    $componentsLayout.ColumnCount = 1
    Add-TableRowStyle $componentsLayout 'Absolute' 46
    Add-TableRowStyle $componentsLayout 'Percent' 100
    Add-TableColumnStyle $componentsLayout 'Percent' 100
    $componentsPage.Controls.Add($componentsLayout)

    $componentsToolbar = New-Object Windows.Forms.FlowLayoutPanel
    $componentsToolbar.Dock = 'Fill'
    $componentsToolbar.FlowDirection = 'LeftToRight'
    $componentsToolbar.WrapContents = $false
    $componentsToolbar.Padding = New-Object Windows.Forms.Padding(0, 5, 0, 5)
    $componentsLayout.Controls.Add($componentsToolbar, 0, 0)

    $btnNewComponent = New-Object Windows.Forms.Button
    $btnNewComponent.Text = 'Новая задача'
    $btnNewComponent.Width = 130
    $btnNewComponent.Height = 30
    Set-PrimaryButtonLook $btnNewComponent
    $componentsToolbar.Controls.Add($btnNewComponent)

    $btnDeleteComponent = New-Object Windows.Forms.Button
    $btnDeleteComponent.Text = 'Удалить'
    $btnDeleteComponent.Width = 100
    $btnDeleteComponent.Height = 30
    Set-SecondaryButtonLook $btnDeleteComponent
    $componentsToolbar.Controls.Add($btnDeleteComponent)

    $btnRefreshComponents = New-Object Windows.Forms.Button
    $btnRefreshComponents.Text = 'Обновить'
    $btnRefreshComponents.Width = 100
    $btnRefreshComponents.Height = 30
    Set-SecondaryButtonLook $btnRefreshComponents
    $componentsToolbar.Controls.Add($btnRefreshComponents)

    $componentsSearchLabel = New-Object Windows.Forms.Label
    $componentsSearchLabel.Text = 'Поиск сделки:'
    $componentsSearchLabel.Width = 92
    $componentsSearchLabel.Height = 30
    $componentsSearchLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $componentsToolbar.Controls.Add($componentsSearchLabel)

    $txtComponentsSearch = New-Object Windows.Forms.TextBox
    $txtComponentsSearch.Width = 220
    $txtComponentsSearch.Height = 30
    $txtComponentsSearch.Margin = New-Object Windows.Forms.Padding(0, 5, 8, 5)
    Set-InputLook $txtComponentsSearch
    $componentsToolbar.Controls.Add($txtComponentsSearch)

    $componentsHint = New-Object Windows.Forms.Label
    $componentsHint.Text = 'Сортировка: нажмите заголовок "Статус" или "Приоритет".'
    $componentsHint.Width = 420
    $componentsHint.Height = 30
    $componentsHint.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $componentsHint.ForeColor = Get-UiColor 'Muted'
    $componentsToolbar.Controls.Add($componentsHint)

    $componentsSplit = New-Object Windows.Forms.SplitContainer
    $componentsSplit.Dock = 'Fill'
    $componentsSplit.Orientation = [Windows.Forms.Orientation]::Horizontal
    $componentsSplit.SplitterWidth = 6
    $componentsSplit.Panel1MinSize = 180
    $componentsSplit.Panel2MinSize = 110
    $componentsLayout.Controls.Add($componentsSplit, 0, 1)
    $componentsGrid = New-Object Windows.Forms.DataGridView
    $componentsGrid.Dock = 'Fill'
    $componentsGrid.AllowUserToAddRows = $false
    $componentsGrid.AllowUserToResizeRows = $false
    $componentsGrid.AllowUserToResizeColumns = $true
    $componentsGrid.RowHeadersVisible = $false
    $componentsGrid.SelectionMode = 'FullRowSelect'
    $componentsGrid.MultiSelect = $false
    $componentsGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $componentsGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both
    $componentsGrid.EditMode = [Windows.Forms.DataGridViewEditMode]::EditOnKeystrokeOrF2
    Set-GridLook $componentsGrid
    $componentsGrid.DefaultCellStyle.SelectionBackColor = Get-UiColor 'SurfaceAlt'
    $componentsGrid.DefaultCellStyle.SelectionForeColor = Get-UiColor 'Text'
    $componentsSplit.Panel1.Controls.Add($componentsGrid)

    $componentNotesPanel = New-Object Windows.Forms.TableLayoutPanel
    $componentNotesPanel.Dock = 'Fill'
    $componentNotesPanel.RowCount = 2
    $componentNotesPanel.ColumnCount = 1
    $componentNotesPanel.Padding = New-Object Windows.Forms.Padding(0, 6, 0, 0)
    Add-TableRowStyle $componentNotesPanel 'Absolute' 36
    Add-TableRowStyle $componentNotesPanel 'Percent' 100
    Add-TableColumnStyle $componentNotesPanel 'Percent' 100
    $componentsSplit.Panel2.Controls.Add($componentNotesPanel)

    $componentNotesHeader = New-Object Windows.Forms.TableLayoutPanel
    $componentNotesHeader.Dock = 'Fill'
    $componentNotesHeader.ColumnCount = 2
    $componentNotesHeader.RowCount = 1
    Add-TableColumnStyle $componentNotesHeader 'Percent' 100
    Add-TableColumnStyle $componentNotesHeader 'Absolute' 150
    Add-TableRowStyle $componentNotesHeader 'Percent' 100
    $componentNotesPanel.Controls.Add($componentNotesHeader, 0, 0)

    $lblComponentNotesTitle = New-Object Windows.Forms.Label
    $lblComponentNotesTitle.Text = 'Записи по сделке'
    $lblComponentNotesTitle.Dock = 'Fill'
    $lblComponentNotesTitle.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $lblComponentNotesTitle.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
    $lblComponentNotesTitle.ForeColor = Get-UiColor 'Text'
    $componentNotesHeader.Controls.Add($lblComponentNotesTitle, 0, 0)

    $btnSaveComponentNotes = New-Object Windows.Forms.Button
    $btnSaveComponentNotes.Text = 'Сохранить записи'
    $btnSaveComponentNotes.Dock = 'Fill'
    $btnSaveComponentNotes.Enabled = $false
    Set-PrimaryButtonLook $btnSaveComponentNotes
    $componentNotesHeader.Controls.Add($btnSaveComponentNotes, 1, 0)

    $txtComponentNotes = New-Object Windows.Forms.TextBox
    $txtComponentNotes.Dock = 'Fill'
    $txtComponentNotes.Multiline = $true
    $txtComponentNotes.AcceptsReturn = $true
    $txtComponentNotes.AcceptsTab = $true
    $txtComponentNotes.ScrollBars = [Windows.Forms.ScrollBars]::Vertical
    $txtComponentNotes.Font = New-Object Drawing.Font('Segoe UI', 10)
    $txtComponentNotes.Enabled = $false
    $componentNotesPanel.Controls.Add($txtComponentNotes, 0, 1)
    $componentsMenu = New-Object Windows.Forms.ContextMenuStrip
    $componentsCopyDealMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $componentsCopyDealMenuItem.Text = 'Копировать номер сделки'
    [void]$componentsMenu.Items.Add($componentsCopyDealMenuItem)
    $componentsGrid.ContextMenuStrip = $componentsMenu
    $componentStatusValues = @('В работе', 'Ожидание ответа', 'На отслеживании', 'Заказано', 'Подано в оплату', 'Выполнено', 'Не актуально')
    $componentStageValues = @('', 'Запросил поставщиков', 'RRFQ отправлено', 'PI отправлен', 'Контроль оплаты', 'Отправлено в РФ')
    $componentPriorityValues = @('Срочно', '1', '2', '3', '4', '5')
    $componentPeriodValues = @('', 'I квартал', 'II квартал', 'III квартал', 'IV квартал')

    $componentEntryCol = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $componentEntryCol.Name = 'EntryDate'
    $componentEntryCol.HeaderText = 'Дата входа'
    $componentEntryCol.Width = 96
    $componentEntryCol.ReadOnly = $true
    $componentEntryCol.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $componentsGrid.Columns.Add($componentEntryCol) | Out-Null

    $componentDealCol = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $componentDealCol.Name = 'DealNumber'
    $componentDealCol.HeaderText = 'Сделка'
    $componentDealCol.Width = 170
    $componentDealCol.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $componentsGrid.Columns.Add($componentDealCol) | Out-Null

    $componentStatusCol = New-Object Windows.Forms.DataGridViewComboBoxColumn
    $componentStatusCol.Name = 'Status'
    $componentStatusCol.HeaderText = 'Статус'
    $componentStatusCol.Width = 145
    $componentStatusCol.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    foreach ($item in $componentStatusValues) { [void]$componentStatusCol.Items.Add($item) }
    $componentStatusCol.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $componentsGrid.Columns.Add($componentStatusCol) | Out-Null

    $componentStageCol = New-Object Windows.Forms.DataGridViewComboBoxColumn
    $componentStageCol.Name = 'Stage'
    $componentStageCol.HeaderText = 'Этап'
    $componentStageCol.Width = 170
    $componentStageCol.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    foreach ($item in $componentStageValues) { [void]$componentStageCol.Items.Add($item) }
    $componentStageCol.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $componentsGrid.Columns.Add($componentStageCol) | Out-Null

    foreach ($columnInfo in @(
        @('Description', 'Описание', 260),
        @('NextAction', 'След.действие', 230)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        $col.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
        $componentsGrid.Columns.Add($col) | Out-Null
    }

    foreach ($columnInfo in @(
        @('ReminderDate', 'Напоминание', 112),
        @('DeadlineDate', 'Дедлайн', 112)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        $col.ReadOnly = $true
        $col.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
        $componentsGrid.Columns.Add($col) | Out-Null
    }

    $componentPriorityCol = New-Object Windows.Forms.DataGridViewComboBoxColumn
    $componentPriorityCol.Name = 'Priority'
    $componentPriorityCol.HeaderText = 'Приоритет'
    $componentPriorityCol.Width = 92
    $componentPriorityCol.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    foreach ($item in $componentPriorityValues) { [void]$componentPriorityCol.Items.Add($item) }
    $componentPriorityCol.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $componentsGrid.Columns.Add($componentPriorityCol) | Out-Null

    $componentPeriodCol = New-Object Windows.Forms.DataGridViewComboBoxColumn
    $componentPeriodCol.Name = 'Period'
    $componentPeriodCol.HeaderText = 'Период'
    $componentPeriodCol.Width = 110
    $componentPeriodCol.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    foreach ($item in $componentPeriodValues) { [void]$componentPeriodCol.Items.Add($item) }
    $componentPeriodCol.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $componentsGrid.Columns.Add($componentPeriodCol) | Out-Null

    Restore-GridColumnWidths $componentsGrid 'ui.components.column_widths'

    $remindersLayout = New-Object Windows.Forms.TableLayoutPanel
    $remindersLayout.Dock = 'Fill'
    $remindersLayout.Padding = New-Object Windows.Forms.Padding(12)
    $remindersLayout.RowCount = 2
    $remindersLayout.ColumnCount = 1
    Add-TableRowStyle $remindersLayout 'Absolute' 46
    Add-TableRowStyle $remindersLayout 'Percent' 100
    $remindersPage.Controls.Add($remindersLayout)

    $remindersToolbar = New-Object Windows.Forms.FlowLayoutPanel
    $remindersToolbar.Dock = 'Fill'
    $remindersToolbar.FlowDirection = 'LeftToRight'
    $remindersToolbar.WrapContents = $false
    $remindersToolbar.Padding = New-Object Windows.Forms.Padding(0, 5, 0, 5)
    $remindersLayout.Controls.Add($remindersToolbar, 0, 0)

    $btnNewReminder = New-Object Windows.Forms.Button
    $btnNewReminder.Text = 'Новое напоминание'
    $btnNewReminder.Width = 150
    $btnNewReminder.Height = 30
    Set-PrimaryButtonLook $btnNewReminder
    $remindersToolbar.Controls.Add($btnNewReminder)

    $btnDoneReminder = New-Object Windows.Forms.Button
    $btnDoneReminder.Text = 'Готово'
    $btnDoneReminder.Width = 90
    $btnDoneReminder.Height = 30
    Set-SecondaryButtonLook $btnDoneReminder
    $remindersToolbar.Controls.Add($btnDoneReminder)
    $btnReceiptArrived = New-Object Windows.Forms.Button
    $btnReceiptArrived.Text = 'Поступило'
    $btnReceiptArrived.Width = 100
    $btnReceiptArrived.Height = 30
    Set-SecondaryButtonLook $btnReceiptArrived
    $remindersToolbar.Controls.Add($btnReceiptArrived)

    $btnRefreshReminders = New-Object Windows.Forms.Button
    $btnRefreshReminders.Text = 'Обновить'
    $btnRefreshReminders.Width = 100
    $btnRefreshReminders.Height = 30
    Set-SecondaryButtonLook $btnRefreshReminders
    $remindersToolbar.Controls.Add($btnRefreshReminders)

    $chkNotificationsEnabled = New-Object Windows.Forms.CheckBox
    $chkNotificationsEnabled.Text = 'Уведомления включены'
    $chkNotificationsEnabled.AutoSize = $true
    $chkNotificationsEnabled.Height = 30
    $chkNotificationsEnabled.Margin = New-Object Windows.Forms.Padding(10, 5, 8, 0)
    $chkNotificationsEnabled.Checked = [bool]$script:NotificationsEnabled
    $remindersToolbar.Controls.Add($chkNotificationsEnabled)

    $remindersSearchLabel = New-Object Windows.Forms.Label
    $remindersSearchLabel.Text = 'Поиск по сделке:'
    $remindersSearchLabel.Width = 112
    $remindersSearchLabel.Height = 30
    $remindersSearchLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $remindersToolbar.Controls.Add($remindersSearchLabel)

    $txtRemindersSearch = New-Object Windows.Forms.TextBox
    $txtRemindersSearch.Width = 220
    $txtRemindersSearch.Height = 30
    $txtRemindersSearch.Margin = New-Object Windows.Forms.Padding(0, 2, 10, 2)
    $remindersToolbar.Controls.Add($txtRemindersSearch)

    $remindersHint = New-Object Windows.Forms.Label
    $remindersHint.Text = 'Автоматические пункты строятся из статусов закупки; ручные можно закрывать кнопкой "Готово". Для "скоро поступление" и "поступление просрочено" используйте "Поступило".'
    $remindersHint.Width = 720
    $remindersHint.Height = 30
    $remindersHint.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $remindersHint.ForeColor = Get-UiColor 'Muted'
    $remindersToolbar.Controls.Add($remindersHint)

    $remindersGrid = New-Object Windows.Forms.DataGridView
    $remindersGrid.Dock = 'Fill'
    $remindersGrid.AllowUserToAddRows = $false
    $remindersGrid.AllowUserToResizeRows = $false
    $remindersGrid.AllowUserToResizeColumns = $true
    $remindersGrid.RowHeadersVisible = $false
    $remindersGrid.SelectionMode = 'FullRowSelect'
    $remindersGrid.MultiSelect = $false
    $remindersGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $remindersGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both
    Set-GridLook $remindersGrid
    foreach ($columnInfo in @(
        @('Severity', 'Важность', 90),
        @('DueDate', 'Дата', 95),
        @('Title', 'Что сделать', 560),
        @('Source', 'Источник', 100)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        $col.ReadOnly = $true
        $remindersGrid.Columns.Add($col) | Out-Null
    }

    $remindersMenu = New-Object Windows.Forms.ContextMenuStrip
    $remindersDeleteMenuItem = New-Object Windows.Forms.ToolStripMenuItem
    $remindersDeleteMenuItem.Text = 'Удалить напоминание'
    [void]$remindersMenu.Items.Add($remindersDeleteMenuItem)
    $remindersGrid.ContextMenuStrip = $remindersMenu
    $remindersGrid.Add_CellMouseDown({
        if ($_.Button -eq [Windows.Forms.MouseButtons]::Right -and $_.RowIndex -ge 0) {
            $remindersGrid.ClearSelection()
            $remindersGrid.Rows[$_.RowIndex].Selected = $true
            if ($_.ColumnIndex -ge 0) {
                $remindersGrid.CurrentCell = $remindersGrid.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
            }
        }
    })
    $remindersLayout.Controls.Add($remindersGrid, 0, 1)

    $quoteBaseLayout = New-Object Windows.Forms.TableLayoutPanel
    $quoteBaseLayout.Dock = 'Fill'
    $quoteBaseLayout.Padding = New-Object Windows.Forms.Padding(12)
    $quoteBaseLayout.RowCount = 2
    $quoteBaseLayout.ColumnCount = 1
    Add-TableRowStyle $quoteBaseLayout 'Absolute' 46
    Add-TableRowStyle $quoteBaseLayout 'Percent' 100
    $quoteBasePage.Controls.Add($quoteBaseLayout)

    $quoteBaseToolbar = New-Object Windows.Forms.TableLayoutPanel
    $quoteBaseToolbar.Dock = 'Fill'
    $quoteBaseToolbar.ColumnCount = 11
    $quoteBaseToolbar.RowCount = 1
    Add-TableColumnStyle $quoteBaseToolbar 'Absolute' 58
    Add-TableColumnStyle $quoteBaseToolbar 'Percent' 100
    Add-TableColumnStyle $quoteBaseToolbar 'Absolute' 72
    Add-TableColumnStyle $quoteBaseToolbar 'Absolute' 118
    Add-TableColumnStyle $quoteBaseToolbar 'Absolute' 34
    Add-TableColumnStyle $quoteBaseToolbar 'Absolute' 118
    Add-TableColumnStyle $quoteBaseToolbar 'Absolute' 100
    Add-TableColumnStyle $quoteBaseToolbar 'Absolute' 120
    Add-TableColumnStyle $quoteBaseToolbar 'Absolute' 150
    Add-TableColumnStyle $quoteBaseToolbar 'Absolute' 120
    Add-TableColumnStyle $quoteBaseToolbar 'Absolute' 145
    Add-TableRowStyle $quoteBaseToolbar 'Percent' 100
    $quoteBaseLayout.Controls.Add($quoteBaseToolbar, 0, 0)

    $quoteBaseToolbar.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'Поиск:'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }), 0, 0)
    $txtQuoteSearch = New-Object Windows.Forms.TextBox
    $txtQuoteSearch.Dock = 'Fill'
    $txtQuoteSearch.Margin = New-Object Windows.Forms.Padding(0, 8, 8, 8)
    $quoteBaseToolbar.Controls.Add($txtQuoteSearch, 1, 0)
    $quoteBaseToolbar.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'С даты:'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }), 2, 0)
    $txtQuoteDateFrom = New-Object Windows.Forms.TextBox
    $txtQuoteDateFrom.Dock = 'Fill'
    $txtQuoteDateFrom.Margin = New-Object Windows.Forms.Padding(0, 8, 8, 8)
    $quoteBaseToolbar.Controls.Add($txtQuoteDateFrom, 3, 0)
    $quoteBaseToolbar.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'по'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleCenter }), 4, 0)
    $txtQuoteDateTo = New-Object Windows.Forms.TextBox
    $txtQuoteDateTo.Dock = 'Fill'
    $txtQuoteDateTo.Margin = New-Object Windows.Forms.Padding(0, 8, 8, 8)
    $quoteBaseToolbar.Controls.Add($txtQuoteDateTo, 5, 0)
    $btnRefreshQuoteBase = New-Object Windows.Forms.Button
    $btnRefreshQuoteBase.Text = 'Обновить'
    $btnRefreshQuoteBase.Dock = 'Fill'
    $btnRefreshQuoteBase.Margin = New-Object Windows.Forms.Padding(0, 8, 8, 8)
    Set-SecondaryButtonLook $btnRefreshQuoteBase
    $quoteBaseToolbar.Controls.Add($btnRefreshQuoteBase, 6, 0)
    $btnExportQuoteBase = New-Object Windows.Forms.Button
    $btnExportQuoteBase.Text = 'Экспорт Excel'
    $btnExportQuoteBase.Dock = 'Fill'
    $btnExportQuoteBase.Margin = New-Object Windows.Forms.Padding(0, 8, 0, 8)
    Set-SecondaryButtonLook $btnExportQuoteBase
    $quoteBaseToolbar.Controls.Add($btnExportQuoteBase, 7, 0)
    $btnDeleteQuoteRows = New-Object Windows.Forms.Button
    $btnDeleteQuoteRows.Text = 'Удалить выделенные'
    $btnDeleteQuoteRows.Dock = 'Fill'
    $btnDeleteQuoteRows.Margin = New-Object Windows.Forms.Padding(8, 8, 8, 8)
    Set-SecondaryButtonLook $btnDeleteQuoteRows
    $quoteBaseToolbar.Controls.Add($btnDeleteQuoteRows, 8, 0)
    $btnClearQuoteBase = New-Object Windows.Forms.Button
    $btnClearQuoteBase.Text = 'Очистить'
    $btnClearQuoteBase.Dock = 'Fill'
    $btnClearQuoteBase.Margin = New-Object Windows.Forms.Padding(0, 8, 0, 8)
    Set-SecondaryButtonLook $btnClearQuoteBase
    $quoteBaseToolbar.Controls.Add($btnClearQuoteBase, 9, 0)
    $btnImportGlobalist = New-Object Windows.Forms.Button
    $btnImportGlobalist.Text = 'Загрузить Globalist'
    $btnImportGlobalist.Dock = 'Fill'
    $btnImportGlobalist.Margin = New-Object Windows.Forms.Padding(4, 6, 0, 6)
    $quoteBaseToolbar.Controls.Add($btnImportGlobalist, 10, 0)

    $globalistGrid = New-Object Windows.Forms.DataGridView
    $globalistGrid.Dock = 'Fill'
    $globalistGrid.AllowUserToAddRows = $false
    $globalistGrid.AllowUserToResizeRows = $false
    $globalistGrid.AllowUserToResizeColumns = $true
    $globalistGrid.RowHeadersVisible = $false
    $globalistGrid.SelectionMode = 'FullRowSelect'
    $globalistGrid.MultiSelect = $true
    $globalistGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $globalistGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both
    Set-GridLook $globalistGrid
    foreach ($columnInfo in @(
        @('ImportedAt', 'Загружено', 135), @('Factory', 'Источник', 95), @('PN', 'PN', 180), @('Comment', 'Комментарий', 240),
        @('PiNumber', 'PI номер', 130), @('Replacement', 'Замена', 170), @('ChineseRemark', 'Кит. примечание', 180), @('Package', 'Корпус', 120),
        @('Brand', 'Бренд', 120), @('Datacode', 'DC', 100), @('Moq', 'MOQ', 90), @('Qty', 'Количество', 100), @('Stock', 'На складе', 100),
        @('NeedSpq', 'Нужен SPQ', 100), @('Spq', 'SPQ', 90), @('UnitPrice', 'Цена', 100), @('TotalAmount', 'Сумма', 110),
        @('LeadTime', 'Срок, нед.', 110), @('Weight', 'Вес, г', 90), @('Target', 'Target', 110), @('SupplierQuoteId', 'ID квоты', 150), @('SheetName', 'Лист', 140)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = [string]$columnInfo[0]; $col.HeaderText = [string]$columnInfo[1]; $col.Width = [int]$columnInfo[2]; $col.ReadOnly = $true
        [void]$globalistGrid.Columns.Add($col)
    }
    $quoteBaseGrid = New-Object Windows.Forms.DataGridView
    $quoteBaseGrid.Dock = 'Fill'
    $quoteBaseGrid.AllowUserToAddRows = $false
    $quoteBaseGrid.AllowUserToResizeRows = $false
    $quoteBaseGrid.AllowUserToResizeColumns = $true
    $quoteBaseGrid.RowHeadersVisible = $false
    $quoteBaseGrid.SelectionMode = 'FullRowSelect'
    $quoteBaseGrid.MultiSelect = $true
    $quoteBaseGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $quoteBaseGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both
    Set-GridLook $quoteBaseGrid
    foreach ($columnInfo in @(
        @('QuoteDate', 'Дата квоты', 145),
        @('Value', 'Value', 260),
        @('PN', 'PN', 180),
        @('Supplier', 'Поставщик', 120),
        @('Price', 'Цена', 90),
        @('Lead', 'Срок', 100),
        @('Mfg', 'MFG', 140),
        @('Winner', 'Побед.', 62),
        @('Reason', 'Почему выбран', 170),
        @('Warning', 'Предупреждение', 280)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        $col.ReadOnly = $true
        $quoteBaseGrid.Columns.Add($col) | Out-Null
    }
    $quoteBaseTabs = New-Object Windows.Forms.TabControl
    $quoteBaseTabs.Dock = 'Fill'
    $quoteBaseOurTab = New-Object Windows.Forms.TabPage
    $quoteBaseOurTab.Text = 'Наши квоты'
    $quoteBaseOurTab.Padding = New-Object Windows.Forms.Padding(3)
    $quoteBaseOurTab.Controls.Add($quoteBaseGrid)
    $quoteBaseGlobalistTab = New-Object Windows.Forms.TabPage
    $quoteBaseGlobalistTab.Text = 'Globalist'
    $quoteBaseGlobalistTab.Padding = New-Object Windows.Forms.Padding(3)
    $quoteBaseGlobalistTab.Controls.Add($globalistGrid)
    [void]$quoteBaseTabs.TabPages.Add($quoteBaseOurTab)
    [void]$quoteBaseTabs.TabPages.Add($quoteBaseGlobalistTab)
    $quoteBaseLayout.Controls.Add($quoteBaseTabs, 0, 1)

    $historyLayout = New-Object Windows.Forms.TableLayoutPanel
    $historyLayout.Dock = 'Fill'
    $historyLayout.Padding = New-Object Windows.Forms.Padding(12)
    $historyLayout.RowCount = 2
    $historyLayout.ColumnCount = 1
    Add-TableRowStyle $historyLayout 'Absolute' 46
    Add-TableRowStyle $historyLayout 'Percent' 100
    $historyPage.Controls.Add($historyLayout)

    $historyToolbar = New-Object Windows.Forms.FlowLayoutPanel
    $historyToolbar.Dock = 'Fill'
    $historyToolbar.FlowDirection = 'LeftToRight'
    $historyToolbar.WrapContents = $false
    $historyToolbar.Padding = New-Object Windows.Forms.Padding(0, 5, 0, 5)
    $historyLayout.Controls.Add($historyToolbar, 0, 0)
    $btnRefreshHistory = New-Object Windows.Forms.Button
    $btnRefreshHistory.Text = 'Обновить'
    $btnRefreshHistory.Width = 100
    $btnRefreshHistory.Height = 30
    Set-SecondaryButtonLook $btnRefreshHistory
    $historyToolbar.Controls.Add($btnRefreshHistory)
    $btnHistorySelectedDeal = New-Object Windows.Forms.Button
    $btnHistorySelectedDeal.Text = 'Текущая сделка'
    $btnHistorySelectedDeal.Width = 130
    $btnHistorySelectedDeal.Height = 30
    Set-SecondaryButtonLook $btnHistorySelectedDeal
    $historyToolbar.Controls.Add($btnHistorySelectedDeal)
    $historySearchLabel = New-Object Windows.Forms.Label
    $historySearchLabel.Text = 'Поиск:'
    $historySearchLabel.AutoSize = $true
    $historySearchLabel.Margin = New-Object Windows.Forms.Padding(12, 8, 4, 0)
    $historyToolbar.Controls.Add($historySearchLabel)
    $txtHistorySearch = New-Object Windows.Forms.TextBox
    $txtHistorySearch.Width = 220
    $txtHistorySearch.Height = 28
    $txtHistorySearch.Margin = New-Object Windows.Forms.Padding(0, 3, 8, 0)
    $historyToolbar.Controls.Add($txtHistorySearch)
    $btnClearHistory = New-Object Windows.Forms.Button
    $btnClearHistory.Text = 'Очистить историю'
    $btnClearHistory.Width = 140
    $btnClearHistory.Height = 30
    Set-SecondaryButtonLook $btnClearHistory
    $historyToolbar.Controls.Add($btnClearHistory)
    $historyHint = New-Object Windows.Forms.Label
    $historyHint.Text = 'История фиксирует основные действия: сделки, поставщиков, документы, импорт RRFQ и изменения статусов.'
    $historyHint.Width = 760
    $historyHint.Height = 30
    $historyHint.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $historyHint.ForeColor = Get-UiColor 'Muted'
    $historyToolbar.Controls.Add($historyHint)

    $historyGrid = New-Object Windows.Forms.DataGridView
    $historyGrid.Dock = 'Fill'
    $historyGrid.AllowUserToAddRows = $false
    $historyGrid.AllowUserToResizeRows = $false
    $historyGrid.AllowUserToResizeColumns = $true
    $historyGrid.RowHeadersVisible = $false
    $historyGrid.SelectionMode = 'FullRowSelect'
    $historyGrid.MultiSelect = $false
    $historyGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $historyGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both
    Set-GridLook $historyGrid
    foreach ($columnInfo in @(
        @('Created', 'Дата', 145),
        @('Action', 'Действие', 210),
        @('Details', 'Детали', 520),
        @('Entity', 'Объект', 120)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        $col.ReadOnly = $true
        $historyGrid.Columns.Add($col) | Out-Null
    }
    $historyLayout.Controls.Add($historyGrid, 0, 1)

    $compelLayout = New-Object Windows.Forms.TableLayoutPanel
    $compelLayout.Dock = 'Fill'
    $compelLayout.Padding = New-Object Windows.Forms.Padding(12)
    $compelLayout.RowCount = 4
    $compelLayout.ColumnCount = 1
    Add-TableRowStyle $compelLayout 'Absolute' 46
    Add-TableRowStyle $compelLayout 'Absolute' 78
    Add-TableRowStyle $compelLayout 'Absolute' 30
    Add-TableRowStyle $compelLayout 'Percent' 100
    $compelParserPage.Controls.Add($compelLayout)

    $compelToolbar = New-Object Windows.Forms.TableLayoutPanel
    $compelToolbar.Dock = 'Fill'
    $compelToolbar.ColumnCount = 7
    $compelToolbar.RowCount = 1
    $compelToolbar.Padding = New-Object Windows.Forms.Padding(0, 5, 0, 5)
    Add-TableColumnStyle $compelToolbar 'Absolute' 48
    Add-TableColumnStyle $compelToolbar 'Percent' 100
    Add-TableColumnStyle $compelToolbar 'Absolute' 130
    Add-TableColumnStyle $compelToolbar 'Absolute' 110
    Add-TableColumnStyle $compelToolbar 'Absolute' 135
    Add-TableColumnStyle $compelToolbar 'Absolute' 125
    Add-TableColumnStyle $compelToolbar 'Absolute' 160
    $compelLayout.Controls.Add($compelToolbar, 0, 0)

    $compelToolbar.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'HTML:'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }), 0, 0)

    $txtCompelHtml = New-Object Windows.Forms.TextBox
    $txtCompelHtml.Dock = 'Fill'
    $txtCompelHtml.ReadOnly = $true
    Set-InputLook $txtCompelHtml
    $compelToolbar.Controls.Add($txtCompelHtml, 1, 0)

    $btnCompelChoose = New-Object Windows.Forms.Button
    $btnCompelChoose.Text = 'Выбрать HTML'
    $btnCompelChoose.Dock = 'Fill'
    Set-SecondaryButtonLook $btnCompelChoose
    $compelToolbar.Controls.Add($btnCompelChoose, 2, 0)

    $btnCompelParse = New-Object Windows.Forms.Button
    $btnCompelParse.Text = 'Разобрать'
    $btnCompelParse.Dock = 'Fill'
    Set-PrimaryButtonLook $btnCompelParse
    $compelToolbar.Controls.Add($btnCompelParse, 3, 0)

    $btnCompelExport = New-Object Windows.Forms.Button
    $btnCompelExport.Text = 'Выгрузить Excel'
    $btnCompelExport.Dock = 'Fill'
    $btnCompelExport.Enabled = $false
    Set-SecondaryButtonLook $btnCompelExport
    $compelToolbar.Controls.Add($btnCompelExport, 4, 0)

    $btnCompelOpenExport = New-Object Windows.Forms.Button
    $btnCompelOpenExport.Text = 'Открыть папку'
    $btnCompelOpenExport.Dock = 'Fill'
    $btnCompelOpenExport.Enabled = $false
    Set-SecondaryButtonLook $btnCompelOpenExport
    $compelToolbar.Controls.Add($btnCompelOpenExport, 5, 0)

    $btnCompelRrfqExport = New-Object Windows.Forms.Button
    $btnCompelRrfqExport.Text = 'Выгрузить в сравнение'
    $btnCompelRrfqExport.Dock = 'Fill'
    $btnCompelRrfqExport.Enabled = $false
    Set-SecondaryButtonLook $btnCompelRrfqExport
    $compelToolbar.Controls.Add($btnCompelRrfqExport, 6, 0)

    $compelSummaryPanel = New-Object Windows.Forms.TableLayoutPanel
    $compelSummaryPanel.Dock = 'Fill'
    $compelSummaryPanel.ColumnCount = 6
    $compelSummaryPanel.RowCount = 1
    for ($i = 0; $i -lt 6; $i++) { Add-TableColumnStyle $compelSummaryPanel 'Percent' 16.666 }
    $compelLayout.Controls.Add($compelSummaryPanel, 0, 1)

    $compelMetricLabels = @{}
    foreach ($metric in @(
        @('TotalRows', 'Строк всего'),
        @('PricedRows', 'Строк с ценой'),
        @('MissingRows', 'Без предложения'),
        @('SelectedTotalUsd', 'USD по строкам'),
        @('CheapTotalUsd', 'USD итог по цене'),
        @('OptimalTotalUsd', 'USD итог по сроку')
    )) {
        $metricBox = New-Object Windows.Forms.TableLayoutPanel
        $metricBox.Dock = 'Fill'
        $metricBox.Margin = New-Object Windows.Forms.Padding(0, 4, 8, 4)
        $metricBox.RowCount = 2
        $metricBox.ColumnCount = 1
        $metricBox.BackColor = Get-UiColor 'Surface'
        $metricBox.CellBorderStyle = [Windows.Forms.TableLayoutPanelCellBorderStyle]::Single
        Add-TableRowStyle $metricBox 'Absolute' 36
        Add-TableRowStyle $metricBox 'Percent' 100

        $valueLabel = New-Object Windows.Forms.Label
        $valueLabel.Text = '0'
        $valueLabel.Dock = 'Fill'
        $valueLabel.TextAlign = [Drawing.ContentAlignment]::MiddleCenter
        $valueLabel.Font = New-Object Drawing.Font('Segoe UI', 14, [Drawing.FontStyle]::Bold)
        $valueLabel.ForeColor = Get-UiColor 'Text'
        $metricBox.Controls.Add($valueLabel, 0, 0)

        $captionLabel = New-Object Windows.Forms.Label
        $captionLabel.Text = [string]$metric[1]
        $captionLabel.Dock = 'Fill'
        $captionLabel.TextAlign = [Drawing.ContentAlignment]::TopCenter
        $captionLabel.ForeColor = Get-UiColor 'Muted'
        $metricBox.Controls.Add($captionLabel, 0, 1)

        $compelMetricLabels[[string]$metric[0]] = $valueLabel
        $compelSummaryPanel.Controls.Add($metricBox)
    }

    $lblCompelStatus = New-Object Windows.Forms.Label
    $lblCompelStatus.Dock = 'Fill'
    $lblCompelStatus.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $lblCompelStatus.ForeColor = Get-UiColor 'Muted'
    $lblCompelStatus.Text = 'Выберите сохраненный HTML-файл расчета SDS Compel.'
    $compelLayout.Controls.Add($lblCompelStatus, 0, 2)

    $compelGrid = New-Object Windows.Forms.DataGridView
    $compelGrid.Dock = 'Fill'
    $compelGrid.AllowUserToAddRows = $false
    $compelGrid.AllowUserToDeleteRows = $false
    $compelGrid.AllowUserToResizeRows = $false
    $compelGrid.AllowUserToResizeColumns = $true
    $compelGrid.RowHeadersVisible = $false
    $compelGrid.ReadOnly = $true
    $compelGrid.SelectionMode = 'FullRowSelect'
    $compelGrid.MultiSelect = $false
    $compelGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $compelGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both
    Set-GridLook $compelGrid
    foreach ($columnInfo in @(
        @('Number', 'N', 58),
        @('SourceName', 'Исходное наименование', 230),
        @('MatchedPart', 'Подобранный товар', 250),
        @('Manufacturer', 'Производитель', 140),
        @('LeadTime', 'Срок', 105),
        @('Quantity', 'Кол-во', 88),
        @('UnitPrice', 'Цена за шт', 105),
        @('Currency', 'Валюта', 80),
        @('Total', 'Сумма', 110),
        @('TotalCurrency', 'Валюта суммы', 110)
    )) {
        $col = New-Object Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $columnInfo[0]
        $col.HeaderText = $columnInfo[1]
        $col.Width = [int]$columnInfo[2]
        $col.SortMode = [Windows.Forms.DataGridViewColumnSortMode]::Automatic
        $col.ReadOnly = $true
        $compelGrid.Columns.Add($col) | Out-Null
    }
    $compelLayout.Controls.Add($compelGrid, 0, 3)

    $script:CompelRows = @()
    $script:CompelSummary = $null
    $script:CompelSourcePath = ''
    $script:CompelLastExportPath = ''

    function Set-CompelStatus {
        param([string]$Text, [bool]$IsError = $false)
        $lblCompelStatus.Text = $Text
        $lblCompelStatus.ForeColor = if ($IsError) { [Drawing.Color]::FromArgb(160, 50, 50) } else { Get-UiColor 'Muted' }
    }

    function Refresh-CompelSummaryView {
        param($Summary)

        if ($null -eq $Summary) {
            $compelMetricLabels['TotalRows'].Text = '0'
            $compelMetricLabels['PricedRows'].Text = '0'
            $compelMetricLabels['MissingRows'].Text = '0'
            $compelMetricLabels['SelectedTotalUsd'].Text = '0.00'
            $compelMetricLabels['CheapTotalUsd'].Text = ''
            $compelMetricLabels['OptimalTotalUsd'].Text = ''
            return
        }

        $compelMetricLabels['TotalRows'].Text = [string]$Summary.TotalRows
        $compelMetricLabels['PricedRows'].Text = [string]$Summary.PricedRows
        $compelMetricLabels['MissingRows'].Text = [string]$Summary.MissingRows.Count
        $compelMetricLabels['SelectedTotalUsd'].Text = Format-CompelNumber $Summary.SelectedTotalUsd 2
        $compelMetricLabels['CheapTotalUsd'].Text = Format-CompelNumber $Summary.CheapTotalUsd 2
        $compelMetricLabels['OptimalTotalUsd'].Text = Format-CompelNumber $Summary.OptimalTotalUsd 2
    }

    function Refresh-CompelGridView {
        param([object[]]$Rows)

        $compelGrid.Rows.Clear()
        foreach ($item in @($Rows)) {
            $rowIndex = $compelGrid.Rows.Add(
                $item.Number,
                $item.SourceName,
                $item.MatchedPart,
                $item.Manufacturer,
                $item.LeadTime,
                $item.Quantity,
                (Format-CompelNumber $item.UnitPrice 5),
                $item.Currency,
                (Format-CompelNumber $item.Total 2),
                $item.TotalCurrency
            )
            if ($null -eq $item.Total) {
                $compelGrid.Rows[$rowIndex].DefaultCellStyle.BackColor = Get-UiColor 'Warn'
            }
        }
    }

    function Invoke-CompelHtmlParse {
        param([string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path)) {
            throw 'Выберите HTML-файл расчета SDS Compel.'
        }

        $btnCompelParse.Enabled = $false
        $btnCompelExport.Enabled = $false
        Set-CompelStatus 'Разбираю HTML...'
        $form.Refresh()
        try {
            $result = Parse-CompelHtml $Path
            $script:CompelRows = @($result.Rows)
            $script:CompelSummary = $result.Summary
            $script:CompelSourcePath = $Path
            $txtCompelHtml.Text = $Path
            Refresh-CompelSummaryView $script:CompelSummary
            Refresh-CompelGridView $script:CompelRows
            $btnCompelExport.Enabled = ($script:CompelRows.Count -gt 0)
            $btnCompelRrfqExport.Enabled = ($script:CompelRows.Count -gt 0)
            Set-CompelStatus ("Готово: найдено {0} строк, {1} с ценой." -f $script:CompelSummary.TotalRows, $script:CompelSummary.PricedRows)
        } finally {
            $btnCompelParse.Enabled = $true
        }
    }

    $btnCompelChoose.Add_Click({
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Title = 'Выберите HTML-файл SDS Compel'
        $dialog.Filter = 'HTML files (*.html;*.htm)|*.html;*.htm|All files (*.*)|*.*'
        $dialog.Multiselect = $false
        if ($dialog.ShowDialog($form) -eq [Windows.Forms.DialogResult]::OK) {
            try {
                Invoke-CompelHtmlParse $dialog.FileName
            } catch {
                Refresh-CompelSummaryView $null
                Refresh-CompelGridView @()
                Set-CompelStatus $_.Exception.Message $true
                [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Компэл Парсер') | Out-Null
            }
        }
    })

    $btnCompelParse.Add_Click({
        try {
            Invoke-CompelHtmlParse $txtCompelHtml.Text
        } catch {
            Set-CompelStatus $_.Exception.Message $true
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Компэл Парсер') | Out-Null
        }
    })

    $btnCompelExport.Add_Click({
        try {
            if ($null -eq $script:CompelSummary -or @($script:CompelRows).Count -eq 0) {
                throw 'Сначала разберите HTML-файл Компэла.'
            }

            $dialog = New-Object Windows.Forms.SaveFileDialog
            $dialog.Title = 'Сохранить Excel по Компэлу'
            $dialog.Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
            $dialog.DefaultExt = 'xlsx'
            $dialog.AddExtension = $true
            $sourceStem = if ([string]::IsNullOrWhiteSpace($script:CompelSourcePath)) { 'compel' } else { [IO.Path]::GetFileNameWithoutExtension($script:CompelSourcePath) }
            $dialog.FileName = "${sourceStem}_compel_prices.xlsx"
            if (-not [string]::IsNullOrWhiteSpace($script:CompelSourcePath)) {
                $dialog.InitialDirectory = Split-Path -Parent $script:CompelSourcePath
            }
            if ($dialog.ShowDialog($form) -ne [Windows.Forms.DialogResult]::OK) { return }

            $btnCompelExport.Enabled = $false
            Set-CompelStatus 'Создаю Excel...'
            $form.Refresh()
            $path = Write-CompelWorkbook $dialog.FileName $script:CompelRows $script:CompelSummary
            $script:CompelLastExportPath = $path
            $btnCompelOpenExport.Enabled = $true
            Set-CompelStatus "Excel готов: $path"
            [Windows.Forms.MessageBox]::Show("Excel готов:`r`n$path", 'Компэл Парсер') | Out-Null
        } catch {
            Set-CompelStatus $_.Exception.Message $true
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Компэл Парсер') | Out-Null
        } finally {
            $btnCompelExport.Enabled = ($null -ne $script:CompelSummary -and @($script:CompelRows).Count -gt 0)
            $btnCompelRrfqExport.Enabled = ($null -ne $script:CompelSummary -and @($script:CompelRows).Count -gt 0)
        }
    })

    $btnCompelRrfqExport.Add_Click({
        try {
            if (@($script:CompelRows).Count -eq 0) { throw 'Сначала разберите HTML-файл Компэла.' }
            $downloads = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'
            if (-not (Test-Path -LiteralPath $downloads -PathType Container)) { [void](New-Item -ItemType Directory -Path $downloads -Force) }
            $baseName = [IO.Path]::GetFileNameWithoutExtension([string]$script:CompelSourcePath)
            if ([string]::IsNullOrWhiteSpace($baseName)) { $baseName = 'Compel_RRFQ' }
            foreach ($invalid in [IO.Path]::GetInvalidFileNameChars()) { $baseName = $baseName.Replace([string]$invalid, '_') }
            $fileName = '{0}_Compel.xlsx' -f $baseName
            $path = Write-CompelWorkbook (Join-Path $downloads $fileName) $script:CompelRows $script:CompelSummary
            foreach ($supplierRow in @($supplierGrid.Rows)) {
                if ([string]$supplierRow.Cells['Path'].Value -eq 'System.Object[]') { $supplierGrid.Rows.Remove($supplierRow) }
            }
            $existing = @($supplierGrid.Rows | Where-Object { [string]$_.Cells['Path'].Value -eq $path })
            if ($existing.Count -eq 0) {
                [void]$supplierGrid.Rows.Add($path, 'Компэл')
            }
            Set-CompelStatus ('Файл добавлен в поставщики RRFQ: {0}' -f $path)
        } catch {
            Set-CompelStatus $_.Exception.Message $true
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Выгрузка Компэла') | Out-Null
        }
    })

    $btnCompelOpenExport.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($script:CompelLastExportPath) -or -not (Test-Path -LiteralPath $script:CompelLastExportPath -PathType Leaf)) {
                throw 'Сначала выгрузите Excel.'
            }
            $arg = "/select,`"$script:CompelLastExportPath`""
            Start-Process -FilePath explorer.exe -ArgumentList $arg | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Компэл Парсер') | Out-Null
        }
    })
    $priceSearchLayout = New-Object Windows.Forms.TableLayoutPanel
    $priceSearchLayout.Dock = 'Fill'
    $priceSearchLayout.Padding = New-Object Windows.Forms.Padding(12)
    $priceSearchLayout.RowCount = 2
    $priceSearchLayout.ColumnCount = 1
    Add-TableRowStyle $priceSearchLayout 'Absolute' 58
    Add-TableRowStyle $priceSearchLayout 'Percent' 100
    $priceSearchPage.Controls.Add($priceSearchLayout)
    $priceSearchToolbar = New-Object Windows.Forms.FlowLayoutPanel
    $priceSearchToolbar.Dock = 'Fill'
    $priceSearchToolbar.WrapContents = $false
    $priceSearchToolbar.Padding = New-Object Windows.Forms.Padding(0, 6, 0, 6)
    $priceSearchLayout.Controls.Add($priceSearchToolbar, 0, 0)
    $lblPriceSearchPn = New-Object Windows.Forms.Label
    $lblPriceSearchPn.Text = 'PN:'; $lblPriceSearchPn.AutoSize = $true; $lblPriceSearchPn.Margin = New-Object Windows.Forms.Padding(0, 9, 5, 0)
    $priceSearchToolbar.Controls.Add($lblPriceSearchPn)
    $txtPriceSearchPn = New-Object Windows.Forms.TextBox
    $txtPriceSearchPn.Width = 220; $txtPriceSearchPn.Margin = New-Object Windows.Forms.Padding(0, 3, 14, 0)
    $priceSearchToolbar.Controls.Add($txtPriceSearchPn)
    $lblPriceSearchQty = New-Object Windows.Forms.Label
    $lblPriceSearchQty.Text = 'Количество:'; $lblPriceSearchQty.AutoSize = $true; $lblPriceSearchQty.Margin = New-Object Windows.Forms.Padding(0, 9, 5, 0)
    $priceSearchToolbar.Controls.Add($lblPriceSearchQty)
    $txtPriceSearchQty = New-Object Windows.Forms.TextBox
    $txtPriceSearchQty.Width = 90; $txtPriceSearchQty.Margin = New-Object Windows.Forms.Padding(0, 3, 14, 0)
    $priceSearchToolbar.Controls.Add($txtPriceSearchQty)
    $btnPriceSearch = New-Object Windows.Forms.Button
    $btnPriceSearch.Text = 'Поиск цены'; $btnPriceSearch.Width = 120; $btnPriceSearch.Height = 32; $btnPriceSearch.Margin = New-Object Windows.Forms.Padding(0, 2, 8, 0)
    Set-PrimaryButtonLook $btnPriceSearch; $priceSearchToolbar.Controls.Add($btnPriceSearch)
    $btnPriceSearchGpt = New-Object Windows.Forms.Button
    $btnPriceSearchGpt.Text = 'Поиск в GPT'; $btnPriceSearchGpt.Width = 120; $btnPriceSearchGpt.Height = 32; $btnPriceSearchGpt.Margin = New-Object Windows.Forms.Padding(0, 2, 8, 0)
    Set-SecondaryButtonLook $btnPriceSearchGpt; $priceSearchToolbar.Controls.Add($btnPriceSearchGpt)
    $priceSearchGrid = New-Object Windows.Forms.DataGridView
    $priceSearchGrid.Dock = 'Fill'; $priceSearchGrid.AllowUserToAddRows = $false; $priceSearchGrid.RowHeadersVisible = $false; $priceSearchGrid.SelectionMode = 'FullRowSelect'; $priceSearchGrid.MultiSelect = $false; $priceSearchGrid.ScrollBars = [Windows.Forms.ScrollBars]::Both; $priceSearchGrid.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    Set-GridLook $priceSearchGrid
    foreach($columnInfo in @(@('Source','Источник',125),@('PN','PN',170),@('Manufacturer','Производитель',150),@('Price','Цена',90),@('Currency','Валюта',80),@('MOQ','MOQ',90),@('Stock','Запас',100),@('LeadTime','Lead Time',120),@('Date','Дата',145),@('Link','Ссылка',330))){
        $col=New-Object Windows.Forms.DataGridViewTextBoxColumn; $col.Name=$columnInfo[0]; $col.HeaderText=$columnInfo[1]; $col.Width=[int]$columnInfo[2]; $col.ReadOnly=$true; [void]$priceSearchGrid.Columns.Add($col)
    }
    $priceSearchLayout.Controls.Add($priceSearchGrid, 0, 1)

    $notesLayout = New-Object Windows.Forms.TableLayoutPanel
    $notesLayout.Dock = 'Fill'
    $notesLayout.Padding = New-Object Windows.Forms.Padding(12)
    $notesLayout.RowCount = 3
    $notesLayout.ColumnCount = 1
    Add-TableRowStyle $notesLayout 'Absolute' 44
    Add-TableRowStyle $notesLayout 'Absolute' 28
    Add-TableRowStyle $notesLayout 'Percent' 100
    $notesPage.Controls.Add($notesLayout)

    $notesToolbar = New-Object Windows.Forms.FlowLayoutPanel
    $notesToolbar.Dock = 'Fill'
    $notesToolbar.FlowDirection = 'LeftToRight'
    $notesToolbar.WrapContents = $false
    $notesToolbar.Padding = New-Object Windows.Forms.Padding(0, 5, 0, 5)
    $notesLayout.Controls.Add($notesToolbar, 0, 0)

    $notesToolbar.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'Заметка:'; Width = 62; Height = 30; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }))

    $cmbNotesFiles = New-Object Windows.Forms.ComboBox
    $cmbNotesFiles.DropDownStyle = [Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbNotesFiles.Width = 260
    $cmbNotesFiles.Height = 30
    $cmbNotesFiles.DisplayMember = 'Name'
    $notesToolbar.Controls.Add($cmbNotesFiles)

    $btnNewNote = New-Object Windows.Forms.Button
    $btnNewNote.Text = 'Новая заметка'
    $btnNewNote.Width = 130
    $btnNewNote.Height = 30
    Set-SecondaryButtonLook $btnNewNote
    $notesToolbar.Controls.Add($btnNewNote)

    $btnSaveNotes = New-Object Windows.Forms.Button
    $btnSaveNotes.Text = 'Сохранить'
    $btnSaveNotes.Width = 120
    $btnSaveNotes.Height = 30
    Set-PrimaryButtonLook $btnSaveNotes
    $notesToolbar.Controls.Add($btnSaveNotes)

    $btnReloadNotes = New-Object Windows.Forms.Button
    $btnReloadNotes.Text = 'Обновить'
    $btnReloadNotes.Width = 110
    $btnReloadNotes.Height = 30
    Set-SecondaryButtonLook $btnReloadNotes
    $notesToolbar.Controls.Add($btnReloadNotes)

    $notesPathLabel = New-Object Windows.Forms.Label
    $notesPathLabel.Text = Get-NotesFilePath
    $notesPathLabel.Dock = 'Fill'
    $notesPathLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $notesPathLabel.ForeColor = Get-UiColor 'Muted'
    $notesLayout.Controls.Add($notesPathLabel, 0, 1)

    $txtProjectNotes = New-Object Windows.Forms.TextBox
    $txtProjectNotes.Dock = 'Fill'
    $txtProjectNotes.Multiline = $true
    $txtProjectNotes.AcceptsReturn = $true
    $txtProjectNotes.AcceptsTab = $true
    $txtProjectNotes.ScrollBars = [Windows.Forms.ScrollBars]::Both
    $txtProjectNotes.WordWrap = $true
    $txtProjectNotes.Font = New-Object Drawing.Font('Segoe UI', 10)
    $notesLayout.Controls.Add($txtProjectNotes, 0, 2)

    $script:NotesLoading = $false

    function Get-SelectedNotePathFromCombo {
        if ($null -eq $cmbNotesFiles.SelectedItem) { return '' }
        $item = $cmbNotesFiles.SelectedItem
        if ($item.PSObject.Properties['Path']) { return [string]$item.Path }
        return ''
    }

    function Save-CurrentNoteFromEditor {
        if ([string]::IsNullOrWhiteSpace((Get-NotesFilePath))) { return }
        Save-ProjectNotes $txtProjectNotes.Text
    }

    function Refresh-NotesList {
        param([string]$PreferredPath = '')

        $script:NotesLoading = $true
        try {
            $cmbNotesFiles.Items.Clear()
            $files = @(Get-ProjectNoteFiles)
            foreach ($file in $files) {
                [void]$cmbNotesFiles.Items.Add([pscustomobject]@{
                    Name = Get-NoteDisplayName $file.FullName
                    Path = $file.FullName
                })
            }

            $targetPath = if (-not [string]::IsNullOrWhiteSpace($PreferredPath)) { $PreferredPath } else { Get-NotesFilePath }
            $selectedIndex = -1
            for ($i = 0; $i -lt $cmbNotesFiles.Items.Count; $i++) {
                if ([string]::Equals([string]$cmbNotesFiles.Items[$i].Path, $targetPath, [StringComparison]::OrdinalIgnoreCase)) {
                    $selectedIndex = $i
                    break
                }
            }
            if ($selectedIndex -lt 0 -and $cmbNotesFiles.Items.Count -gt 0) { $selectedIndex = 0 }
            if ($selectedIndex -ge 0) { $cmbNotesFiles.SelectedIndex = $selectedIndex }
        } finally {
            $script:NotesLoading = $false
        }
        Load-ProjectNotesToEditor
    }

    function Load-ProjectNotesToEditor {
        try {
            $selectedPath = Get-SelectedNotePathFromCombo
            if (-not [string]::IsNullOrWhiteSpace($selectedPath)) {
                $script:CurrentNotesFilePath = $selectedPath
            }
            $txtProjectNotes.Text = Read-ProjectNotes
            $notesPathLabel.Text = Get-NotesFilePath
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Заметки') | Out-Null
        }
    }

    $cmbNotesFiles.Add_SelectedIndexChanged({
        if ($script:NotesLoading) { return }
        try {
            Save-CurrentNoteFromEditor
            Load-ProjectNotesToEditor
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Заметки') | Out-Null
        }
    })

    $btnNewNote.Add_Click({
        try {
            Save-CurrentNoteFromEditor
            $title = Show-SimpleInputDialog 'Новая заметка' 'Название заметки' ''
            if ([string]::IsNullOrWhiteSpace($title)) { return }
            $path = New-ProjectNoteFile $title
            Refresh-NotesList $path
            $txtProjectNotes.Focus()
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Новая заметка') | Out-Null
        }
    })

    $btnSaveNotes.Add_Click({
        try {
            Save-CurrentNoteFromEditor
            Refresh-NotesList (Get-NotesFilePath)
            [Windows.Forms.MessageBox]::Show('Заметка сохранена.', 'Заметки') | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Заметки') | Out-Null
        }
    })
    $btnReloadNotes.Add_Click({ Refresh-NotesList (Get-NotesFilePath) })
    Refresh-NotesList
    $instructionsLayout = New-Object Windows.Forms.TableLayoutPanel
    $instructionsLayout.Dock = 'Fill'
    $instructionsLayout.Padding = New-Object Windows.Forms.Padding(24)
    $instructionsLayout.RowCount = 2
    $instructionsLayout.ColumnCount = 1
    Add-TableRowStyle $instructionsLayout 'Absolute' 44
    Add-TableRowStyle $instructionsLayout 'Percent' 100
    $instructionsPage.Controls.Add($instructionsLayout)

    $instructionsTitle = New-Object Windows.Forms.Label
    $instructionsTitle.Text = 'Инструкции'
    $instructionsTitle.Dock = 'Fill'
    $instructionsTitle.Font = New-Object Drawing.Font('Segoe UI', 15, [Drawing.FontStyle]::Bold)
    $instructionsTitle.ForeColor = Get-UiColor 'Text'
    $instructionsLayout.Controls.Add($instructionsTitle, 0, 0)

    $instructionsBody = New-Object Windows.Forms.TableLayoutPanel
    $instructionsBody.Dock = 'Fill'
    $instructionsBody.ColumnCount = 2
    $instructionsBody.RowCount = 1
    Add-TableColumnStyle $instructionsBody 'Absolute' 230
    Add-TableColumnStyle $instructionsBody 'Percent' 100
    Add-TableRowStyle $instructionsBody 'Percent' 100
    $instructionsLayout.Controls.Add($instructionsBody, 0, 1)

    $instructionsList = New-Object Windows.Forms.ListBox
    $instructionsList.Dock = 'Fill'
    $instructionsList.BackColor = Get-UiColor 'SurfaceAlt'
    $instructionsList.ForeColor = Get-UiColor 'Text'
    $instructionsList.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
    $instructionsList.Font = New-Object Drawing.Font('Segoe UI', 10)
    $instructionsList.Margin = New-Object Windows.Forms.Padding(0, 0, 12, 0)
    $instructionsBody.Controls.Add($instructionsList, 0, 0)

    $instructionsText = New-Object Windows.Forms.RichTextBox
    $instructionsText.Dock = 'Fill'
    $instructionsText.ReadOnly = $true
    $instructionsText.BackColor = Get-UiColor 'Surface'
    $instructionsText.ForeColor = Get-UiColor 'Text'
    $instructionsText.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
    $instructionsText.Font = New-Object Drawing.Font('Segoe UI', 10)
    $instructionsText.DetectUrls = $false
    $instructionsBody.Controls.Add($instructionsText, 1, 0)

    $instructionSections = [ordered]@{
        'FAQ' = @'
FAQ по приложению

Что делает приложение?
Приложение сравнивает RRFQ поставщиков, ведет контроль закупки, хранит документы, историю действий, задачи и базу квот.

Где лучше работать?
Сейчас используется только основная стабильная desktop-версия.

Как тестировать без мусора?
В "База квот" можно удалить выбранные строки или очистить базу. В "История" можно очистить журнал действий.
'@
        'Сравнение RRFQ' = @'
Сравнение RRFQ

1. Выберите исходный RFQ.
2. Добавьте файлы поставщиков.
3. Выберите приоритет: цена или срок.
4. Нажмите "Сформировать preview".
5. Проверьте победителей, предупреждения и несопоставленные строки.
6. При необходимости выберите квоту вручную или заполните предложение вручную.
7. Нажмите "Создать результирующий RRFQ".

База квот заполняется после анализа RRFQ: сохраняются все найденные квоты, а не только победители.
'@
        'Контроль закупки' = @'
Контроль закупки

1. Создайте сделку вручную или импортируйте результирующий RRFQ.
2. Заполните клиента, период, статус и комментарий.
3. Добавьте поставщиков.
4. Загрузите PI/ERP/прочие документы.
5. Отмечайте оплату, ERP, подтверждение инвойса и поступление.

Документы копируются в локальное хранилище приложения. При удалении сделки удаляется и папка сделки.
'@
        'Задачи' = @'
Задачи

Раздел нужен для оперативного отслеживания задач.

При создании строки автоматически создается папка в data\components\files.
Если заполнить номер сделки, папка переименуется с учетом этого номера.
При удалении строки задачи ее папка также удаляется.
'@
        'Перенос' = @'
Перенос на другой компьютер

Скопируйте всю папку проекта целиком: app, config, data, Start.bat и Start.vbs.

На другом компьютере нужен Microsoft Excel.
Для считывания сумм из PDF желательно наличие Microsoft Word.

Если в настройках документов указан внешний путь, перенесите и эту папку либо верните путь внутрь папки приложения.
'@
        'Outlook позже' = @'
Outlook

Интеграция с Outlook запланирована позже отдельным этапом.
Здесь будут инструкции по письмам, задачам и автоматическим напоминаниям, когда мы дойдем до этого блока.
'@
    }
    foreach ($key in $instructionSections.Keys) {
        [void]$instructionsList.Items.Add($key)
    }
    $instructionsList.Add_SelectedIndexChanged({
        $key = [string]$instructionsList.SelectedItem
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $instructionsText.Text = [string]$instructionSections[$key]
            $instructionsText.SelectionStart = 0
            $instructionsText.ScrollToCaret()
        }
    })
    $instructionsList.SelectedIndex = 0

    function Resize-PurchasePanels {
        $purchaseContent.PerformLayout()
        if ($script:PurchaseSplitterInitialized) { return }
        if ($purchaseMainSplit.Height -gt 430) {
            $purchaseMainSplit.SplitterDistance = [Math]::Max(170, [int]($purchaseMainSplit.Height * 0.60))
        }
        if ($purchaseLowerSplit.Height -gt 220) {
            $purchaseLowerSplit.SplitterDistance = [Math]::Max(90, [int]($purchaseLowerSplit.Height * 0.50))
        }
        if ($purchaseMainSplit.Height -gt 430 -and $purchaseLowerSplit.Height -gt 220) {
            $script:PurchaseSplitterInitialized = $true
        }
    }

    $settingsLayout = New-Object Windows.Forms.TableLayoutPanel
    $settingsLayout.Dock = 'Top'
    $settingsLayout.Padding = New-Object Windows.Forms.Padding(18)
    $settingsLayout.RowCount = 3
    $settingsLayout.ColumnCount = 3
    Add-TableColumnStyle $settingsLayout 'Absolute' 180
    Add-TableColumnStyle $settingsLayout 'Percent' 100
    Add-TableColumnStyle $settingsLayout 'Absolute' 150
    Add-TableRowStyle $settingsLayout 'Absolute' 38
    Add-TableRowStyle $settingsLayout 'Absolute' 38
    Add-TableRowStyle $settingsLayout 'Absolute' 48
    $settingsPage.Controls.Add($settingsLayout)

    $settingsTitle = New-Object Windows.Forms.Label
    $settingsTitle.Text = 'Настройки контроля закупки'
    $settingsTitle.Dock = 'Fill'
    $settingsTitle.Font = New-Object Drawing.Font('Segoe UI', 12, [Drawing.FontStyle]::Bold)
    $settingsLayout.Controls.Add($settingsTitle, 0, 0)
    $settingsLayout.SetColumnSpan($settingsTitle, 3)

    $settingsLayout.Controls.Add((New-Object Windows.Forms.Label -Property @{ Text = 'Папка документов'; Dock = 'Fill'; TextAlign = [Drawing.ContentAlignment]::MiddleLeft }), 0, 1)
    $txtDocumentsRoot = New-Object Windows.Forms.TextBox
    $txtDocumentsRoot.Dock = 'Fill'
    $txtDocumentsRoot.Text = Get-PurchaseDocumentsRoot
    $settingsLayout.Controls.Add($txtDocumentsRoot, 1, 1)

    $btnChooseDocumentsRoot = New-Object Windows.Forms.Button
    $btnChooseDocumentsRoot.Text = 'Выбрать...'
    $btnChooseDocumentsRoot.Dock = 'Fill'
    Set-SecondaryButtonLook $btnChooseDocumentsRoot
    $settingsLayout.Controls.Add($btnChooseDocumentsRoot, 2, 1)

    $btnSaveSettings = New-Object Windows.Forms.Button
    $btnSaveSettings.Text = 'Сохранить настройки'
    $btnSaveSettings.Width = 180
    $btnSaveSettings.Height = 30
    Set-PrimaryButtonLook $btnSaveSettings
    $settingsLayout.Controls.Add($btnSaveSettings, 1, 2)

    $script:PurchaseLoading = $false
    $script:PurchaseSplitterInitialized = $false
    $script:RestoringDealColumnWidths = $false
    $script:ComponentsLoading = $false
    $script:CurrentComponentNotesId = 0
    $script:CurrentComponentNotesText = ''
    $script:LoadingComponentNotes = $false

    function Get-DealFilterCode {
        switch ([string]$cmbDealFilter.SelectedItem) {
            'Активные' { return 'Active' }
            'На отслеживании' { return 'Tracking' }
            'Китай' { return 'China' }
            'Закупка' { return 'Purchase' }
            'Пауза' { return 'Pause' }
            'Есть незавершенное' { return 'Open' }
            'Нет инвойсов' { return 'NoInvoices' }
            'Ожидается поступление' { return 'WaitingReceipt' }
            'Готово' { return 'Done' }
            'Архив' { return 'Archived' }
            'Все включая архив' { return 'AllRecords' }
            default { return 'All' }
        }
    }

    function Get-SelectedDealId {
        if ($dealsGrid.SelectedRows.Count -eq 0 -or $null -eq $dealsGrid.SelectedRows[0].Tag) { return 0 }
        return [int]$dealsGrid.SelectedRows[0].Tag
    }

    function Get-SelectedSupplierId {
        if ($suppliersGrid.SelectedRows.Count -eq 0 -or $null -eq $suppliersGrid.SelectedRows[0].Tag) { return 0 }
        return [int]$suppliersGrid.SelectedRows[0].Tag
    }

    function Get-SelectedDealRow {
        if ($dealsGrid.SelectedRows.Count -eq 0) { return $null }
        return $dealsGrid.SelectedRows[0]
    }

    function Get-SelectedSupplierRow {
        if ($suppliersGrid.SelectedRows.Count -eq 0) { return $null }
        return $suppliersGrid.SelectedRows[0]
    }

    function Get-SelectedDocumentInfo {
        if ($docsGrid.SelectedRows.Count -eq 0 -or $null -eq $docsGrid.SelectedRows[0].Tag) {
            return $null
        }

        $tag = $docsGrid.SelectedRows[0].Tag
        if ($tag.PSObject.Properties['Id'] -and $tag.PSObject.Properties['Path']) {
            return $tag
        }

        return [pscustomobject]@{
            Id = 0
            Path = [string]$tag
        }
    }

    function Test-SupplierErpNotRequired {
        param([string]$Supplier)

        $name = (([string]$Supplier).Trim().ToLowerInvariant()) -replace '\s+', ''
        if ([string]::IsNullOrWhiteSpace($name)) { return $false }
        return (
            $name -match '^(комп[эе]л|compel)(_\d+)?$' -or
            $name -match '^(чид|chid)(_\d+)?$' -or
            $name -match '^(промэлектроника|promelec|promelektronika)(_\d+)?$'
        )
    }

    function Test-SupplierRowDone {
        param($Row)

        if ($null -eq $Row) { return $false }
        $erpDone = $true
        if (-not (Test-SupplierErpNotRequired ([string]$Row.Cells['Supplier'].Value))) {
            $erpDone = ([bool]$Row.Cells['ErpSupplier'].Value -and [bool]$Row.Cells['ErpRoger'].Value)
        }
        return (
            [bool]$Row.Cells['InvoiceReceived'].Value -and
            [bool]$Row.Cells['InvoiceConfirmed'].Value -and
            $erpDone
        )
    }

    function Apply-SupplierRowStyle {
        param($Row)

        if ($null -eq $Row) { return }

        $hasPiAmount = -not [string]::IsNullOrWhiteSpace([string]$Row.Cells['PiUsd'].Value) -or -not [string]::IsNullOrWhiteSpace([string]$Row.Cells['PiCny'].Value) -or -not [string]::IsNullOrWhiteSpace([string]$Row.Cells['PiRub'].Value)
        $invoiceReceived = [bool]$Row.Cells['InvoiceReceived'].Value
        $paymentSubmitted = [bool]$Row.Cells['PaymentSubmitted'].Value
        $paid = [bool]$Row.Cells['Paid'].Value
        $rowColor = Get-UiColor 'Surface'

        if (Test-SupplierRowDone $Row) {
            $rowColor = Get-UiColor 'Success'
        } elseif (-not $invoiceReceived) {
            $rowColor = Get-UiColor 'Warn'
        } elseif ($invoiceReceived -and -not $hasPiAmount) {
            $rowColor = [Drawing.Color]::FromArgb(255, 249, 221)
        } elseif ($hasPiAmount -and -not $paymentSubmitted) {
            $rowColor = Get-UiColor 'Info'
        } elseif ($paymentSubmitted -and -not $paid) {
            $rowColor = Get-UiColor 'Attention'
        }

        $Row.DefaultCellStyle.BackColor = $rowColor
        foreach ($cell in $Row.Cells) {
            $cell.Style.BackColor = $rowColor
            $cell.Style.ForeColor = Get-UiColor 'Text'
        }
        if ($Row.DataGridView.Columns.Contains('ActualReceiptDate') -and -not [string]::IsNullOrWhiteSpace([string]$Row.Cells['ActualReceiptDate'].Value)) {
            $Row.Cells['ActualReceiptDate'].Style.BackColor = Get-UiColor 'Success'
            $Row.Cells['ActualReceiptDate'].Style.ForeColor = Get-UiColor 'Text'
            $Row.Cells['ActualReceiptDate'].Style.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
        }
    }

    function Apply-SupplierErpCellRules {
        param($Row)

        if ($null -eq $Row -or -not $Row.DataGridView.Columns.Contains('ErpSupplier') -or -not $Row.DataGridView.Columns.Contains('ErpRoger')) { return }
        $locked = Test-SupplierErpNotRequired ([string]$Row.Cells['Supplier'].Value)
        foreach ($columnName in @('ErpSupplier', 'ErpRoger')) {
            $cell = $Row.Cells[$columnName]
            $cell.ReadOnly = $locked
            if ($locked) {
                $cell.Value = $false
                $cell.Style.BackColor = [Drawing.Color]::FromArgb(230, 230, 230)
                $cell.Style.ForeColor = Get-UiColor 'Muted'
            }
        }
    }

    function Update-AutoReceiptDateForRow {
        param($Row)

        if ($null -eq $Row) { return }
        $invoiceDate = [string]$Row.Cells['InvoiceDate'].Value
        $deliveryWeeks = [string]$Row.Cells['DeliveryWeeks'].Value
        $supplier = [string]$Row.Cells['Supplier'].Value
        $receiptDate = Get-AutoReceiptDate $invoiceDate $deliveryWeeks $supplier
        if (-not [string]::IsNullOrWhiteSpace($receiptDate)) {
            $Row.Cells['ReceiptDate'].Value = $receiptDate
        }
    }

    function Load-SelectedDealInfo {
        $row = Get-SelectedDealRow
        $script:PurchaseLoading = $true
        try {
            if ($null -eq $row) {
                $txtDealClient.Text = ''
                $cmbDealStatusEdit.SelectedIndex = 0
                $cmbDealPeriodEdit.Text = ''
                $txtDealComment.Text = ''
                return
            }

            $txtDealClient.Text = [string]$row.Cells['Client'].Value
            $cmbDealPeriodEdit.Text = [string]$row.Cells['Period'].Value
            $status = [string]$row.Cells['Stage'].Value
            if ([string]::IsNullOrWhiteSpace($status)) { $status = 'RFQ' }
            if ($cmbDealStatusEdit.Items.Contains($status)) {
                $cmbDealStatusEdit.SelectedItem = $status
            } else {
                $cmbDealStatusEdit.SelectedIndex = 0
            }
            $txtDealComment.Text = [string]$row.Cells['Comment'].Value
        } finally {
            $script:PurchaseLoading = $false
        }
    }


    function Update-PurchaseCockpitMetrics {
        $rows = @()
        try { $rows = @(Get-PurchaseDeals '' 'AllRecords') } catch { $rows = @() }
        $today = (Get-Date).Date
        $active = 0; $overdue = 0; $dueToday = 0; $attention = 0
        foreach ($row in $rows) {
            $archived = Convert-DbBool $row.archived
            $isDone = ([string]$row.status -match 'Done|Completed')
            if (-not $archived -and -not $isDone) { $active++ }
            $receipt = Convert-PurchaseDateText ([string]$row.completion_receipt_date)
            if ($null -ne $receipt -and -not $archived -and -not $isDone) {
                if ($receipt.Date -lt $today) { $overdue++ }
                if ($receipt.Date -eq $today) { $dueToday++ }
            }
            if (-not $archived -and [int]$row.supplier_count -gt 0 -and [int]$row.invoice_confirmed_count -lt [int]$row.supplier_count) { $attention++ }
        }
        Set-CockpitMetricCard $metricActive ([string]$active) 'Current operational queue'
        Set-CockpitMetricCard $metricOverdue ([string]$overdue) 'Requires immediate action'
        Set-CockpitMetricCard $metricToday ([string]$dueToday) 'Expected receipt date'
        Set-CockpitMetricCard $metricAttention ([string]$attention) 'Needs manual review'
    }

    function Refresh-PurchaseDeals {
        $selectedId = Get-SelectedDealId
        $script:PurchaseLoading = $true
        try {
            $dealsGrid.Rows.Clear()
            $table = Get-PurchaseDeals $txtDealSearch.Text (Get-DealFilterCode)
            foreach ($row in $table.Rows) {
                $idx = $dealsGrid.Rows.Add(
                    [string]$row.deal_number,
                    [string]$row.board_count,
                    [string]$row.client,
                    [string]$row.priority,
                    [string]$row.status,
                    [string]$row.supplier_count,
                    [string]$row.invoice_count,
                    ('{0}/{1}' -f $row.paid_count, $row.supplier_count),
                    ('{0}/{1}' -f $row.done_count, $row.supplier_count),
                    $(if ([string]::IsNullOrWhiteSpace([string]$row.masks) -or [int]$row.masks -eq 2) { '' } elseif (Convert-DbBool $row.masks) { 'Да' } else { 'Нет' }),
                    [string]$row.comment,
                    (Format-PurchaseDate $row.completion_receipt_date),
                    [string]$row.updated_at,
                    [string]$row.period,
                    [string]$row.executor,
                    [string]$row.assembly_location
                )
                $dealsGrid.Rows[$idx].Tag = [int]$row.id
                Apply-PriorityCellStyle $dealsGrid.Rows[$idx]
                if ([string]$row.status -in @('RRFQ', 'PI')) {
                    $dealsGrid.Rows[$idx].DefaultCellStyle.BackColor = Get-UiColor 'Warn'
                } elseif (Convert-DbBool $row.archived) {
                    $dealsGrid.Rows[$idx].DefaultCellStyle.BackColor = Get-UiColor 'SurfaceAlt'
                    $dealsGrid.Rows[$idx].DefaultCellStyle.ForeColor = Get-UiColor 'Muted'
                } elseif ([string]$row.assembly_location -eq 'Китай' -and [string]$row.status -eq 'Заказано') {
                    $dealsGrid.Rows[$idx].DefaultCellStyle.BackColor = Get-UiColor 'Success'
                } elseif ([string]$row.status -eq 'Заказано' -and [int]$row.supplier_count -gt 0 -and [int]$row.invoice_confirmed_count -eq [int]$row.supplier_count) {
                    $dealsGrid.Rows[$idx].DefaultCellStyle.BackColor = Get-UiColor 'Success'
                } else {
                    $dealsGrid.Rows[$idx].DefaultCellStyle.BackColor = Get-UiColor 'Danger'
                }
            }
            if ($dealsGrid.Rows.Count -gt 0) {
                $target = $dealsGrid.Rows[0]
                foreach ($gridRow in $dealsGrid.Rows) {
                    if ($selectedId -gt 0 -and [int]$gridRow.Tag -eq $selectedId) {
                        $target = $gridRow
                        break
                    }
                }
                $dealsGrid.ClearSelection()
                $target.Selected = $true
                $dealsGrid.CurrentCell = $target.Cells[0]
            }
        } finally {
            $script:PurchaseLoading = $false
        }
        Load-SelectedDealInfo
    }

    function Refresh-PurchaseSuppliers {
        $script:PurchaseLoading = $true
        try {
            $selectedSupplierId = Get-SelectedSupplierId
            $suppliersGrid.Rows.Clear()
            $dealId = Get-SelectedDealId
            if ($dealId -le 0) { return }
            $table = Get-PurchaseSuppliers $dealId
            foreach ($row in $table.Rows) {
                $idx = $suppliersGrid.Rows.Add(
                    [string]$row.supplier,
                    (Format-PurchaseAmount $row.pi_amount_usd),
                    (Format-PurchaseAmount $row.pi_amount_cny),
                    (Format-PurchaseAmount $row.pi_amount_rub),
                    [string]$row.paid_amount,
                    [string]$row.delivery_weeks,
                    (Convert-DbBool $row.payment_submitted),
                    (Convert-DbBool $row.paid),
                    (Convert-DbBool $row.invoice_received),
                    (Convert-DbBool $row.invoice_confirmed),
                    (Convert-DbBool $row.erp_supplier_sent),
                    (Convert-DbBool $row.erp_roger_sent),
                    (Format-PurchaseDate $row.invoice_confirmed_date),
                    (Format-PurchaseDate $row.components_receipt_date),
                    (Format-PurchaseDate $row.actual_receipt_date),
                    [string]$row.comment
                )
                $suppliersGrid.Rows[$idx].Tag = [int]$row.id
                Apply-SupplierRowStyle $suppliersGrid.Rows[$idx]
                Apply-SupplierErpCellRules $suppliersGrid.Rows[$idx]
            }
            if ($suppliersGrid.Rows.Count -gt 0) {
                $target = $suppliersGrid.Rows[0]
                foreach ($gridRow in $suppliersGrid.Rows) {
                    if ($selectedSupplierId -gt 0 -and [int]$gridRow.Tag -eq $selectedSupplierId) {
                        $target = $gridRow
                        break
                    }
                }
                $suppliersGrid.ClearSelection()
                $target.Selected = $true
                $suppliersGrid.CurrentCell = $target.Cells[0]
            }
        } finally {
            $script:PurchaseLoading = $false
        }
    }

    function Refresh-PurchaseDocuments {
        $docsGrid.Rows.Clear()
        $dealId = Get-SelectedDealId
        if ($dealId -le 0) { return }
        $table = Get-PurchaseDocuments $dealId
        foreach ($row in $table.Rows) {
            $resolvedDoc = Resolve-PurchaseDocumentFile ([int]$row.id) ([string]$row.stored_path) ([string]$row.original_name) ([string]$row.file_hash)
            $idx = $docsGrid.Rows.Add(
                [string]$row.document_type,
                [string]$row.supplier,
                [string]$resolvedDoc.Name,
                [string]$row.uploaded_at,
                [string]$resolvedDoc.Path
            )
            $docsGrid.Rows[$idx].Tag = [pscustomobject]@{
                Id = [int]$row.id
                Path = [string]$resolvedDoc.Path
                Name = [string]$resolvedDoc.Name
                FileHash = [string]$resolvedDoc.FileHash
            }
        }
    }

    function Refresh-PurchaseDetails {
        Load-SelectedDealInfo
        Refresh-PurchaseSuppliers
        Refresh-PurchaseDocuments
    }

    function Refresh-PurchaseAll {
        Refresh-PurchaseDeals
        Refresh-PurchaseDetails
    }

    function Refresh-PurchaseDealsDeferred {
        if ($null -ne $form -and -not $form.IsDisposed -and $form.IsHandleCreated) {
            [void]$form.BeginInvoke([Action]{
                try {
                    Refresh-PurchaseDeals
                } catch {
                    [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Обновить сделки') | Out-Null
                }
            })
        } else {
            Refresh-PurchaseDeals
        }
    }

    function Refresh-PurchaseDetailsDeferred {
        if ($null -ne $form -and -not $form.IsDisposed -and $form.IsHandleCreated) {
            [void]$form.BeginInvoke([Action]{
                try {
                    Refresh-PurchaseDetails
                } catch {
                    [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Обновить контроль закупки') | Out-Null
                }
            })
        } else {
            Refresh-PurchaseDetails
        }
    }

    function Get-ComponentPriorityRank {
        param([string]$Priority)
        switch ($Priority) {
            'Срочно' { return 0 }
            '1' { return 1 }
            '2' { return 2 }
            '3' { return 3 }
            '4' { return 4 }
            '5' { return 5 }
            default { return 99 }
        }
    }

    function Get-ComponentStatusRank {
        param([string]$Status)
        switch ($Status) {
            'В работе' { return 0 }
            'Ожидание ответа' { return 1 }
            'На отслеживании' { return 2 }
            'Заказано' { return 3 }
            'Подано в оплату' { return 4 }
            'Выполнено' { return 5 }
            'Не актуально' { return 6 }
            default { return 99 }
        }
    }

    function Apply-PriorityCellStyle {
        param($Row)

        if ($null -eq $Row -or $null -eq $Row.DataGridView -or -not $Row.DataGridView.Columns.Contains('Priority')) { return }
        $priority = [string]$Row.Cells['Priority'].Value
        $priorityColor = switch ($priority) {
            'Срочно' { [Drawing.Color]::FromArgb(235, 70, 75) }
            '1' { [Drawing.Color]::FromArgb(255, 176, 176) }
            '2' { [Drawing.Color]::FromArgb(255, 218, 179) }
            '3' { [Drawing.Color]::FromArgb(255, 245, 215) }
            '4' { [Drawing.Color]::FromArgb(221, 235, 255) }
            '5' { [Drawing.Color]::FromArgb(226, 238, 226) }
            default { Get-UiColor 'Surface' }
        }
        $Row.Cells['Priority'].Style.BackColor = $priorityColor
        if ($priority -eq 'Срочно') {
            $Row.Cells['Priority'].Style.ForeColor = [Drawing.Color]::White
            $Row.Cells['Priority'].Style.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
        } else {
            $Row.Cells['Priority'].Style.ForeColor = Get-UiColor 'Text'
            $Row.Cells['Priority'].Style.Font = New-Object Drawing.Font('Segoe UI', 9)
        }
    }

    function Apply-ComponentRowStyle {
        param($Row)

        if ($null -eq $Row) { return }
        $status = [string]$Row.Cells['Status'].Value
        $rowColor = switch ($status) {
            'В работе' { [Drawing.Color]::FromArgb(255, 226, 226) }
            'Ожидание ответа' { [Drawing.Color]::FromArgb(255, 245, 215) }
            'На отслеживании' { [Drawing.Color]::FromArgb(221, 235, 255) }
            'Заказано' { [Drawing.Color]::FromArgb(221, 246, 228) }
            'Подано в оплату' { [Drawing.Color]::FromArgb(214, 232, 255) }
            'Выполнено' { [Drawing.Color]::FromArgb(198, 239, 206) }
            default { Get-UiColor 'Surface' }
        }

        $Row.DefaultCellStyle.BackColor = $rowColor
        foreach ($cell in $Row.Cells) {
            $cell.Style.BackColor = $rowColor
            $cell.Style.ForeColor = Get-UiColor 'Text'
        }
        if ($Row.DataGridView.Columns.Contains('ActualReceiptDate') -and -not [string]::IsNullOrWhiteSpace([string]$Row.Cells['ActualReceiptDate'].Value)) {
            $Row.Cells['ActualReceiptDate'].Style.BackColor = Get-UiColor 'Success'
            $Row.Cells['ActualReceiptDate'].Style.ForeColor = Get-UiColor 'Text'
            $Row.Cells['ActualReceiptDate'].Style.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
        }

        $priority = [string]$Row.Cells['Priority'].Value
        $priorityColor = switch ($priority) {
            'Срочно' { [Drawing.Color]::FromArgb(235, 70, 75) }
            '1' { [Drawing.Color]::FromArgb(255, 176, 176) }
            '2' { [Drawing.Color]::FromArgb(186, 213, 255) }
            '3' { [Drawing.Color]::FromArgb(207, 226, 255) }
            '4' { [Drawing.Color]::FromArgb(224, 228, 235) }
            '5' { [Drawing.Color]::FromArgb(235, 238, 243) }
            default { $rowColor }
        }
        $Row.Cells['Priority'].Style.BackColor = $priorityColor
        if ($priority -eq 'Срочно') {
            $Row.Cells['Priority'].Style.ForeColor = [Drawing.Color]::White
            $Row.Cells['Priority'].Style.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
        } else {
            $Row.Cells['Priority'].Style.ForeColor = Get-UiColor 'Text'
            $Row.Cells['Priority'].Style.Font = New-Object Drawing.Font('Segoe UI', 9)
        }
    }

    function Apply-ComponentReminderCellStyle {
        param($Row)

        if ($null -eq $Row -or $null -eq $Row.DataGridView -or -not $Row.DataGridView.Columns.Contains('ReminderDate')) { return }
        $reminderCell = $Row.Cells['ReminderDate']
        $reminderCell.Style.Font = New-Object Drawing.Font('Segoe UI', 9)
        $reminderDate = Convert-PurchaseDateText ([string]$reminderCell.Value)
        if ($null -ne $reminderDate -and $reminderDate.Date -le (Get-Date).Date) {
            $reminderCell.Style.BackColor = [Drawing.Color]::Red
            $reminderCell.Style.ForeColor = [Drawing.Color]::White
            $reminderCell.Style.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
        }
    }

    function Get-SelectedComponentRow {
        if ($componentsGrid.SelectedRows.Count -gt 0) { return $componentsGrid.SelectedRows[0] }
        if ($null -ne $componentsGrid.CurrentRow) { return $componentsGrid.CurrentRow }
        return $null
    }

    function Load-SelectedComponentNotes {
        $row = Get-SelectedComponentRow
        $script:LoadingComponentNotes = $true
        try {
            if ($null -eq $row -or $null -eq $row.Tag) {
                $script:CurrentComponentNotesId = 0
                $script:CurrentComponentNotesText = ''
                $lblComponentNotesTitle.Text = 'Записи по сделке'
                $txtComponentNotes.Text = ''
                $txtComponentNotes.Enabled = $false
                $btnSaveComponentNotes.Enabled = $false
                return
            }

            $id = [int]$row.Tag
            $dealNumber = ([string]$row.Cells['DealNumber'].Value).Trim()
            if ([string]::IsNullOrWhiteSpace($dealNumber)) { $dealNumber = "задача #$id" }
            $notes = Get-ComponentDealNotes $id
            $script:CurrentComponentNotesId = $id
            $script:CurrentComponentNotesText = $notes
            $lblComponentNotesTitle.Text = "Записи по сделке: $dealNumber"
            $txtComponentNotes.Text = $notes
            $txtComponentNotes.Enabled = $true
            $btnSaveComponentNotes.Enabled = $true
        } finally {
            $script:LoadingComponentNotes = $false
        }
    }

    function Save-CurrentComponentNotes {
        param([bool]$ShowMessage = $false)

        if ($script:LoadingComponentNotes -or $script:CurrentComponentNotesId -le 0) { return }
        $notes = [string]$txtComponentNotes.Text
        if ($notes -eq $script:CurrentComponentNotesText) {
            if ($ShowMessage) {
                [Windows.Forms.MessageBox]::Show('Изменений нет.', 'Задачи') | Out-Null
            }
            return
        }

        Set-ComponentDealNotes ([int]$script:CurrentComponentNotesId) $notes
        $script:CurrentComponentNotesText = $notes
        if ($ShowMessage) {
            [Windows.Forms.MessageBox]::Show('Записи сохранены.', 'Задачи') | Out-Null
        }
    }
    function Refresh-Components {
        Save-CurrentComponentNotes
        $selectedId = if ($componentsGrid.SelectedRows.Count -gt 0 -and $null -ne $componentsGrid.SelectedRows[0].Tag) { [int]$componentsGrid.SelectedRows[0].Tag } else { 0 }
        $script:ComponentsLoading = $true
        try {
            $componentsGrid.Rows.Clear()
            $table = Get-ComponentDeals ([string]$txtComponentsSearch.Text)
            foreach ($row in $table.Rows) {
                [void](Ensure-ComponentDealFolder ([int]$row.id) ([string]$row.deal_number))
                $idx = $componentsGrid.Rows.Add(
                    (Format-PurchaseDate $row.entry_date),
                    [string]$row.deal_number,
                    [string]$row.status,
                    [string]$row.stage,
                    [string]$row.description,
                    [string]$row.next_action,
                    (Format-PurchaseDate $row.reminder_date),
                    (Format-PurchaseDate $row.deadline_date),
                    [string]$row.priority,
                    [string]$row.period
                )
                $componentsGrid.Rows[$idx].Tag = [int]$row.id
                Apply-ComponentRowStyle $componentsGrid.Rows[$idx]
                Apply-ComponentReminderCellStyle $componentsGrid.Rows[$idx]
            }

            if ($componentsGrid.Rows.Count -gt 0) {
                $target = $componentsGrid.Rows[0]
                foreach ($gridRow in $componentsGrid.Rows) {
                    if ($selectedId -gt 0 -and [int]$gridRow.Tag -eq $selectedId) {
                        $target = $gridRow
                        break
                    }
                }
                $componentsGrid.ClearSelection()
                $target.Selected = $true
                $componentsGrid.CurrentCell = $target.Cells[0]
            }
        } finally {
            $script:ComponentsLoading = $false
        }
        Load-SelectedComponentNotes
    }

    function Get-SelectedReminderInfo {
        if ($remindersGrid.SelectedRows.Count -eq 0 -or $null -eq $remindersGrid.SelectedRows[0].Tag) { return $null }
        return $remindersGrid.SelectedRows[0].Tag
    }

    function Open-SelectedReminderTarget {
        $info = Get-SelectedReminderInfo
        if ($null -eq $info -or [int]$info.DealId -le 0) {
            throw 'В заявке отсутствует или поврежден номер.'
        }

        $source = ([string]$info.Source).Trim().ToLowerInvariant()
        if ($source -eq 'component' -or $source -eq 'task' -or $source -eq 'components') {
            Refresh-Components
            $target = $null
            foreach ($gridRow in $componentsGrid.Rows) {
                if ([int]$gridRow.Tag -eq [int]$info.DealId) { $target = $gridRow; break }
            }
            if ($null -eq $target) {
                $txtComponentsSearch.Text = ''
                Refresh-Components
                foreach ($gridRow in $componentsGrid.Rows) {
                    if ([int]$gridRow.Tag -eq [int]$info.DealId) { $target = $gridRow; break }
                }
            }
            if ($null -eq $target) { throw 'Нужная сделка не найдена в таблице "Сделки".' }
            Show-AppPage 'Components'
            $componentsGrid.ClearSelection()
            $target.Selected = $true
            $componentsGrid.CurrentCell = $target.Cells[0]
            $target.DataGridView.FirstDisplayedScrollingRowIndex = $target.Index
            return
        }

        $target = $null
        foreach ($gridRow in $dealsGrid.Rows) {
            if ([int]$gridRow.Tag -eq [int]$info.DealId) { $target = $gridRow; break }
        }
        if ($null -eq $target) {
            # A reminder can point to a deal hidden by the current filter or
            # search text. Reset both before looking it up.
            $txtDealSearch.Text = ''
            $cmbDealFilter.SelectedItem = 'Все включая архив'
            Refresh-PurchaseDeals
            foreach ($gridRow in $dealsGrid.Rows) {
                if ([int]$gridRow.Tag -eq [int]$info.DealId) { $target = $gridRow; break }
            }
        }
        if ($null -eq $target) { throw 'Нужная закупка не найдена в таблице "Контроль закупки".' }
        Show-AppPage 'Purchase'
        $dealsGrid.ClearSelection()
        $target.Selected = $true
        $dealsGrid.CurrentCell = $target.Cells[0]
        $target.DataGridView.FirstDisplayedScrollingRowIndex = $target.Index
        Refresh-PurchaseDetails
    }

    function Confirm-SelectedReminderReceipt {
        $info = Get-SelectedReminderInfo
        if ($null -eq $info -or [int]$info.SupplierId -le 0 -or [string]$info.Source -ne 'auto' -or -not ([string]$info.Title).ToLowerInvariant().Contains('поступление')) {
            throw 'Выберите автоматическое напоминание по поступлению.'
        }

        $selectedDate = Show-DatePickerDialog ((Get-Date).ToString('dd.MM.yyyy'))
        if ($null -eq $selectedDate -or [string]::IsNullOrWhiteSpace([string]$selectedDate)) { return }

        Set-PurchaseSupplierActualReceiptDate ([int]$info.SupplierId) ([string]$selectedDate)
        Refresh-PurchaseSuppliers
        Refresh-PurchaseDeals
        Refresh-Reminders
        Refresh-History
        [Windows.Forms.MessageBox]::Show(("Поступление подтверждено`r`n" + ('Факт: {0}' -f [string]$selectedDate)), 'Напоминания') | Out-Null
    }
    function Refresh-Reminders {
        $remindersGrid.Rows.Clear()
        $search = ([string]$txtRemindersSearch.Text).Trim()
        foreach ($item in @(Get-PurchaseActionItems)) {
            if (-not [string]::IsNullOrWhiteSpace($search)) {
                $dealIdText = [string]$item.DealId
                $titleText = [string]$item.Title
                if ($titleText.IndexOf($search, [StringComparison]::OrdinalIgnoreCase) -lt 0 -and $dealIdText.IndexOf($search, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
                    continue
                }
            }
            $idx = $remindersGrid.Rows.Add(
                [string]$item.Severity,
                [string]$item.DueDate,
                [string]$item.Title,
                [string]$item.Source
            )
            $remindersGrid.Rows[$idx].Tag = [pscustomobject]@{
                ReminderId = if ($null -ne $item.PSObject.Properties['ReminderId']) { [int]$item.ReminderId } else { 0 }
                SupplierId = if ($null -ne $item.PSObject.Properties['SupplierId']) { [int]$item.SupplierId } else { 0 }
                DealId = if ($null -ne $item.PSObject.Properties['DealId']) { [int]$item.DealId } else { 0 }
                Title = [string]$item.Title
                Source = [string]$item.Source
                Severity = [string]$item.Severity
            }
            $color = switch ([string]$item.Severity) {
                'Danger' { Get-UiColor 'Danger' }
                'Attention' { Get-UiColor 'Attention' }
                'Warn' { Get-UiColor 'Warn' }
                default { Get-UiColor 'Info' }
            }
            $remindersGrid.Rows[$idx].DefaultCellStyle.BackColor = $color
        }
    }

    function Refresh-GlobalistQuotes {
        $globalistGrid.Rows.Clear()
        $table = Get-GlobalistQuotes $txtQuoteSearch.Text 20000
        foreach ($row in $table.Rows) {
            [void]$globalistGrid.Rows.Add(
                [string]$row.imported_at, [string]$row.factory, [string]$row.pn, [string]$row.comment,
                [string]$row.pi_number, [string]$row.replacement, [string]$row.chinese_remark, [string]$row.package,
                [string]$row.brand, [string]$row.datacode, [string]$row.moq, [string]$row.qty, [string]$row.stock,
                [string]$row.need_spq, [string]$row.spq,
                $(if ($null -eq $row.unit_price -or $row.unit_price -is [DBNull]) { '' } else { Format-Price $row.unit_price }),
                $(if ($null -eq $row.total_amount -or $row.total_amount -is [DBNull]) { '' } else { Format-Price $row.total_amount }),
                [string]$row.lead_time, [string]$row.weight, [string]$row.target, [string]$row.supplier_quote_id, [string]$row.sheet_name
            )
        }
    }
    function Refresh-QuoteBase {
        $quoteBaseGrid.Rows.Clear()
        $dateFrom = Convert-PurchaseDateForSort $txtQuoteDateFrom.Text
        $dateTo = Convert-PurchaseDateForSort $txtQuoteDateTo.Text
        $table = Get-QuoteHistory $txtQuoteSearch.Text $dateFrom $dateTo 2000
        foreach ($row in $table.Rows) {
            $idx = $quoteBaseGrid.Rows.Add(
                [string]$row.quote_date,
                [string]$row.rfq_value,
                [string]$row.pn,
                [string]$row.supplier,
                $(if ($null -eq $row.unit_price -or $row.unit_price -is [DBNull]) { '' } else { Format-Price $row.unit_price }),
                [string]$row.lead_time,
                [string]$row.mfg,
                $(if (Convert-DbBool $row.is_winner) { 'Да' } else { '' }),
                [string]$row.winner_reason,
                [string]$row.warning
            )
            $quoteBaseGrid.Rows[$idx].Tag = [int]$row.id
            if (Convert-DbBool $row.is_winner) {
                $quoteBaseGrid.Rows[$idx].DefaultCellStyle.BackColor = Get-UiColor 'Success'
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$row.warning)) {
                $quoteBaseGrid.Rows[$idx].Cells['Warning'].Style.BackColor = Get-UiColor 'Danger'
            }
        }
    }

    function Refresh-History {
        param([int]$DealId = 0)
        $historyGrid.Rows.Clear()
        $table = Get-ActivityLog $DealId 500 ([string]$txtHistorySearch.Text)
        foreach ($row in $table.Rows) {
            $idx = $historyGrid.Rows.Add(
                [string]$row.created_at,
                [string]$row.action,
                [string]$row.details,
                [string]$row.entity_type
            )
            $historyGrid.Rows[$idx].Tag = [int]$row.id
        }
    }

    function Export-QuoteBaseToExcel {
        $dateFrom = Convert-PurchaseDateForSort $txtQuoteDateFrom.Text
        $dateTo = Convert-PurchaseDateForSort $txtQuoteDateTo.Text
        $table = Get-QuoteHistory $txtQuoteSearch.Text $dateFrom $dateTo 100000
        if ($table.Rows.Count -eq 0) { throw 'Нет строк для экспорта.' }
        $path = Join-Path (Get-PurchaseDataRoot) ("quote_history_export_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $excel = Open-ExcelApp
        $wb = $null
        try {
            $wb = $excel.Workbooks.Add()
            $ws = $wb.Worksheets.Item(1)
            $headers = @('Дата квоты','Value','PN','Поставщик','Цена','Срок','Lead time total','MFG','Победитель','Почему выбран','Предупреждение','RFQ')
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $ws.Cells.Item(1, $c + 1).Value2 = $headers[$c]
            }
            $r = 2
            foreach ($row in $table.Rows) {
                $values = @(
                    [string]$row.quote_date,
                    [string]$row.rfq_value,
                    [string]$row.pn,
                    [string]$row.supplier,
                    $(if ($null -eq $row.unit_price -or $row.unit_price -is [DBNull]) { '' } else { [double]$row.unit_price }),
                    [string]$row.lead_time,
                    [string]$row.lead_time_total,
                    [string]$row.mfg,
                    $(if (Convert-DbBool $row.is_winner) { 'Да' } else { '' }),
                    [string]$row.winner_reason,
                    [string]$row.warning,
                    [string]$row.rfq_path
                )
                for ($c = 0; $c -lt $values.Count; $c++) {
                    $ws.Cells.Item($r, $c + 1).Value2 = $values[$c]
                }
                $r++
            }
            [void]$ws.Columns.AutoFit() | Out-Null
            $wb.SaveAs($path)
        } finally {
            if ($null -ne $wb) {
                $wb.Close($true)
                Release-ComObject $wb
            }
        }
        return $path
    }

    function Remove-SelectedQuoteBaseRows {
        if ($quoteBaseGrid.SelectedRows.Count -eq 0) { throw 'Выделите одну или несколько квот.' }
        $ids = New-Object System.Collections.ArrayList
        foreach ($row in $quoteBaseGrid.SelectedRows) {
            if ($null -ne $row.Tag) { [void]$ids.Add([int]$row.Tag) }
        }
        if ($ids.Count -eq 0) { throw 'Не удалось определить выбранные квоты.' }
        $answer = [Windows.Forms.MessageBox]::Show(
            ("Удалить выбранные квоты: {0}?" -f $ids.Count),
            'База квот',
            [Windows.Forms.MessageBoxButtons]::YesNo,
            [Windows.Forms.MessageBoxIcon]::Question
        )
        if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
        [void](Remove-QuoteHistoryItems ([int[]]@($ids)))
        Refresh-QuoteBase
    }

    function Clear-AllQuoteBaseRows {
        $answer = [Windows.Forms.MessageBox]::Show(
            'Очистить всю базу квот? Это удобно для тестов, но действие нельзя отменить.',
            'База квот',
            [Windows.Forms.MessageBoxButtons]::YesNo,
            [Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
        Clear-QuoteHistory
        Refresh-QuoteBase
    }

    function Clear-AllHistoryRows {
        $answer = [Windows.Forms.MessageBox]::Show(
            'Очистить всю историю действий? Это действие нельзя отменить.',
            'История',
            [Windows.Forms.MessageBoxButtons]::YesNo,
            [Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
        Clear-ActivityLog
        Refresh-History
    }

    function Save-ComponentGridRow {
        param($Row)

        if ($script:ComponentsLoading -or $null -eq $Row -or $null -eq $Row.Tag) { return }
        $status = [string]$Row.Cells['Status'].Value
        if ([string]::IsNullOrWhiteSpace($status)) { $status = 'В работе' }
        $stage = ([string]$Row.Cells['Stage'].Value).Trim()
        $priority = [string]$Row.Cells['Priority'].Value
        if ([string]::IsNullOrWhiteSpace($priority)) { $priority = '3' }

        Update-ComponentDeal `
            ([int]$Row.Tag) `
            ([string]$Row.Cells['DealNumber'].Value) `
            $status `
            $stage `
            ([string]$Row.Cells['Description'].Value) `
            ([string]$Row.Cells['NextAction'].Value) `
            ([string]$Row.Cells['ReminderDate'].Value) `
            ([string]$Row.Cells['DeadlineDate'].Value) `
            $priority `
            ([string]$Row.Cells['Period'].Value)
        Apply-ComponentRowStyle $Row
        Apply-ComponentReminderCellStyle $Row
    }

    function Save-PurchaseSupplierGridRow {
        param($row, [bool]$AutoUpdateReceipt = $false)

        if ($script:PurchaseLoading -or $null -eq $row) { return }
        if ($null -eq $row.Tag) { return }
        if ($AutoUpdateReceipt) {
            Update-AutoReceiptDateForRow $row
        }
        $erpNotRequired = Test-SupplierErpNotRequired ([string]$row.Cells['Supplier'].Value)
        $erpSupplierSent = if ($erpNotRequired) { $false } else { [bool]$row.Cells['ErpSupplier'].Value }
        $erpRogerSent = if ($erpNotRequired) { $false } else { [bool]$row.Cells['ErpRoger'].Value }
        Update-PurchaseSupplier `
            ([int]$row.Tag) `
            ([bool]$row.Cells['InvoiceReceived'].Value) `
            ([bool]$row.Cells['InvoiceConfirmed'].Value) `
            $false `
            $erpSupplierSent `
            $erpRogerSent `
            ([string]$row.Cells['PiUsd'].Value) `
            ([string]$row.Cells['PiCny'].Value) `
            ([string]$row.Cells['PiRub'].Value) `
            ([string]$row.Cells['PaidAmount'].Value) `
            ([string]$row.Cells['DeliveryWeeks'].Value) `
            ([bool]$row.Cells['PaymentSubmitted'].Value) `
            ([bool]$row.Cells['Paid'].Value) `
            ([string]$row.Cells['InvoiceDate'].Value) `
            ([string]$row.Cells['ReceiptDate'].Value) `
            ([string]$row.Cells['ActualReceiptDate'].Value) `
            ([string]$row.Cells['Comment'].Value)
        Apply-SupplierRowStyle $row
        Apply-SupplierErpCellRules $row
    }

    function Save-PurchaseDealGridRow {
        param($row)

        if ($script:PurchaseLoading -or $null -eq $row -or $null -eq $row.Tag) { return }
        $dealNumber = ([string]$row.Cells['DealNumber'].Value).Trim()
        if ([string]::IsNullOrWhiteSpace($dealNumber)) {
            $dealNumber = ([string](Invoke-PurchaseScalar 'SELECT deal_number FROM deals WHERE id = @id' @{ '@id' = [int]$row.Tag })).Trim()
            if ([string]::IsNullOrWhiteSpace($dealNumber)) {
                return
            }
            $row.Cells['DealNumber'].Value = $dealNumber
        }
        $boardCount = ([string]$row.Cells['BoardCount'].Value).Trim()
        if (-not [string]::IsNullOrWhiteSpace($boardCount) -and $boardCount -notmatch '^\d+$') {
            throw 'В колонке "Кол-во плат" нужно указать целое число.'
        }

        $executorValue = ([string]$row.Cells['Executor'].Value).Trim()
        if ([string]::IsNullOrWhiteSpace($executorValue)) {
            $executorValue = ([string](Invoke-PurchaseScalar 'SELECT executor FROM deals WHERE id = @id' @{ '@id' = [int]$row.Tag })).Trim()
        }
        Update-PurchaseDeal `
            ([int]$row.Tag) `
            ([string]$row.Cells['Client'].Value) `
            ([string]$row.Cells['Stage'].Value) `
            ([string]$row.Cells['Comment'].Value) `
            $boardCount `
            $dealNumber `
            $null `
            ([string]$row.Cells['Period'].Value) `
            ([string]$row.Cells['Masks'].Value) `
            $null `
            ([string]$row.Cells['Priority'].Value) `
            $null `
            $executorValue `
            ([string]$row.Cells['AssemblyLocation'].Value)
        Load-SelectedDealInfo
    }

    function Save-SelectedPurchaseSupplierRow {
        if ($suppliersGrid.SelectedRows.Count -eq 0) { return }
        Save-PurchaseSupplierGridRow $suppliersGrid.SelectedRows[0] $false
    }

    function Get-SelectedPreviewFilter {
        switch ([string]$cmbFilter.SelectedItem) {
            'Только победители' { return 'Winners' }
            'Без квот' { return 'NoQuotes' }
            'Предупреждения' { return 'Warnings' }
            'Ручная проверка' { return 'Manual' }
            default { return 'All' }
        }
    }

    function Refresh-CurrentPreview {
        if ($null -eq $script:LastAnalysis) {
            return
        }

        $selectedId = ''
        if ($previewGrid.SelectedRows.Count -gt 0 -and $null -ne $previewGrid.SelectedRows[0].Tag) {
            $selectedId = [string]$previewGrid.SelectedRows[0].Tag
        }
        Refresh-PreviewGrid $script:LastAnalysis $previewGrid $logBox (Get-SelectedPreviewFilter) $txtSearch.Text
        if (-not [string]::IsNullOrWhiteSpace($selectedId)) {
            foreach ($row in $previewGrid.Rows) {
                if (-not $row.IsNewRow -and [string]$row.Tag -eq $selectedId) {
                    $previewGrid.ClearSelection()
                    $row.Selected = $true
                    $previewGrid.CurrentCell = $row.Cells[0]
                    break
                }
            }
        }
        Refresh-DecisionDetails $script:LastAnalysis $previewGrid $detailsBox $quoteDetailsGrid
    }

    function Copy-SelectedPurchaseDealNumber {
        try {
            $row = Get-SelectedDealRow
            if ($null -eq $row) { throw 'Выберите сделку.' }
            $dealNumber = ([string]$row.Cells['DealNumber'].Value).Trim()
            if ([string]::IsNullOrWhiteSpace($dealNumber) -and $null -ne $row.Tag) {
                $dealNumber = ([string](Invoke-PurchaseScalar 'SELECT deal_number FROM deals WHERE id = @id' @{ '@id' = [int]$row.Tag })).Trim()
            }
            if ([string]::IsNullOrWhiteSpace($dealNumber)) { throw 'Номер сделки пустой.' }
            [Windows.Forms.Clipboard]::SetText($dealNumber)
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Копировать номер сделки') | Out-Null
        }
    }

    function Open-SelectedPurchaseDocumentInExplorer {
        try {
            $info = Get-SelectedDocumentInfo
            if ($null -eq $info) { throw 'Выберите документ.' }
            $name = if ($info.PSObject.Properties['Name']) { [string]$info.Name } else { '' }
            $hash = if ($info.PSObject.Properties['FileHash']) { [string]$info.FileHash } else { '' }
            $resolvedDoc = Resolve-PurchaseDocumentFile ([int]$info.Id) ([string]$info.Path) $name $hash
            $path = [string]$resolvedDoc.Path
            if ([bool]$resolvedDoc.Changed) { Refresh-PurchaseDocuments }
            if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path -PathType Leaf)) {
                Start-Process -FilePath explorer.exe -ArgumentList ('/select,"{0}"' -f $path)
                return
            }
            $directory = if ([string]::IsNullOrWhiteSpace($path)) { '' } else { [IO.Path]::GetDirectoryName($path) }
            if (-not [string]::IsNullOrWhiteSpace($directory) -and (Test-Path -LiteralPath $directory -PathType Container)) {
                Start-Process -FilePath explorer.exe -ArgumentList ('"{0}"' -f $directory)
                return
            }
            throw "Файл не найден:`r`n$path"
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Открыть в проводнике') | Out-Null
        }
    }

    function Remove-SelectedPurchaseDeal {
        try {
            $dealId = Get-SelectedDealId
            if ($dealId -le 0) { throw 'Выберите сделку.' }
            $row = Get-SelectedDealRow
            $dealNumber = if ($null -ne $row) { [string]$row.Cells['DealNumber'].Value } else { '' }
            $answer = [Windows.Forms.MessageBox]::Show(
                "Удалить сделку, всех поставщиков и загруженные документы?`r`n$dealNumber",
                'Удалить сделку',
                [Windows.Forms.MessageBoxButtons]::YesNo,
                [Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }

            Delete-PurchaseDeal $dealId
            Refresh-PurchaseAll
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Удалить сделку') | Out-Null
        }
    }

    function Toggle-SelectedPurchaseDealArchive {
        try {
            $dealId = Get-SelectedDealId
            if ($dealId -le 0) { throw 'Выберите сделку.' }
            $current = Invoke-PurchaseScalar 'SELECT archived FROM deals WHERE id = @id' @{ '@id' = $dealId }
            $archive = -not (Convert-DbBool $current)
            Set-PurchaseDealArchived $dealId $archive
            Refresh-PurchaseAll
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Архив') | Out-Null
        }
    }

    function Get-SelectedPreviewDecision {
        if ($null -eq $script:LastAnalysis -or $previewGrid.SelectedRows.Count -eq 0) {
            return $null
        }

        return Get-DecisionById $script:LastAnalysis ([string]$previewGrid.SelectedRows[0].Tag)
    }

    function Set-SelectedQuoteAsWinner {
        try {
            $decision = Get-SelectedPreviewDecision
            if ($null -eq $decision) {
                throw 'Выберите строку preview.'
            }
            if ($quoteDetailsGrid.SelectedRows.Count -eq 0 -or $null -eq $quoteDetailsGrid.SelectedRows[0].Tag) {
                throw 'Выберите квоту в таблице справа.'
            }

            Update-DecisionsFromGrid $script:LastAnalysis $previewGrid
            Set-RfqRowManualWinner $script:LastAnalysis $decision.Id ([string]$quoteDetailsGrid.SelectedRows[0].Tag)
            Refresh-CurrentPreview
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Ручной выбор квоты') | Out-Null
        }
    }

    function Compare-ExcelWorkbooks {
        param([string]$FirstPath, [string]$SecondPath, [string]$OutputPath)
        if (-not (Test-Path -LiteralPath $FirstPath -PathType Leaf)) { throw 'Первый Excel-файл не найден.' }
        if (-not (Test-Path -LiteralPath $SecondPath -PathType Leaf)) { throw 'Второй Excel-файл не найден.' }
        $excel = Open-ExcelApp
        $first = $null
        $second = $null
        $result = $null
        $diffCount = 0
        try {
            $excel.ScreenUpdating = $false
            $excel.DisplayAlerts = $false
            $first = $excel.Workbooks.Open($FirstPath, 0, $true)
            $second = $excel.Workbooks.Open($SecondPath, 0, $true)
            New-Item -ItemType Directory -Force -Path (Get-PurchaseDataRoot) | Out-Null
            Copy-Item -LiteralPath $FirstPath -Destination $OutputPath -Force
            $result = $excel.Workbooks.Open($OutputPath, 0, $false)
            foreach ($sheet in $result.Worksheets) {
                $other = $null
                try { $other = $second.Worksheets.Item([string]$sheet.Name) } catch { $other = $null }
                $used = $sheet.UsedRange
                $firstRow = [int]$used.Row
                $firstCol = [int]$used.Column
                $firstLastRow = $firstRow + [int]$used.Rows.Count - 1
                $firstLastCol = $firstCol + [int]$used.Columns.Count - 1
                if ($null -eq $other) {
                    $sheet.Range($sheet.Cells.Item($firstRow, $firstCol), $sheet.Cells.Item($firstLastRow, $firstLastCol)).Interior.Color = 255
                    $diffCount += [int]$used.Rows.Count * [int]$used.Columns.Count
                    continue
                }
                $otherUsed = $other.UsedRange
                $secondRow = [int]$otherUsed.Row
                $secondCol = [int]$otherUsed.Column
                $secondLastRow = $secondRow + [int]$otherUsed.Rows.Count - 1
                $secondLastCol = $secondCol + [int]$otherUsed.Columns.Count - 1
                $maxRow = [Math]::Max($firstLastRow, $secondLastRow)
                $maxCol = [Math]::Max($firstLastCol, $secondLastCol)
                $firstValues = $used.Value2
                $secondValues = $otherUsed.Value2
                $firstIsArray = ($firstValues -is [Array])
                $secondIsArray = ($secondValues -is [Array])
                for ($r = [Math]::Min($firstRow, $secondRow); $r -le $maxRow; $r++) {
                    for ($c = [Math]::Min($firstCol, $secondCol); $c -le $maxCol; $c++) {
                        $a = $null
                        $b = $null
                        if ($r -ge $firstRow -and $r -le $firstLastRow -and $c -ge $firstCol -and $c -le $firstLastCol) {
                            if ($firstIsArray) { $a = $firstValues.GetValue($r - $firstRow + 1, $c - $firstCol + 1) } elseif ($r -eq $firstRow -and $c -eq $firstCol) { $a = $firstValues }
                        }
                        if ($r -ge $secondRow -and $r -le $secondLastRow -and $c -ge $secondCol -and $c -le $secondLastCol) {
                            if ($secondIsArray) { $b = $secondValues.GetValue($r - $secondRow + 1, $c - $secondCol + 1) } elseif ($r -eq $secondRow -and $c -eq $secondCol) { $b = $secondValues }
                        }
                        $aText = if ($null -eq $a) { '' } else { ([string]$a).Trim() }
                        $bText = if ($null -eq $b) { '' } else { ([string]$b).Trim() }
                        if ($aText -ne $bText) { $sheet.Cells.Item($r, $c).Interior.Color = 255; $diffCount++ }
                    }
                }
            }
            $result.SaveAs($OutputPath, 51)
            return $diffCount
        } finally {
            if ($null -ne $result) { $result.Close($false); Release-ComObject $result }
            if ($null -ne $second) { $second.Close($false); Release-ComObject $second }
            if ($null -ne $first) { $first.Close($false); Release-ComObject $first }
            $excel.ScreenUpdating = $true
        }
    }    $btnNavExcelCompare.Add_Click({ Show-AppPage 'ExcelCompare' })
    $btnExcelPick1.Add_Click({
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*'
        if ($dialog.ShowDialog($form) -eq [Windows.Forms.DialogResult]::OK) { $excelCompareFile1.Text = $dialog.FileName }
        $dialog.Dispose()
    })
    $btnExcelPick2.Add_Click({
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*'
        if ($dialog.ShowDialog($form) -eq [Windows.Forms.DialogResult]::OK) { $excelCompareFile2.Text = $dialog.FileName }
        $dialog.Dispose()
    })
    $btnExcelCompare.Add_Click({
        try {
            $output = Join-Path (Get-PurchaseDataRoot) ('excel_compare_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            $count = Compare-ExcelWorkbooks $excelCompareFile1.Text $excelCompareFile2.Text $output
            $excelCompareStatus.Text = "Готово. Несовпадений: $count. Результат: $output"
            [Windows.Forms.MessageBox]::Show($form, "Несовпадений: $count`r`nРезультат сохранен:`r`n$output", 'Сравнение Excel') | Out-Null
        } catch { [Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, 'Сравнение Excel') | Out-Null }
    })
    $btnNavRrfq.Add_Click({ Show-AppPage 'RRFQ' })
    $btnNavBitrix.Add_Click({ Show-AppPage 'Bitrix'; if ($script:BitrixIntegrationEnabled -and $bitrixGrid.Rows.Count -eq 0) { $btnBitrixRefresh.PerformClick() } })
    $btnNavPurchase.Add_Click({
        Refresh-PurchaseAll
        Show-AppPage 'Purchase'
        Resize-PurchasePanels
    })
    $btnNavComponents.Add_Click({
        Refresh-Components
        Show-AppPage 'Components'
    })
    $btnNavReminders.Add_Click({
        Refresh-Reminders
        Show-AppPage 'Reminders'
    })
    $btnNavQuoteBase.Add_Click({
        Refresh-QuoteBase
        Show-AppPage 'QuoteBase'
    })
    $btnNavHistory.Add_Click({
        Refresh-History
        Show-AppPage 'History'
    })
    $priceSearchGrid.Add_CellDoubleClick({
        param($sender, $eventArgs)
        if ($eventArgs.RowIndex -lt 0 -or $eventArgs.RowIndex -ge $priceSearchGrid.Rows.Count) { return }
        $link = [string]$priceSearchGrid.Rows[$eventArgs.RowIndex].Cells['Link'].Value
        if (-not [string]::IsNullOrWhiteSpace($link) -and $link -match '^https?://') {
            try { Start-Process $link } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Открытие ссылки') | Out-Null }
        }
    })
    $btnPriceSearch.Add_Click({
        try {
            $pn=[string]$txtPriceSearchPn.Text
            if([string]::IsNullOrWhiteSpace($pn)){throw 'Укажите PN.'}
            $quantity=0
            if(-not [string]::IsNullOrWhiteSpace($txtPriceSearchQty.Text) -and -not [int]::TryParse($txtPriceSearchQty.Text,[ref]$quantity)){throw 'Количество должно быть целым числом.'}
            $priceSearchGrid.Rows.Clear()
            $btnPriceSearch.Enabled=$false
            foreach($item in @(Get-PriceSearchResults $pn $quantity)){
                $price = if($null -eq $item.Price -or [string]::IsNullOrWhiteSpace([string]$item.Price)){''}else{Format-Price $item.Price}
                [void]$priceSearchGrid.Rows.Add([string]$item.Source,[string]$item.PN,[string]$item.Manufacturer,$price,[string]$item.Currency,[string]$item.MOQ,[string]$item.Stock,[string]$item.LeadTime,[string]$item.Date,[string]$item.Link)
            }
            if($priceSearchGrid.Rows.Count -eq 0){[Windows.Forms.MessageBox]::Show('По этому PN предложения не найдены.','Поиск цены')|Out-Null}
        } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Поиск цены')|Out-Null } finally { $btnPriceSearch.Enabled=$true }
    })
    $btnPriceSearchGpt.Add_Click({
        try {
            $pn=[string]$txtPriceSearchPn.Text
            if([string]::IsNullOrWhiteSpace($pn)){throw 'Укажите PN.'}
            $quantity=0; [void][int]::TryParse($txtPriceSearchQty.Text,[ref]$quantity)
            $prompt=Get-PriceSearchGptPrompt $pn $quantity
            [Windows.Forms.Clipboard]::SetText($prompt)
            Start-Process 'https://chatgpt.com/'
            [Windows.Forms.MessageBox]::Show('Запрос отправлен в чат GPT.','Поиск в GPT')|Out-Null
        } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Поиск в GPT')|Out-Null }
    })
    $btnNavCompelParser.Add_Click({ Show-AppPage 'CompelParser' })
    $btnNavPriceSearch.Add_Click({ Show-AppPage 'PriceSearch' })
    $btnNavNotes.Add_Click({
        Refresh-NotesList (Get-NotesFilePath)
        Show-AppPage 'Notes'
    })
    $btnNavInstructions.Add_Click({ Show-AppPage 'Instructions' })
    $btnNavSettings.Add_Click({
        $txtDocumentsRoot.Text = Get-PurchaseDocumentsRoot
        Show-AppPage 'Settings'
    })

    $btnSaveDealInfo.Add_Click({
        try {
            if ($dealsGrid.IsCurrentCellInEditMode -or $dealsGrid.IsCurrentRowDirty) {
                [void]$dealsGrid.CancelEdit()
            }
            $dealId = Get-SelectedDealId
            if ($dealId -le 0) { throw 'Выберите сделку.' }
            $currentDeal = Invoke-PurchaseQuery 'SELECT deal_number, board_count, priority, masks, executor, assembly_location, tracking_status FROM deals WHERE id = @id' @{ '@id' = $dealId }
            if ($currentDeal.Rows.Count -eq 0) { throw 'Сделка не найдена.' }
            $currentDealRow = $currentDeal.Rows[0]
            $masksValue = $null
            if (-not [DBNull]::Value.Equals($currentDealRow.masks)) {
                $masksValue = [string]$currentDealRow.masks
            }
            Update-PurchaseDeal `
                $dealId $txtDealClient.Text ([string]$cmbDealStatusEdit.SelectedItem) $txtDealComment.Text `
                ([string]$currentDealRow.board_count) ([string]$currentDealRow.deal_number) $null $cmbDealPeriodEdit.Text `
                $masksValue $null `
                ([string]$currentDealRow.priority) ([string]$currentDealRow.tracking_status) `
                ([string]$currentDealRow.executor) ([string]$currentDealRow.assembly_location)
            Refresh-PurchaseDeals
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Сохранить сделку') | Out-Null
        }
    })

    $btnExtractPiAmounts.Add_Click({
        try {
            $supplierId = Get-SelectedSupplierId
            if ($supplierId -le 0) { throw 'Выберите поставщика.' }
            $supplierName = [string](Get-PurchaseSupplierName $supplierId)
            $docs = Get-PurchaseSupplierPiDocuments $supplierId
            if ($docs.Rows.Count -eq 0) {
                throw 'У выбранного поставщика нет загруженных документов типа PI/Invoice.'
            }

            $messages = New-Object System.Collections.ArrayList
        $rrfqFiles = New-Object System.Collections.ArrayList
            foreach ($docRow in $docs.Rows) {
                $resolvedDoc = Resolve-PurchaseDocumentFile ([int]$docRow.id) ([string]$docRow.stored_path) ([string]$docRow.original_name) ([string]$docRow.file_hash)
                $path = [string]$resolvedDoc.Path
                if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                    [void]$messages.Add("Файл не найден: $([string]$docRow.original_name)")
                    continue
                }
                $result = Import-PiAmountsFromPdfForSupplier $supplierId $supplierName $path
                if (-not [string]::IsNullOrWhiteSpace([string]$result.Message)) {
                    [void]$messages.Add("$([string]$docRow.original_name): $([string]$result.Message)")
                }
                if ([bool]$result.Updated) {
                    break
                }
            }

            Refresh-PurchaseAll
            [Windows.Forms.MessageBox]::Show(([string]::Join("`r`n", @($messages))), 'Считать PI') | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Считать PI') | Out-Null
        }
    })

    Add-DebouncedTextChanged $txtDealSearch { Refresh-PurchaseAll }
    $cmbDealFilter.Add_SelectedIndexChanged({ Refresh-PurchaseAll })
    $btnRefreshPurchase.Add_Click({ Refresh-PurchaseAll })
    $btnArchiveDeal.Add_Click({ Toggle-SelectedPurchaseDealArchive })
    $btnDeleteDeal.Add_Click({ Remove-SelectedPurchaseDeal })
    $remindersGrid.Add_CellDoubleClick({
        try { Open-SelectedReminderTarget } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Напоминания') | Out-Null
        }
    })
    $remindersDeleteMenuItem.Add_Click({
        try {
            $info = Get-SelectedReminderInfo
            if ($null -eq $info -or [int]$info.ReminderId -le 0) {
                throw 'Автоматическое напоминание нельзя удалить вручную.'
            }
            $answer = [Windows.Forms.MessageBox]::Show('Удалить выбранное напоминание?', 'Напоминания', [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Question)
            if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
            Remove-Reminder ([int]$info.ReminderId)
            Refresh-Reminders
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Напоминания') | Out-Null
        }
    })
    Add-DebouncedTextChanged $txtRemindersSearch { Refresh-Reminders }
    $btnRefreshReminders.Add_Click({ Refresh-Reminders })
    $btnDoneReminder.Add_Click({
        try {
            $info = Get-SelectedReminderInfo
            if ($null -eq $info -or [int]$info.ReminderId -le 0) {
                throw 'Выберите ручное напоминание.'
            }
            Set-ReminderDone ([int]$info.ReminderId) $true
            Refresh-Reminders
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Напоминания') | Out-Null
        }
    })
    $btnReceiptArrived.Add_Click({
        try {
            Confirm-SelectedReminderReceipt
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Напоминания') | Out-Null
        }
    })
    $btnNewReminder.Add_Click({
        try {
            $title = Show-SimpleInputDialog 'Новое напоминание' 'Что нужно сделать?' ''
            if ([string]::IsNullOrWhiteSpace($title)) { return }
            $due = Show-SimpleInputDialog 'Новое напоминание' 'Дата (можно пусто, формат дд.мм.гггг)' ''
            Save-Reminder $title $due (Get-SelectedDealId) (Get-SelectedSupplierId) 0 'manual'
            Refresh-Reminders
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Напоминания') | Out-Null
        }
    })
    $quoteBaseTabs.Add_SelectedIndexChanged({
        if ($quoteBaseTabs.SelectedIndex -eq 1) { Refresh-GlobalistQuotes }
    })
    $btnImportGlobalist.Add_Click({
        try {
            $dialog = New-Object Windows.Forms.OpenFileDialog
            $dialog.Filter = 'Файлы Globalist (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*'
            $dialog.Title = 'Выберите файл Globalist'
            if ($dialog.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) { return }
            $count = Import-GlobalistWorkbook $dialog.FileName
            Refresh-GlobalistQuotes
            $quoteBaseTabs.SelectedIndex = 1
            Show-Toast (('Файлов Globalist загружено: {0}' -f $count)) 'Success'
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Ошибка Globalist') | Out-Null
        }
    })
    Add-DebouncedTextChanged $txtQuoteSearch { Refresh-QuoteBase; if ($quoteBaseTabs.SelectedIndex -eq 1) { Refresh-GlobalistQuotes } }
    $btnRefreshQuoteBase.Add_Click({ Refresh-QuoteBase })
    $btnExportQuoteBase.Add_Click({
        try {
            $path = Export-QuoteBaseToExcel
            [Windows.Forms.MessageBox]::Show("Экспорт готов:`r`n$path", 'База квот') | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'База квот') | Out-Null
        }
    })
    $btnDeleteQuoteRows.Add_Click({
        try {
            Remove-SelectedQuoteBaseRows
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'База квот') | Out-Null
        }
    })
    $btnClearQuoteBase.Add_Click({
        try {
            Clear-AllQuoteBaseRows
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'База квот') | Out-Null
        }
    })
    Add-DebouncedTextChanged $txtHistorySearch { Refresh-History }
    $btnRefreshHistory.Add_Click({ Refresh-History })
    $btnHistorySelectedDeal.Add_Click({ Refresh-History (Get-SelectedDealId) })
    $btnClearHistory.Add_Click({
        try {
            Clear-AllHistoryRows
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'История') | Out-Null
        }
    })
    $dealsGrid.Add_SelectionChanged({
        if ($script:PurchaseLoading) { return }
        Refresh-PurchaseDetails
    })
    $dealsGrid.Add_ColumnWidthChanged({
        if (-not $script:PurchaseLoading) {
            Save-GridColumnWidths $dealsGrid 'ui.purchase.deals.column_widths'
        }
    })
    $dealsGrid.Add_CellValidating({
        if ($script:PurchaseLoading -or $_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) { return }
        $columnName = $dealsGrid.Columns[$_.ColumnIndex].Name
        $value = ([string]$_.FormattedValue).Trim()
        if ($columnName -eq 'DealNumber' -and [string]::IsNullOrWhiteSpace($value)) {
            return
        } elseif ($columnName -eq 'BoardCount' -and -not [string]::IsNullOrWhiteSpace($value) -and $value -notmatch '^\d+$') {
            $_.Cancel = $true
            [Windows.Forms.MessageBox]::Show('В колонке "Кол-во плат" нужно указать целое число.', 'Кол-во плат') | Out-Null
        }
    })
    $dealsGrid.Add_CellEndEdit({
        if ($script:PurchaseLoading -or $_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) { return }
        $columnName = $dealsGrid.Columns[$_.ColumnIndex].Name
        if ($columnName -ne 'DealNumber' -and $columnName -ne 'BoardCount' -and $columnName -ne 'Client' -and $columnName -ne 'Period') { return }
        try {
            Save-PurchaseDealGridRow $dealsGrid.Rows[$_.RowIndex]
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Сохранить сделку') | Out-Null
            Refresh-PurchaseDeals
        }
    })
    $dealsGrid.Add_CurrentCellDirtyStateChanged({
        if ($script:PurchaseLoading -or -not $dealsGrid.IsCurrentCellDirty -or $null -eq $dealsGrid.CurrentCell) { return }
        $columnName = $dealsGrid.Columns[$dealsGrid.CurrentCell.ColumnIndex].Name
        if ($columnName -eq 'Masks' -or $columnName -eq 'Priority' -or $columnName -eq 'Executor' -or $columnName -eq 'AssemblyLocation') {
            [void]$dealsGrid.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })
    $dealsGrid.Add_CellValueChanged({
        if ($script:PurchaseLoading -or $_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) { return }
        $columnName = $dealsGrid.Columns[$_.ColumnIndex].Name
        if ($columnName -ne 'Masks' -and $columnName -ne 'Priority' -and $columnName -ne 'Executor' -and $columnName -ne 'AssemblyLocation') { return }
        try {
            Save-PurchaseDealGridRow $dealsGrid.Rows[$_.RowIndex]
            if ($columnName -eq 'Executor' -or $columnName -eq 'AssemblyLocation') { Refresh-PurchaseDeals }
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Сохранить сделку') | Out-Null
            Refresh-PurchaseDeals
        }
    })
    $dealsGrid.Add_SortCompare({
        if ($_.Column.Name -eq 'Priority') {
            $_.SortResult = (Get-ComponentPriorityRank ([string]$_.CellValue1)).CompareTo((Get-ComponentPriorityRank ([string]$_.CellValue2)))
            $_.Handled = $true
        }
    })
    $dealsGrid.Add_CellMouseDown({
        if ($_.Button -ne [Windows.Forms.MouseButtons]::Right -or $_.RowIndex -lt 0) { return }
        $dealsGrid.ClearSelection()
        $dealsGrid.Rows[$_.RowIndex].Selected = $true
        if ($_.ColumnIndex -ge 0) {
            $dealsGrid.CurrentCell = $dealsGrid.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
        }
    })
    $suppliersGrid.Add_CellMouseDown({
        if ($_.Button -ne [Windows.Forms.MouseButtons]::Right -or $_.RowIndex -lt 0) { return }
        $suppliersGrid.ClearSelection()
        $suppliersGrid.Rows[$_.RowIndex].Selected = $true
        if ($_.ColumnIndex -ge 0) {
            $suppliersGrid.CurrentCell = $suppliersGrid.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
        }
    })
    $dealsCopyNumberMenuItem.Add_Click({ Copy-SelectedPurchaseDealNumber })
    $dealsArchiveMenuItem.Add_Click({ Toggle-SelectedPurchaseDealArchive })
    $dealsDeleteMenuItem.Add_Click({ Remove-SelectedPurchaseDeal })
    $suppliersDeleteMenuItem.Add_Click({
        try {
            $supplierId = Get-SelectedSupplierId
            if ($supplierId -le 0) { throw 'Выберите поставщика.' }
            $supplierName = [string](Invoke-PurchaseScalar 'SELECT supplier FROM deal_suppliers WHERE id = @id' @{ '@id' = $supplierId })
            $answer = [Windows.Forms.MessageBox]::Show(
                ('Удалить поставщика и его документы? ' + $supplierName),
                'Удалить поставщика',
                [Windows.Forms.MessageBoxButtons]::YesNo,
                [Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
            Delete-PurchaseSupplier $supplierId
            Refresh-PurchaseAll
        } catch {
                [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Удалить поставщика') | Out-Null
        }
    })
    $purchasePage.Add_Resize({ Resize-PurchasePanels })
    $form.Add_Shown({ Resize-PurchasePanels })

    $btnNewDeal.Add_Click({
        try {
            $dealNumber = Show-SimpleInputDialog 'Новая сделка' 'Введите номер сделки / заказа'
            if ([string]::IsNullOrWhiteSpace($dealNumber)) { return }
            [void](New-PurchaseDeal $dealNumber)
            Refresh-PurchaseAll
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Новая сделка') | Out-Null
        }
    })

    $btnAddPurchaseSupplier.Add_Click({
        try {
            $dealId = Get-SelectedDealId
            if ($dealId -le 0) { throw 'Выберите сделку.' }
            $supplier = Show-SimpleInputDialog 'Поставщик' 'Введите наименование поставщика'
            if ([string]::IsNullOrWhiteSpace($supplier)) { return }
            Add-PurchaseSupplier $dealId $supplier
            Refresh-PurchaseAll
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Добавить поставщика') | Out-Null
        }
    })

    $btnImportRrfq.Add_Click({
        try {
            $dialog = New-Object Windows.Forms.OpenFileDialog
            $dialog.Filter = 'Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*'
            $dialog.Title = 'Выберите результирующий RRFQ'
            if ($dialog.ShowDialog($form) -ne 'OK') { return }

            $defaultDeal = ([IO.Path]::GetFileNameWithoutExtension($dialog.FileName)) -replace '_RRFQ_result.*$', ''
            $dealNumber = Show-SimpleInputDialog 'Импорт из RRFQ' 'Введите номер сделки / заказа' $defaultDeal
            if ([string]::IsNullOrWhiteSpace($dealNumber)) { return }

            $supplierNames = @(Get-SupplierNamesFromRrfqResult $dialog.FileName)
            if ($supplierNames.Count -eq 0) {
                throw 'В файле не найдены поставщики-победители. Проверьте, что это результирующий RRFQ с колонкой Supplier.'
            }

            $dealId = New-PurchaseDeal $dealNumber '' '' '' 'RRFQ'
            foreach ($supplierName in $supplierNames) {
                Add-PurchaseSupplier $dealId $supplierName
            }
            [void](Add-PurchaseDocument $dealId 0 'RRFQ' $dialog.FileName 'RRFQ')
            $quoteImport = Import-PurchaseRrfqQuotesFromFiles @($dialog.FileName) 'Первоначальный импорт RRFQ'
            Refresh-PurchaseAll
            $toastText = "Поставщиков обработано: $($supplierNames.Count)"
            if ($null -ne $quoteImport -and [bool]$quoteImport.Imported) {
                $toastText = "$toastText; квот добавлено: $($quoteImport.Count)"
            }
            Show-Toast $toastText 'Success'
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Импорт из RRFQ') | Out-Null
        }
    })

    $suppliersGrid.Add_CurrentCellDirtyStateChanged({
        if ($suppliersGrid.IsCurrentCellDirty -and $null -ne $suppliersGrid.CurrentCell) {
            $column = $suppliersGrid.CurrentCell.OwningColumn
            if ($column -is [Windows.Forms.DataGridViewCheckBoxColumn]) {
                try {
                    $suppliersGrid.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit) | Out-Null
                } catch {
                    [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Сохранить поставщика') | Out-Null
                }
            }
        }
    })
    $suppliersGrid.Add_CellValidating({
        if ($script:PurchaseLoading -or $_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) { return }
        $columnName = $suppliersGrid.Columns[$_.ColumnIndex].Name
        if ($columnName -eq 'DeliveryWeeks') {
            $value = ([string]$_.FormattedValue).Trim()
            if (-not [string]::IsNullOrWhiteSpace($value) -and $value -notmatch '^\d+$') {
                $_.Cancel = $true
                [Windows.Forms.MessageBox]::Show('В колонке "Срок" нужно указать целое число недель.', 'Срок') | Out-Null
            }
        }
    })
    $suppliersGrid.Add_CellClick({
        if ($script:PurchaseLoading -or $_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) { return }
        $columnName = $suppliersGrid.Columns[$_.ColumnIndex].Name
        if ($columnName -ne 'InvoiceDate' -and $columnName -ne 'ReceiptDate' -and $columnName -ne 'ActualReceiptDate') { return }

        $row = $suppliersGrid.Rows[$_.RowIndex]
        $suppliersGrid.ClearSelection()
        $row.Selected = $true
        $suppliersGrid.CurrentCell = $row.Cells[$_.ColumnIndex]

        $selectedDate = Show-DatePickerDialog ([string]$row.Cells[$columnName].Value)
        if ($null -eq $selectedDate) { return }

        $script:PurchaseLoading = $true
        try {
            $row.Cells[$columnName].Value = [string]$selectedDate
            if ($columnName -eq 'InvoiceDate') {
                Update-AutoReceiptDateForRow $row
            }
        } finally {
            $script:PurchaseLoading = $false
        }

        Save-PurchaseSupplierGridRow $row $false
        Refresh-PurchaseDealsDeferred
    })
    $suppliersGrid.Add_CellValueChanged({
        if (-not $script:PurchaseLoading) {
            $row = if ($_.RowIndex -ge 0) { $suppliersGrid.Rows[$_.RowIndex] } else { $null }
            $columnName = if ($_.ColumnIndex -ge 0) { $suppliersGrid.Columns[$_.ColumnIndex].Name } else { '' }
            if ($_.ColumnIndex -ge 0 -and $suppliersGrid.Columns[$_.ColumnIndex] -is [Windows.Forms.DataGridViewCheckBoxColumn]) {
                Save-PurchaseSupplierGridRow $row $false
                Refresh-PurchaseDealsDeferred
            }
        }
    })
    $suppliersGrid.Add_CellEndEdit({
        if (-not $script:PurchaseLoading) {
            $row = if ($_.RowIndex -ge 0) { $suppliersGrid.Rows[$_.RowIndex] } else { $null }
            $columnName = if ($_.ColumnIndex -ge 0) { $suppliersGrid.Columns[$_.ColumnIndex].Name } else { '' }
            if ($_.ColumnIndex -ge 0 -and -not ($suppliersGrid.Columns[$_.ColumnIndex] -is [Windows.Forms.DataGridViewCheckBoxColumn])) {
                try {
                    Save-PurchaseSupplierGridRow $row ($columnName -eq 'DeliveryWeeks')
                    Refresh-PurchaseDealsDeferred
                } catch {
                    [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Сохранить поставщика') | Out-Null
                    Refresh-PurchaseDetailsDeferred
                }
            }
        }
    })

    $btnNewComponent.Add_Click({
        try {
            if ($componentsGrid.IsCurrentCellInEditMode -or $componentsGrid.IsCurrentRowDirty) {
                [void]$componentsGrid.EndEdit()
            }
            if ($null -ne $componentsGrid.CurrentRow -and $null -ne $componentsGrid.CurrentRow.Tag) {
                Save-ComponentGridRow $componentsGrid.CurrentRow
            }
            Save-CurrentComponentNotes
            $newId = New-ComponentDeal
            Refresh-Components
            foreach ($row in $componentsGrid.Rows) {
                if ($null -ne $row.Tag -and [int]$row.Tag -eq $newId) {
                    $componentsGrid.ClearSelection()
                    $row.Selected = $true
                    $componentsGrid.CurrentCell = $row.Cells['DealNumber']
                    break
                }
            }
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Задачи') | Out-Null
        }
    })

    $btnDeleteComponent.Add_Click({
        try {
            if ($componentsGrid.SelectedRows.Count -eq 0 -or $null -eq $componentsGrid.SelectedRows[0].Tag) {
                throw 'Выберите строку.'
            }
            $deal = [string]$componentsGrid.SelectedRows[0].Cells['DealNumber'].Value
            $answer = [Windows.Forms.MessageBox]::Show(
                "Удалить задачу?`r`n$deal",
                'Задачи',
                [Windows.Forms.MessageBoxButtons]::YesNo,
                [Windows.Forms.MessageBoxIcon]::Question
            )
            if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
            Delete-ComponentDeal ([int]$componentsGrid.SelectedRows[0].Tag)
            Refresh-Components
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Задачи') | Out-Null
        }
    })

    $btnRefreshComponents.Add_Click({ Refresh-Components })
    Add-DebouncedTextChanged $txtComponentsSearch { Refresh-Components }

    $btnSaveComponentNotes.Add_Click({
        try {
            Save-CurrentComponentNotes $true
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Задачи') | Out-Null
        }
    })

    $componentsGrid.Add_SelectionChanged({
        if ($script:ComponentsLoading) { return }
        try {
            Save-CurrentComponentNotes
            Load-SelectedComponentNotes
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Задачи') | Out-Null
        }
    })
    $componentsGrid.Add_CellMouseDown({
        if ($_.Button -eq [Windows.Forms.MouseButtons]::Right -and $_.RowIndex -ge 0) {
            $componentsGrid.ClearSelection()
            $componentsGrid.Rows[$_.RowIndex].Selected = $true
            if ($_.ColumnIndex -ge 0) {
                $componentsGrid.CurrentCell = $componentsGrid.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
            }
        }
    })

    $componentsCopyDealMenuItem.Add_Click({
        try {
            if ($componentsGrid.SelectedRows.Count -eq 0) { throw 'Выберите задачу.' }
            $dealNumber = ([string]$componentsGrid.SelectedRows[0].Cells['DealNumber'].Value).Trim()
            if ([string]::IsNullOrWhiteSpace($dealNumber)) { throw 'Номер сделки пустой.' }
            [Windows.Forms.Clipboard]::SetText($dealNumber)
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Задачи') | Out-Null
        }
    })
    $componentsGrid.Add_ColumnWidthChanged({
        if (-not $script:ComponentsLoading) {
            Save-GridColumnWidths $componentsGrid 'ui.components.column_widths'
        }
    })

    $componentsGrid.Add_CurrentCellDirtyStateChanged({
        if ($componentsGrid.IsCurrentCellDirty -and $null -ne $componentsGrid.CurrentCell) {
            $column = $componentsGrid.CurrentCell.OwningColumn
            if ($column -is [Windows.Forms.DataGridViewComboBoxColumn]) {
                try {
                    $componentsGrid.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit) | Out-Null
                } catch {
                    [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Задачи') | Out-Null
                }
            }
        }
    })

    $componentsGrid.Add_CellValueChanged({
        if ($script:ComponentsLoading -or $_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) { return }
        if ($componentsGrid.Columns[$_.ColumnIndex] -is [Windows.Forms.DataGridViewComboBoxColumn]) {
            try {
                Save-ComponentGridRow $componentsGrid.Rows[$_.RowIndex]
            } catch {
                [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Задачи') | Out-Null
                Refresh-Components
            }
        }
    })

    $componentsGrid.Add_CellEndEdit({
        if ($script:ComponentsLoading -or $_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) { return }
        if ($componentsGrid.Columns[$_.ColumnIndex] -is [Windows.Forms.DataGridViewComboBoxColumn]) { return }
        try {
            Save-ComponentGridRow $componentsGrid.Rows[$_.RowIndex]
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Задачи') | Out-Null
            Refresh-Components
        }
    })

    $componentsGrid.Add_CellClick({
        if ($script:ComponentsLoading -or $_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) { return }
        $columnName = $componentsGrid.Columns[$_.ColumnIndex].Name
        if ($columnName -ne 'ReminderDate' -and $columnName -ne 'DeadlineDate') { return }

        $row = $componentsGrid.Rows[$_.RowIndex]
        $selectedDate = Show-DatePickerDialog ([string]$row.Cells[$columnName].Value)
        if ($null -eq $selectedDate) { return }
        $script:ComponentsLoading = $true
        try {
            $row.Cells[$columnName].Value = [string]$selectedDate
        } finally {
            $script:ComponentsLoading = $false
        }
        Save-ComponentGridRow $row
    })

    $componentsGrid.Add_SortCompare({
        $columnName = $_.Column.Name
        if ($columnName -eq 'Priority') {
            $_.SortResult = (Get-ComponentPriorityRank ([string]$_.CellValue1)).CompareTo((Get-ComponentPriorityRank ([string]$_.CellValue2)))
            $_.Handled = $true
        } elseif ($columnName -eq 'Status') {
            $_.SortResult = (Get-ComponentStatusRank ([string]$_.CellValue1)).CompareTo((Get-ComponentStatusRank ([string]$_.CellValue2)))
            $_.Handled = $true
        } elseif ($columnName -eq 'EntryDate' -or $columnName -eq 'ReminderDate' -or $columnName -eq 'DeadlineDate') {
            $d1 = Convert-PurchaseDateText ([string]$_.CellValue1)
            $d2 = Convert-PurchaseDateText ([string]$_.CellValue2)
            if ($null -eq $d1 -and $null -eq $d2) {
                $_.SortResult = 0
            } elseif ($null -eq $d1) {
                $_.SortResult = 1
            } elseif ($null -eq $d2) {
                $_.SortResult = -1
            } else {
                $_.SortResult = $d1.CompareTo($d2)
            }
            $_.Handled = $true
        }
    })

    function Confirm-PurchaseRrfqQuotes {
        param([object[]]$Quotes, [string[]]$Files, $Owner)
        $lines = New-Object System.Collections.ArrayList
        [void]$lines.Add(('Найдено квот: {0}' -f @($Quotes).Count))
        [void]$lines.Add('')
        foreach ($quote in @($Quotes)) {
            [void]$lines.Add(('{0} | {1} | {2} | цена: {3} | срок: {4}' -f $quote.Supplier, $quote.Key, $quote.PN, (Format-Price $quote.UnitPrice), $quote.LeadTime))
        }
        # A standard MessageBox cannot scroll long RRFQ results, so use a small
        # modal form with a scrollable read-only text area and fixed action buttons.
        $dialog = New-Object Windows.Forms.Form
        $dialog.Text = 'Квоты из RRFQ - добавить в базу?'
        $dialog.StartPosition = 'CenterParent'
        $dialog.FormBorderStyle = [Windows.Forms.FormBorderStyle]::Sizable
        $dialog.MinimizeBox = $false
        $dialog.MaximizeBox = $false
        $dialog.MinimumSize = New-Object Drawing.Size(560, 260)
        $dialog.Font = New-Object Drawing.Font('Segoe UI', 9)
        $dialog.BackColor = Get-UiColor 'Canvas'

        $icon = New-Object Windows.Forms.PictureBox
        $icon.Left = 16
        $icon.Top = 16
        $icon.Width = 32
        $icon.Height = 32
        $icon.SizeMode = [Windows.Forms.PictureBoxSizeMode]::StretchImage
        $icon.Image = [System.Drawing.SystemIcons]::Question.ToBitmap()
        $dialog.Controls.Add($icon)

        $summary = New-Object Windows.Forms.Label
        $summary.Text = ('Найдено квот: {0}' -f @($Quotes).Count)
        $summary.Left = 60
        $summary.Top = 20
        $summary.AutoSize = $true
        $dialog.Controls.Add($summary)

        $details = New-Object Windows.Forms.RichTextBox
        $details.Left = 60
        $details.Top = 48
        $details.Width = 620
        $details.Height = 420
        $details.ReadOnly = $true
        $details.WordWrap = $false
        $details.ScrollBars = [Windows.Forms.RichTextBoxScrollBars]::Both
        $details.Text = [string]::Join("`r`n", @($lines | Select-Object -Skip 2))
        $dialog.Controls.Add($details)

        $btnYes = New-Object Windows.Forms.Button
        $btnYes.Text = 'Да'
        $btnYes.Width = 92
        Set-PrimaryButtonLook $btnYes
        $dialog.Controls.Add($btnYes)

        $btnNo = New-Object Windows.Forms.Button
        $btnNo.Text = 'Нет'
        $btnNo.Width = 92
        Set-SecondaryButtonLook $btnNo
        $dialog.Controls.Add($btnNo)

        $btnYes.Add_Click({ $dialog.DialogResult = [Windows.Forms.DialogResult]::Yes; $dialog.Close() })
        $btnNo.Add_Click({ $dialog.DialogResult = [Windows.Forms.DialogResult]::No; $dialog.Close() })
        $dialog.AcceptButton = $btnYes
        $dialog.CancelButton = $btnNo

        $dialog.Add_Shown({
            $workArea = [Windows.Forms.Screen]::FromControl($dialog).WorkingArea
            $dialog.Width = [Math]::Min(760, [Math]::Max(560, $workArea.Width - 80))
            $dialog.Height = [Math]::Min(620, [Math]::Max(260, $workArea.Height - 80))
        })
        $dialog.Add_Resize({
            $details.Width = $dialog.ClientSize.Width - $details.Left - 20
            $details.Height = $dialog.ClientSize.Height - $details.Top - 62
            $btnNo.Left = $dialog.ClientSize.Width - $btnNo.Width - 12
            $btnYes.Left = $btnNo.Left - $btnYes.Width - 8
            $btnYes.Top = $dialog.ClientSize.Height - $btnYes.Height - 12
            $btnNo.Top = $btnYes.Top
        })

        Apply-ControlTreeLook $dialog
        [void]$dialog.ShowDialog($Owner)
        return ($dialog.DialogResult -eq [Windows.Forms.DialogResult]::Yes)
    }
    function Import-PurchaseRrfqQuotesFromFiles {
        param(
            [string[]]$Files,
            [string]$Comment = 'Импорт из документов сделки'
        )

        $filesToImport = @($Files | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($filesToImport.Count -eq 0) {
            return [pscustomobject]@{ Imported = $false; Count = 0; Message = 'Файлы RRFQ не выбраны.' }
        }

        $parsedQuotes = New-Object System.Collections.ArrayList
        foreach ($rrfqFile in $filesToImport) {
            $supplierName = Get-SupplierFromFileName $rrfqFile
            if ([string]::IsNullOrWhiteSpace($supplierName)) {
                $supplierName = [IO.Path]::GetFileNameWithoutExtension($rrfqFile)
            }
            foreach ($quote in @(Get-QuotesFromWorkbook $rrfqFile $supplierName)) {
                [void]$parsedQuotes.Add($quote)
            }
        }

        if ($parsedQuotes.Count -eq 0) {
            return [pscustomobject]@{ Imported = $false; Count = 0; Message = 'В RRFQ не найдено квот.' }
        }

        $quoteAnalysis = [pscustomobject]@{
            RfqPath = [string]::Join('; ', [string[]]$filesToImport)
            Suppliers = @()
            RfqRows = @()
            Quotes = @($parsedQuotes)
            Decisions = @()
            Priority = 'Price'
        }
        if (-not (Confirm-PurchaseRrfqQuotes @($parsedQuotes) $filesToImport $form)) {
            return [pscustomobject]@{ Imported = $false; Count = 0; Message = 'Импорт квот отменен пользователем.' }
        }

        [void](Save-QuoteHistory $quoteAnalysis $Comment)
        return [pscustomobject]@{ Imported = $true; Count = $parsedQuotes.Count; Message = ('В базу квот добавлено: {0}' -f $parsedQuotes.Count) }
    }

    function Add-PurchaseDocumentsFromPaths {
        param([string[]]$Files)

        $dealId = Get-SelectedDealId
        $selectedSupplierId = Get-SelectedSupplierId
        if ($dealId -le 0) { throw 'Выберите сделку.' }
        $files = @($Files | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($files.Count -eq 0) { return }
        $documentType = Show-DocumentTypeDialog
        if ([string]::IsNullOrWhiteSpace($documentType)) { return }
        $isDealLevelDocument = Test-DealLevelDocumentType $documentType
        $supplierId = if ($isDealLevelDocument) { 0 } else { $selectedSupplierId }
        if (-not $isDealLevelDocument -and $supplierId -le 0) { throw 'Выберите поставщика для привязки файлов.' }
        $folderName = if ($isDealLevelDocument) { $documentType } else { '' }

        $messages = New-Object System.Collections.ArrayList
        $rrfqFiles = New-Object System.Collections.ArrayList
        foreach ($file in $files) {
            [void](Add-PurchaseDocument $dealId $supplierId $documentType $file $folderName)
            if ($documentType -eq 'RRFQ') { [void]$rrfqFiles.Add($file) }
            if ($documentType -eq 'PI/Invoice' -and $supplierId -gt 0) {
                Update-PurchaseSupplierPiAmounts $supplierId $null $null $true
            }
        }
        if ($documentType -eq 'PI/Invoice' -and $supplierId -gt 0) {
            [void]$messages.Add('Файлы PI/Invoice загружены быстро, без авто-чтения PDF. Чтобы заполнить суммы PI, выберите поставщика и нажмите "Считать PI".')
        }
        if ($documentType -eq 'RRFQ' -and $rrfqFiles.Count -gt 0) {
            try {
                $quoteImport = Import-PurchaseRrfqQuotesFromFiles @($rrfqFiles) 'Импорт из документов сделки'
                if ($null -ne $quoteImport -and -not [string]::IsNullOrWhiteSpace([string]$quoteImport.Message)) {
                    [void]$messages.Add([string]$quoteImport.Message)
                }
            } catch { [void]$messages.Add(('Ошибка разбора RRFQ: {0}' -f $_.Exception.Message)) }
        }
        Refresh-PurchaseSuppliers
        Refresh-PurchaseDeals
        Refresh-PurchaseDocuments
        Refresh-History
        $message = "Загружено файлов: $($files.Count)"
        if ($messages.Count -gt 0) {
            $message = "$message`r`n`r`n$([string]::Join("`r`n", @($messages)))"
        }
        [Windows.Forms.MessageBox]::Show($message, 'Документы') | Out-Null
    }

    $purchaseDocsDragEnter = {
        if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
            $_.Effect = [Windows.Forms.DragDropEffects]::Copy
        } else {
            $_.Effect = [Windows.Forms.DragDropEffects]::None
        }
    }
    $docsGrid.Add_DragEnter($purchaseDocsDragEnter)
    $docsPanel.Add_DragEnter($purchaseDocsDragEnter)
    $docsHeaderPanel.Add_DragEnter($purchaseDocsDragEnter)
    $docsLabel.Add_DragEnter($purchaseDocsDragEnter)

    $purchaseDocsDragDrop = {
        try {
            $files = @($_.Data.GetData([Windows.Forms.DataFormats]::FileDrop))
            Add-PurchaseDocumentsFromPaths @($files)
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Загрузка документов') | Out-Null
        }
    }
    $docsGrid.Add_DragDrop($purchaseDocsDragDrop)
    $docsPanel.Add_DragDrop($purchaseDocsDragDrop)
    $docsHeaderPanel.Add_DragDrop($purchaseDocsDragDrop)
    $docsLabel.Add_DragDrop($purchaseDocsDragDrop)
    $btnPickPurchaseDocuments.Add_Click({
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Title = 'Выберите документы'
        $dialog.Filter = 'All files (*.*)|*.*'
        $dialog.Multiselect = $true
        try {
            if ($dialog.ShowDialog($form) -eq [Windows.Forms.DialogResult]::OK) {
                Add-PurchaseDocumentsFromPaths @($dialog.FileNames)
            }
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Загрузка документов') | Out-Null
        } finally {
            $dialog.Dispose()
        }
    })
    $docsGrid.Add_CellMouseDown({
        if ($_.Button -ne [Windows.Forms.MouseButtons]::Right -or $_.RowIndex -lt 0) { return }
        $docsGrid.ClearSelection()
        $docsGrid.Rows[$_.RowIndex].Selected = $true
        if ($_.ColumnIndex -ge 0) {
            $docsGrid.CurrentCell = $docsGrid.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
        }
    })
    $docsOpenExplorerMenuItem.Add_Click({ Open-SelectedPurchaseDocumentInExplorer })
    $docsLoadQuotesMenuItem.Add_Click({
        try {
            $row = if ($docsGrid.SelectedRows.Count -gt 0) { $docsGrid.SelectedRows[0] } else { $null }
            if ($null -eq $row) { throw 'Выберите документ.' }
            if ([string]$row.Cells['Type'].Value -ne 'RRFQ') {
                throw 'Пункт "Загрузить квоты" доступен только для документов типа RRFQ.'
            }

            $info = Get-SelectedDocumentInfo
            if ($null -eq $info) { throw 'Выберите документ.' }
            $name = if ($info.PSObject.Properties['Name']) { [string]$info.Name } else { '' }
            $hash = if ($info.PSObject.Properties['FileHash']) { [string]$info.FileHash } else { '' }
            $resolvedDoc = Resolve-PurchaseDocumentFile ([int]$info.Id) ([string]$info.Path) $name $hash
            if ([bool]$resolvedDoc.Changed) { Refresh-PurchaseDocuments }
            $path = [string]$resolvedDoc.Path
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                throw "Файл не найден:`r`n$path"
            }

            $quoteImport = Import-PurchaseRrfqQuotesFromFiles @($path) 'Повторная загрузка квот из документа сделки'
            Refresh-History
            [Windows.Forms.MessageBox]::Show([string]$quoteImport.Message, 'Загрузка квот') | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Загрузка квот') | Out-Null
        }
    })
    $docsDeleteMenuItem.Add_Click({
        try {
            $info = Get-SelectedDocumentInfo
            if ($null -eq $info -or [int]$info.Id -le 0) { throw 'Выберите документ.' }
            $fileName = if ($docsGrid.SelectedRows.Count -gt 0) { [string]$docsGrid.SelectedRows[0].Cells['File'].Value } else { '' }
            $answer = [Windows.Forms.MessageBox]::Show(
                "Удалить документ из приложения и файл из хранилища?`r`n$fileName",
                'Удалить документ',
                [Windows.Forms.MessageBoxButtons]::YesNo,
                [Windows.Forms.MessageBoxIcon]::Question
            )
            if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }

            Delete-PurchaseDocument ([int]$info.Id)
            Refresh-PurchaseDocuments
            Refresh-PurchaseDeals
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Удалить документ') | Out-Null
        }
    })
    $docsGrid.Add_CellDoubleClick({
        $info = Get-SelectedDocumentInfo
        if ($null -eq $info) { return }
        $name = if ($info.PSObject.Properties['Name']) { [string]$info.Name } else { '' }
        $hash = if ($info.PSObject.Properties['FileHash']) { [string]$info.FileHash } else { '' }
        $resolvedDoc = Resolve-PurchaseDocumentFile ([int]$info.Id) ([string]$info.Path) $name $hash
        $path = [string]$resolvedDoc.Path
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            if ([bool]$resolvedDoc.Changed) { Refresh-PurchaseDocuments }
            Start-Process -FilePath $path
        } else {
            [Windows.Forms.MessageBox]::Show("Файл не найден:`r`n$path", 'Документы') | Out-Null
        }
    })

    $btnChooseDocumentsRoot.Add_Click({
        $dialog = New-Object Windows.Forms.FolderBrowserDialog
        $dialog.Description = 'Выберите папку для хранения документов'
        $dialog.SelectedPath = $txtDocumentsRoot.Text
        if ($dialog.ShowDialog($form) -eq 'OK') {
            $txtDocumentsRoot.Text = $dialog.SelectedPath
        }
    })
    $btnSaveSettings.Add_Click({
        try {
            if ([string]::IsNullOrWhiteSpace($txtDocumentsRoot.Text)) {
                throw 'Укажите папку документов.'
            }
            $documentsRoot = Set-PurchaseDocumentsRoot $txtDocumentsRoot.Text
            $txtDocumentsRoot.Text = $documentsRoot
            [Windows.Forms.MessageBox]::Show('Настройки сохранены.', 'Настройки') | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Настройки') | Out-Null
        }
    })

    $notificationTimer = New-Object Windows.Forms.Timer
    $notificationTimer.Interval = 60000
    $notificationTimer.Add_Tick({ Check-Notifications })
    $chkNotificationsEnabled.Add_CheckedChanged({
        $script:NotificationsEnabled = [bool]$chkNotificationsEnabled.Checked
        Set-PurchaseSetting 'notifications.enabled' ($(if ($script:NotificationsEnabled) { '1' } else { '0' }))
        if ($script:NotificationsEnabled) {
            $notificationTimer.Start()
            Check-Notifications
        } else {
            $notificationTimer.Stop()
            $script:NotificationQueue.Clear()
            $script:NotificationKeys.Clear()
            Close-NotificationCard
        }
    })
    if ($script:NotificationsEnabled) { $notificationTimer.Start() }
    $form.Add_FormClosed({
        $notificationTimer.Stop()
        Close-NotificationCard
    })
    $form.Add_Shown({ if ($script:NotificationsEnabled) { Check-Notifications } })

    Apply-ControlTreeLook $form
    $navPanel.BackColor = Get-UiColor 'Surface'
    $navTitle.ForeColor = Get-UiColor 'Text'
    $navHint.ForeColor = Get-UiColor 'Muted'
    $docsPanel.BackColor = Get-UiColor 'Surface'
    $docsLabel.BackColor = Get-UiColor 'Surface'
    $docsLabel.ForeColor = Get-UiColor 'Text'
    $docsLabel.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
    Set-NavButtonLook $btnNavRrfq $false
    Set-NavButtonLook $btnNavPurchase $true
    Set-NavButtonLook $btnNavComponents $false
    $btnNavComponents.Image = New-NavIcon 'Components'
    Set-NavButtonLook $btnNavReminders $false
    $btnNavReminders.Image = New-NavIcon 'Reminders'
    Set-NavButtonLook $btnNavQuoteBase $false
    $btnNavQuoteBase.Image = New-NavIcon 'QuoteBase'
    Set-NavButtonLook $btnNavHistory $false
    $btnNavHistory.Image = New-NavIcon 'History'
    Set-NavButtonLook $btnNavCompelParser $false
    $btnNavCompelParser.Image = New-NavIcon 'CompelParser'
    Set-NavButtonLook $btnNavPriceSearch $false
    $btnNavPriceSearch.Image = New-NavIcon 'PriceSearch'
    Set-NavButtonLook $btnNavInstructions $false
    $btnNavInstructions.Image = New-NavIcon 'Instructions'
    Set-NavButtonLook $btnNavSettings $false
    $btnNavSettings.Image = New-NavIcon 'Settings'

    Refresh-PurchaseAll
    Refresh-Components
    Show-AppPage 'Purchase'

    $btnRfq.Add_Click({
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*'
        $dialog.Multiselect = $false
        if ($dialog.ShowDialog() -eq 'OK') {
            $txtRfq.Text = $dialog.FileName
        }
    })

    $btnAdd.Add_Click({
        $dialog = New-Object Windows.Forms.OpenFileDialog
        $dialog.Filter = 'Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*'
        $dialog.Multiselect = $true
        if ($dialog.ShowDialog() -eq 'OK') {
            foreach ($file in $dialog.FileNames) {
                $supplier = Get-SupplierFromFileName $file
                [void]$supplierGrid.Rows.Add($file, $supplier)
            }
        }
    })

    $btnRemove.Add_Click({
        foreach ($row in @($supplierGrid.SelectedRows)) {
            if (-not $row.IsNewRow) {
                $supplierGrid.Rows.Remove($row)
            }
        }
    })

    $btnPreview.Add_Click({
        try {
            $suppliers = New-Object System.Collections.ArrayList
            foreach ($row in $supplierGrid.Rows) {
                if ($row.IsNewRow) { continue }
                $path = [string]$row.Cells['Path'].Value
                $supplier = [string]$row.Cells['Supplier'].Value
                if ([string]::IsNullOrWhiteSpace($path)) { continue }
                if ([string]::IsNullOrWhiteSpace($supplier)) { $supplier = Get-SupplierFromFileName $path }
                [void]$suppliers.Add([pscustomobject]@{ Path = $path; Supplier = $supplier })
            }

            if ([string]::IsNullOrWhiteSpace($txtRfq.Text)) {
                throw 'Выберите RFQ.'
            }
            if ($suppliers.Count -eq 0) {
                throw 'Добавьте хотя бы один файл поставщика.'
            }

            $priorityMode = if ($radioLead.Checked) { 'LeadTime' } else { 'Price' }
            $btnPreview.Enabled = $false
            $logBox.Text = 'Читаю Excel-файлы...'
            $form.Refresh()
            $script:LastAnalysis = Invoke-Analysis $txtRfq.Text @($suppliers) $priorityMode
            Refresh-CurrentPreview
            $btnSave.Enabled = $true
            $btnManual.Enabled = $true
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Ошибка') | Out-Null
        } finally {
            $btnPreview.Enabled = $true
        }
    })

    $btnManual.Add_Click({
        if ($null -eq $script:LastAnalysis) {
            return
        }
        Show-ManualMatchDialog $script:LastAnalysis $previewGrid $logBox
        Refresh-CurrentPreview
    })

    $cmbFilter.Add_SelectedIndexChanged({ Refresh-CurrentPreview })
    $txtSearch.Add_TextChanged({ Refresh-CurrentPreview })
    $previewGrid.Add_SelectionChanged({
        Refresh-DecisionDetails $script:LastAnalysis $previewGrid $detailsBox $quoteDetailsGrid
    })
    $quoteDetailsGrid.Add_CellDoubleClick({
        Set-SelectedQuoteAsWinner
    })

    $btnChooseQuote.Add_Click({
        Set-SelectedQuoteAsWinner
    })

    $btnAddManualQuote.Add_Click({
        try {
            $decision = Get-SelectedPreviewDecision
            if ($null -eq $decision) {
                throw 'Выберите строку preview.'
            }

            Update-DecisionsFromGrid $script:LastAnalysis $previewGrid
            if (Show-ManualQuoteDialog $script:LastAnalysis $decision $form) {
                Refresh-CurrentPreview
            }
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Ручной ввод квоты') | Out-Null
        }
    })

    $btnIncludeSelected.Add_Click({
        if ($null -eq $script:LastAnalysis -or $previewGrid.SelectedRows.Count -eq 0) { return }
        $decision = Get-DecisionById $script:LastAnalysis ([string]$previewGrid.SelectedRows[0].Tag)
        if ($null -ne $decision) {
            $decision.Include = $true
            Refresh-CurrentPreview
        }
    })

    $btnExcludeSelected.Add_Click({
        if ($null -eq $script:LastAnalysis -or $previewGrid.SelectedRows.Count -eq 0) { return }
        $decision = Get-DecisionById $script:LastAnalysis ([string]$previewGrid.SelectedRows[0].Tag)
        if ($null -ne $decision) {
            $decision.Include = $false
            Refresh-CurrentPreview
        }
    })

    $btnSave.Add_Click({
        try {
            if ($null -eq $script:LastAnalysis) {
                throw 'Сначала сформируйте preview.'
            }

            Update-DecisionsFromGrid $script:LastAnalysis $previewGrid
            $btnSave.Enabled = $false
            $logBox.Text = "$($logBox.Text)`r`n`r`nСоздаю результирующий RRFQ..."
            $form.Refresh()
            $path = Write-ResultWorkbook $script:LastAnalysis
            $logBox.Text = "$($logBox.Text)`r`nГотово: $path"
            [Windows.Forms.MessageBox]::Show("Готово.`r`n$path", 'RRFQ создан') | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Ошибка сохранения') | Out-Null
        } finally {
            $btnSave.Enabled = ($null -ne $script:LastAnalysis)
        }
    })

    $statusBar = New-Object Windows.Forms.Panel
    $statusBar.Dock = 'Bottom'
    $statusBar.Height = 30
    $statusBar.BackColor = Get-UiColor 'Surface'
    $statusText = New-Object Windows.Forms.Label
    $statusText.Dock = 'Fill'
    $statusText.ForeColor = Get-UiColor 'Muted'
    $statusText.Font = New-Object Drawing.Font('Segoe UI', 8.5)
    $statusText.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $statusText.Padding = New-Object Windows.Forms.Padding(14, 0, 0, 0)
    $statusText.Text = 'Procurement control'
    $statusBar.Controls.Add($statusText)
    $statusClock = New-Object Windows.Forms.Label
    $statusClock.Dock = 'Right'
    $statusClock.AutoSize = $true
    $statusClock.ForeColor = Get-UiColor 'Muted'
    $statusClock.Font = New-Object Drawing.Font('Segoe UI', 8.5)
    $statusClock.TextAlign = [Drawing.ContentAlignment]::MiddleRight
    $statusClock.Padding = New-Object Windows.Forms.Padding(0, 0, 14, 0)
    $statusClock.Text = Get-Date -Format 'dd.MM.yyyy HH:mm'
    $statusBar.Controls.Add($statusClock)
    $statusTop = New-Object Windows.Forms.Panel
    $statusTop.Dock = 'Top'
    $statusTop.Height = 1
    $statusTop.BackColor = Get-UiColor 'Line'
    $statusBar.Controls.Add($statusTop)
    $form.Controls.Add($statusBar)
    $statusTimer = New-Object Windows.Forms.Timer
    $statusTimer.Interval = 30000
    $statusTimer.Add_Tick({ $statusClock.Text = Get-Date -Format 'dd.MM.yyyy HH:mm' })
    $statusTimer.Start()
    $form.Add_FormClosed({ Close-ExcelApp })
    if ($SmokeTest) {
        $form.Dispose()
        return
    }

    [Windows.Forms.Application]::Run($form)
}

$script:Config = Get-AppConfig

if ($ImportOnly) {
    return
} elseif ($UiSmokeTest) {
    try {
        Show-MainFormV2 -SmokeTest
        Write-Host 'UI smoke OK'
    } finally {
        Close-ExcelApp
    }
} elseif ($SelfTest) {
    try {
        if ([string]::IsNullOrWhiteSpace($Rfq)) {
            throw 'Для SelfTest укажите -Rfq.'
        }

        $suppliers = @()
        foreach ($file in $SupplierFiles) {
            $suppliers += [pscustomobject]@{
                Path = $file
                Supplier = Get-SupplierFromFileName $file
            }
        }

        $analysis = Invoke-Analysis $Rfq $suppliers $Priority
        $unresolved = @(Get-UnresolvedQuotes $analysis)
        $winners = @($analysis.Decisions | Where-Object { $_.Winner })
        Write-Host "RFQ rows: $($analysis.RfqRows.Count)"
        Write-Host "Quotes: $($analysis.Quotes.Count)"
        Write-Host "Winners: $($winners.Count)"
        Write-Host "Unresolved: $($unresolved.Count)"
        foreach ($decision in @($analysis.Decisions | Where-Object { $_.Winner } | Select-Object -First 10)) {
            Write-Host ("{0} R{1}: {2} -> {3} price={4} lead={5} total={6}" -f $decision.SheetName, $decision.Row, $decision.Key, $decision.WinnerSupplier, (Format-Price $decision.WinnerPrice), $decision.WinnerLead, $decision.WinnerLeadTotal)
        }

        if ($WriteResult) {
            $path = Write-ResultWorkbook $analysis
            Write-Host "Result: $path"
        }
    } finally {
        Close-ExcelApp
    }
} else {
    try {
        Show-MainFormV2
    } finally {
        Close-ExcelApp
    }
}












































































