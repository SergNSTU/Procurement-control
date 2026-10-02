# Compel HTML parsing and Excel export functions.
# Includes reusable price formatting and supplier extraction from RRFQ results.
# Dot-sourced by RRFQComparer.ps1; no application startup code belongs here.

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
