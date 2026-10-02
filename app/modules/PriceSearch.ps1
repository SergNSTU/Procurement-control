# Price-search data sources, result normalization, ranking, and GPT prompt generation.
# Dot-sourced by RRFQComparer.ps1; no UI controls or startup code belong here.

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
