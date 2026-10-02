# Purchase dashboard calculations. Dot-sourced by RRFQComparer.ps1.
# Data access is supplied by PurchaseStore functions in the main application.

function Get-PurchaseOrderTrackingStartedAt {
    $raw = [string](Get-PurchaseSetting 'dashboard.orders_tracking_started_at')
    $date = [DateTime]::MinValue
    if (-not [string]::IsNullOrWhiteSpace($raw) -and [DateTime]::TryParse($raw, [ref]$date)) {
        return $date
    }
    return $null
}

function Get-PurchaseDashboardCount {
    param(
        [ValidateSet('deals', 'component_deals')]
        [string]$TableName,
        [ValidateSet('created_at', 'ordered_at')]
        [string]$DateColumn,
        [DateTime]$Start,
        [DateTime]$End
    )

    $orderFilter = if ($DateColumn -eq 'ordered_at') { " AND IFNULL(ordered_at, '') <> ''" } else { '' }
    return [int](Invoke-PurchaseScalar "SELECT COUNT(*) FROM $TableName WHERE $DateColumn >= @start AND $DateColumn < @end$orderFilter" @{
        '@start' = $Start.ToString('yyyy-MM-dd HH:mm:ss')
        '@end' = $End.ToString('yyyy-MM-dd HH:mm:ss')
    })
}

function Get-PurchaseDashboardConversion {
    param([int]$Ordered, [int]$Created)

    if ($Created -le 0) { return $null }
    return [Math]::Round((100.0 * $Ordered) / $Created, 1)
}

function Get-PurchaseDashboardDirection {
    param(
        [ValidateSet('deals', 'component_deals')]
        [string]$TableName,
        [DateTime]$MonthStart,
        [DateTime]$MonthEnd,
        [DateTime]$TrackingStartedAt,
        [DateTime]$Now
    )

    $currentCohortStart = if ($TrackingStartedAt -gt $MonthStart) { $TrackingStartedAt } else { $MonthStart }
    $matureCutoff = $Now.Date.AddDays(-60)
    $trackingText = $TrackingStartedAt.ToString('yyyy-MM-dd HH:mm:ss')
    $currentCohortText = $currentCohortStart.ToString('yyyy-MM-dd HH:mm:ss')
    $monthEndText = $MonthEnd.ToString('yyyy-MM-dd HH:mm:ss')
    $matureCutoffText = $matureCutoff.ToString('yyyy-MM-dd HH:mm:ss')

    $createdThisMonth = Get-PurchaseDashboardCount $TableName 'created_at' $MonthStart $MonthEnd
    $orderedThisMonth = Get-PurchaseDashboardCount $TableName 'ordered_at' $MonthStart $MonthEnd
    $currentCohortCreated = [int](Invoke-PurchaseScalar "SELECT COUNT(*) FROM $TableName WHERE created_at >= @start AND created_at < @end" @{
        '@start' = $currentCohortText; '@end' = $monthEndText
    })
    $currentCohortOrdered = [int](Invoke-PurchaseScalar "SELECT COUNT(*) FROM $TableName WHERE created_at >= @start AND created_at < @end AND IFNULL(ordered_at, '') <> ''" @{
        '@start' = $currentCohortText; '@end' = $monthEndText
    })
    $matureCohortCreated = [int](Invoke-PurchaseScalar "SELECT COUNT(*) FROM $TableName WHERE created_at >= @tracking_start AND created_at < @mature_cutoff" @{
        '@tracking_start' = $trackingText; '@mature_cutoff' = $matureCutoffText
    })
    $matureCohortOrdered = [int](Invoke-PurchaseScalar "SELECT COUNT(*) FROM $TableName WHERE created_at >= @tracking_start AND created_at < @mature_cutoff AND IFNULL(ordered_at, '') <> ''" @{
        '@tracking_start' = $trackingText; '@mature_cutoff' = $matureCutoffText
    })

    return [pscustomobject]@{
        AddedThisMonth = $createdThisMonth
        OrderedThisMonth = $orderedThisMonth
        CurrentConversion = Get-PurchaseDashboardConversion $currentCohortOrdered $currentCohortCreated
        MatureConversion = Get-PurchaseDashboardConversion $matureCohortOrdered $matureCohortCreated
        CurrentCohortCreated = $currentCohortCreated
        MatureCohortCreated = $matureCohortCreated
    }
}

function Get-PurchaseDashboardHistory {
    param([DateTime]$MonthStart, [DateTime]$TrackingStartedAt)

    $culture = [Globalization.CultureInfo]::GetCultureInfo('ru-RU')
    $history = New-Object System.Collections.ArrayList
    foreach ($offset in 5..0) {
        $start = $MonthStart.AddMonths(-$offset)
        $end = $start.AddMonths(1)
        $hasOrderData = ($end -gt $TrackingStartedAt)
        [void]$history.Add([pscustomobject]@{
            MonthStart = $start
            Label = $start.ToString('MMM yy', $culture)
            AssemblyAdded = Get-PurchaseDashboardCount 'deals' 'created_at' $start $end
            ComponentAdded = Get-PurchaseDashboardCount 'component_deals' 'created_at' $start $end
            AssemblyOrdered = if ($hasOrderData) { Get-PurchaseDashboardCount 'deals' 'ordered_at' $start $end } else { $null }
            ComponentOrdered = if ($hasOrderData) { Get-PurchaseDashboardCount 'component_deals' 'ordered_at' $start $end } else { $null }
            HasOrderData = $hasOrderData
        })
    }
    return @($history)
}

function Get-PurchaseDashboard {
    $now = Get-Date
    $monthStart = [DateTime]::new($now.Year, $now.Month, 1)
    $monthEnd = $monthStart.AddMonths(1)
    $trackingStartedAt = Get-PurchaseOrderTrackingStartedAt
    if ($null -eq $trackingStartedAt) { $trackingStartedAt = $now }
    $culture = [Globalization.CultureInfo]::GetCultureInfo('ru-RU')

    return [pscustomobject]@{
        MonthLabel = $monthStart.ToString('MMMM yyyy', $culture)
        TrackingStartedAt = $trackingStartedAt
        Assemblies = Get-PurchaseDashboardDirection 'deals' $monthStart $monthEnd $trackingStartedAt $now
        Components = Get-PurchaseDashboardDirection 'component_deals' $monthStart $monthEnd $trackingStartedAt $now
        History = Get-PurchaseDashboardHistory $monthStart $trackingStartedAt
    }
}
