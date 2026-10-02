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
$configModulePath = Join-Path $PSScriptRoot 'modules\AppConfig.ps1'
$uiThemeModulePath = Join-Path $PSScriptRoot 'modules\UiTheme.ps1'
$uiFoundationModulePath = Join-Path $PSScriptRoot 'modules\UiFoundation.ps1'
$rrfqUiModulePath = Join-Path $PSScriptRoot 'modules\RrfqUi.ps1'
$rrfqEngineModulePath = Join-Path $PSScriptRoot 'modules\RrfqEngine.ps1'
$purchaseStoreModulePath = Join-Path $PSScriptRoot 'modules\PurchaseStore.ps1'
$dataSafetyModulePath = Join-Path $PSScriptRoot 'modules\DataSafety.ps1'
$quoteHistoryModulePath = Join-Path $PSScriptRoot 'modules\QuoteHistory.ps1'
$compelParserModulePath = Join-Path $PSScriptRoot 'modules\CompelParser.ps1'
$priceSearchModulePath = Join-Path $PSScriptRoot 'modules\PriceSearch.ps1'
$dashboardModulePath = Join-Path $PSScriptRoot 'modules\PurchaseDashboard.ps1'
foreach ($modulePath in @($configModulePath, $uiThemeModulePath, $uiFoundationModulePath, $rrfqUiModulePath, $rrfqEngineModulePath, $purchaseStoreModulePath, $dataSafetyModulePath, $quoteHistoryModulePath, $compelParserModulePath, $priceSearchModulePath, $dashboardModulePath)) {
    if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
        throw "Required application module was not found: $modulePath"
    }
    . $modulePath
}

$bitrixIntegrationPath = Join-Path $PSScriptRoot 'BitrixIntegration.ps1'
if (Test-Path -LiteralPath $bitrixIntegrationPath -PathType Leaf) {
    . $bitrixIntegrationPath
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
    $navPanel.RowCount = 19
    Add-TableRowStyle $navPanel 'Absolute' 46
    Add-TableRowStyle $navPanel 'Absolute' 42
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

    $btnNavDashboard = New-Object Windows.Forms.Button
    $btnNavDashboard.Text = 'Дашборд'
    $btnNavDashboard.Dock = 'Fill'
    Set-NavButtonLook $btnNavDashboard $false
    $btnNavDashboard.Image = New-NavIcon 'Dashboard'
    $navPanel.Controls.Add($btnNavDashboard, 0, 1)

    $btnNavRrfq = New-Object Windows.Forms.Button
    $btnNavRrfq.Text = 'Сравнение RRFQ'
    $btnNavRrfq.Dock = 'Fill'
    Set-NavButtonLook $btnNavRrfq $true
    $btnNavRrfq.Image = New-NavIcon 'RRFQ'
    $navPanel.Controls.Add($btnNavRrfq, 0, 2)


    $btnNavBitrix = New-Object Windows.Forms.Button

    $btnNavBitrix.Text = 'Bitrix API Test'

    $btnNavBitrix.Dock = 'Fill'

    Set-NavButtonLook $btnNavBitrix $false
    $btnNavBitrix.Image = New-NavIcon 'Bitrix'

    $navPanel.Controls.Add($btnNavBitrix, 0, 3)

    $btnNavPurchase = New-Object Windows.Forms.Button
    $btnNavPurchase.Text = 'Контроль закупки'
    $btnNavPurchase.Dock = 'Fill'
    Set-NavButtonLook $btnNavPurchase $false
    $btnNavPurchase.Image = New-NavIcon 'Purchase'
    $navPanel.Controls.Add($btnNavPurchase, 0, 4)

    $btnNavComponents = New-Object Windows.Forms.Button
    $btnNavComponents.Text = 'Задачи'
    $btnNavComponents.Dock = 'Fill'
    Set-NavButtonLook $btnNavComponents $false
    $btnNavComponents.Image = New-NavIcon 'Components'
    $navPanel.Controls.Add($btnNavComponents, 0, 5)

    $btnNavReminders = New-Object Windows.Forms.Button
    $btnNavReminders.Text = 'Напоминания'
    $btnNavReminders.Dock = 'Fill'
    Set-NavButtonLook $btnNavReminders $false
    $btnNavReminders.Image = New-NavIcon 'Reminders'
    $navPanel.Controls.Add($btnNavReminders, 0, 7)

    $btnNavQuoteBase = New-Object Windows.Forms.Button
    $btnNavQuoteBase.Text = 'База квот'
    $btnNavQuoteBase.Dock = 'Fill'
    Set-NavButtonLook $btnNavQuoteBase $false
    $btnNavQuoteBase.Image = New-NavIcon 'QuoteBase'
    $navPanel.Controls.Add($btnNavQuoteBase, 0, 8)

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
    $navPanel.Controls.Add($btnNavCompelParser, 0, 9)

    $btnNavPriceSearch = New-Object Windows.Forms.Button
    $btnNavPriceSearch.Text = 'Поиск цен'
    $btnNavPriceSearch.Dock = 'Fill'
    Set-NavButtonLook $btnNavPriceSearch $false
    $btnNavPriceSearch.Image = New-NavIcon 'PriceSearch'
    $navPanel.Controls.Add($btnNavPriceSearch, 0, 10)
    $navDivider1 = New-Object Windows.Forms.Panel
    $navDivider1.Dock = 'Fill'
    $navDivider1.BackColor = Get-UiColor 'Line'
    $navDivider1.Margin = New-Object Windows.Forms.Padding(16, 4, 16, 4)
    $navPanel.Controls.Add($navDivider1, 0, 6)
    $navDivider2 = New-Object Windows.Forms.Panel
    $navDivider2.Dock = 'Fill'
    $navDivider2.BackColor = Get-UiColor 'Line'
    $navDivider2.Margin = New-Object Windows.Forms.Padding(16, 4, 16, 4)
    $navPanel.Controls.Add($navDivider2, 0, 11)

    $btnNavInstructions = New-Object Windows.Forms.Button
    $btnNavInstructions.Text = 'Инструкции'
    $btnNavInstructions.Dock = 'Fill'
    Set-NavButtonLook $btnNavInstructions $false
    $btnNavInstructions.Image = New-NavIcon 'Instructions'
    $navPanel.Controls.Add($btnNavInstructions, 0, 12)

    $btnNavSettings = New-Object Windows.Forms.Button
    $btnNavSettings.Text = 'Настройки'
    $btnNavSettings.Dock = 'Fill'
    Set-NavButtonLook $btnNavSettings $false
    $btnNavSettings.Image = New-NavIcon 'Settings'
    $navPanel.Controls.Add($btnNavSettings, 0, 13)

    $btnNavNotes = New-Object Windows.Forms.Button
    $btnNavNotes.Text = 'Заметки'
    $btnNavNotes.Dock = 'Fill'
    Set-NavButtonLook $btnNavNotes $false
    $btnNavNotes.Image = New-NavIcon 'Notes'
    $navPanel.Controls.Add($btnNavNotes, 0, 14)

    $navPanel.Controls.Add($btnNavHistory, 0, 15)

    $btnNavExcelCompare = New-Object Windows.Forms.Button
    $btnNavExcelCompare.Text = 'Сравнение Excel'
    $btnNavExcelCompare.Dock = 'Fill'
    Set-NavButtonLook $btnNavExcelCompare $false
    $btnNavExcelCompare.Image = New-NavIcon 'ExcelCompare'
    $navPanel.Controls.Add($btnNavExcelCompare, 0, 16)

    $navHint = New-Object Windows.Forms.Label
    $navHint.Text = "Данные закупки хранятся внутри папки приложения."
    $navHint.Dock = 'Fill'
    $navHint.ForeColor = Get-UiColor 'Muted'
    $navHint.TextAlign = [Drawing.ContentAlignment]::BottomLeft
    $navPanel.Controls.Add($navHint, 0, 18)

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

    $dashboardPage = New-Object Windows.Forms.Panel
    $dashboardPage.Dock = 'Fill'
    $dashboardPage.Visible = $false
    $pageHost.Controls.Add($dashboardPage)

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
        $dashboardPage.Visible = ($Page -eq 'Dashboard')
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
        Set-NavButtonLook $btnNavDashboard ($Page -eq 'Dashboard')
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
        } elseif ($Page -eq 'Dashboard') {
            $dashboardPage.BringToFront()
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
        } elseif ($Notification.Source -eq 'deal') {
            $txtDealSearch.Text = ''
            $cmbDealFilter.SelectedItem = 'Все включая архив'
            Refresh-PurchaseDeals
            $target = $null
            foreach ($row in $dealsGrid.Rows) {
                if ($null -ne $row.Tag -and [int]$row.Tag -eq [int]$Notification.SourceId) {
                    $target = $row
                    break
                }
            }
            Show-AppPage 'Purchase'
            if ($null -ne $target) {
                $dealsGrid.ClearSelection(); $target.Selected = $true; $dealsGrid.CurrentCell = $target.Cells[0]
                $target.DataGridView.FirstDisplayedScrollingRowIndex = $target.Index
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
            } elseif ($notification.Source -eq 'deal') {
                Clear-PurchaseDealReminder ([int]$notification.SourceId)
                Mark-NotificationHandled $notification
                Refresh-PurchaseDeals
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
        $hasTaskDetails = $notification.Source -eq 'component'
        $card.Width = 440; $card.Height = if ($hasTaskDetails) { 264 } else { 190 }
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

        $dueTop = 111
        $buttonsTop = 148
        if ($hasTaskDetails) {
            $description = New-Object Windows.Forms.Label
            $description.Text = 'Описание: ' + $(if ([string]::IsNullOrWhiteSpace([string]$notification.Description)) { '—' } else { [string]$notification.Description })
            $description.Left = 92; $description.Top = 109; $description.Width = 320; $description.Height = 30
            $description.ForeColor = Get-UiColor 'Muted'; $description.AutoEllipsis = $true
            $description.Font = New-Object Drawing.Font('Segoe UI', 8.5)
            $card.Controls.Add($description)

            $nextAction = New-Object Windows.Forms.Label
            $nextAction.Text = 'Следующее действие: ' + $(if ([string]::IsNullOrWhiteSpace([string]$notification.NextAction)) { '—' } else { [string]$notification.NextAction })
            $nextAction.Left = 92; $nextAction.Top = 140; $nextAction.Width = 320; $nextAction.Height = 30
            $nextAction.ForeColor = Get-UiColor 'Muted'; $nextAction.AutoEllipsis = $true
            $nextAction.Font = New-Object Drawing.Font('Segoe UI', 8.5)
            $card.Controls.Add($nextAction)

            $dueTop = 177
            $buttonsTop = 222
        }

        $dueIcon = New-Object Windows.Forms.Panel
        $dueIcon.Left = 92; $dueIcon.Top = ($dueTop + 2); $dueIcon.Width = 18; $dueIcon.Height = 18
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
        $dueLabel.Text = 'Срок:'; $dueLabel.Left = 118; $dueLabel.Top = $dueTop; $dueLabel.Width = 42; $dueLabel.Height = 22
        $dueLabel.ForeColor = Get-UiColor 'Muted'; $card.Controls.Add($dueLabel)
        $dueValue = New-Object Windows.Forms.Label
        $dueValue.Text = $dueDate; $dueValue.Left = 160; $dueValue.Top = $dueTop; $dueValue.Width = 210; $dueValue.Height = 22
        $dueValue.ForeColor = Get-UiColor 'Primary'; $dueValue.Font = New-Object Drawing.Font('Segoe UI Semibold', 9)
        $card.Controls.Add($dueValue)
        $open = New-Object Windows.Forms.Button
        $open.Text = 'Открыть'; $open.Left = 16; $open.Top = $buttonsTop; $open.Width = 114; $open.Height = 30
        Set-SecondaryButtonLook $open; $card.Controls.Add($open)
        $later = New-Object Windows.Forms.Button
        $later.Text = 'Напомнить позже'; $later.Left = 138; $later.Top = $buttonsTop; $later.Width = 138; $later.Height = 30
        Set-SecondaryButtonLook $later; $card.Controls.Add($later)
        $done = New-Object Windows.Forms.Button
        $done.Text = 'Выполнено'; $done.Left = 284; $done.Top = $buttonsTop; $done.Width = 122; $done.Height = 30
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
    Add-TableRowStyle $root 'Absolute' 248
    Add-TableRowStyle $root 'Percent' 100
    $rrfqPage.Controls.Add($root)

    $topLayout = New-Object Windows.Forms.TableLayoutPanel
    $topLayout.Dock = 'Fill'
    $topLayout.ColumnCount = 2
    $topLayout.RowCount = 1
    Add-TableColumnStyle $topLayout 'Percent' 68
    Add-TableColumnStyle $topLayout 'Percent' 32
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
    Add-TableRowStyle $filesLayout 'Absolute' 30
    Add-TableRowStyle $filesLayout 'Percent' 100
    Add-TableRowStyle $filesLayout 'Absolute' 34
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
    $fileButtons.Padding = New-Object Windows.Forms.Padding(0)
    $filesLayout.Controls.Add($fileButtons, 0, 2)
    $filesLayout.SetColumnSpan($fileButtons, 3)

    $btnAdd = New-Object Windows.Forms.Button
    $btnAdd.Text = 'Добавить поставщиков'
    $btnAdd.Width = 180
    $btnAdd.Height = 28
    $btnAdd.Margin = New-Object Windows.Forms.Padding(0, 3, 8, 3)
    Set-SecondaryButtonLook $btnAdd
    $fileButtons.Controls.Add($btnAdd)

    $btnRemove = New-Object Windows.Forms.Button
    $btnRemove.Text = 'Удалить выбранные'
    $btnRemove.Width = 160
    $btnRemove.Height = 28
    $btnRemove.Margin = New-Object Windows.Forms.Padding(0, 3, 8, 3)
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
    $actionsLayout.RowCount = 5
    Add-TableRowStyle $actionsLayout 'Absolute' 40
    Add-TableRowStyle $actionsLayout 'Absolute' 34
    Add-TableRowStyle $actionsLayout 'Absolute' 34
    Add-TableRowStyle $actionsLayout 'Absolute' 34
    Add-TableRowStyle $actionsLayout 'Percent' 100
    $actionsGroup.Controls.Add($actionsLayout)
    $actionsGroup.Controls.Add($actionsGroupTitle)

    $groupPriority = New-CardPanel
    $groupPriorityTitle = New-Object Windows.Forms.Label
    $groupPriorityTitle.Text = 'Приоритет:'
    $groupPriorityTitle.AutoSize = $true
    $groupPriorityTitle.Location = New-Object Drawing.Point(12, 11)
    $groupPriorityTitle.Font = New-Object Drawing.Font('Segoe UI', 8.5, [Drawing.FontStyle]::Bold)
    $groupPriorityTitle.ForeColor = Get-UiColor 'Muted'
    $groupPriority.Controls.Add($groupPriorityTitle)
    $groupPriority.Dock = 'Fill'
    $actionsLayout.Controls.Add($groupPriority, 0, 0)

    $radioPrice = New-Object Windows.Forms.RadioButton
    $radioPrice.Text = 'Цена'
    $radioPrice.Left = 108
    $radioPrice.Top = 9
    $radioPrice.Width = 90
    $radioPrice.Checked = $true
    $groupPriority.Controls.Add($radioPrice)

    $radioLead = New-Object Windows.Forms.RadioButton
    $radioLead.Text = 'Срок'
    $radioLead.Left = 204
    $radioLead.Top = 9
    $radioLead.Width = 90
    $groupPriority.Controls.Add($radioLead)

    $btnPreview = New-Object Windows.Forms.Button
    $btnPreview.Text = 'Сформировать preview'
    $btnPreview.Dock = 'Fill'
    Set-PrimaryButtonLook $btnPreview
    $btnPreview.ForeColor = Get-UiColor 'Text'
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
    $btnSave.ForeColor = Get-UiColor 'Text'
    $actionsLayout.Controls.Add($btnSave, 0, 3)

    $logBox = New-Object Windows.Forms.TextBox
    $logBox.Multiline = $true
    $logBox.ReadOnly = $true
    $logBox.Dock = 'Fill'
    $logBox.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
    $logBox.BackColor = Get-UiColor 'SurfaceAlt'
    $logBox.Text = "Выберите RFQ и добавьте файлы поставщиков.`r`nПосле preview здесь появится сводка."
    $actionsLayout.Controls.Add($logBox, 0, 4)

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
    $previewLayout.RowCount = 2
    $previewLayout.ColumnCount = 1
    Add-TableRowStyle $previewLayout 'Absolute' 34
    Add-TableRowStyle $previewLayout 'Percent' 100
    $previewGroup.Controls.Add($previewLayout)
    $previewGroup.Controls.Add($previewGroupTitle)

    $filterPanel = New-Object Windows.Forms.FlowLayoutPanel
    $filterPanel.Dock = 'Fill'
    $filterPanel.FlowDirection = 'LeftToRight'
    $filterPanel.WrapContents = $false
    $filterPanel.Padding = New-Object Windows.Forms.Padding(0)
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
    $previewSplit = New-Object Windows.Forms.SplitContainer
    $previewSplit.Dock = 'Fill'
    $previewSplit.Orientation = 'Horizontal'
    $previewSplit.SplitterWidth = 6
    $previewLayout.Controls.Add($previewSplit, 0, 1)
    $previewSplit.Panel1.Controls.Add($previewGrid)
    $script:LastPreviewGrid = $previewGrid

    # Start with more room for supplier quotes; the splitter remains user-adjustable.
    $form.Add_Shown({
        $availableHeight = $previewSplit.Height - $previewSplit.SplitterWidth
        if ($availableHeight -lt 300) { return }
        $previewSplit.Panel1MinSize = 140
        $previewSplit.Panel2MinSize = 140
        $targetDistance = [int][Math]::Round($availableHeight * 0.45)
        $targetDistance = [Math]::Max($previewSplit.Panel1MinSize, $targetDistance)
        $targetDistance = [Math]::Min(($availableHeight - $previewSplit.Panel2MinSize), $targetDistance)
        if ($targetDistance -ge $previewSplit.Panel1MinSize) {
            $previewSplit.SplitterDistance = $targetDistance
        }
    })

    $quotePanel = New-Object Windows.Forms.TableLayoutPanel
    $quotePanel.Dock = 'Fill'
    $quotePanel.RowCount = 2
    $quotePanel.ColumnCount = 1
    Add-TableRowStyle $quotePanel 'Absolute' 36
    Add-TableRowStyle $quotePanel 'Percent' 100
    $previewSplit.Panel2.Controls.Add($quotePanel)

    $quoteButtons = New-Object Windows.Forms.FlowLayoutPanel
    $quoteButtons.Dock = 'Fill'
    $quoteButtons.FlowDirection = 'LeftToRight'
    $quoteButtons.WrapContents = $false
    $quoteButtons.Padding = New-Object Windows.Forms.Padding(0)
    $quotePanel.Controls.Add($quoteButtons, 0, 0)

    $quoteTitle = New-Object Windows.Forms.Label
    $quoteTitle.Text = 'Квоты выбранной позиции'
    $quoteTitle.Width = 188
    $quoteTitle.Height = 28
    $quoteTitle.Margin = New-Object Windows.Forms.Padding(0, 3, 12, 3)
    $quoteTitle.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $quoteTitle.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
    $quoteTitle.ForeColor = Get-UiColor 'Text'
    $quoteButtons.Controls.Add($quoteTitle)

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
    $quoteDetailsGrid.AutoSizeRowsMode = [Windows.Forms.DataGridViewAutoSizeRowsMode]::None
    $quoteDetailsGrid.RowTemplate.Height = 28
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

    $dashboardLayout = New-Object Windows.Forms.TableLayoutPanel
    $dashboardLayout.Dock = 'Fill'
    $dashboardLayout.Padding = New-Object Windows.Forms.Padding(12)
    $dashboardLayout.RowCount = 5
    $dashboardLayout.ColumnCount = 1
    Add-TableRowStyle $dashboardLayout 'Absolute' 48
    Add-TableRowStyle $dashboardLayout 'Absolute' 126
    Add-TableRowStyle $dashboardLayout 'Absolute' 34
    Add-TableRowStyle $dashboardLayout 'Absolute' 142
    Add-TableRowStyle $dashboardLayout 'Absolute' 26
    Add-TableColumnStyle $dashboardLayout 'Percent' 100
    $dashboardPage.Controls.Add($dashboardLayout)

    $dashboardHeader = New-Object Windows.Forms.TableLayoutPanel
    $dashboardHeader.Dock = 'Fill'
    $dashboardHeader.ColumnCount = 2
    $dashboardHeader.RowCount = 1
    Add-TableColumnStyle $dashboardHeader 'Percent' 100
    Add-TableColumnStyle $dashboardHeader 'Absolute' 120
    Add-TableRowStyle $dashboardHeader 'Percent' 100
    $dashboardLayout.Controls.Add($dashboardHeader, 0, 0)

    $dashboardTitle = New-Object Windows.Forms.Label
    $dashboardTitle.Text = 'Дашборд эффективности'
    $dashboardTitle.Dock = 'Fill'
    $dashboardTitle.Font = New-Object Drawing.Font('Segoe UI', 14, [Drawing.FontStyle]::Bold)
    $dashboardTitle.ForeColor = Get-UiColor 'Text'
    $dashboardTitle.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $dashboardHeader.Controls.Add($dashboardTitle, 0, 0)

    $btnRefreshDashboard = New-Object Windows.Forms.Button
    $btnRefreshDashboard.Text = 'Обновить'
    $btnRefreshDashboard.Dock = 'Fill'
    $btnRefreshDashboard.Margin = New-Object Windows.Forms.Padding(4, 6, 0, 6)
    Set-SecondaryButtonLook $btnRefreshDashboard
    $dashboardHeader.Controls.Add($btnRefreshDashboard, 1, 0)

    $dashboardDirectionsPanel = New-Object Windows.Forms.TableLayoutPanel
    $dashboardDirectionsPanel.Dock = 'Fill'
    $dashboardDirectionsPanel.ColumnCount = 2
    $dashboardDirectionsPanel.RowCount = 1
    Add-TableColumnStyle $dashboardDirectionsPanel 'Percent' 50
    Add-TableColumnStyle $dashboardDirectionsPanel 'Percent' 50
    Add-TableRowStyle $dashboardDirectionsPanel 'Percent' 100
    $dashboardLayout.Controls.Add($dashboardDirectionsPanel, 0, 1)

    $assemblyMetricsLayout = New-Object Windows.Forms.TableLayoutPanel
    $assemblyMetricsLayout.Dock = 'Fill'
    $assemblyMetricsLayout.Margin = New-Object Windows.Forms.Padding(0, 0, 6, 0)
    $assemblyMetricsLayout.ColumnCount = 1
    $assemblyMetricsLayout.RowCount = 2
    Add-TableRowStyle $assemblyMetricsLayout 'Absolute' 28
    Add-TableRowStyle $assemblyMetricsLayout 'Percent' 100
    $dashboardDirectionsPanel.Controls.Add($assemblyMetricsLayout, 0, 0)
    $assemblyMetricsTitle = New-Object Windows.Forms.Label
    $assemblyMetricsTitle.Text = 'Сборки · Контроль закупки'
    $assemblyMetricsTitle.Dock = 'Fill'
    $assemblyMetricsTitle.Font = New-Object Drawing.Font('Segoe UI Semibold', 10)
    $assemblyMetricsTitle.ForeColor = Get-UiColor 'Text'
    $assemblyMetricsTitle.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $assemblyMetricsLayout.Controls.Add($assemblyMetricsTitle, 0, 0)
    $assemblyMetricsCards = New-Object Windows.Forms.TableLayoutPanel
    $assemblyMetricsCards.Dock = 'Fill'
    $assemblyMetricsCards.ColumnCount = 3
    $assemblyMetricsCards.RowCount = 1
    foreach ($i in 1..3) { Add-TableColumnStyle $assemblyMetricsCards 'Percent' (100.0 / 3.0) }
    Add-TableRowStyle $assemblyMetricsCards 'Percent' 100
    $assemblyMetricsLayout.Controls.Add($assemblyMetricsCards, 0, 1)
    $dashboardAssemblyAdded = New-CockpitMetricCard 'Добавлено' '0' 'В текущем месяце' 'Info'
    $dashboardAssemblyOrdered = New-CockpitMetricCard 'Заказано' '0' 'В текущем месяце' 'Success'
    $dashboardAssemblyConversion = New-CockpitMetricCard 'Конверсия' '—' 'Зрелая 60+ дн.: —' 'Attention'
    $dashboardAssemblyConversion.Margin = New-Object Windows.Forms.Padding(0)
    $assemblyMetricsCards.Controls.Add($dashboardAssemblyAdded, 0, 0)
    $assemblyMetricsCards.Controls.Add($dashboardAssemblyOrdered, 1, 0)
    $assemblyMetricsCards.Controls.Add($dashboardAssemblyConversion, 2, 0)

    $componentMetricsLayout = New-Object Windows.Forms.TableLayoutPanel
    $componentMetricsLayout.Dock = 'Fill'
    $componentMetricsLayout.Margin = New-Object Windows.Forms.Padding(6, 0, 0, 0)
    $componentMetricsLayout.ColumnCount = 1
    $componentMetricsLayout.RowCount = 2
    Add-TableRowStyle $componentMetricsLayout 'Absolute' 28
    Add-TableRowStyle $componentMetricsLayout 'Percent' 100
    $dashboardDirectionsPanel.Controls.Add($componentMetricsLayout, 1, 0)
    $componentMetricsTitle = New-Object Windows.Forms.Label
    $componentMetricsTitle.Text = 'Компоненты · Задачи'
    $componentMetricsTitle.Dock = 'Fill'
    $componentMetricsTitle.Font = New-Object Drawing.Font('Segoe UI Semibold', 10)
    $componentMetricsTitle.ForeColor = Get-UiColor 'Text'
    $componentMetricsTitle.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $componentMetricsLayout.Controls.Add($componentMetricsTitle, 0, 0)
    $componentMetricsCards = New-Object Windows.Forms.TableLayoutPanel
    $componentMetricsCards.Dock = 'Fill'
    $componentMetricsCards.ColumnCount = 3
    $componentMetricsCards.RowCount = 1
    foreach ($i in 1..3) { Add-TableColumnStyle $componentMetricsCards 'Percent' (100.0 / 3.0) }
    Add-TableRowStyle $componentMetricsCards 'Percent' 100
    $componentMetricsLayout.Controls.Add($componentMetricsCards, 0, 1)
    $dashboardComponentAdded = New-CockpitMetricCard 'Добавлено' '0' 'В текущем месяце' 'Info'
    $dashboardComponentOrdered = New-CockpitMetricCard 'Заказано' '0' 'В текущем месяце' 'Success'
    $dashboardComponentConversion = New-CockpitMetricCard 'Конверсия' '—' 'Зрелая 60+ дн.: —' 'Attention'
    $dashboardComponentConversion.Margin = New-Object Windows.Forms.Padding(0)
    $componentMetricsCards.Controls.Add($dashboardComponentAdded, 0, 0)
    $componentMetricsCards.Controls.Add($dashboardComponentOrdered, 1, 0)
    $componentMetricsCards.Controls.Add($dashboardComponentConversion, 2, 0)

    $dashboardHistoryTitle = New-Object Windows.Forms.Label
    $dashboardHistoryTitle.Text = 'Показатели за 6 месяцев'
    $dashboardHistoryTitle.Dock = 'Fill'
    $dashboardHistoryTitle.Font = New-Object Drawing.Font('Segoe UI Semibold', 10)
    $dashboardHistoryTitle.ForeColor = Get-UiColor 'Text'
    $dashboardHistoryTitle.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $dashboardLayout.Controls.Add($dashboardHistoryTitle, 0, 2)

    $dashboardHistoryTiles = New-Object Windows.Forms.TableLayoutPanel
    $dashboardHistoryTiles.Dock = 'Fill'
    $dashboardHistoryTiles.ColumnCount = 6
    $dashboardHistoryTiles.RowCount = 1
    $dashboardHistoryTiles.Padding = New-Object Windows.Forms.Padding(0, 0, 0, 4)
    foreach ($i in 1..6) { Add-TableColumnStyle $dashboardHistoryTiles 'Percent' (100.0 / 6.0) }
    Add-TableRowStyle $dashboardHistoryTiles 'Percent' 100
    $dashboardLayout.Controls.Add($dashboardHistoryTiles, 0, 3)

    $dashboardHistoryTileCards = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt 6; $i++) {
        $historyTile = New-Object Windows.Forms.Panel
        $historyTile.Dock = 'Fill'
        $historyTile.BackColor = Get-UiColor 'Surface'
        $historyTile.BorderStyle = [Windows.Forms.BorderStyle]::FixedSingle
        $historyTile.Padding = New-Object Windows.Forms.Padding(9, 8, 9, 8)
        $historyTile.Margin = if ($i -lt 5) { New-Object Windows.Forms.Padding(0, 0, 7, 0) } else { New-Object Windows.Forms.Padding(0) }

        $historyAccent = New-Object Windows.Forms.Panel
        $historyAccent.Dock = 'Top'
        $historyAccent.Height = 3
        $historyAccent.BackColor = Get-UiColor 'Primary'
        $historyTile.Controls.Add($historyAccent)

        $historyTileLayout = New-Object Windows.Forms.TableLayoutPanel
        $historyTileLayout.Dock = 'Fill'
        $historyTileLayout.ColumnCount = 1
        $historyTileLayout.RowCount = 3
        Add-TableColumnStyle $historyTileLayout 'Percent' 100
        Add-TableRowStyle $historyTileLayout 'Absolute' 27
        Add-TableRowStyle $historyTileLayout 'Percent' 50
        Add-TableRowStyle $historyTileLayout 'Percent' 50
        $historyTile.Controls.Add($historyTileLayout)

        $historyMonthLabel = New-Object Windows.Forms.Label
        $historyMonthLabel.Dock = 'Fill'
        $historyMonthLabel.Font = New-Object Drawing.Font('Segoe UI Semibold', 9)
        $historyMonthLabel.ForeColor = Get-UiColor 'Text'
        $historyMonthLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
        $historyTileLayout.Controls.Add($historyMonthLabel, 0, 0)

        $historyAssemblyLabel = New-Object Windows.Forms.Label
        $historyAssemblyLabel.Dock = 'Fill'
        $historyAssemblyLabel.Font = New-Object Drawing.Font('Segoe UI', 8.25)
        $historyAssemblyLabel.ForeColor = [Drawing.Color]::FromArgb(30, 64, 175)
        $historyAssemblyLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
        $historyTileLayout.Controls.Add($historyAssemblyLabel, 0, 1)

        $historyComponentLabel = New-Object Windows.Forms.Label
        $historyComponentLabel.Dock = 'Fill'
        $historyComponentLabel.Font = New-Object Drawing.Font('Segoe UI', 8.25)
        $historyComponentLabel.ForeColor = [Drawing.Color]::FromArgb(109, 40, 217)
        $historyComponentLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
        $historyTileLayout.Controls.Add($historyComponentLabel, 0, 2)

        $historyTile.Tag = [pscustomobject]@{
            Month = $historyMonthLabel
            Assemblies = $historyAssemblyLabel
            Components = $historyComponentLabel
        }
        [void]$dashboardHistoryTileCards.Add($historyTile)
        $dashboardHistoryTiles.Controls.Add($historyTile, $i, 0)
    }

    $dashboardHistoryHint = New-Object Windows.Forms.Label
    $dashboardHistoryHint.Dock = 'Fill'
    $dashboardHistoryHint.ForeColor = Get-UiColor 'Muted'
    $dashboardHistoryHint.Font = New-Object Drawing.Font('Segoe UI', 8.5)
    $dashboardHistoryHint.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $dashboardLayout.Controls.Add($dashboardHistoryHint, 0, 4)

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
        @('ReminderDate', 'Напоминание', 112),
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
    $componentPriorityValues = @('', 'Срочно', '1', '2', '3', '4', '5')
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
    $settingsLayout.Dock = 'Fill'
    $settingsLayout.Padding = New-Object Windows.Forms.Padding(18)
    $settingsLayout.RowCount = 9
    $settingsLayout.ColumnCount = 3
    Add-TableColumnStyle $settingsLayout 'Absolute' 180
    Add-TableColumnStyle $settingsLayout 'Percent' 100
    Add-TableColumnStyle $settingsLayout 'Absolute' 150
    Add-TableRowStyle $settingsLayout 'Absolute' 38
    Add-TableRowStyle $settingsLayout 'Absolute' 38
    Add-TableRowStyle $settingsLayout 'Absolute' 38
    Add-TableRowStyle $settingsLayout 'Absolute' 38
    Add-TableRowStyle $settingsLayout 'Absolute' 48
    Add-TableRowStyle $settingsLayout 'Absolute' 34
    Add-TableRowStyle $settingsLayout 'Absolute' 42
    Add-TableRowStyle $settingsLayout 'Percent' 100
    Add-TableRowStyle $settingsLayout 'Absolute' 34
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

    $quickCaptureLabel = New-Object Windows.Forms.Label
    $quickCaptureLabel.Text = 'Быстрый захват'
    $quickCaptureLabel.Dock = 'Fill'
    $quickCaptureLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $settingsLayout.Controls.Add($quickCaptureLabel, 0, 2)
    $chkQuickCaptureEnabled = New-Object Windows.Forms.CheckBox
    $chkQuickCaptureEnabled.Text = 'Разрешить создание из браузера и чатов'
    $chkQuickCaptureEnabled.AutoSize = $true
    $chkQuickCaptureEnabled.Checked = Get-PurchaseBooleanSetting 'quick_capture.enabled' $true
    $settingsLayout.Controls.Add($chkQuickCaptureEnabled, 1, 2)
    $quickCaptureStatus = New-Object Windows.Forms.Label
    $quickCaptureStatus.Text = 'Хоткеи: Ctrl+Shift+T / Ctrl+Shift+R; агент: порт 8765'
    $quickCaptureStatus.Dock = 'Fill'
    $quickCaptureStatus.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $quickCaptureStatus.ForeColor = Get-UiColor 'Muted'
    $settingsLayout.Controls.Add($quickCaptureStatus, 1, 3)
    $settingsLayout.SetColumnSpan($quickCaptureStatus, 2)

    $btnSaveSettings = New-Object Windows.Forms.Button
    $btnSaveSettings.Text = 'Сохранить настройки'
    $btnSaveSettings.Width = 180
    $btnSaveSettings.Height = 30
    Set-PrimaryButtonLook $btnSaveSettings
    $settingsLayout.Controls.Add($btnSaveSettings, 1, 4)

    $recoveryTitle = New-Object Windows.Forms.Label
    $recoveryTitle.Text = 'Корзина и резервные копии'
    $recoveryTitle.Dock = 'Fill'
    $recoveryTitle.Font = New-Object Drawing.Font('Segoe UI Semibold', 10)
    $recoveryTitle.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $settingsLayout.Controls.Add($recoveryTitle, 0, 5)
    $settingsLayout.SetColumnSpan($recoveryTitle, 3)

    $recoveryToolbar = New-Object Windows.Forms.FlowLayoutPanel
    $recoveryToolbar.Dock = 'Fill'
    $recoveryToolbar.FlowDirection = 'LeftToRight'
    $recoveryToolbar.WrapContents = $false
    $settingsLayout.Controls.Add($recoveryToolbar, 0, 6)
    $settingsLayout.SetColumnSpan($recoveryToolbar, 3)
    $btnCreateBackup = New-Object Windows.Forms.Button
    $btnCreateBackup.Text = 'Создать копию'
    $btnCreateBackup.Width = 135
    Set-PrimaryButtonLook $btnCreateBackup
    $recoveryToolbar.Controls.Add($btnCreateBackup)
    $btnRestoreRecovery = New-Object Windows.Forms.Button
    $btnRestoreRecovery.Text = 'Восстановить выбранное'
    $btnRestoreRecovery.Width = 190
    Set-SecondaryButtonLook $btnRestoreRecovery
    $recoveryToolbar.Controls.Add($btnRestoreRecovery)
    $btnDeleteRecovery = New-Object Windows.Forms.Button
    $btnDeleteRecovery.Text = 'Удалить безвозвратно'
    $btnDeleteRecovery.Width = 125
    Set-SecondaryButtonLook $btnDeleteRecovery
    $recoveryToolbar.Controls.Add($btnDeleteRecovery)
    $btnOpenRecoveryFolder = New-Object Windows.Forms.Button
    $btnOpenRecoveryFolder.Text = 'Папка копий'
    $btnOpenRecoveryFolder.Width = 125
    Set-SecondaryButtonLook $btnOpenRecoveryFolder
    $recoveryToolbar.Controls.Add($btnOpenRecoveryFolder)

    $recoveryGrid = New-Object Windows.Forms.DataGridView
    $recoveryGrid.Dock = 'Fill'
    $recoveryGrid.ReadOnly = $true
    $recoveryGrid.AllowUserToAddRows = $false
    $recoveryGrid.AllowUserToDeleteRows = $false
    $recoveryGrid.SelectionMode = 'FullRowSelect'
    $recoveryGrid.MultiSelect = $false
    $recoveryGrid.AutoSizeRowsMode = 'AllCells'
    Set-GridLook $recoveryGrid
    [void]$recoveryGrid.Columns.Add('Kind', 'Тип')
    [void]$recoveryGrid.Columns.Add('CreatedAt', 'Создано')
    [void]$recoveryGrid.Columns.Add('Reason', 'Что сохранено')
    [void]$recoveryGrid.Columns.Add('Size', 'Размер')
    $recoveryGrid.Columns['Kind'].Width = 135
    $recoveryGrid.Columns['CreatedAt'].Width = 145
    $recoveryGrid.Columns['Reason'].AutoSizeMode = [Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $recoveryGrid.Columns['Size'].Width = 90
    $settingsLayout.Controls.Add($recoveryGrid, 0, 7)
    $settingsLayout.SetColumnSpan($recoveryGrid, 3)

    $recoveryHint = New-Object Windows.Forms.Label
    $recoveryHint.Text = 'Корзина хранит только удалённые записи и их файлы. Полные копии создаются вручную и перед восстановлением данных.'
    $recoveryHint.Dock = 'Fill'
    $recoveryHint.ForeColor = Get-UiColor 'Muted'
    $recoveryHint.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $settingsLayout.Controls.Add($recoveryHint, 0, 8)
    $settingsLayout.SetColumnSpan($recoveryHint, 3)

    $script:PurchaseLoading = $false
    $script:PurchaseSplitterInitialized = $false
    $script:RestoringDealColumnWidths = $false
    $script:ComponentsLoading = $false
    $script:CurrentComponentNotesId = 0
    $script:CurrentComponentNotesText = ''
    $script:LoadingComponentNotes = $false

    function Refresh-RecoverySnapshots {
        $recoveryGrid.Rows.Clear()
        foreach ($item in (Get-PurchaseTrashItems).Rows) {
            $idx = $recoveryGrid.Rows.Add(
                'Корзина',
                ([datetime]::Parse([string]$item.deleted_at)).ToString('dd.MM.yyyy HH:mm'),
                [string]$item.title,
                $(if ([int]$item.has_files -eq 1) { 'с файлами' } else { 'запись' })
            )
            $recoveryGrid.Rows[$idx].Tag = [pscustomobject]@{ Type = 'Trash'; Id = [int]$item.id; Path = '' }
        }
        foreach ($snapshot in @(Get-PurchaseRecoverySnapshots | Where-Object { $_.Kind -eq 'Backup' })) {
            $idx = $recoveryGrid.Rows.Add(
                'Резервная копия',
                ([datetime]$snapshot.CreatedAt).ToLocalTime().ToString('dd.MM.yyyy HH:mm'),
                [string]$snapshot.Reason,
                [string]$snapshot.Size
            )
            $recoveryGrid.Rows[$idx].Tag = [pscustomobject]@{ Type = 'Backup'; Id = 0; Path = [string]$snapshot.Path }
        }
        $recoveryHint.Text = ('Записей и копий: {0}. Корзина хранит только удалённые записи и их файлы; полные копии — не более 20 последних.' -f $recoveryGrid.Rows.Count)
    }

    function Get-SelectedRecoveryItem {
        if ($recoveryGrid.SelectedRows.Count -eq 0 -or $null -eq $recoveryGrid.SelectedRows[0].Tag) { return '' }
        return $recoveryGrid.SelectedRows[0].Tag
    }

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
        try { $rows = (Get-PurchaseDeals '' 'AllRecords').Rows } catch { $rows = @() }
        $today = (Get-Date).Date
        $active = 0; $overdue = 0; $dueToday = 0; $attention = 0
        try { $active = Get-ActiveCockpitDealCount } catch { $active = 0 }
        foreach ($row in $rows) {
            $archived = Convert-DbBool $row.archived
            $isDone = ([string]$row.status -match 'Done|Completed')
            $receipt = Convert-PurchaseDateText ([string]$row.completion_receipt_date)
            if ($null -ne $receipt -and -not $archived -and -not $isDone) {
                if ($receipt.Date -lt $today) { $overdue++ }
                if ($receipt.Date -eq $today) { $dueToday++ }
            }
            if (-not $archived -and [int]$row.supplier_count -gt 0 -and [int]$row.invoice_confirmed_count -lt [int]$row.supplier_count) { $attention++ }
        }
        Set-CockpitMetricCard $metricActive ([string]$active) 'Активные, отслеживание, Китай и задачи'
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
                    (Format-PurchaseDate $row.reminder_date),
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
                Apply-ComponentReminderCellStyle $dealsGrid.Rows[$idx]
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
                Apply-PriorityCellStyle $dealsGrid.Rows[$idx]
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
        Update-PurchaseCockpitMetrics
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

    function Format-DashboardPercent {
        param($Value)

        if ($null -eq $Value) { return '—' }
        return ('{0:0.#}%' -f [double]$Value)
    }

    function Refresh-DashboardHistoryTiles {
        param($History)

        $historyItems = @($History)
        for ($i = 0; $i -lt $dashboardHistoryTileCards.Count; $i++) {
            $historyTile = $dashboardHistoryTileCards[$i]
            $tileLabels = $historyTile.Tag
            if ($i -ge $historyItems.Count) {
                $historyTile.Visible = $false
                continue
            }

            $item = $historyItems[$i]
            $assemblyOrders = if ([bool]$item.HasOrderData) { [string]$item.AssemblyOrdered } else { '—' }
            $componentOrders = if ([bool]$item.HasOrderData) { [string]$item.ComponentOrdered } else { '—' }
            $tileLabels.Month.Text = ([string]$item.Label).ToUpperInvariant()
            $tileLabels.Assemblies.Text = "Сборки`r`n+{0} добавлено · {1} заказано" -f [int]$item.AssemblyAdded, $assemblyOrders
            $tileLabels.Components.Text = "Компоненты`r`n+{0} добавлено · {1} заказано" -f [int]$item.ComponentAdded, $componentOrders
            $historyTile.Visible = $true
        }
    }

    function Refresh-Dashboard {
        $dashboard = Get-PurchaseDashboard
        $dashboardTitle.Text = 'Дашборд эффективности · ' + $dashboard.MonthLabel
        $assemblyCurrent = Format-DashboardPercent $dashboard.Assemblies.CurrentConversion
        $assemblyMature = Format-DashboardPercent $dashboard.Assemblies.MatureConversion
        $componentCurrent = Format-DashboardPercent $dashboard.Components.CurrentConversion
        $componentMature = Format-DashboardPercent $dashboard.Components.MatureConversion

        Set-CockpitMetricCard $dashboardAssemblyAdded ([string]$dashboard.Assemblies.AddedThisMonth) 'В текущем месяце'
        Set-CockpitMetricCard $dashboardAssemblyOrdered ([string]$dashboard.Assemblies.OrderedThisMonth) 'Перешли в «Заказано»'
        Set-CockpitMetricCard $dashboardAssemblyConversion $assemblyCurrent ('Зрелая 60+ дн.: ' + $assemblyMature)
        Set-CockpitMetricCard $dashboardComponentAdded ([string]$dashboard.Components.AddedThisMonth) 'В текущем месяце'
        Set-CockpitMetricCard $dashboardComponentOrdered ([string]$dashboard.Components.OrderedThisMonth) 'Перешли в «Заказано»'
        Set-CockpitMetricCard $dashboardComponentConversion $componentCurrent ('Зрелая 60+ дн.: ' + $componentMature)
        Refresh-DashboardHistoryTiles $dashboard.History
        $dashboardHistoryHint.Text = 'Заказы отслеживаются с {0}. В предыдущих периодах данные по заказам не отображаются.' -f $dashboard.TrackingStartedAt.ToString('dd.MM.yyyy')
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
        if ([string]::IsNullOrWhiteSpace($priority)) {
            # An unspecified priority must inherit the deal row color.
            $Row.Cells['Priority'].Style.BackColor = $Row.DefaultCellStyle.BackColor
            $Row.Cells['Priority'].Style.ForeColor = Get-UiColor 'Text'
            $Row.Cells['Priority'].Style.Font = New-Object Drawing.Font('Segoe UI', 9)
            return
        }
        $priorityColor = switch ($priority) {
            'Срочно' { [Drawing.Color]::FromArgb(235, 70, 75) }
            '1' { [Drawing.Color]::FromArgb(255, 176, 176) }
            '2' { [Drawing.Color]::FromArgb(255, 218, 179) }
            '3' { [Drawing.Color]::FromArgb(255, 245, 215) }
            '4' { [Drawing.Color]::FromArgb(221, 235, 255) }
            '5' { [Drawing.Color]::FromArgb(226, 238, 226) }
            default { $Row.DefaultCellStyle.BackColor }
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
            ("Переместить выбранные квоты в корзину: {0}?" -f $ids.Count),
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
            'Очистить всю базу квот? Перед этим будет создана запись в корзине, откуда данные можно восстановить.',
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
            'Очистить всю историю действий? Перед этим будет создана запись в корзине, откуда данные можно восстановить.',
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
        $reminderDateValue = ([string]$row.Cells['ReminderDate'].Value).Trim()
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
            $reminderDateValue `
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
        Refresh-DecisionDetails $script:LastAnalysis $previewGrid $quoteDetailsGrid
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
                "Переместить сделку, поставщиков и документы в корзину?`r`n$dealNumber`r`nВосстановление — в Настройках.",
                'Корзина',
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
    $btnNavDashboard.Add_Click({
        try {
            Refresh-Dashboard
            Show-AppPage 'Dashboard'
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Дашборд') | Out-Null
        }
    })
    $btnRefreshDashboard.Add_Click({
        try {
            $btnRefreshDashboard.Enabled = $false
            Refresh-Dashboard
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Дашборд') | Out-Null
        } finally {
            $btnRefreshDashboard.Enabled = $true
        }
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
        Refresh-RecoverySnapshots
        Show-AppPage 'Settings'
    })

    $btnSaveDealInfo.Add_Click({
        try {
            if ($dealsGrid.IsCurrentCellInEditMode -or $dealsGrid.IsCurrentRowDirty) {
                [void]$dealsGrid.CancelEdit()
            }
            $dealId = Get-SelectedDealId
            if ($dealId -le 0) { throw 'Выберите сделку.' }
            $currentDeal = Invoke-PurchaseQuery 'SELECT deal_number, board_count, priority, masks, executor, reminder_date, assembly_location, tracking_status FROM deals WHERE id = @id' @{ '@id' = $dealId }
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
                ([string]$currentDealRow.executor) ([string]$currentDealRow.reminder_date) ([string]$currentDealRow.assembly_location)
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
            $answer = [Windows.Forms.MessageBox]::Show('Переместить выбранное напоминание в корзину?', 'Корзина', [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Question)
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
            if ($null -eq $info) {
                throw 'Выберите ручное напоминание.'
            }
            if ([string]$info.Source -eq 'deal') {
                Clear-PurchaseDealReminder ([int]$info.DealId)
            } elseif ([int]$info.ReminderId -gt 0) {
                Set-ReminderDone ([int]$info.ReminderId) $true
            } else {
                throw 'Выберите ручное напоминание.'
            }
            Refresh-Reminders
            Refresh-PurchaseDeals
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
    $dealsGrid.Add_CellClick({
        if ($script:PurchaseLoading -or $_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) { return }
        if ($dealsGrid.Columns[$_.ColumnIndex].Name -ne 'ReminderDate') { return }
        $row = $dealsGrid.Rows[$_.RowIndex]
        $selectedDate = Show-DatePickerDialog ([string]$row.Cells['ReminderDate'].Value)
        if ($null -eq $selectedDate) { return }
        $script:PurchaseLoading = $true
        try {
            $row.Cells['ReminderDate'].Value = [string]$selectedDate
        } finally {
            $script:PurchaseLoading = $false
        }
        try {
            Save-PurchaseDealGridRow $row
            Apply-ComponentReminderCellStyle $row
            Refresh-Reminders
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Сохранить напоминание') | Out-Null
            Refresh-PurchaseDeals
        }
    })
    $dealsGrid.Add_SortCompare({
        if ($_.Column.Name -eq 'Priority') {
            $_.SortResult = (Get-ComponentPriorityRank ([string]$_.CellValue1)).CompareTo((Get-ComponentPriorityRank ([string]$_.CellValue2)))
            $_.Handled = $true
        } elseif ($_.Column.Name -eq 'ReminderDate') {
            $d1 = Convert-PurchaseDateText ([string]$_.CellValue1)
            $d2 = Convert-PurchaseDateText ([string]$_.CellValue2)
            if ($null -eq $d1 -and $null -eq $d2) { $_.SortResult = 0 }
            elseif ($null -eq $d1) { $_.SortResult = 1 }
            elseif ($null -eq $d2) { $_.SortResult = -1 }
            else { $_.SortResult = $d1.CompareTo($d2) }
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
                ('Переместить поставщика и его документы в корзину? ' + $supplierName),
                'Корзина',
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
                "Переместить задачу в корзину?`r`n$deal",
                'Корзина',
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
                "Переместить документ в корзину?`r`n$fileName`r`nЕго можно будет восстановить в Настройках.",
                'Корзина',
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
            Set-PurchaseSetting 'quick_capture.enabled' ($(if ($chkQuickCaptureEnabled.Checked) { '1' } else { '0' }))
            [Windows.Forms.MessageBox]::Show('Настройки сохранены.', 'Настройки') | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Настройки') | Out-Null
        }
    })
    $btnCreateBackup.Add_Click({
        try {
            $btnCreateBackup.Enabled = $false
            $path = New-PurchaseRecoverySnapshot -Kind 'Backup' -Reason 'Создано вручную'
            Refresh-RecoverySnapshots
            [Windows.Forms.MessageBox]::Show("Резервная копия готова:`r`n$path", 'Резервная копия') | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Резервная копия') | Out-Null
        } finally {
            $btnCreateBackup.Enabled = $true
        }
    })
    $btnRestoreRecovery.Add_Click({
        try {
            $selected = Get-SelectedRecoveryItem
            if ($null -eq $selected -or [string]::IsNullOrWhiteSpace([string]$selected.Type)) { throw 'Выберите резервную копию или запись в корзине.' }
            $isTrash = ([string]$selected.Type -eq 'Trash')
            $answer = [Windows.Forms.MessageBox]::Show(
                $(if ($isTrash) { 'Восстановить удалённую запись и её файлы из корзины?' } else { "Восстановить данные из выбранного архива?`r`nТекущие данные будут заменены, но перед этим приложение создаст страховочную копию." }),
                'Восстановление данных',
                [Windows.Forms.MessageBoxButtons]::YesNo,
                [Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
            $btnRestoreRecovery.Enabled = $false
            $safetyCopy = ''
            if ($isTrash) {
                Restore-PurchaseTrashItem ([int]$selected.Id)
            } else {
                $safetyCopy = Restore-PurchaseRecoverySnapshot ([string]$selected.Path)
            }
            Refresh-PurchaseAll
            Refresh-Components
            Refresh-Reminders
            Refresh-History
            Refresh-QuoteBase
            Refresh-RecoverySnapshots
            $message = if ($isTrash) { 'Удалённые данные восстановлены.' } else { "Данные восстановлены.`r`nСтраховочная копия: $safetyCopy" }
            [Windows.Forms.MessageBox]::Show($message, 'Восстановление данных') | Out-Null
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Восстановление данных') | Out-Null
        } finally {
            $btnRestoreRecovery.Enabled = $true
        }
    })
    $btnDeleteRecovery.Add_Click({
        try {
            $selected = Get-SelectedRecoveryItem
            if ($null -eq $selected -or [string]::IsNullOrWhiteSpace([string]$selected.Type)) { throw 'Выберите архив или запись корзины.' }
            $isTrash = ([string]$selected.Type -eq 'Trash')
            $answer = [Windows.Forms.MessageBox]::Show($(if ($isTrash) { 'Удалить выбранную запись из корзины без возможности восстановления?' } else { 'Удалить выбранный архив без возможности восстановления?' }), 'Удалить безвозвратно', [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Warning)
            if ($answer -ne [Windows.Forms.DialogResult]::Yes) { return }
            if ($isTrash) { Remove-PurchaseTrashItem ([int]$selected.Id) } else { Remove-PurchaseRecoverySnapshot ([string]$selected.Path) }
            Refresh-RecoverySnapshots
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Удалить архив') | Out-Null
        }
    })
    $btnOpenRecoveryFolder.Add_Click({
        try {
            Initialize-PurchaseRecoveryStore
            Start-Process -FilePath explorer.exe -ArgumentList ('"{0}"' -f (Get-PurchaseRecoveryRoot))
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Резервные копии') | Out-Null
        }
    })

    function Get-QuickCaptureViewStamp {
        return [string](Invoke-PurchaseScalar @'
SELECT
    (SELECT IFNULL(MAX(id), 0) FROM component_deals) || ':' ||
    (SELECT IFNULL(MAX(updated_at), '') FROM component_deals) || '|' ||
    (SELECT IFNULL(MAX(id), 0) FROM reminders) || ':' ||
    (SELECT IFNULL(MAX(updated_at), '') FROM reminders)
'@ @{})
    }

    $script:QuickCaptureViewStamp = Get-QuickCaptureViewStamp
    $quickCaptureRefreshTimer = New-Object Windows.Forms.Timer
    $quickCaptureRefreshTimer.Interval = 1500
    $quickCaptureRefreshTimer.Add_Tick({
        try {
            $stamp = Get-QuickCaptureViewStamp
            if ($stamp -eq $script:QuickCaptureViewStamp) { return }
            $script:QuickCaptureViewStamp = $stamp
            if ($componentsPage.Visible) { Refresh-Components }
            if ($remindersPage.Visible) { Refresh-Reminders }
            if ($historyPage.Visible) { Refresh-History }
            if ($script:NotificationsEnabled) { Check-Notifications }
        } catch { }
    })
    $quickCaptureRefreshTimer.Start()

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
        $quickCaptureRefreshTimer.Stop()
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
    Set-NavButtonLook $btnNavDashboard $false
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
        Refresh-DecisionDetails $script:LastAnalysis $previewGrid $quoteDetailsGrid
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












































































