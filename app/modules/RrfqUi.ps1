# RRFQ page presentation helpers and manual matching dialogs.
# Dot-sourced by RRFQComparer.ps1; no application startup code belongs here.

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
    param($Analysis, $PreviewGrid, $QuotesGrid)

    $QuotesGrid.Rows.Clear()
    if ($null -eq $Analysis -or $PreviewGrid.SelectedRows.Count -eq 0) {
        return
    }

    $decision = Get-DecisionById $Analysis ([string]$PreviewGrid.SelectedRows[0].Tag)
    if ($null -eq $decision) {
        return
    }

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
