# Shared application functions. Dot-sourced by RRFQComparer.ps1.
# Keep this file free of application startup code.

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
