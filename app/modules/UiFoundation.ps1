# Shared WinForms layout, navigation icon, card, and toast helpers.
# Dot-sourced by RRFQComparer.ps1; no application startup code belongs here.

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
        'Dashboard' {
            $g.DrawRectangle($pen, 3, 3, 5, 5)
            $g.DrawRectangle($pen, 12, 3, 5, 5)
            $g.DrawRectangle($pen, 3, 12, 5, 5)
            $g.DrawRectangle($pen, 12, 12, 5, 5)
        }
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
