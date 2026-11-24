#Requires -Version 5.1
# VIEW-ID: Buttons3D
# VIEW-NAME: 3D buttons
# VIEW-NOTE: Raised, bevelled tiles. Big targets, grouped by section.
# VIEW-FUNCTION: Show-Console3DView
# VIEW-ORDER: 60
<#
    View_3DButtons.ps1
    Admin Console view: 3D buttons

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    The same tools as large raised tiles, grouped by section. Big targets
    for a touch screen or a laptop being driven from across a desk.

    New-Console3DTile lives here rather than in the console because
    nothing else draws one -- it is this view's own chrome.

    THIS FILE IS NOT RUN. It is DOT-SOURCED by AdminConsole.ps1, which is why
    it defines a function and does nothing else. Running it on its own defines
    Show-Console3DView into that prompt and exits; nothing appears, and nothing
    is wrong. Open the console instead.

    WHAT IT MAY USE. Everything the console has already defined by the time it
    is sourced: Get-ConsoleToolModel, New-ConsoleViewShell,
    Update-ConsoleViewShell, Add-ConsoleViewFontStepper, Get-ConsoleViewColors,
    Invoke-ConsoleViewTool, New-ConsoleToolTile, the theme helpers in
    Common.ps1, and $form -- the console window, used as the dialog owner.

    ADDING A VIEW. Copy this file, change the five VIEW- lines at the top and
    the function name, and drop it in this folder. The console finds it by
    reading those lines; there is no list to edit. Keep the VIEW-ID short and
    unique -- it is what a per-user default view is stored as.
#>

function New-Console3DTile {
    <#
        One tool as a raised, bevelled button.

        OWNER-DRAWN, AND THE CAPTION IS DRAWN HERE TOO. The button's own Text
        is left empty on purpose: a Flat button still paints its caption, so
        leaving it set would draw the words twice -- once by WinForms in the
        system position, once by this code in the bevelled one -- and the
        offset between them reads as a rendering fault.

        EVERY COLOUR IS PRECOMPUTED ONTO THE TAG, NOT LOOKED UP IN THE PAINT
        HANDLER. A Paint handler runs from the message loop, where the scope it
        was written in is not the scope it executes in -- the lesson this
        console has now learned from the status strip, the settings cache and
        the palette view. $s.Tag is a property on the sender, so it means the
        same thing from anywhere, and painting cannot fail because a variable
        went out of reach.
    #>
    param(
        [Parameter(Mandatory)]$Row,
        [Parameter(Mandatory)]$Colors,
        [Parameter(Mandatory)][int]$Size,
        [Parameter(Mandatory)]$Tip
    )
    $face = $Colors.TileBack
    $b = New-Object System.Windows.Forms.Button
    $b.Text = ''
    $b.Dock = [System.Windows.Forms.DockStyle]::Fill
    $b.Height = [int]($Size * 4.2)
    $b.Margin = New-Object System.Windows.Forms.Padding(5)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderSize = 0
    $b.FlatAppearance.MouseOverBackColor = $face
    $b.FlatAppearance.MouseDownBackColor = $face
    $b.UseVisualStyleBackColor = $false
    $b.BackColor = $face

    $b.Tag = [PSCustomObject]@{
        Row     = $Row
        Tool    = $Row.Tool
        Name    = $Row.Name
        Blocked = "$($Row.Blocked)"
        Caption = "$(if ($Row.Pinned) { '* ' } else { '' })$($Row.Name)"
        Font    = New-Object System.Drawing.Font('Microsoft Sans Serif', $Size, [System.Drawing.FontStyle]::Bold)
        # A raised edge is the face lightened; a shadow is the face darkened
        # -- derived rather than fixed, so the bevel reads correctly on a dark
        # theme and a light one without either being a special case.
        Top     = Get-BlendedColor -From $face -To ([System.Drawing.Color]::White) -Amount 0.22
        Bottom  = Get-BlendedColor -From $face -To ([System.Drawing.Color]::Black) -Amount 0.20
        Light   = Get-BlendedColor -From $face -To ([System.Drawing.Color]::White) -Amount 0.55
        Dark    = Get-BlendedColor -From $face -To ([System.Drawing.Color]::Black) -Amount 0.45
        Fore    = $(if ($Row.Blocked) { $Colors.Dim } else { $Colors.TileFore })
        Pressed = $false
        Hot     = $false
    }

    $b.Add_Paint({
        param($s, $e)
        $t = $s.Tag
        if (-not $t) { return }
        $g = $e.Graphics
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::None
        $w = $s.ClientSize.Width; $h = $s.ClientSize.Height
        if ($w -le 2 -or $h -le 2) { return }
        $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)

        # PRESSED INVERTS THE BEVEL RATHER THAN MOVING THE BUTTON. Shifting the
        # control itself would relayout the whole table on every click; swapping
        # which edges are lit costs one repaint and is what the eye actually
        # reads as "pushed in".
        $c1 = $(if ($t.Pressed) { $t.Bottom } else { $t.Top })
        $c2 = $(if ($t.Pressed) { $t.Top }    else { $t.Bottom })
        if ($t.Hot -and -not $t.Pressed) {
            $c1 = Get-BlendedColor -From $c1 -To ([System.Drawing.Color]::White) -Amount 0.12
            $c2 = Get-BlendedColor -From $c2 -To ([System.Drawing.Color]::White) -Amount 0.12
        }

        $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $c1, $c2, 90.0)
        try { $g.FillRectangle($br, $rect) } finally { $br.Dispose() }

        $hi = $(if ($t.Pressed) { $t.Dark }  else { $t.Light })
        $lo = $(if ($t.Pressed) { $t.Light } else { $t.Dark })
        $pHi = New-Object System.Drawing.Pen($hi, 1)
        $pLo = New-Object System.Drawing.Pen($lo, 1)
        try {
            # Two-pixel bevel: the outer line carries the contrast, the inner
            # one softens it. One pixel alone reads as a border, not a shape.
            $g.DrawLine($pHi, 0, 0, ($w - 1), 0)
            $g.DrawLine($pHi, 0, 0, 0, ($h - 1))
            $g.DrawLine($pHi, 1, 1, ($w - 2), 1)
            $g.DrawLine($pHi, 1, 1, 1, ($h - 2))
            $g.DrawLine($pLo, 0, ($h - 1), ($w - 1), ($h - 1))
            $g.DrawLine($pLo, ($w - 1), 0, ($w - 1), ($h - 1))
            $g.DrawLine($pLo, 1, ($h - 2), ($w - 2), ($h - 2))
            $g.DrawLine($pLo, ($w - 2), 1, ($w - 2), ($h - 2))
        } finally { $pHi.Dispose(); $pLo.Dispose() }

        $pad = [int]($s.Font.Size * 1.2)
        $off = $(if ($t.Pressed) { 1 } else { 0 })
        $tr = New-Object System.Drawing.Rectangle(($pad + $off), $off, ($w - ($pad * 2)), $h)
        $flags = [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
                 [System.Windows.Forms.TextFormatFlags]::Left -bor
                 [System.Windows.Forms.TextFormatFlags]::EndEllipsis -bor
                 [System.Windows.Forms.TextFormatFlags]::NoPrefix
        [System.Windows.Forms.TextRenderer]::DrawText($g, "$($t.Caption)", $t.Font, $tr, $t.Fore, $flags)

        if ($s.Focused) {
            $fr = New-Object System.Drawing.Rectangle(3, 3, ($w - 7), ($h - 7))
            [System.Windows.Forms.ControlPaint]::DrawFocusRectangle($g, $fr)
        }
    })

    $b.Add_MouseDown({ param($s, $e) $s.Tag.Pressed = $true;  $s.Invalidate() })
    $b.Add_MouseUp({   param($s, $e) $s.Tag.Pressed = $false; $s.Invalidate() })
    $b.Add_MouseEnter({ $this.Tag.Hot = $true;  $this.Invalidate() })
    $b.Add_MouseLeave({ $this.Tag.Hot = $false; $this.Tag.Pressed = $false; $this.Invalidate() })
    $b.Add_GotFocus({  $this.Invalidate() })
    $b.Add_LostFocus({ $this.Invalidate() })

    $seen = ''
    if ("$($Row.LastRun)") { $seen = "`r`nlast run $($Row.LastRun)  ($($Row.RunCount)x)" }
    $Tip.SetToolTip($b, $(if ($Row.Blocked) { "$($Row.Name)`r`n$($Row.Detail)$seen`r`n`r`nNOT AVAILABLE: $($Row.Blocked)" } else { "$($Row.Name)`r`n$($Row.Detail)$seen" }))
    return $b
}

function Show-Console3DView {
    <#
        THE SAME TOOLS AS RAISED BUTTONS, full screen and grouped by section.

        Fewer across than the list view by default, because the whole point is
        that each one looks like something you press -- and a bevel needs room
        to read as a bevel. Four columns of chunky buttons is a different thing
        from ten columns of labels, and squeezing these to the same density
        would just be the list view with decoration.
    #>
    $sh = New-ConsoleViewShell -Title '3D buttons' -Hint 'Click or press Enter to run.  RIGHT-CLICK to pin (pinned tools show a *).  Esc returns to the console.' -ViewId 'Buttons3D'
    $v = $sh.Form
    $v | Add-Member -NotePropertyName 'Cols' -NotePropertyValue 4 -Force
    $v | Add-Member -NotePropertyName 'Rows' -NotePropertyValue @() -Force
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.AutoPopDelay = 20000

    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($n in 2, 3, 4, 5, 6) { [void]$cbo.Items.Add("$n") }
    $cbo.SelectedItem = '4'
    $sh.Bar.Controls.Add($cbo)
    $lblCols = New-Object System.Windows.Forms.Label
    $lblCols.Text = 'Columns:'
    $sh.Bar.Controls.Add($lblCols)
    [void]$sh.Right.Add($cbo); [void]$sh.Right.Add($lblCols)

    $lblFind = New-Object System.Windows.Forms.Label; $lblFind.Text = 'Find:'
    $txtFind = New-Object System.Windows.Forms.TextBox
    $lblCount = New-Object System.Windows.Forms.Label
    foreach ($c in @($lblFind, $txtFind, $lblCount)) { $sh.Bar.Controls.Add($c); [void]$sh.Left.Add($c) }

    $scroll = New-Object System.Windows.Forms.Panel
    $scroll.Dock = [System.Windows.Forms.DockStyle]::Fill
    $scroll.AutoScroll = $true
    $sh.Body.Controls.Add($scroll)
    $tbl = New-Object System.Windows.Forms.TableLayoutPanel
    $tbl.Dock = [System.Windows.Forms.DockStyle]::Top
    $tbl.AutoSize = $true
    $tbl.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $tbl.ColumnCount = 4
    $tbl.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 28)
    $scroll.Controls.Add($tbl)

    $build = {
        param([string]$Filter)
        $sz = [int]$v.FontSize
        $cols = [int]$v.Cols
        $col = Get-ConsoleViewColors -Form $v
        $fh = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 2)), [System.Drawing.FontStyle]::Bold)
        $tbl.SuspendLayout()
        try {
            # Disposed, not just cleared. Each tile owns a Font and a Tag and
            # is rebuilt on every keystroke in Find; clearing the collection
            # alone leaks one handle per tool per character typed.
            $old = @($tbl.Controls); $tbl.Controls.Clear(); $tbl.RowStyles.Clear()
            foreach ($c in $old) { try { $c.Dispose() } catch { } }
            $tbl.ColumnCount = $cols
            $tbl.ColumnStyles.Clear()
            for ($i = 0; $i -lt $cols; $i++) {
                [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, (100.0 / $cols))))
            }
            $flt = "$Filter".Trim()
            $row = 0; $c0 = 0; $shown = 0; $lastSec = ''
            foreach ($m in $v.Rows) {
                if ($flt -and "$($m.Name)" -notmatch [regex]::Escape($flt)) { continue }
                if ($m.Section -ne $lastSec) {
                    if ($c0 -ne 0) { $row++; $c0 = 0 }
                    $h = New-Object System.Windows.Forms.Label
                    $h.Text = "  $($m.Section)"; $h.Font = $fh
                    $h.BackColor = $col.HeaderBack; $h.ForeColor = $col.HeaderFore
                    $h.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $h.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $h.Height = [int]($sz * 2.8)
                    $h.Margin = New-Object System.Windows.Forms.Padding(5, ([int]($sz * 1.1)), 5, 4)
                    $tbl.Controls.Add($h, 0, $row); $tbl.SetColumnSpan($h, $cols)
                    $row++; $c0 = 0; $lastSec = $m.Section
                }
                $b = New-Console3DTile -Row $m -Colors $col -Size $sz -Tip $tip
                $b.Add_Click({ Invoke-ConsoleViewTool -ViewForm $v -Tool $this.Tag.Tool -Blocked $this.Tag.Blocked -Rebuild { $v.Rows = Get-ConsoleToolModel; & $build $txtFind.Text } })
                $b.Add_MouseUp({
                    param($s, $e)
                    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
                    Set-ConsoleToolPinned -Name $s.Tag.Name -Pinned (-not $s.Tag.Row.Pinned)
                    $v.Rows = Get-ConsoleToolModel
                    & $build $txtFind.Text
                })
                $tbl.Controls.Add($b, $c0, $row)
                $shown++; $c0++
                if ($c0 -ge $cols) { $c0 = 0; $row++ }
            }
            $total = @($v.Rows).Count
            $lblCount.Text = $(if ($flt) { "$shown of $total tools" } else { "$total tools" })
        } finally { $tbl.ResumeLayout() }
        Update-ConsoleViewShell -Shell $sh
    }

    $txtFind.Add_TextChanged({ & $build $txtFind.Text })
    $cbo.Add_SelectedIndexChanged({ $v.Cols = [int]"$($cbo.SelectedItem)"; & $build $txtFind.Text })
    Add-ConsoleViewFontStepper -Shell $sh -OnChange { & $build $txtFind.Text }
    $v.Add_Resize({ try { Update-ConsoleViewShell -Shell $sh; $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })
    $scroll.Add_Resize({ try { $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })

    Set-ConsoleTheme -Form $v
    $v.Rows = Get-ConsoleToolModel
    Update-ConsoleViewShell -Shell $sh
    & $build ''
    $sh.Title.Text = "3D buttons  --  $(@($v.Rows).Count) tools"
    [void]$txtFind.Focus()
    [void]$v.ShowDialog($form)
}
