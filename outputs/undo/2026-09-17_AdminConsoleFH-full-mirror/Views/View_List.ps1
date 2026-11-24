#Requires -Version 5.1
# VIEW-ID: List
# VIEW-NAME: Tool list - large text, columns
# VIEW-NOTE: Every tool as a full-width row, grouped by section, in a bigger font.
# VIEW-FUNCTION: Show-ConsoleListView
# VIEW-ORDER: 10
<#
    View_List.ps1
    Admin Console view: Tool list - large text, columns

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    Built for the operator who does not want to hunt a small button. Every
    tool is a full-width row in a larger face, grouped under its section
    heading, with a find box at the top. Nothing is hidden behind a group
    button here -- a grouped tool appears as itself, because the point of
    this view is to see the whole inventory at once.

    THIS FILE IS NOT RUN. It is DOT-SOURCED by AdminConsole.ps1, which is why
    it defines a function and does nothing else. Running it on its own defines
    Show-ConsoleListView into that prompt and exits; nothing appears, and nothing
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

function Show-ConsoleListView {
    <#
        FULL SCREEN, BIG TEXT, TOOLS IN COLUMNS, GROUPED BY SECTION.

        NAMES ONLY. The first cut carried 'Runs on' and 'Status' columns beside
        every tool. They were honest and they were wasted space: a list whose
        job is to be readable across the room should not spend two thirds of
        its width on text nobody is reading. Both facts are still there -- in
        the tooltip and in the colour -- and dropping them buys the width to
        run three tools across, which cuts the scrolling to a third.

        A TableLayoutPanel RATHER THAN A ListView. A ListView in Details mode
        gives one item per ROW; three tools across a row would mean three tools
        sharing one selection, and clicking would have to hit-test which cell
        was meant. A table of real buttons has none of that.
    #>
    $sh = New-ConsoleViewShell -Title 'Tool list' -Hint 'Click or press Enter on a tool to run it.  RIGHT-CLICK to pin it (pinned tools show a *).  Tab moves between tools.  Esc returns to the console.' -ViewId 'List'
    $v = $sh.Form
    $v | Add-Member -NotePropertyName 'Cols' -NotePropertyValue 3 -Force
    $v | Add-Member -NotePropertyName 'Rows' -NotePropertyValue @() -Force
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.AutoPopDelay = 20000

    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($n in 1, 2, 3, 4, 5) { [void]$cbo.Items.Add("$n") }
    $cbo.SelectedItem = '3'
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
    # Docked TOP, not Fill: filled it would be clipped to the visible area and
    # AutoScroll would have nothing to scroll to.
    $tbl = New-Object System.Windows.Forms.TableLayoutPanel
    $tbl.Dock = [System.Windows.Forms.DockStyle]::Top
    $tbl.AutoSize = $true
    $tbl.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $tbl.ColumnCount = 3
    $tbl.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 24)
    $scroll.Controls.Add($tbl)

    $build = {
        param([string]$Filter)
        $sz = [int]$v.FontSize
        $cols = [int]$v.Cols
        $col = Get-ConsoleViewColors -Form $v
        $fh = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 1)), [System.Drawing.FontStyle]::Bold)
        $tbl.SuspendLayout()
        try {
            # Dispose, not just Clear. These are real controls with real
            # handles; clearing the collection alone leaks one per rebuild, and
            # this rebuilds on every keystroke in the Find box.
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
                    $h.Text = "$($m.Section)"; $h.Font = $fh
                    $h.BackColor = $col.HeaderBack; $h.ForeColor = $col.HeaderFore
                    $h.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $h.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $h.Height = [int]($sz * 2.6)
                    $h.Margin = New-Object System.Windows.Forms.Padding(2, ([int]($sz * 0.9)), 2, 2)
                    $tbl.Controls.Add($h, 0, $row); $tbl.SetColumnSpan($h, $cols)
                    $row++; $c0 = 0; $lastSec = $m.Section
                }
                $b = New-ConsoleToolTile -Row $m -Colors $col -Size $sz -Tip $tip -Prefix $(if ($m.Pinned) { '* ' } else { '' })
                $b.Tag = $m
                $b.Add_Click({ Invoke-ConsoleViewTool -ViewForm $v -Tool $this.Tag.Tool -Blocked $this.Tag.Blocked -Rebuild { $v.Rows = Get-ConsoleToolModel; & $build $txtFind.Text } })
                # PIN FROM HERE TOO. The Recent view's hint told operators they
                # could right-click a tool "in this view or the list view", and
                # in the list view nothing happened -- the handler was only ever
                # wired in Recent. A hint that names a gesture the window does
                # not support is worse than no hint.
                $b.Add_MouseUp({
                    param($s, $e)
                    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
                    Set-ConsoleToolPinned -Name $s.Tag.Name -Pinned (-not $s.Tag.Pinned)
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
    $sh.Title.Text = "Tool list  --  $(@($v.Rows).Count) tools"
    [void]$txtFind.Focus()
    [void]$v.ShowDialog($form)
}
