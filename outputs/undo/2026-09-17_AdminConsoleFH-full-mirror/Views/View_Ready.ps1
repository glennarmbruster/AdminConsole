#Requires -Version 5.1
# VIEW-ID: Ready
# VIEW-NAME: What I can run right now
# VIEW-NOTE: Grouped by what each tool needs, connected groups first.
# VIEW-FUNCTION: Show-ConsoleReadyView
# VIEW-ORDER: 70
<#
    View_Ready.ps1
    Admin Console view: What I can run right now

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    The same tools grouped by what they NEED rather than what they are:
    Exchange Online, on-premises Exchange, AD, SharePoint. Groups whose
    requirement is already satisfied come first, so the top of the window
    is always the part that will work.

    THIS FILE IS NOT RUN. It is DOT-SOURCED by AdminConsole.ps1, which is why
    it defines a function and does nothing else. Running it on its own defines
    Show-ConsoleReadyView into that prompt and exits; nothing appears, and nothing
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

function Show-ConsoleReadyView {
    <#
        THE SAME TOOLS, GROUPED BY WHAT THEY NEED, WITH THE CONNECTED GROUPS
        FIRST.

        Built for a console with a hundred tools across three kinds of
        connection. An operator who has only signed in to Exchange Online does
        not want to read the grid and work out which buttons are grey; they
        want the tools they can use right now at the top and everything else
        below, labelled with what it is waiting for.

        THE GROUPING COMES FROM THE SAME ANSWER THAT GREYS THE BUTTONS.
        Get-ToolAvailability is what the grid asks, so this view cannot say a
        tool is ready while the grid says it is not. It is the list view with
        a different sort key and nothing else changed, on purpose: the tiles,
        the pinning, the find box and the font stepper are all inherited.
    #>
    $sh = New-ConsoleViewShell -Title 'What I can run' -Hint 'Tools you can run right now are at the top. Each group says what it needs and whether that is connected.  Click to run.  RIGHT-CLICK to pin.  Esc returns to the console.' -ViewId 'Ready'
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
    $tbl = New-Object System.Windows.Forms.TableLayoutPanel
    $tbl.Dock = [System.Windows.Forms.DockStyle]::Top
    $tbl.AutoSize = $true
    $tbl.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $tbl.ColumnCount = 3
    $tbl.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 24)
    $scroll.Controls.Add($tbl)

    # One row per tool, each carrying its Needs label and whether every need
    # is met. Unknown availability (never classified) counts as ready, which
    # is the grid's fail-open rule too.
    $classify = {
        $map = $null
        try { $map = Get-CapabilityStateMap } catch { }
        $rows = @()
        foreach ($m in @(Get-ConsoleToolModel)) {
            $needs = @(); $ok = $true
            $a = $null
            try { $a = $script:ToolAvail["$($m.Name)"] } catch { }
            if ($a) {
                foreach ($id in @($a.Requires)) {
                    $lbl = "$id"; $up = $true
                    try {
                        if ($map -and $map[$id]) {
                            if ("$($map[$id].Label)") { $lbl = "$($map[$id].Label)" }
                            $up = [bool]$map[$id].Available
                        }
                    } catch { }
                    $needs += $lbl
                    if (-not $up) { $ok = $false }
                }
                if (-not $a.Enabled) { $ok = $false }
            }
            $needText = $(if ($needs.Count) { ($needs | Sort-Object -Unique) -join ' + ' } else { 'No connection needed' })
            $m | Add-Member -NotePropertyName 'Needs' -NotePropertyValue $needText -Force
            $m | Add-Member -NotePropertyName 'Ready' -NotePropertyValue $ok -Force
            $rows += $m
        }
        # Ready groups first, then by what they need, then by name.
        return @($rows | Sort-Object @{ Expression = { -not $_.Ready } }, Needs, Name)
    }

    $build = {
        param([string]$Filter)
        $sz = [int]$v.FontSize
        $cols = [int]$v.Cols
        $col = Get-ConsoleViewColors -Form $v
        $fh = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 1)), [System.Drawing.FontStyle]::Bold)
        $tbl.SuspendLayout()
        try {
            $old = @($tbl.Controls); $tbl.Controls.Clear(); $tbl.RowStyles.Clear()
            foreach ($c in $old) { try { $c.Dispose() } catch { } }
            $tbl.ColumnCount = $cols
            $tbl.ColumnStyles.Clear()
            for ($i = 0; $i -lt $cols; $i++) {
                [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, (100.0 / $cols))))
            }
            $flt = "$Filter".Trim()
            $row = 0; $c0 = 0; $shown = 0; $ready = 0; $lastKey = $null
            foreach ($m in $v.Rows) {
                if ($flt -and "$($m.Name)" -notmatch [regex]::Escape($flt)) { continue }
                $key = "$($m.Ready)|$($m.Needs)"
                if ($key -ne $lastKey) {
                    if ($c0 -ne 0) { $row++; $c0 = 0 }
                    $h = New-Object System.Windows.Forms.Label
                    $h.Text = $(if ($m.Ready) { "READY  --  needs: $($m.Needs)" } else { "NOT CONNECTED  --  needs: $($m.Needs)" })
                    $h.Font = $fh
                    $h.BackColor = $col.HeaderBack
                    $h.ForeColor = $(if ($m.Ready) { $col.HeaderFore } else { $col.Dim })
                    $h.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $h.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $h.Height = [int]($sz * 2.6)
                    $h.Margin = New-Object System.Windows.Forms.Padding(2, ([int]($sz * 0.9)), 2, 2)
                    $tbl.Controls.Add($h, 0, $row); $tbl.SetColumnSpan($h, $cols)
                    $row++; $c0 = 0; $lastKey = $key
                }
                $b = New-ConsoleToolTile -Row $m -Colors $col -Size $sz -Tip $tip -Prefix $(if ($m.Pinned) { '* ' } else { '' })
                $b.Tag = $m
                $b.Add_Click({ Invoke-ConsoleViewTool -ViewForm $v -Tool $this.Tag.Tool -Blocked $this.Tag.Blocked -Rebuild { $v.Rows = & $classify; & $build $txtFind.Text } })
                $b.Add_MouseUp({
                    param($s, $e)
                    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
                    Set-ConsoleToolPinned -Name $s.Tag.Name -Pinned (-not $s.Tag.Pinned)
                    $v.Rows = & $classify
                    & $build $txtFind.Text
                })
                $tbl.Controls.Add($b, $c0, $row)
                $shown++; $c0++
                if ($m.Ready) { $ready++ }
                if ($c0 -ge $cols) { $c0 = 0; $row++ }
            }
            $total = @($v.Rows).Count
            $lblCount.Text = $(if ($flt) { "$shown of $total tools" } else { "$ready of $total ready" })
        } finally { $tbl.ResumeLayout() }
        Update-ConsoleViewShell -Shell $sh
    }

    $txtFind.Add_TextChanged({ & $build $txtFind.Text })
    $cbo.Add_SelectedIndexChanged({ $v.Cols = [int]"$($cbo.SelectedItem)"; & $build $txtFind.Text })
    Add-ConsoleViewFontStepper -Shell $sh -OnChange { & $build $txtFind.Text }
    $v.Add_Resize({ try { Update-ConsoleViewShell -Shell $sh; $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })
    $scroll.Add_Resize({ try { $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })

    Set-ConsoleTheme -Form $v
    $v.Rows = & $classify
    Update-ConsoleViewShell -Shell $sh
    & $build ''
    $readyN = @($v.Rows | Where-Object { $_.Ready }).Count
    $sh.Title.Text = "What I can run  --  $readyN of $(@($v.Rows).Count) ready"
    [void]$txtFind.Focus()
    [void]$v.ShowDialog($form)
}
