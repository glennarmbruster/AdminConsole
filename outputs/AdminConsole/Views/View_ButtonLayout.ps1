#Requires -Version 5.1
# VIEW-ID: ButtonLayout
# VIEW-NAME: Button Layout
# VIEW-NOTE: The tool grid as a view -- the console's shape, in a file you can edit.
# VIEW-FUNCTION: Show-ConsoleButtonLayoutView
# VIEW-ORDER: 5
<#
    View_ButtonLayout.ps1
    Admin Console view: Button Layout

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    THE GRID, AS A FILE SOMEBODY CAN CHANGE.

    The console window IS a button grid, and it is excellent at being one --
    but it is also the lookup box, the connect buttons, the PIM clock, the
    status strip, drag-to-arrange, collapsible sections and the zoom walk, all
    wound together in one 5,800-line script. Handing an engineer that file and
    saying "make the buttons the way you like" is not a small ask.

    So this is the grid on its own: the same tools, the same sections, the same
    order, drawn as ordinary console buttons, in ninety lines that do nothing
    else. Copy it in View Manager and change what you like -- the columns, the
    height, the colours, what goes on the caption, which tools appear at all --
    and the console is untouched. That is the whole point of it existing.

    WHAT IT IS NOT. It is not the console window. There is no lookup box here,
    no connect buttons, no status strip, and arranging is not drag-and-drop --
    those live on the console and stay there. Esc returns to it.

    WHAT IT HONOURS. Sections and their order, this operator's ButtonsPerRow,
    parked tools (they never reach the model), and the capability greying: a
    tool this window cannot run is dim and says why when you hover it. All of
    that comes from Get-ConsoleToolModel, which is the single answer to "what
    tools are there and which can run" -- so this view cannot disagree with the
    console about it.

    WHY THE COLUMNS ARE A DROPDOWN. On the console the column count reshapes
    the whole window and is a saved setting. Here it is a view control, seeded
    from that saved setting: somebody trying four columns against six is
    answering a question, not making a decision, and should not have to save
    anything to do it.

    THIS FILE IS NOT RUN. It is DOT-SOURCED by AdminConsole.ps1, which is why
    it defines a function and does nothing else. Running it on its own defines
    Show-ConsoleButtonLayoutView into that prompt and exits; nothing appears,
    and nothing is wrong. Open the console instead.

    ADDING A VIEW. Copy this file from View Manager -- it renames the functions
    and the header for you. By hand: change the five VIEW- lines at the top and
    the function name, and drop it in this folder. Keep the VIEW-ID short and
    unique; it is what a per-user default view is stored as.
#>

function Get-ConsoleButtonLayoutColumns {
    <#
        How many across this operator has chosen for the console, or 5.

        Profile first, then the shared Sections.json -- the same order the
        console resolves it in. A view that opened at its own column count
        while the console was set to six would look like a different toolkit.

        Clamped to what the dropdown offers, so a hand-edited 12 does not
        produce a selector with nothing selected in it.
    #>
    $n = 0
    try { $n = [int](Get-UserLayoutProperty -Name 'ButtonsPerRow') } catch { $n = 0 }
    if ($n -lt 1) {
        try {
            $sp = Get-ConfigFilePath -Name 'Sections.json'
            if (Test-Path -LiteralPath $sp) {
                $sj = Get-Content -LiteralPath $sp -Raw -Encoding UTF8 | ConvertFrom-Json
                $n = [int]$sj.ButtonsPerRow
            }
        } catch { $n = 0 }
    }
    if ($n -lt 2) { $n = 5 }
    if ($n -gt 8) { $n = 8 }
    return $n
}

function Show-ConsoleButtonLayoutView {
    <#
        The tool grid, full screen, grouped by section.
    #>
    $sh = New-ConsoleViewShell -Title 'Button Layout' -Hint 'Click or press Enter to run.  RIGHT-CLICK to pin (pinned tools show a *).  Esc returns to the console.' -ViewId 'ButtonLayout'
    $v = $sh.Form
    $start = Get-ConsoleButtonLayoutColumns
    $v | Add-Member -NotePropertyName 'Cols' -NotePropertyValue $start -Force
    $v | Add-Member -NotePropertyName 'Rows' -NotePropertyValue @() -Force
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.AutoPopDelay = 20000

    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($n in 2, 3, 4, 5, 6, 7, 8) { [void]$cbo.Items.Add("$n") }
    $cbo.SelectedItem = "$start"
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
    $tbl.ColumnCount = $start
    $tbl.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 28)
    $scroll.Controls.Add($tbl)

    $build = {
        param([string]$Filter)
        $sz = [int]$v.FontSize
        $cols = [int]$v.Cols
        $col = Get-ConsoleViewColors -Form $v

        # RESOLVED BEFORE ANYTHING IS ASSIGNED, for the reason spelled out in
        # Get-ConsoleViewColors: ConvertTo-ThemeColor returns $null for a key a
        # theme does not define, and assigning $null to a Color property throws
        # where it is assigned, not where it is tested.
        $edge = $null
        if ($script:Theme) { $edge = ConvertTo-ThemeColor $script:Theme.ButtonBorder }
        if (-not $edge) { $edge = Get-BlendedColor -From $col.TileBack -To ([System.Drawing.Color]::Black) -Amount 0.25 }

        $fh = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 2)), [System.Drawing.FontStyle]::Bold)
        $fb = New-Object System.Drawing.Font('Microsoft Sans Serif', $sz)
        $tbl.SuspendLayout()
        try {
            # Disposed, not just cleared. Each button owns a Font and a Tag and
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

                $b = New-Object System.Windows.Forms.Button
                $b.Text = $(if ($m.Pinned) { "* $($m.Name)" } else { "$($m.Name)" })
                $b.Font = $fb
                $b.Dock = [System.Windows.Forms.DockStyle]::Fill
                $b.Height = [int]($sz * 3.4)
                $b.Margin = New-Object System.Windows.Forms.Padding(4)
                $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $b.FlatAppearance.BorderSize = 1
                $b.FlatAppearance.BorderColor = $edge
                $b.BackColor = $col.TileBack
                # DIM, NOT DISABLED. A disabled button cannot raise a tooltip,
                # and the tooltip is where the REASON lives -- so greying one
                # out the WinForms way would remove the only explanation of why
                # it is grey. Invoke-ConsoleViewTool refuses the click instead.
                $b.ForeColor = $(if ($m.Blocked) { $col.Dim } else { $col.TileFore })
                $b.Tag = [PSCustomObject]@{ Tool = $m.Tool; Blocked = $m.Blocked; Name = $m.Name; Row = $m }

                $b.Add_Click({ Invoke-ConsoleViewTool -ViewForm $v -Tool $this.Tag.Tool -Blocked $this.Tag.Blocked -Rebuild { $v.Rows = Get-ConsoleToolModel; & $build $txtFind.Text } })
                $b.Add_MouseUp({
                    param($s, $e)
                    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
                    Set-ConsoleToolPinned -Name $s.Tag.Name -Pinned (-not $s.Tag.Row.Pinned)
                    $v.Rows = Get-ConsoleToolModel
                    & $build $txtFind.Text
                })

                $seen = ''
                if ("$($m.LastRun)") { $seen = "`r`nlast run $($m.LastRun)  ($($m.RunCount)x)" }
                $tip.SetToolTip($b, $(if ($m.Blocked) { "$($m.Name)`r`n$($m.Detail)$seen`r`n`r`nNOT AVAILABLE: $($m.Blocked)" } else { "$($m.Name)`r`n$($m.Detail)$seen" }))

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
    $sh.Title.Text = "Button Layout  --  $(@($v.Rows).Count) tools"
    [void]$txtFind.Focus()
    [void]$v.ShowDialog($form)
}
