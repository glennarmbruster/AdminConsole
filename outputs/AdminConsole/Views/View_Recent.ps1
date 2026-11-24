#Requires -Version 5.1
# VIEW-ID: Recent
# VIEW-NAME: Recent and pinned
# VIEW-NOTE: The handful you actually use. Right-click to pin.
# VIEW-FUNCTION: Show-ConsoleRecentView
# VIEW-ORDER: 20
<#
    View_Recent.ps1
    Admin Console view: Recent and pinned

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    Most operators use a dozen of the tools most days and meet the rest
    twice a year. This view is those dozen: what was run recently, and
    what was pinned deliberately. Right-click a tile to pin or unpin it;
    the list lives in the console state file, per operator.

    THIS FILE IS NOT RUN. It is DOT-SOURCED by AdminConsole.ps1, which is why
    it defines a function and does nothing else. Running it on its own defines
    Show-ConsoleRecentView into that prompt and exits; nothing appears, and nothing
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

function Show-ConsoleRecentView {
    <#
        THE FIFTEEN TOOLS YOU ACTUALLY USE, out of a hundred and twenty seven.

        THREE BANDS, AND THE ORDER IS THE ARGUMENT. Pinned first, because it is
        the only one the operator chose deliberately. Recent second, because
        "what was I just doing" is the commonest question after a tool closes.
        Most used last, because a frequency count is a description of the past
        few months and ages slowly -- useful, but never the thing you wanted
        this second.

        PIN IS A RIGHT-CLICK, not a column of checkboxes. Pinning happens about
        twice a year; giving it permanent screen furniture would cost width on
        every row forever to serve a decision nobody is making today.
    #>
    $sh = New-ConsoleViewShell -Title 'Recent and pinned' -Hint 'Click to run.  RIGHT-CLICK a tool to pin or unpin it.  Esc returns to the console.' -ViewId 'Recent'
    $v = $sh.Form
    $v | Add-Member -NotePropertyName 'Cols' -NotePropertyValue 3 -Force
    $tip = New-Object System.Windows.Forms.ToolTip; $tip.AutoPopDelay = 20000

    $scroll = New-Object System.Windows.Forms.Panel
    $scroll.Dock = [System.Windows.Forms.DockStyle]::Fill; $scroll.AutoScroll = $true
    $sh.Body.Controls.Add($scroll)
    $tbl = New-Object System.Windows.Forms.TableLayoutPanel
    $tbl.Dock = [System.Windows.Forms.DockStyle]::Top
    $tbl.AutoSize = $true; $tbl.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $tbl.ColumnCount = 3
    $tbl.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 24)
    $scroll.Controls.Add($tbl)

    $build = {
        $sz = [int]$v.FontSize
        $cols = [int]$v.Cols
        $col = Get-ConsoleViewColors -Form $v
        $fh = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 1)), [System.Drawing.FontStyle]::Bold)
        $model = Get-ConsoleToolModel

        $pinned = @($model | Where-Object { $_.Pinned } | Sort-Object Name)
        $recent = @($model | Where-Object { $_.LastRun } | Sort-Object { $_.LastRun } -Descending | Select-Object -First 15)
        $often  = @($model | Where-Object { $_.RunCount -gt 0 } | Sort-Object RunCount -Descending | Select-Object -First 15)

        $bands = @()
        $bands += [PSCustomObject]@{ Title = 'Pinned'; Rows = $pinned; Empty = 'Nothing pinned yet -- right-click any tool in this view or the list view to pin it.' }
        $bands += [PSCustomObject]@{ Title = 'Recently used'; Rows = $recent; Empty = 'No history yet. Tools you run from any view appear here.' }
        $bands += [PSCustomObject]@{ Title = 'Most used'; Rows = $often; Empty = '' }

        $tbl.SuspendLayout()
        try {
            $old = @($tbl.Controls); $tbl.Controls.Clear(); $tbl.RowStyles.Clear()
            foreach ($c in $old) { try { $c.Dispose() } catch { } }
            $tbl.ColumnCount = $cols; $tbl.ColumnStyles.Clear()
            for ($i = 0; $i -lt $cols; $i++) {
                [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, (100.0 / $cols))))
            }
            $row = 0
            foreach ($band in $bands) {
                if (@($band.Rows).Count -eq 0 -and -not $band.Empty) { continue }
                $h = New-Object System.Windows.Forms.Label
                $h.Text = "$($band.Title)"; $h.Font = $fh
                $h.BackColor = $col.HeaderBack; $h.ForeColor = $col.HeaderFore
                $h.Dock = [System.Windows.Forms.DockStyle]::Fill
                $h.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                $h.Height = [int]($sz * 2.6)
                $h.Margin = New-Object System.Windows.Forms.Padding(2, ([int]($sz * 0.9)), 2, 2)
                $tbl.Controls.Add($h, 0, $row); $tbl.SetColumnSpan($h, $cols); $row++
                if (@($band.Rows).Count -eq 0) {
                    $e = New-Object System.Windows.Forms.Label
                    $e.Text = "   $($band.Empty)"; $e.ForeColor = $col.Dim
                    $e.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $e.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $e.Height = [int]($sz * 2.4)
                    $tbl.Controls.Add($e, 0, $row); $tbl.SetColumnSpan($e, $cols); $row++
                    continue
                }
                $c0 = 0
                foreach ($m in $band.Rows) {
                    $pre = $(if ($m.Pinned) { '* ' } else { '' })
                    $b = New-ConsoleToolTile -Row $m -Colors $col -Size $sz -Tip $tip -Prefix $pre
                    $b.Tag = $m
                    $b.Add_Click({ Invoke-ConsoleViewTool -ViewForm $v -Tool $this.Tag.Tool -Blocked $this.Tag.Blocked -Rebuild { & $build } })
                    $b.Add_MouseUp({
                        param($s, $e)
                        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
                        Set-ConsoleToolPinned -Name $s.Tag.Name -Pinned (-not $s.Tag.Pinned)
                        & $build
                    })
                    $tbl.Controls.Add($b, $c0, $row)
                    $c0++
                    if ($c0 -ge $cols) { $c0 = 0; $row++ }
                }
                if ($c0 -ne 0) { $row++ }
            }
        } finally { $tbl.ResumeLayout() }
        Update-ConsoleViewShell -Shell $sh
        $sh.Title.Text = "Recent and pinned  --  $(@($pinned).Count) pinned, $(@($recent).Count) recent"
    }

    Add-ConsoleViewFontStepper -Shell $sh -OnChange { & $build }
    $v.Add_Resize({ try { Update-ConsoleViewShell -Shell $sh; $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })
    $scroll.Add_Resize({ try { $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })

    Set-ConsoleTheme -Form $v
    Update-ConsoleViewShell -Shell $sh
    & $build
    [void]$v.ShowDialog($form)
}
