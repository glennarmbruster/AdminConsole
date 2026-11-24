#Requires -Version 5.1
# VIEW-ID: Catalogue
# VIEW-NAME: Catalogue - what each tool does
# VIEW-NOTE: Every tool with the help text from its own header.
# VIEW-FUNCTION: Show-ConsoleCatalogueView
# VIEW-ORDER: 50
<#
    View_Catalogue.ps1
    Admin Console view: Catalogue - what each tool does

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    What does this toolkit actually do. Every tool listed with the help
    text read from its own header, so the description cannot drift from
    the script the way a separate document would. This is the view to
    hand a new starter.

    THIS FILE IS NOT RUN. It is DOT-SOURCED by AdminConsole.ps1, which is why
    it defines a function and does nothing else. Running it on its own defines
    Show-ConsoleCatalogueView into that prompt and exits; nothing appears, and nothing
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

function Show-ConsoleCatalogueView {
    <#
        WHAT DOES THIS TOOLKIT ACTUALLY DO. Every tool with the help text from
        its own header, for training somebody new or for the tool you have
        never opened.

        HELP IS READ ON SELECTION, NOT UP FRONT. Reading the TOOLKIT-HELP block
        out of 127 files to build this view would take seconds and would need a
        cache, and a cache needs invalidating -- which is a bug waiting on a
        script somebody edits. One file read per click is instant and is always
        current. The cheapest correct design here is also the fastest one.
    #>
    $sh = New-ConsoleViewShell -Title 'Catalogue' -Hint 'Pick a tool to read what it does.  Run opens it.  Esc returns to the console.' -ViewId 'Catalogue'
    $v = $sh.Form
    $col = Get-ConsoleViewColors -Form $v

    $lblFind = New-Object System.Windows.Forms.Label; $lblFind.Text = 'Find:'
    $txtFind = New-Object System.Windows.Forms.TextBox
    foreach ($c in @($lblFind, $txtFind)) { $sh.Bar.Controls.Add($c); [void]$sh.Left.Add($c) }

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Run this tool'
    $sh.Bar.Controls.Add($btnRun)
    [void]$sh.Right.Add($btnRun)

    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = [System.Windows.Forms.DockStyle]::Fill
    $split.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $sh.Body.Controls.Add($split)

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lst.IntegralHeight = $false
    $split.Panel1.Controls.Add($lst)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true; $txt.ReadOnly = $true; $txt.ScrollBars = 'Vertical'
    $txt.Dock = [System.Windows.Forms.DockStyle]::Fill
    $split.Panel2.Controls.Add($txt)

    # ON THE FORM, NOT BESIDE IT -- see the note in the palette. A '$shown'
    # assigned inside $fillList never reaches $showHelp, which is why the help
    # pane kept showing whichever tool it was built with and the Run button
    # kept its caption no matter what was clicked.
    # SORT, AND THE DEFAULT IS NAMED SO IT IS NOT A MYSTERY.
    #
    # The catalogue has always been in console order -- section by section,
    # then the order the buttons sit in -- which is a defensible default and
    # was invisible: nothing on screen said so, and a list of a hundred tools
    # in no announced order reads as no order at all. Reported 2026-09-12 as
    # "I have no idea how it's sorted".
    #
    # Modified is the file's LastWriteTime. It answers "what changed recently"
    # without leaving the console, which is the question after any delivery.
    $lblSort = New-Object System.Windows.Forms.Label
    $lblSort.Text = 'Sort:'
    $sh.Bar.Controls.Add($lblSort)

    $cboSort = New-Object System.Windows.Forms.ComboBox
    $cboSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cboSort.Items.Add('Console order')
    [void]$cboSort.Items.Add('Name')
    [void]$cboSort.Items.Add('Modified')
    $cboSort.SelectedIndex = 0
    $sh.Bar.Controls.Add($cboSort)

    # THE BOX BEFORE THE LABEL, and that is not a typo. The shell places its
    # right-hand cluster right to left, so the first thing added sits
    # furthest right. Adding the label first would read "[Console order] Sort:"
    # on screen.
    [void]$sh.Right.Add($cboSort)
    [void]$sh.Right.Add($lblSort)

    $v | Add-Member -NotePropertyName 'Model' -NotePropertyValue (Get-ConsoleToolModel) -Force
    $v | Add-Member -NotePropertyName 'Shown' -NotePropertyValue @() -Force

    # STAMPED ONCE, HERE. One stat per tool is nothing; one per keystroke in
    # the filter box, on a hundred files, is a stutter the operator would feel
    # and never connect to a sort they are not using.
    foreach ($m in @($v.Model)) {
        $when = $null
        try { $when = (Get-Item -LiteralPath $m.Path -ErrorAction Stop).LastWriteTime } catch { }
        $m | Add-Member -NotePropertyName 'Modified' -NotePropertyValue $when -Force
    }

    $fillList = {
        param([string]$Filter)
        $f = "$Filter".Trim()
        $rows = @($v.Model | Where-Object { -not $f -or "$($_.Name)" -imatch [regex]::Escape($f) -or "$($_.Section)" -imatch [regex]::Escape($f) })

        # Console order is the model's own order, so it is the one case that
        # sorts by nothing. Modified is newest first -- the reason to ask is
        # almost always "what moved", not "what is oldest" -- and a file that
        # could not be read sorts last rather than crashing the comparison.
        switch ("$($cboSort.SelectedItem)") {
            'Name'     { $rows = @($rows | Sort-Object { "$($_.Name)" }) }
            'Modified' {
                $rows = @($rows | Sort-Object -Descending { if ($_.Modified) { $_.Modified } else { [datetime]::MinValue } })
            }
        }
        $v.Shown = $rows

        $byDate = ("$($cboSort.SelectedItem)" -eq 'Modified')
        $lst.BeginUpdate()
        try {
            $lst.Items.Clear()
            foreach ($m in $v.Shown) {
                # The date is IN the row when that is what you sorted by, and
                # absent otherwise: a date on every line when it is not the
                # ordering is a column of noise.
                $label = "$($m.Name)"
                if ($byDate) {
                    $d = $(if ($m.Modified) { $m.Modified.ToString('yyyy-MM-dd') } else { '----------' })
                    $label = "$d   $label"
                }
                [void]$lst.Items.Add($label)
            }
            if ($lst.Items.Count) { $lst.SelectedIndex = 0 }
            else { $txt.Text = 'Nothing matches that.' }
        } finally { $lst.EndUpdate() }
        $sh.Title.Text = "Catalogue  --  $(@($v.Shown).Count) of $(@($v.Model).Count) tools, by $("$($cboSort.SelectedItem)".ToLower())"
    }

    # SelectedIndexChanged, not Click: a dropdown list changes by keyboard too.
    $cboSort.Add_SelectedIndexChanged({
        try { & $fillList $txtFind.Text } catch { }
    })

    $showHelp = {
        if ($lst.SelectedIndex -lt 0 -or $lst.SelectedIndex -ge @($v.Shown).Count) { return }
        $m = $v.Shown[$lst.SelectedIndex]
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("$($m.Name)")
        [void]$sb.AppendLine(('=' * [Math]::Max(8, "$($m.Name)".Length)))
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Section : $($m.Section)")
        [void]$sb.AppendLine("Script  : $(Split-Path -Leaf $m.Path)")
        [void]$sb.AppendLine("Modified: $(if ($m.Modified) { $m.Modified.ToString('yyyy-MM-dd HH:mm') } else { 'unknown' })")
        [void]$sb.AppendLine("Applies : $($m.Detail)")
        if ($m.LastRun) { [void]$sb.AppendLine("History : last run $($m.LastRun), $($m.RunCount) time(s)") }
        [void]$sb.AppendLine("Status  : $(if ($m.Blocked) { "NOT AVAILABLE -- $($m.Blocked)" } else { 'available' })")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('---------------------------------------------------------------')
        [void]$sb.AppendLine('')
        $help = $null
        try { $help = Get-ToolHelpText -Path $m.Path } catch { }
        if ($help) { [void]$sb.AppendLine(($help -join "`r`n")) }
        else {
            [void]$sb.AppendLine('No overview has been written for this tool yet.')
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Add one between TOOLKIT-HELP-BEGIN and TOOLKIT-HELP-END in the')
            [void]$sb.AppendLine('script header and it will appear here.')
        }
        $txt.Text = $sb.ToString()
        $txt.Select(0, 0)
        $btnRun.Text = "Run $($m.Name)"
        Update-ConsoleViewShell -Shell $sh
    }

    $lst.Add_SelectedIndexChanged({ & $showHelp })
    $txtFind.Add_TextChanged({ & $fillList $txtFind.Text })
    $btnRun.Add_Click({
        if ($lst.SelectedIndex -lt 0 -or $lst.SelectedIndex -ge @($v.Shown).Count) { return }
        $m = $v.Shown[$lst.SelectedIndex]
        Invoke-ConsoleViewTool -ViewForm $v -Tool $m.Tool -Blocked $m.Blocked -Rebuild {
            $v.Model = Get-ConsoleToolModel
            & $fillList $txtFind.Text
        }
    })
    Add-ConsoleViewFontStepper -Shell $sh -OnChange {
        $f = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]$v.FontSize))
        $lst.Font = $f
        $txt.Font = New-Object System.Drawing.Font('Consolas', ([int]$v.FontSize))
    }
    $v.Add_Resize({
        try {
            Update-ConsoleViewShell -Shell $sh
            if ($split.Width -gt 400) { $split.SplitterDistance = [int]($split.Width * 0.32) }
        } catch { }
    })

    Set-ConsoleTheme -Form $v
    Update-ConsoleViewShell -Shell $sh
    $lst.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]$v.FontSize))
    $txt.Font = New-Object System.Drawing.Font('Consolas', ([int]$v.FontSize))
    & $fillList ''
    & $showHelp
    try { if ($split.Width -gt 400) { $split.SplitterDistance = [int]($split.Width * 0.32) } } catch { }
    [void]$v.ShowDialog($form)
}
