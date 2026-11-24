#Requires -Version 5.1
# VIEW-ID: Palette
# VIEW-NAME: Command palette
# VIEW-NOTE: Type three letters, press Enter. For when you know the name.
# VIEW-FUNCTION: Show-ConsolePaletteView
# VIEW-ORDER: 30
<#
    View_Palette.ps1
    Admin Console view: Command palette

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    For the operator who already knows the name of the tool and does not
    want to look for it. Type, press Enter, the tool runs. Matching is on
    the tool name and its section, so 'cal' finds the calendar tools
    whatever they are called.

    THIS FILE IS NOT RUN. It is DOT-SOURCED by AdminConsole.ps1, which is why
    it defines a function and does nothing else. Running it on its own defines
    Show-ConsolePaletteView into that prompt and exits; nothing appears, and nothing
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

function Show-ConsolePaletteView {
    <#
        TYPE THREE LETTERS, PRESS ENTER.

        THE ONE VIEW THAT IS NOT MAXIMISED, and that is the whole design. The
        list view is for browsing; this is for an operator who already knows
        the tool's name and wants it gone from the screen in under a second. A
        full-screen window would make that slower, not faster.

        RANKED, NOT FILTERED ALPHABETICALLY. A name that STARTS with what was
        typed beats one that merely contains it, and among equals the tool used
        most often wins. Typing 'ma' should offer Mailbox Health before Group
        Manager, and after a fortnight it should offer whichever of those this
        operator actually opens.
    #>
    $v = New-Object System.Windows.Forms.Form
    $v.Text = 'Admin Console -- Command palette'
    $v.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $v.StartPosition = 'CenterScreen'
    $v.MaximizeBox = $false; $v.MinimizeBox = $false
    $v.KeyPreview = $true
    $sz = 14
    try { if ($form.Font.Size -gt 0) { $sz = [int]$form.Font.Size + 5 } } catch { }
    $v.ClientSize = New-Object System.Drawing.Size(760, 520)
    $v.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', $sz)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(14, 14)
    $txt.Size = New-Object System.Drawing.Size(732, ([int]($sz * 2.4)))
    $txt.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 4)))
    $v.Controls.Add($txt)

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = New-Object System.Drawing.Point(14, ([int]($sz * 2.4) + 22))
    $lst.Size = New-Object System.Drawing.Size(732, (520 - [int]($sz * 2.4) - 22 - 46))
    $lst.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', $sz)
    $lst.IntegralHeight = $false
    $v.Controls.Add($lst)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Location = New-Object System.Drawing.Point(14, (520 - 34))
    $lblHint.Size = New-Object System.Drawing.Size(732, 26)
    $lblHint.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', ([Math]::Max(8, $sz - 4)))
    $lblHint.Text = 'Up / Down to choose.  Enter to run.  Esc to close.'
    $v.Controls.Add($lblHint)

    # STATE LIVES ON THE FORM, NOT IN A VARIABLE BESIDE IT.
    #
    # THIS IS WHY Enter DID NOTHING. '$shown = ...' inside the $fill
    # scriptblock creates $shown in $fill's OWN scope; the moment $fill
    # returns that copy is gone, and the outer one is still the empty array it
    # was initialised to. $run then read an empty list, decided the selection
    # was out of range, and returned -- silently, every time.
    #
    # A PROPERTY ASSIGNMENT ON AN OBJECT HAS NO SUCH PROBLEM: $v.Shown = ...
    # mutates the form that both scriptblocks can see. The list view worked
    # from its first run for exactly this reason -- it happened to use $v.Rows.
    $v | Add-Member -NotePropertyName 'Model' -NotePropertyValue (Get-ConsoleToolModel) -Force
    $v | Add-Member -NotePropertyName 'Shown' -NotePropertyValue @() -Force

    $fill = {
        param([string]$Filter)
        $f = "$Filter".Trim()
        $rank = @()
        foreach ($m in $v.Model) {
            if (-not $f) { $rank += [PSCustomObject]@{ M = $m; R = 2 }; continue }
            $n = "$($m.Name)"
            if ($n -imatch ('^' + [regex]::Escape($f))) { $rank += [PSCustomObject]@{ M = $m; R = 0 }; continue }
            if ($n -imatch [regex]::Escape($f))         { $rank += [PSCustomObject]@{ M = $m; R = 1 }; continue }
            # Section is searched too, so 'report' finds everything filed under
            # Reporting even when the tool's own name never says so.
            if ("$($m.Section)" -imatch [regex]::Escape($f)) { $rank += [PSCustomObject]@{ M = $m; R = 3 } }
        }
        $v.Shown = @($rank | Sort-Object @{ Expression = 'R' }, @{ Expression = { -$_.M.RunCount } }, @{ Expression = { $_.M.Name } } | ForEach-Object { $_.M })
        $lst.BeginUpdate()
        try {
            $lst.Items.Clear()
            foreach ($m in $v.Shown) {
                $mark = $(if ($m.Blocked) { '  (not available)' } else { '' })
                [void]$lst.Items.Add("$($m.Name)   --   $($m.Section)$mark")
            }
            if ($lst.Items.Count) { $lst.SelectedIndex = 0 }
        } finally { $lst.EndUpdate() }
        $lblHint.Text = "$($lst.Items.Count) match(es).   Up / Down to choose.  Enter to run.  Esc to close."
    }

    $run = {
        if ($lst.SelectedIndex -lt 0 -or $lst.SelectedIndex -ge @($v.Shown).Count) { return }
        $m = $v.Shown[$lst.SelectedIndex]
        Invoke-ConsoleViewTool -ViewForm $v -Tool $m.Tool -Blocked $m.Blocked -Rebuild {
            $v.Model = Get-ConsoleToolModel
            & $fill $txt.Text
        }
    }

    # ARROWS ARE HANDLED ON THE TEXTBOX so the caret never leaves it. Moving
    # focus to the list to scroll it would mean the next keystroke typed into
    # nothing, which is the classic way a palette feels broken.
    $txt.Add_KeyDown({
        param($s, $e)
        switch ($e.KeyCode) {
            ([System.Windows.Forms.Keys]::Down) {
                if ($lst.Items.Count) { $lst.SelectedIndex = [Math]::Min($lst.Items.Count - 1, $lst.SelectedIndex + 1) }
                $e.Handled = $true; $e.SuppressKeyPress = $true
            }
            ([System.Windows.Forms.Keys]::Up) {
                if ($lst.Items.Count) { $lst.SelectedIndex = [Math]::Max(0, $lst.SelectedIndex - 1) }
                $e.Handled = $true; $e.SuppressKeyPress = $true
            }
            ([System.Windows.Forms.Keys]::Enter) { $e.Handled = $true; $e.SuppressKeyPress = $true; & $run }
            ([System.Windows.Forms.Keys]::Escape) { $v.Close() }
        }
    })
    $txt.Add_TextChanged({ & $fill $txt.Text })
    $lst.Add_DoubleClick({ & $run })
    $v.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $v.Close() }
    })

    Set-ConsoleTheme -Form $v
    & $fill ''
    [void]$txt.Focus()
    [void]$v.ShowDialog($form)
}
