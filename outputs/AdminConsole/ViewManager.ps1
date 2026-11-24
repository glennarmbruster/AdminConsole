#Requires -Version 5.1
# TOOLKIT-VARIANT: Any
<#
    ViewManager.ps1
    Which view the Admin Console opens into, for YOU.

    Author   : Glenn Armbruster
    Revised  : 2026-09-16
#>
# TOOLKIT-HELP-RUNS: Reads Views\ and writes one key in your own profile. Touches no mailbox, no shared file, and nobody else's settings.
# TOOLKIT-HELP-BEGIN
# The console can show the same tools several ways. This says which one YOU
# get when you launch it.
#
# WHAT A VIEW IS. A window over the console showing the same tools drawn
# differently -- as a large-text list, as a command box you type into, as
# raised tiles, as a board of what is connected right now. They are not modes
# and they do not change anything: every view launches the same scripts from
# the same Tools.json.
#
# THE CONSOLE IS ONE OF THE CHOICES. It is the window itself -- the grid plus
# the lookup box and the connect buttons -- and it is what everyone gets until
# they choose otherwise.
#
# BUTTON LAYOUT IS THE GRID AS A FILE. Same tools, same sections, same order,
# drawn as buttons, in a View_*.ps1 anyone can copy and change. That is the one
# to start from if somebody wants the grid to look different: the console
# window itself is not a file and cannot be copied.
#
# THIS IS YOURS ALONE. The choice goes in Config\Users\<your account>.json
# beside your theme and your groups. Nobody else's console changes, and you
# are not changing what any tool does -- only which screen you land on.
#
# THE WAY BACK IS ALWAYS THERE. A default view opens OVER the console, so
# closing it puts you back on the console, with the lookup box and the connect
# buttons. Nothing is hidden and nothing is locked.
#
# NEW VIEWS ARRIVE AS FILES. Each one is a single View_*.ps1 in Views\. Drop
# one in and it appears in this list and on the console's View menu; there is
# no list to edit and no config to change.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')
if (-not (Test-ConsoleLaunch)) { return }

#region Data -------------------------------------------------------------------

# The console itself, as a row in the same list as the files. It is a real
# choice and has to look like one -- a list of seven alternatives with no way
# to say "actually, the normal one" would be a trap for anyone who tried a view
# and wanted out.
$script:ButtonsRow = [PSCustomObject]@{
    Id    = 'Buttons'
    Name  = 'Console'
    Note  = 'The console window itself -- the grid, the lookup box and the connect buttons. The default.'
    File  = '(the console)'
    Error = ''
}

function Get-ViewRow {
    <#
        The console first, then every View_*.ps1 in menu order.

        A FILE WITH A PROBLEM IS LISTED WITH ITS PROBLEM. Get-ConsoleViewCatalogue
        returns broken files rather than hiding them, and so does this window:
        somebody looking for a view that is not on the menu needs to see the
        reason, not an absence.
    #>
    $rows = @($script:ButtonsRow)
    foreach ($v in @(Get-ConsoleViewCatalogue)) {
        $rows += [PSCustomObject]@{
            Id    = "$($v.Id)"
            Name  = "$($v.Name)"
            Note  = "$($v.Note)"
            File  = "$($v.File)"
            Error = "$($v.Error)"
        }
    }
    return $rows
}

function ConvertTo-ViewName {
    <#
        Turn whatever somebody typed into a view name this toolkit will accept.

        ONE ANSWER, WHATEVER THEY TYPE. 'Bobsview', 'View_Bobsview',
        'view_bobsview.ps1' and 'VIEW_View_Bobsview' all become 'Bobsview' --
        so the file is always View_Bobsview.ps1 and never View_View_Bobsview.ps1.
        The prefix is stripped REPEATEDLY, because somebody copying a copy is
        exactly how a double prefix would arrive.

        Returns '' when nothing usable is left, which the caller reports rather
        than writing a file called View_.ps1.
    #>
    param([string]$Text)
    $n = "$Text".Trim()
    $n = $n -replace '\.ps1$', ''
    # Repeatedly, not once: 'View_View_X' has to come out as 'X'.
    while ($n -match '^(view[_\-\s]+)(.+)$') { $n = $Matches[2].Trim() }
    # LETTERS AND DIGITS ONLY -- not hyphens, and that is not fussiness. The
    # name becomes a suffix on a FUNCTION name, and the scanner that finds
    # functions to rename matches Verb-Noun with ONE hyphen. 'My-Copy' produced
    # Show-ConsoleListViewMy-Copy, which the next copy of that file read as
    # 'Show-ConsoleListViewMy', wrote a blank VIEW-FUNCTION, and turned into a
    # file the console reports as 'not a view'. It also let a bare '-' through
    # as a name.
    $n = ($n -replace '[^A-Za-z0-9]', '')
    return $n
}

function Copy-ConsoleView {
    <#
        Duplicate a view file under a new name, ready to edit.

        WHY COPY AND NOT 'NEW'. Nobody wants to write a view from nothing --
        they want the tool list with three changes. A copy of a working file
        starts from something that already draws, which is the difference
        between a five-minute experiment and an afternoon.

        THE FUNCTIONS INSIDE ARE RENAMED TOO, and this is the part that would
        be a real bug if skipped. Both files get dot-sourced into the same
        scope, so a copy that still defined Show-ConsoleListView would simply
        overwrite the original -- two entries on the View menu, one behind
        both, and no error anywhere. So EVERY function the file defines is
        suffixed with the new name: Show-ConsoleListView becomes
        Show-ConsoleListViewBobsview, New-Console3DTile becomes
        New-Console3DTileBobsview. One rule, no special cases, and every
        mention inside the file moves with it.

        Returns the new path, or throws with a reason.
    #>
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$NewName
    )

    if (-not (Test-Path -LiteralPath $SourcePath)) { throw "The source view is gone: $SourcePath" }
    $dir = Get-ConsoleViewFolder
    if (-not (Test-Path -LiteralPath $dir)) { throw "There is no Views folder at $dir." }

    $target = Join-Path $dir "View_$NewName.ps1"
    if (Test-Path -LiteralPath $target) { throw "View_$NewName.ps1 already exists. Pick another name." }

    # THE FILE NAME IS NOT THE ID, so a free file name is not a free id.
    # View_3DButtons.ps1 declares VIEW-ID 'Buttons3D' -- nothing requires them
    # to match. A copy named 'Buttons3D' would write a second file claiming
    # that id, and the catalogue resolves duplicates by whichever FILE NAME
    # sorts first: the original survives or does not depending on collation,
    # which is not a design.
    if ($NewName -eq 'Buttons') {
        throw "'Buttons' is the console's own id -- it is reserved. Pick another name."
    }
    $clash = @(Get-ConsoleViewCatalogue | Where-Object { "$($_.Id)" -eq $NewName })
    if ($clash.Count -gt 0) {
        throw "$($clash[0].File) already uses the id '$NewName'. Pick another name."
    }

    # Read as UTF8 and keep the line endings: these files are BOM + CRLF by
    # toolkit rule, and a copy that lost either would be a file the next
    # encoding check flags.
    $text = [System.IO.File]::ReadAllText($SourcePath, [System.Text.Encoding]::UTF8)

    # Every function this file defines, longest first -- so renaming
    # Show-ConsoleListView cannot half-match something that contains it.
    $fns = @()
    foreach ($m in [regex]::Matches($text, '(?m)^\s*function\s+([A-Za-z]+-[A-Za-z0-9]+)')) {
        $fns += $m.Groups[1].Value
    }
    # Select-Object -Unique FIRST. 'Sort-Object -Unique { -$_.Length }' would
    # have deduped by the SORT KEY -- the length -- so of two functions whose
    # names happen to be the same length, one would silently not be renamed,
    # and the copy would quietly redefine it for the original.
    $fns = @($fns | Select-Object -Unique | Sort-Object { - $_.Length })
    if ($fns.Count -eq 0) { throw "$(Split-Path -Leaf $SourcePath) defines no functions -- there is nothing to copy." }

    $mainOld = ''
    $mm = [regex]::Match($text, '(?m)^\s*#\s*VIEW-FUNCTION:\s*(\S+)\s*$')
    if ($mm.Success) { $mainOld = $mm.Groups[1].Value }
    if (-not $mainOld) { $mainOld = $fns[0] }

    # THE OLD ID, CAPTURED BEFORE THE MANIFEST IS REWRITTEN. It is not only in
    # the header: each view passes it to New-ConsoleViewShell as a literal
    # (-ViewId 'List'), and that string is the key the operator's remembered
    # font size is stored under. A copy that kept the original's literal would
    # share that key -- step the font in the copy and the original's font moves
    # too, which is exactly the shared-state problem views are meant not to
    # have.
    $idOld = ''
    $im = [regex]::Match($text, '(?m)^\s*#\s*VIEW-ID:\s*(\S+)\s*$')
    if ($im.Success) { $idOld = $im.Groups[1].Value }

    $map = @{}
    foreach ($f in $fns) { $map[$f] = "$f$NewName" }
    foreach ($f in $fns) {
        $text = [regex]::Replace($text, "(?i)\b$([regex]::Escape($f))\b", $map[$f])
    }

    # The manifest. Replaced rather than appended, so the copy declares itself
    # and does not claim the original's id.
    $text = [regex]::Replace($text, '(?m)^\s*#\s*VIEW-ID:.*$',       "# VIEW-ID: $NewName")
    $text = [regex]::Replace($text, '(?m)^\s*#\s*VIEW-NAME:.*$',     "# VIEW-NAME: $NewName")
    $text = [regex]::Replace($text, '(?m)^\s*#\s*VIEW-FUNCTION:.*$', "# VIEW-FUNCTION: $($map[$mainOld])")
    # Last on the menu until somebody decides otherwise. A copy that landed in
    # the middle of the list would look like it had replaced something.
    $text = [regex]::Replace($text, '(?m)^\s*#\s*VIEW-ORDER:.*$',    '# VIEW-ORDER: 900')

    # And the literal the view hands to the shell -- see $idOld above.
    if ($idOld) {
        $text = [regex]::Replace($text, "(?i)(-ViewId\s+')$([regex]::Escape($idOld))(')", "`${1}$NewName`${2}")
    }

    $lead = "View_$NewName.ps1"
    $text = [regex]::Replace($text, '(?m)^(\s*)View_[A-Za-z0-9\-]+\.ps1\s*$', "`${1}$lead")
    $text = [regex]::Replace($text, '(?m)^(\s*Revised\s*:\s*).*$', "`${1}$(Get-Date -Format 'yyyy-MM-dd')")

    $note = @(
        ''
        "    COPIED from $(Split-Path -Leaf $SourcePath) on $(Get-Date -Format 'yyyy-MM-dd')."
        '    The functions in here were renamed so this file and the one it came'
        '    from can both be loaded. Edit freely -- nothing else reads them.'
    ) -join "`r`n"
    # NO COUNT ARGUMENT. The 4-argument static overload takes RegexOptions,
    # not a count -- passing 1 would quietly mean IgnoreCase. There is one
    # Revised line in a view file, so one match is what this gets.
    $text = [regex]::Replace($text, '(?m)^(\s*Revised\s*:\s*\d{4}-\d{2}-\d{2})\s*$', "`${1}$note")

    $text = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
    $enc  = New-Object System.Text.UTF8Encoding($true)

    # TEMP FILE THEN RENAME. Views\ is SHARED and the console dot-sources
    # everything in it, so a write that dies part way -- a dropped share, a
    # quota -- would leave a truncated View_X.ps1 that every operator's console
    # then reports as a broken view. A rename is atomic; a half file never
    # exists under the real name.
    $tmp = "$target.tmp"
    [System.IO.File]::WriteAllText($tmp, $text, $enc)
    try {
        Move-Item -LiteralPath $tmp -Destination $target -Force -ErrorAction Stop
    } catch {
        try { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } catch { }
        throw
    }
    return $target
}

#endregion

#region UI ---------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'View Manager'
$form.ClientSize = New-Object System.Drawing.Size(760, 470)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWho = New-Object System.Windows.Forms.Label
$lblWho.Location = New-Object System.Drawing.Point(12, 12)
$lblWho.Size = New-Object System.Drawing.Size(736, 20)
$lblWho.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWho.Text = $(if (Get-ActiveTenantLabel) {
                    "Views for '$(Get-UserProfileKey)' in '$(Get-ActiveTenantLabel)'"
                } else {
                    "Views for '$(Get-UserProfileKey)'"
                })
$form.Controls.Add($lblWho)

$lblNow = New-Object System.Windows.Forms.Label
$lblNow.Location = New-Object System.Drawing.Point(12, 34)
$lblNow.Size = New-Object System.Drawing.Size(736, 18)
$lblNow.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblNow)

$lv = New-Object System.Windows.Forms.ListView
$lv.Location = New-Object System.Drawing.Point(12, 58)
$lv.Size = New-Object System.Drawing.Size(736, 250)
$lv.View = 'Details'
$lv.FullRowSelect = $true
$lv.MultiSelect = $false
$lv.HideSelection = $false
[void]$lv.Columns.Add('Opens into', 78)
[void]$lv.Columns.Add('View', 210)
[void]$lv.Columns.Add('What it shows', 330)
[void]$lv.Columns.Add('File', 110)
$form.Controls.Add($lv)

$btnSet = New-Object System.Windows.Forms.Button
$btnSet.Text = 'Open into this one'
$btnSet.Location = New-Object System.Drawing.Point(12, 318)
$btnSet.Size = New-Object System.Drawing.Size(160, 30)
$form.Controls.Add($btnSet)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'Back to the console'
$btnClear.Location = New-Object System.Drawing.Point(182, 318)
$btnClear.Size = New-Object System.Drawing.Size(170, 30)
$form.Controls.Add($btnClear)

$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = 'Copy to a new view...'
$btnCopy.Location = New-Object System.Drawing.Point(362, 318)
$btnCopy.Size = New-Object System.Drawing.Size(170, 30)
$form.Controls.Add($btnCopy)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(638, 318)
$btnClose.Size = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(12, 358)
$txtOut.Size = New-Object System.Drawing.Size(736, 100)
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$form.CancelButton = $btnClose

$script:Refresh = {
    $cur = ''
    try { $cur = Get-UserDefaultView } catch { }
    if (-not $cur) { $cur = 'Buttons' }

    $rows = Get-ViewRow
    # THE CURRENT DEFAULT MAY NAME A VIEW THAT IS NOT HERE -- a file removed, a
    # profile roamed from a machine with more views on it. Saying so is the
    # point: the console silently falls back to its own window, and without
    # this line the operator has no way to find out why.
    $known = @($rows | Where-Object { $_.Id -eq $cur }).Count -gt 0

    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        foreach ($r in $rows) {
            $it = New-Object System.Windows.Forms.ListViewItem($(if ($r.Id -eq $cur) { 'YES' } else { '' }))
            [void]$it.SubItems.Add($r.Name)
            [void]$it.SubItems.Add($(if ($r.Error) { $r.Error } else { $r.Note }))
            [void]$it.SubItems.Add($r.File)
            $it.Tag = $r
            if ($r.Error) { $it.ForeColor = [System.Drawing.Color]::Firebrick }
            elseif ($r.Id -eq $cur) { $it.Font = New-Object System.Drawing.Font($lv.Font, [System.Drawing.FontStyle]::Bold) }
            [void]$lv.Items.Add($it)
        }
    } finally { $lv.EndUpdate() }

    $name = $(if ($known) { @($rows | Where-Object { $_.Id -eq $cur })[0].Name } else { "$cur (not installed here)" })
    $lblNow.Text = "The Admin Console opens into: $name"
    $btnClear.Enabled = ($cur -ne 'Buttons')
}

function Save-DefaultView {
    <#
        Write the choice, and say exactly where it went.

        THE SCOPE QUESTION IS ONLY ASKED WHEN IT HAS TWO ANSWERS. In a
        single-tenant install there are no tenant blocks, so Save-UserLayout
        writes Default whatever it is told -- asking would be a dialog whose
        answers do the same thing, which is how operators learn to click
        through dialogs without reading them.
    #>
    param([string]$Id, [string]$Label)

    $all = $true
    if ((Test-MultiTenantMode) -and (Get-ActiveTenantKey)) {
        $here = "'$(Get-ActiveTenantLabel)'"
        $a = Show-ToolDialog(
            ("Open into $Label -- where?" + "`r`n`r`n" +
             "YES     in $here only" + "`r`n" +
             "NO      in every tenant" + "`r`n" +
             "CANCEL  change nothing" + "`r`n`r`n" +
             "This is yours alone either way. Nobody else's console changes."),
            'View Manager', 'YesNoCancel', 'Question')
        if ($a -eq [System.Windows.Forms.DialogResult]::Cancel) { return $null }
        $all = ($a -eq [System.Windows.Forms.DialogResult]::No)
    }
    return (Set-UserDefaultView -Id $Id -AllTenants:$all)
}

$btnSet.Add_Click({
    try {
        if ($lv.SelectedItems.Count -lt 1) {
            $txtOut.Text = 'Pick a view in the list first.'
            return
        }
        $r = $lv.SelectedItems[0].Tag
        if ($r.Error) {
            $txtOut.Text = "$($r.File) cannot be used as your default:" + [Environment]::NewLine +
                           "  $($r.Error)" + [Environment]::NewLine + [Environment]::NewLine +
                           'Fix the file, or pick another view.'
            return
        }
        $p = Save-DefaultView -Id $r.Id -Label "'$($r.Name)'"
        if (-not $p) { $txtOut.Text = 'Nothing changed.'; return }
        & $script:Refresh
        $txtOut.Text = "SAVED" + [Environment]::NewLine + [Environment]::NewLine +
            "  The Admin Console will open into '$($r.Name)'." + [Environment]::NewLine +
            "  Written to your profile: $p" + [Environment]::NewLine + [Environment]::NewLine +
            'Close and reopen the Admin Console to see it. The view opens OVER the console --' + [Environment]::NewLine +
            'close the view and the console is underneath it, exactly as it is now.'
        Write-CompletedLog -Action 'Configure' -Target $p -Result 'Success' `
            -Detail "Default console view set to '$($r.Id)' for $(Get-UserProfileKey)" -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Setting your default view'
        Write-FailedLog -Action 'Configure' -Target 'Default console view' -ErrorRecord $_ -Detail 'View Manager'
    }
})

$btnClear.Add_Click({
    try {
        $p = Save-DefaultView -Id 'Buttons' -Label 'the console'
        if (-not $p) { $txtOut.Text = 'Nothing changed.'; return }
        & $script:Refresh
        $txtOut.Text = "SAVED" + [Environment]::NewLine + [Environment]::NewLine +
            '  The Admin Console will open into the console window itself.' + [Environment]::NewLine +
            "  Written to your profile: $p" + [Environment]::NewLine + [Environment]::NewLine +
            'Close and reopen the Admin Console to see it.'
        Write-CompletedLog -Action 'Configure' -Target $p -Result 'Success' `
            -Detail "Default console view cleared for $(Get-UserProfileKey)" -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Clearing your default view'
    }
})

$btnCopy.Add_Click({
    try {
        if ($lv.SelectedItems.Count -lt 1) { $txtOut.Text = 'Pick a view in the list first.'; return }
        $r = $lv.SelectedItems[0].Tag
        if ($r.Id -eq 'Buttons') {
            $txtOut.Text = 'The console is a window, not a file -- there is nothing to copy.' + [Environment]::NewLine + [Environment]::NewLine +
                           'To customise the GRID, copy Button Layout instead: it is the same tools in' + [Environment]::NewLine +
                           'the same sections, as a file you can edit.'
            return
        }
        if ($r.Error -or -not $r.File) {
            $txtOut.Text = "That row is not a usable view file:" + [Environment]::NewLine + "  $($r.Error)"
            return
        }

        # THREE LINES, NOT FOUR. Show-InputDialog's prompt label is a fixed
        # height and has no parameter for it; a fourth line is simply not drawn.
        $typed = Show-InputDialog -Title 'Copy view' -Width 520 -Prompt (
            "Copy '$($r.Name)' to a new view." + [Environment]::NewLine +
            'Name it -- View_ and .ps1 are added for you, so Bobsview' + [Environment]::NewLine +
            'is saved as View_Bobsview.ps1.')
        if (-not $typed) { $txtOut.Text = 'Nothing copied.'; return }

        $clean = ConvertTo-ViewName -Text $typed
        if (-not $clean) {
            $txtOut.Text = "'$typed' leaves nothing usable as a name." + [Environment]::NewLine +
                           'Letters, digits and hyphens; the View_ prefix is added for you.'
            return
        }

        $src = Join-Path (Get-ConsoleViewFolder) $r.File
        $new = Copy-ConsoleView -SourcePath $src -NewName $clean
        & $script:Refresh
        $txtOut.Text = "COPIED" + [Environment]::NewLine + [Environment]::NewLine +
            "  $(Split-Path -Leaf $src)   ->   $(Split-Path -Leaf $new)" + [Environment]::NewLine + [Environment]::NewLine +
            $(if ($clean -ne "$typed".Trim()) { "  You typed '$typed'; it was saved as View_$clean.ps1." + [Environment]::NewLine + [Environment]::NewLine } else { '' }) +
            'It is on the View menu now -- reopen the Admin Console to see it. Edit the file' + [Environment]::NewLine +
            'to change what it draws; the functions inside were renamed so both copies load.' + [Environment]::NewLine + [Environment]::NewLine +
            'THIS IS A SHARED FILE. A view in Views\ is code, so everyone who uses this' + [Environment]::NewLine +
            'folder gets it on their menu. Only which view you OPEN INTO is personal.'
        Write-CompletedLog -Action 'Configure' -Target $new -Result 'Success' `
            -Detail "View copied from $(Split-Path -Leaf $src) by $(Get-UserProfileKey)" -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Copying the view'
        Write-FailedLog -Action 'Configure' -Target 'Copy view' -ErrorRecord $_ -Detail 'View Manager'
    }
})

# Double-click is the gesture people try first on a list like this.
$lv.Add_DoubleClick({ $btnSet.PerformClick() })

& $script:Refresh
$txtOut.Text = 'Pick a view and choose "Open into this one". Yours alone -- Config\Users\' +
               "$(Get-UserProfileKey).json." + [Environment]::NewLine +
               'To TRY a view without committing to it, use the View button on the console instead.' + [Environment]::NewLine +
               'Views cannot be opened from here: they are drawn by the console and need its window.'

if (@(Get-ConsoleViewCatalogue).Count -eq 0) {
    $txtOut.Text = 'No views are installed.' + [Environment]::NewLine + [Environment]::NewLine +
        "  Expected: $(Get-ConsoleViewFolder)" + [Environment]::NewLine + [Environment]::NewLine +
        'Each view is one View_*.ps1 file in that folder. If this install was updated recently,' + [Environment]::NewLine +
        'the Views folder was probably not copied with it.'
}

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'View Manager' | Out-Null
Set-ConsoleTheme -Form $form
[void]$form.ShowDialog()

#endregion
