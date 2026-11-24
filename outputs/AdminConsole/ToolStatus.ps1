#Requires -Version 5.1
<#
    ToolStatus.ps1
    Which scripts in this folder the console actually uses, and a way to park
    the ones it does not.

    Reasoning: ScriptNotes/ToolStatus.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-08
#>
# TOOLKIT-HELP-BEGIN
# Lists every .ps1 beside the console and says what loads it.
#
#   AdminConsole   -- a button in the console points at it.
#   Dashboard      -- the dashboard can launch it.
#   Not loaded     -- nothing opens it. Safe to park.
#   Parked         -- already moved to DisabledTools.
#
# Green means something loads it. Yellow means nothing does. Grey means it is
# already parked.
#
# Click the ? column, or double-click a row, to read that tool's own help --
# the same text its own ? button shows. Yellow does not mean useless: a script
# may be deliberately shelved, or called by another script rather than opened
# by a button. Read it before parking it.
#
# Last run comes from the audit log, so it is when the tool last did something
# worth recording -- not when it was last opened.
#
# Common.ps1, AdminConsole.ps1, Publish-Toolkit.ps1 and this tool are not
# listed. They are infrastructure and must never be parked.
#
# Parking MOVES a file to DisabledTools. It deletes nothing, and Restore puts
# it back. This deployment keeps only what it uses; the master copy keeps
# everything.
#
# AFTER MOVING ANYTHING, RE-RUN Publish-Toolkit. The version manifest lists
# files by name, so a parked script reads as missing until the toolkit is
# re-stamped.
# TOOLKIT-HELP-END

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:DisabledDir = Join-Path $PSScriptRoot 'DisabledTools'

# NEVER MOVABLE, and the list is short on purpose.
#
# Everything dot-sources Common.ps1 and AdminConsole.ps1 is the console itself,
# so parking either turns the whole toolkit off with no obvious way back.
# ToolStatus.ps1 is here because a tool that can park itself is a tool that can
# strand you -- the window would close and the file would be gone.
#
# Config editors are NOT core. They are genuinely optional in a finished
# deployment, they are merely dangerous to lose, and 'Referenced' already says
# so. A list of things the operator may not touch should contain only the
# things that actually cannot be touched.
# EXCLUDED FROM THE LIST ENTIRELY, rather than shown and protected.
#
# These are infrastructure, not tools: nothing loads them because everything
# does. Listing them only raises a question the answer to which is "never touch
# it", and a row you must not act on is a row worth removing.
#
# SETTINGS.PS1 IS HERE FOR A DIFFERENT REASON, and it was briefly solved the
# wrong way. It shows as an orphan because Dashboard opens the settings editor
# by calling Show-SettingsEditor -- a function in Common.ps1 -- rather than by
# launching the file, so no script anywhere contains the string 'Settings.ps1'
# and the reference scan correctly finds nothing.
#
# The first fix was a $NeverPark list: show it, paint it, and refuse the park.
# That added a third state, a tick guard and a second refusal dialog to protect
# ONE file -- and it argued against the paragraph directly above, which had the
# right answer already. A row whose only correct action is "leave it alone" does
# not need defending. It needs removing.
# Convert-LogHistory.ps1 MOVED to 'Diagnostic tools\' on 2026-09-20 and this
# window scans the root, so the entry no longer matches anything. It stays:
# the list costs nothing, and a copy left behind in the root -- which is what
# happens when someone upgrades over the top -- should still be excluded.
$script:ExcludedFiles = @(
    'Common.ps1', 'AdminConsole.ps1', 'Publish-Toolkit.ps1', 'ToolStatus.ps1',
    'Unblock.ps1', 'Convert-LogHistory.ps1', 'Settings.ps1'
)

# Two states, two colours. Lighter than the originals because the console theme
# is dark, and the old green was near-invisible against it.
$script:ColLoaded    = [System.Drawing.Color]::FromArgb(120, 215, 140)
$script:ColNotLoaded = [System.Drawing.Color]::FromArgb(232, 196, 80)

# Parked is a third STATE, not a third warning. Grey reads as "set aside" and
# stops the yellow count -- the thing you still have to act on -- from being
# padded by files already dealt with.
$script:ColParked    = [System.Drawing.Color]::FromArgb(150, 150, 150)

#region Scan --------------------------------------------------------------------
function Get-ListedScript {
    <#
        Every script Tools.json points at, plus PROOF that it was read.

        This used to return an empty hashtable on any failure -- file missing,
        JSON malformed -- and say nothing. With an empty set every script falls
        through to "not in Tools.json", so the tool reported confidently on a
        file it had never successfully opened, and scripts that plainly DO have
        buttons showed as unreferenced.

        Returns the error and the count so the window can show both. A scan
        that cannot read its input must say so, not answer anyway.

        Only Script is read. This used to read ScriptOnPrem too, because a
        paired tool named two files and reading one would have reported every
        AD twin as having no button. That feature was removed on 2026-08-13
        (ARCHITECTURE section 30), so one field is now the whole truth -- and
        an AD- script that no longer has a button will correctly show as not
        loaded, which is the answer this window exists to give.
    #>
    $path = Get-ConfigFilePath -Name 'Tools.json'
    $result = [PSCustomObject]@{
        Path = $path; Map = @{}; Entries = 0; Error = $null; Shape = ''; Names = 0
    }
    if (-not (Test-Path -LiteralPath $path)) {
        $result.Error = "Tools.json not found at '$path'."
        return $result
    }
    # Parsed defensively, and it REPORTS WHAT IT SAW.
    #
    # An earlier version read this file and announced "1 entry" for a 55-entry
    # array, while AdminConsole parsed the same file into 55 buttons with
    # equivalent code. Reading the source did not explain that, so the tool now
    # states the shape it received rather than assuming one. A scan that can
    # only say "1" leaves you guessing; one that says "1 of type PSCustomObject"
    # names the problem.
    #
    # Two shapes are accepted because two exist in the wild: a bare array, and
    # an object with a Tools property. Test-Toolkit already handles both.
    $parsed = $null
    try {
        $json   = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
        $parsed = ConvertFrom-Json $json -ErrorAction Stop
    } catch {
        $result.Error = "Tools.json could not be parsed: $($_.Exception.Message)"
        return $result
    }

    $tools = @()
    if ($null -ne $parsed -and $parsed.PSObject.Properties['Tools'] -and $parsed.Tools) {
        $tools = @($parsed.Tools)
        $result.Shape = 'object with a Tools property'
    } else {
        $tools = @($parsed)
        $result.Shape = "$(if ($parsed -is [array]) { 'array' } else { $parsed.GetType().Name })"
    }

    $map = @{}
    foreach ($t in $tools) {
        foreach ($k in @('Script')) {
            $v = "$($t.$k)".Trim()
            if (-not $v) { continue }
            # SPLIT ON BOTH SEPARATORS. No replacement string, so nothing to
            # escape wrongly.
            #
            # This line previously did -replace '/', '\' and the comment above
            # it claimed the problem was solved. It was not: in a -replace the
            # REPLACEMENT is regex-substitution syntax, where a lone backslash
            # starts an escape sequence. Entries written with forward slashes
            # -- './permissionsaudit.ps1', and Tools.json has several -- hit
            # that path and threw, so those tools could read as unreferenced.
            #
            # Same fault, same day, in Test-Toolkit's GRD001 rule. Prefer -split
            # with a character class anywhere a path separator is involved.
            $leaf = ($v -split '[\\/]')[-1]
            if ($leaf) { $map[$leaf.ToLower()] = "$($t.Name)" }
        }
    }
    $result.Map = $map
    $result.Entries = @($tools).Count
    $result.Names = $map.Keys.Count
    return $result
}

function Get-DashboardScript {
    <#
        Which scripts the dashboard can launch.

        Read out of Dashboard.ps1 itself rather than from a list kept here. A
        list kept here would be a second copy of the same fact, and a second
        copy is what has to be remembered -- which is exactly how this tool
        ended up with no launcher of its own for weeks.

        Comments are stripped, strings are NOT. Script names live in string
        literals ('ToolManager.ps1'), so a scanner that stripped strings would
        find nothing at all.
    #>
    $out = @{}
    $path = Join-Path $PSScriptRoot 'Dashboard.ps1'
    if (-not (Test-Path -LiteralPath $path)) { return $out }
    try { $text = Get-Content -LiteralPath $path -Raw -ErrorAction Stop } catch { return $out }

    $text = $text -replace '(?s)<#.*?#>', ''
    $text = $text -replace '(?m)#.*$', ''

    foreach ($m in [regex]::Matches($text, "(?i)'([A-Za-z0-9._-]+\.ps1)'")) {
        $out[$m.Groups[1].Value.ToLower()] = $true
    }
    return $out
}

function Get-PeerReferenceMap {
    <#
        Which scripts are opened BY ANOTHER SCRIPT, and by which one.

        Get-DashboardScript asks this of one file. This asks it of every file in
        the folder, because the toolkit grew a second kind of caller: a tool that
        routes to another tool. Unified-Lookup opens the Mailbox, Room, Group,
        Person and Password editors (the older destinations -- AD-Lookup,
        MailboxPermissions, Group-Manager-Editor, GroupEditor -- are parked).

        THOSE DESTINATIONS ARE NOT INDEPENDENT TOOLS. Nothing on the console
        opens Group-Manager-Editor -- it exists to be handed a group by something
        that already looked it up, and opened on its own it is a window asking a
        question the operator has already answered elsewhere. Listing it invites
        exactly one action, which is the one action that breaks its caller.

        RETURNS THE CALLERS, NOT JUST A FLAG. "Referenced" is a fact the operator
        cannot check; "called by Group-Manager.ps1" is one they can.

        COMMENTS ARE STRIPPED FIRST. This file is full of prose naming other
        scripts -- "see Group-Manager-Editor for the long version" -- and a
        documentation habit must not become a dependency.

        SELF-REFERENCES ARE IGNORED. A script naming itself, usually in a header
        or an error message, is not a caller.
    #>
    # THE LAUNCHERS ARE NOT PEERS, and leaving them in got this wrong.
    #
    # Dashboard.ps1 references every tool it puts a button on. Counting that as
    # "another script calls this" makes every Dashboard button un-parkable and
    # states one fact twice -- the row already says 'Dashboard' in its Where
    # column. Parking a button leaves a DEAD BUTTON: visible, annoying, and
    # undone in ten seconds. That deserves the warning it already has, not a
    # refusal.
    #
    # AdminConsole.ps1 is the same thing for the console's own grid.
    #
    # What is left after excluding them is the case that actually justifies a
    # refusal: one TOOL routing to another tool, where the breakage is invisible
    # until somebody takes that particular path.
    # AND NOR IS THIS FILE, which is the subtler one.
    #
    # ToolStatus NAMES scripts as data all over itself -- $ExcludedFiles and the
    # Dashboard path it parses. Those are inventory, not calls.
    #
    # THIS BIT IMMEDIATELY, in the obvious way. Adding 'Settings.ps1' to a
    # protected list inside this file made the scan see ToolStatus as a CALLER
    # of Settings.ps1 -- so the fix meant to keep it visible hid it instead, as
    # a helper of the tool doing the hiding.
    #
    # A scanner that treats its own inventory as evidence will always find
    # whatever it just wrote down. The list moved to $ExcludedFiles in the end,
    # but this exclusion stays: any filename this file records is data.
    $launchers = @('dashboard.ps1', 'adminconsole.ps1', 'toolstatus.ps1')

    $out = @{}
    $files = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File -ErrorAction SilentlyContinue)
    $known = @{}
    foreach ($f in $files) { $known[$f.Name.ToLower()] = $true }

    foreach ($f in $files) {
        if ($launchers -contains $f.Name.ToLower()) { continue }
        try { $text = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction Stop } catch { continue }
        $text = $text -replace '(?s)<#.*?#>', ''
        $text = $text -replace '(?m)#.*$', ''

        foreach ($m in [regex]::Matches($text, "(?i)'([A-Za-z0-9._-]+\.ps1)'")) {
            $tgt = $m.Groups[1].Value.ToLower()
            if ($tgt -eq $f.Name.ToLower()) { continue }
            if (-not $known.ContainsKey($tgt)) { continue }
            if (-not $out.ContainsKey($tgt)) { $out[$tgt] = New-Object System.Collections.ArrayList }
            if ($out[$tgt] -notcontains $f.Name) { [void]$out[$tgt].Add($f.Name) }
        }
    }
    return $out
}

function Get-LastRunMap {
    <#
        When each script last logged something.

        Read from the audit log rather than from a timestamp the tool would have
        to maintain: the log already records it, and a fact with two sources
        drifts. The Script column holds the leaf filename, so it joins directly.

        Every CompletedTasks*.csv is read, not just the current one -- the
        schema-versioned files are still history.
    #>
    $map = @{}
    try {
        $dir = Get-LogPath
        if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { return $map }
        $files = @(Get-ChildItem -LiteralPath $dir -Filter 'CompletedTasks*.csv' -File -ErrorAction SilentlyContinue)
        foreach ($lf in $files) {
            foreach ($row in (Import-Csv -LiteralPath $lf.FullName -ErrorAction SilentlyContinue)) {
                $s = "$($row.Script)".Trim().ToLower()
                if (-not $s) { continue }
                $when = $null
                try { $when = [datetimeoffset]::Parse("$($row.Timestamp)") } catch { continue }
                if (-not $map.ContainsKey($s) -or $when -gt $map[$s]) { $map[$s] = $when }
            }
        }
    } catch { }
    return $map
}

function Get-ToolStatusRow {
    <#
        One question per script: is anything loading it, and if so what.

        AdminConsole means Tools.json points a button at it. Dashboard means
        Dashboard.ps1 can launch it. Both is possible and is shown as both.
        Neither means nothing in the toolkit opens it.
    #>
    $listed = Get-ListedScript
    $dash   = Get-DashboardScript
    $peers  = Get-PeerReferenceMap
    $runs   = Get-LastRunMap

    $root = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
              Where-Object { $script:ExcludedFiles -notcontains $_.Name } |
              Sort-Object Name)

    $rows = @()
    foreach ($f in $root) {
        $key   = $f.Name.ToLower()
        $where = @()
        if ($listed.Map.ContainsKey($key)) { $where += 'AdminConsole' }
        if ($dash.ContainsKey($key))       { $where += 'Dashboard' }

        # CALLED BY ANOTHER TOOL -- a different fact from "on a button".
        $callers = @()
        if ($peers.ContainsKey($key)) { $callers = @($peers[$key]) }

        # A HELPER IS HIDDEN. A TOOL THAT IS ALSO A DESTINATION IS NOT.
        #
        # The distinction matters and the obvious rule gets it wrong. "Called by
        # another script, so hide it" would remove MailboxPermissions -- which
        # Unified-Lookup opens AND which is a console button in its own right --
        # along with LogViewer and Disable-User, both Dashboard buttons that
        # Dashboard naturally references.
        #
        # A HELPER IS A DESTINATION AND NOTHING ELSE: called by a peer, and on
        # no button anywhere. Group-Manager-Editor is the clean example --
        # nothing opens it but Group-Manager and Unified-Lookup, and opening it
        # cold asks a question the operator already answered.
        #
        # Same reasoning as $script:ExcludedFiles above: a row whose only
        # available action is the one action that breaks something is a row
        # worth removing.
        if ($callers.Count -gt 0 -and $where.Count -eq 0) { continue }

        $rows += [PSCustomObject]@{
            Name     = $f.Name
            Path     = $f.FullName
            Where    = $(if ($where.Count) { $where -join ', ' } else { 'Not loaded' })
            Loaded   = ($where.Count -gt 0)

            # WHO CALLS THIS, kept on the row so the park gate can name them
            # rather than saying "something does".
            Callers  = $callers

            LastRun  = $(if ($runs.ContainsKey($key)) { $runs[$key].ToString('yyyy-MM-dd HH:mm') } else { '--' })

            # Sorted on the real value, not the printed one. '--' would sort
            # between dates as text, and a never-run tool would land in the
            # middle of the list instead of at one end.
            LastRunSort  = $(if ($runs.ContainsKey($key)) { $runs[$key].UtcDateTime } else { [datetime]::MinValue })
            Modified     = $f.LastWriteTime.ToString('yyyy-MM-dd')
            ModifiedSort = $f.LastWriteTime
            Parked   = $false
        }
    }

    if (Test-Path -LiteralPath $script:DisabledDir) {
        foreach ($f in @(Get-ChildItem -LiteralPath $script:DisabledDir -Filter '*.ps1' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
            $key = $f.Name.ToLower()
            $rows += [PSCustomObject]@{
                Name     = $f.Name
                Path     = $f.FullName
                Where    = 'Parked'
                Loaded   = $false
                LastRun  = $(if ($runs.ContainsKey($key)) { $runs[$key].ToString('yyyy-MM-dd HH:mm') } else { '--' })
                LastRunSort  = $(if ($runs.ContainsKey($key)) { $runs[$key].UtcDateTime } else { [datetime]::MinValue })
                Modified     = $f.LastWriteTime.ToString('yyyy-MM-dd')
                ModifiedSort = $f.LastWriteTime
                Parked   = $true
            }
        }
    }
    return [PSCustomObject]@{ Rows = $rows; Listed = $listed }
}
#endregion

#region Form --------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Tool Status'
$form.Size = New-Object System.Drawing.Size(900, 668)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblTop = New-Object System.Windows.Forms.Label
$lblTop.Location = New-Object System.Drawing.Point(12, 10)
$lblTop.Size = New-Object System.Drawing.Size(860, 32)
$lblTop.Text = 'Green means something loads it, yellow means nothing does. Click ? (or double-click a row) for that tool''s own help before deciding to park it. Parking MOVES a file to DisabledTools and deletes nothing -- re-run Publish-Toolkit afterwards.'
$lblTop.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblTop)

$lv = New-Object System.Windows.Forms.ListView
$lv.Location = New-Object System.Drawing.Point(12, 46)
$lv.Size = New-Object System.Drawing.Size(860, 470)
$lv.View = 'Details'
$lv.FullRowSelect = $true
$lv.CheckBoxes = $true
$lv.GridLines = $true
[void]$lv.Columns.Add('Script', 280)
[void]$lv.Columns.Add('Loaded in', 200)
[void]$lv.Columns.Add('Last run', 160)
[void]$lv.Columns.Add('Modified', 120)
[void]$lv.Columns.Add('Help', 50)

# Index of the Help column, derived rather than written as 4. A column inserted
# above it would otherwise silently move the click target somewhere else.
$script:HelpCol = $lv.Columns.Count - 1
$form.Controls.Add($lv)

# Proof the scan read its input. Without this the window looks identical
# whether Tools.json was parsed or silently failed -- which is exactly how the
# first version reported every script as having no button.
$lblSrc = New-Object System.Windows.Forms.Label
$lblSrc.Location = New-Object System.Drawing.Point(12, 520)
$lblSrc.Size = New-Object System.Drawing.Size(860, 18)
$lblSrc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblSrc)

# 860 wide, so it spans the whole form. It has to sit clear of the button row
# below it: this label used to end at y=560 while the buttons began at y=548,
# and because it is added to the form FIRST it drew on top of all five of them
# -- WinForms z-order is creation order, earlier on top. The buttons appeared
# clipped and the top 12 pixels of each one swallowed the click.
$lblSum = New-Object System.Windows.Forms.Label
$lblSum.Location = New-Object System.Drawing.Point(12, 540)
$lblSum.Size = New-Object System.Drawing.Size(860, 20)
$lblSum.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblSum)

# DEFAULT: loaded first. The list exists to be worked through, and the loaded
# ones are the ones worth reading; the yellow tail is what you come back to.
$script:SortKey  = 'Loaded'
$script:SortDesc = $false
$script:ColBase  = @('Script', 'Loaded in', 'Last run', 'Modified', 'Help')

function Sort-ToolRow {
    <#
        Sorts the DATA and rebuilds, rather than installing a ListViewItemSorter.

        A comparer would mean a PowerShell class implementing IComparer and
        reaching into ListViewItem subitems -- comparing printed strings. Sorting
        the rows instead lets Last run compare as a date and Loaded compare as a
        boolean, which is what those columns actually mean.
    #>
    param($Rows)

    switch ($script:SortKey) {
        'Name'     { $s = @($Rows | Sort-Object Name) }
        'Where'    { $s = @($Rows | Sort-Object @{ Expression = { -not $_.Loaded } }, Where, Name) }
        'LastRun'  { $s = @($Rows | Sort-Object LastRunSort, Name) }
        'Modified' { $s = @($Rows | Sort-Object ModifiedSort, Name) }
        default    { $s = @($Rows | Sort-Object @{ Expression = { -not $_.Loaded } }, Name) }
    }
    if ($script:SortDesc) { [array]::Reverse($s) }
    return @($s)
}

function Update-ColumnHeader {
    <# The arrow goes on the sorted column and nowhere else. #>
    for ($i = 0; $i -lt $lv.Columns.Count; $i++) { $lv.Columns[$i].Text = $script:ColBase[$i] }
    $idx = switch ($script:SortKey) {
        'Name'     { 0 } 'Where' { 1 } 'LastRun' { 2 } 'Modified' { 3 } default { 1 }
    }
    $lv.Columns[$idx].Text = "$($script:ColBase[$idx]) $(if ($script:SortDesc) { [char]0x25BC } else { [char]0x25B2 })"
}

function Update-List {
    $lv.Items.Clear()
    $scan = Get-ToolStatusRow
    $rows = Sort-ToolRow -Rows @($scan.Rows)
    Update-ColumnHeader

    if ($scan.Listed.Error) {
        $lblSrc.Text = "Tools.json COULD NOT BE READ -- $($scan.Listed.Error)  Every script below will show no button; that is this error, not the truth."
        $lblSrc.ForeColor = [System.Drawing.Color]::Firebrick
    } else {
        $lblSrc.Text = "Read $($scan.Listed.Entries) entry/entries ($($scan.Listed.Shape)), $($scan.Listed.Names) script name(s), from $($scan.Listed.Path)"
        if ($scan.Listed.Names -eq 0) { $lblSrc.ForeColor = [System.Drawing.Color]::Firebrick }
        $lblSrc.ForeColor = [System.Drawing.Color]::DimGray
    }

    foreach ($r in $rows) {
        $it = New-Object System.Windows.Forms.ListViewItem($r.Name)
        [void]$it.SubItems.Add($r.Where)
        [void]$it.SubItems.Add($r.LastRun)
        [void]$it.SubItems.Add($r.Modified)
        [void]$it.SubItems.Add('?')
        $it.Tag = $r

        # Green loads, yellow does not, grey is already parked. Parked is
        # checked FIRST: a parked file is by definition not loaded, and would
        # otherwise be painted yellow as if it still needed a decision.
        $it.ForeColor = $(if ($r.Parked)    { $script:ColParked }
                          elseif ($r.Loaded) { $script:ColLoaded }
                          else               { $script:ColNotLoaded })
        [void]$lv.Items.Add($it)
    }
    $loaded = @($rows | Where-Object { $_.Loaded }).Count
    $parked = @($rows | Where-Object { $_.Parked }).Count
    $idle   = @($rows | Where-Object { -not $_.Loaded -and -not $_.Parked }).Count
    $lblSum.Text = "Loaded: $loaded   Not loaded: $idle   Parked: $parked"
}

# No tick guard: the files that must never be parked are excluded from the
# list rather than shown and defended.

function Show-RowHelp {
    <#
        The tool's own help, from the tool's own header -- the same text its ?
        button shows. Deciding whether to park something requires knowing what
        it does, and that answer already exists in the file.
    #>
    param($Row)
    if (-not $Row) { return }
    if (-not (Test-Path -LiteralPath $Row.Path)) {
        $null = Show-ToolDialog("$($Row.Name) is no longer at $($Row.Path).", 'Tool Status')
        return
    }
    try { Show-ToolHelp -Path $Row.Path -Title $Row.Name }
    catch {
        $null = Show-ToolDialog(
            "Could not read help from $($Row.Name):`r`n$($_.Exception.Message)", 'Tool Status')
    }
}

# HitTest rather than a column-width guess: it reports which subitem was struck,
# so the target stays correct if the columns are resized or reordered.
$lv.Add_MouseClick({
    param($s, $e)
    $hit = $lv.HitTest($e.Location)
    if (-not $hit.Item) { return }
    $idx = $hit.Item.SubItems.IndexOf($hit.SubItem)
    if ($idx -eq $script:HelpCol) { Show-RowHelp -Row $hit.Item.Tag }
})

# Double-click anywhere on the row does the same thing. The column is the
# discoverable route; this is the one people actually use.
$lv.Add_DoubleClick({
    if ($lv.SelectedItems.Count -gt 0) { Show-RowHelp -Row $lv.SelectedItems[0].Tag }
})

# Help is not sortable -- there is nothing to sort, every row says '?'.
$lv.Add_ColumnClick({
    param($s, $e)
    $key = switch ($e.Column) {
        0 { 'Name' } 1 { 'Where' } 2 { 'LastRun' } 3 { 'Modified' } default { $null }
    }
    if (-not $key) { return }
    if ($key -eq $script:SortKey) { $script:SortDesc = -not $script:SortDesc }
    else { $script:SortKey = $key; $script:SortDesc = $false }
    Update-List
})

function Get-Checked {
    $out = @()
    foreach ($it in $lv.Items) { if ($it.Checked -and $it.Tag) { $out += $it.Tag } }
    return $out
}

# Below lblSum's bottom edge (540 + 20), not overlapping it.
$y = 566
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = 'Refresh'
$btnRefresh.Location = New-Object System.Drawing.Point(12, $y)
$btnRefresh.Size = New-Object System.Drawing.Size(100, 30)
$btnRefresh.Add_Click({ Update-List })
$form.Controls.Add($btnRefresh)

$btnPark = New-Object System.Windows.Forms.Button
$btnPark.Text = 'Park selected'
$btnPark.Location = New-Object System.Drawing.Point(120, $y)
$btnPark.Size = New-Object System.Drawing.Size(130, 30)
$btnPark.Add_Click({
    $sel = @(Get-Checked | Where-Object { -not $_.Parked })
    if ($sel.Count -eq 0) { $null = Show-ToolDialog('Nothing ticked.', 'Tool Status'); return }

    # CALLED BY ANOTHER TOOL MEANS REFUSED, NOT WARNED.
    #
    # The warning below covers a file that is on a button: park it and you get a
    # dead button, which is visible, annoying and reversible in ten seconds.
    #
    # THIS IS A DIFFERENT FAILURE. Parking something another SCRIPT opens leaves
    # a caller that looks completely fine until the moment somebody uses the one
    # path that needs it -- Group-Manager finding a Microsoft 365 Group, say --
    # and then fails with a file-not-found from a tool nobody edited. Nothing
    # about the tool that broke points at the tool that was moved.
    #
    # A confirmation box does not help here, because the operator ticking the
    # box has no way to know that Unified-Lookup routes to this file. So the
    # answer is no, and the reason names the callers.
    $wired = @($sel | Where-Object { @($_.Callers).Count -gt 0 })
    if ($wired.Count -gt 0) {
        $lines = ($wired | ForEach-Object { "  $($_.Name)  --  called by $(@($_.Callers) -join ', ')" }) -join "`r`n"
        $null = Show-ToolDialog(
            ("These cannot be parked -- another tool opens them:`r`n`r`n$lines`r`n`r`n" +
             "Parking one leaves its caller looking healthy until somebody takes the path that needs it.`r`n`r`n" +
             "Remove the reference from the caller first, then park this."),
            'In use by another tool',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Referenced files are called out separately and by name. 'Are you sure?'
    # over a list of twenty is a question nobody can answer; 'these three are
    # called by something' is.
    $risky = @($sel | Where-Object { $_.Loaded })
    $msg = "Move $($sel.Count) file(s) to DisabledTools:`r`n`r`n" +
           (($sel | ForEach-Object { "  $($_.Name)  [$($_.Where)]" }) -join "`r`n")
    if ($risky.Count -gt 0) {
        $msg += "`r`n`r`nWARNING -- $($risky.Count) of these are still loaded and will leave a dead button:`r`n" +
                (($risky | ForEach-Object { "  $($_.Name) -- loaded in $($_.Where)" }) -join "`r`n")
    }
    $msg += "`r`n`r`nNothing is deleted. Restore puts them back.`r`n`r`nContinue?"
    $ans = Show-ToolDialog($msg, 'Park tools',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        $(if ($risky.Count) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Question }),
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    if (-not (Test-Path -LiteralPath $script:DisabledDir)) {
        New-Item -Path $script:DisabledDir -ItemType Directory -Force | Out-Null
    }
    $done = 0; $failed = @()
    foreach ($r in $sel) {
        $src = Join-Path $PSScriptRoot $r.Name
        $dst = Join-Path $script:DisabledDir $r.Name
        try {
            # An existing file in DisabledTools is an OLDER parked copy. Keeping
            # it under a timestamp beats overwriting: the whole point of this
            # folder is that nothing is destroyed.
            if (Test-Path -LiteralPath $dst) {
                $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
                $dst = Join-Path $script:DisabledDir "$([System.IO.Path]::GetFileNameWithoutExtension($r.Name)).$stamp.ps1"
            }
            Move-Item -LiteralPath $src -Destination $dst -ErrorAction Stop
            $done++
        } catch { $failed += "$($r.Name): $($_.Exception.Message)" }
    }
    Update-List
    $out = "Parked $done file(s)."
    if ($failed.Count) { $out += "`r`n`r`nFailed:`r`n" + ($failed -join "`r`n") }
    $out += "`r`n`r`nRe-run Publish-Toolkit so the version manifest matches."
    $null = Show-ToolDialog($out, 'Park tools')
    try { Write-CompletedLog -Action 'ToolStatus' -Target 'Park' -Result 'Success' -Detail "$done parked" } catch { }
})
$form.Controls.Add($btnPark)

$btnRestore = New-Object System.Windows.Forms.Button
$btnRestore.Text = 'Restore selected'
$btnRestore.Location = New-Object System.Drawing.Point(258, $y)
$btnRestore.Size = New-Object System.Drawing.Size(130, 30)
$btnRestore.Add_Click({
    $sel = @(Get-Checked | Where-Object { $_.Parked })
    if ($sel.Count -eq 0) { $null = Show-ToolDialog('Tick something in DisabledTools first.', 'Tool Status'); return }
    $done = 0; $failed = @()
    foreach ($r in $sel) {
        $src = Join-Path $script:DisabledDir $r.Name
        $dst = Join-Path $PSScriptRoot $r.Name
        if (Test-Path -LiteralPath $dst) { $failed += "$($r.Name): a file of that name is already in the console folder."; continue }
        try { Move-Item -LiteralPath $src -Destination $dst -ErrorAction Stop; $done++ }
        catch { $failed += "$($r.Name): $($_.Exception.Message)" }
    }
    Update-List
    $out = "Restored $done file(s)."
    if ($failed.Count) { $out += "`r`n`r`nSkipped:`r`n" + ($failed -join "`r`n") }
    if ($done -gt 0) { $out += "`r`n`r`nAdd a Tools.json entry in ToolManager if it needs a button." }
    $null = Show-ToolDialog($out, 'Restore tools')
})
$form.Controls.Add($btnRestore)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = 'Open DisabledTools'
$btnOpen.Location = New-Object System.Drawing.Point(396, $y)
$btnOpen.Size = New-Object System.Drawing.Size(150, 30)
$btnOpen.Add_Click({
    if (-not (Test-Path -LiteralPath $script:DisabledDir)) {
        New-Item -Path $script:DisabledDir -ItemType Directory -Force | Out-Null
    }
    Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$script:DisabledDir`""
})
$form.Controls.Add($btnOpen)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(772, $y)
$btnClose.Size = New-Object System.Drawing.Size(100, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

Update-List
try { Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath | Out-Null } catch { }
try { Set-ConsoleTheme -Form $form } catch { }
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
[void]$form.ShowDialog()
#endregion
