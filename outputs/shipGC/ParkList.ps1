#Requires -Version 5.1
# TOOLKIT-VARIANT: Any
<#
    ParkList.ps1
    Park the scripts a build has made redundant, from a curated list that
    ships with the build.

    Author   : Glenn Armbruster
    Created  : 2026-09-21
#>
# TOOLKIT-HELP-RUNS: nothing against any tenant. It reads Config\ParkList.json and writes the Parked block in your own profile, or in the tenant file if you choose that scope.
# TOOLKIT-HELP-BEGIN
# Parks the tools a new build has absorbed, in one pass, from a list that came
# with the build.
#
# WHY THIS EXISTS. When the Mailbox, Room, Group and Person editors took over
# a tool's job, the old tool keeps its button until somebody hides it -- and
# working out WHICH ones are now redundant is a job for whoever built them,
# not for each engineer to rediscover. Config\ParkList.json is that answer,
# written once and shipped.
#
# IT RECOMMENDS; YOU DECIDE. Every entry is listed with the reason it is
# redundant. Entries the build is sure about open TICKED. Entries that are a
# judgement call open UNTICKED and say so. Untick anything you want to keep --
# nothing is parked until you press Park.
#
# IT DOES NOT ASK WHETHER ANOTHER SCRIPT CALLS THE FILE. Tool Status refuses a
# park when something else references the script, which is the right caution
# for a one-off decision and the wrong one here: these tools were superseded
# deliberately, the references are comments and dead hand-off rows, and being
# stopped 25 times in a row teaches you to ignore the warning. Parking hides a
# BUTTON. It does not delete, move or rename anything, so nothing that runs
# can break -- the file is still there and still runs if something calls it.
#
# PARKING IS A LIST, NOT A SWITCH, AND THIS ADDS TO IT. Whatever you have
# already parked stays parked: the existing entries are read, the new ones are
# added, and the whole list is written back. That matters because both stores
# are replaced wholesale on save -- writing only the new names would silently
# un-park everything else.
#
# SCOPE. 'Just me' writes your own profile and affects nobody else; it is the
# default. 'This tenant' writes the tenant file and hides the tools for every
# engineer on it -- use that one when the tenant genuinely does not have the
# capability any more.
#
# IT CLEARS ITSELF, PER PERSON. What you action is recorded in your own
# profile, so a second run shows nothing to do. When a later build tops the
# file up, you are offered only the entries you have not seen. Nothing is
# deleted from the shipped file -- your colleagues have not run it yet.
#
# A SCRIPT THAT IS NOT INSTALLED IS NOT AN ERROR. The same list ships to every
# tenant and not every tenant has every tool. Those rows say 'not installed'
# and are skipped.
#
# Close and reopen the console to see the buttons change.
# TOOLKIT-HELP-END

param(
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:ParkListFile = Join-Path (Join-Path $PSScriptRoot 'Config') 'ParkList.json'

#region Model --------------------------------------------------------------------
function Get-ParkListEntry {
    <#
        The shipped list, with each entry's current state worked out.

        State is one of:
          Pending        not parked, the file is here, and you have not
                         actioned it -- the only state that does anything
          Already parked hidden already, by you or by the tenant
          Done           you actioned it on an earlier run; left alone so a
                         deliberate un-park is not undone on the next run
          Not installed  the list ships to every tenant; not every tenant has
                         every tool
    #>
    if (-not (Test-Path -LiteralPath $script:ParkListFile)) { return @() }
    $j = $null
    try { $j = Get-Content -LiteralPath $script:ParkListFile -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop }
    catch { throw "Config\ParkList.json is not valid JSON: $($_.Exception.Message)" }
    if (-not $j -or -not $j.PSObject.Properties['Scripts']) { return @() }

    $applied = Get-ParkListApplied
    $mine    = Get-UserParked
    $tenant  = @{}
    foreach ($e in @(Get-ParkedScript)) {
        $k = Get-ParkedKey -Script $(if ($e.PSObject.Properties['Script']) { "$($e.Script)" } else { "$e" })
        if ($k) { $tenant[$k] = $true }
    }

    $rows = @()
    foreach ($x in @($j.Scripts)) {
        $name = "$($x.Script)".Trim()
        if (-not $name) { continue }
        $key = Get-ParkedKey -Script $name
        $onDisk = Test-Path -LiteralPath (Join-Path $PSScriptRoot (Split-Path -Leaf $name))

        $state = 'Pending'
        if (-not $onDisk)                                   { $state = 'Not installed' }
        elseif ($mine.Scripts.ContainsKey($key) -or $tenant.ContainsKey($key)) { $state = 'Already parked' }
        elseif ($applied -contains $key)                    { $state = 'Done' }

        # RECOMMEND DECIDES THE TICK, STATE DECIDES WHETHER IT CAN BE TICKED.
        # A missing Recommend property counts as recommended: a list written
        # before this field existed should still do the obvious thing.
        $rec = $true
        if ($x.PSObject.Properties['Recommend']) { $rec = [bool]$x.Recommend }

        $rows += [PSCustomObject]@{
            Script    = $name
            Reason    = "$($x.Reason)"
            Recommend = $rec
            State     = $state
            Key       = $key
            Actionable = ($state -eq 'Pending')
        }
    }
    return @($rows)
}

function Get-ParkListApplied {
    # The keys this person has already actioned, from their own profile.
    $l = Get-UserLayout
    if ($l -and $l.PSObject.Properties['ParkListApplied']) {
        return @(@($l.ParkListApplied) | ForEach-Object { "$_".Trim() } | Where-Object { $_ })
    }
    return @()
}

function Save-ParkListApplied {
    <#
        Add keys to this person's actioned list.

        READ-MODIFY-WRITE THROUGH Get-UserLayoutOwn, not Get-UserLayout. The
        merged view would drag the Default block into the tenant block and
        shadow it from then on -- the trap Get-UserLayoutOwn's own comment
        exists to describe.
    #>
    param([string[]]$Keys)
    $layout = Get-UserLayoutOwn
    if (-not $layout) { $layout = [PSCustomObject]@{} }
    $have = @()
    if ($layout.PSObject.Properties['ParkListApplied']) { $have = @($layout.ParkListApplied) }
    $all = @(@($have) + @($Keys) | ForEach-Object { "$_".Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    Add-Member -InputObject $layout -NotePropertyName 'ParkListApplied' -NotePropertyValue $all -Force
    return (Save-UserLayout -Layout $layout)
}
#endregion

#region Park ---------------------------------------------------------------------
function Invoke-ParkListApply {
    <#
        Park the chosen scripts, ADDED to whatever is parked already.

        THE WHOLE LIST GOES BACK EVERY TIME. Save-ParkedScript and
        Save-UserLayout both REPLACE the Parked block -- they take what the
        list should be, not a delta. Sending only the new names would un-park
        everything the person or the tenant had parked before, which is a data
        loss that looks like the tool working.

        SECTIONS ARE CARRIED THROUGH UNTOUCHED for the same reason: they live
        in the same block, and omitting them would un-park every parked
        section.
    #>
    param([Parameter(Mandatory)]$Rows, [bool]$ForMe)

    $entries = New-Object System.Collections.ArrayList
    $seen    = @{}
    $secs    = @()

    if ($ForMe) {
        $mine = Get-UserParked
        foreach ($k in $mine.Scripts.Keys) { [void]$entries.Add($mine.Scripts[$k]); $seen[$k] = $true }
        $secs = @($mine.Sections)
    } else {
        foreach ($e in @(Get-ParkedScript)) {
            $k = Get-ParkedKey -Script $(if ($e.PSObject.Properties['Script']) { "$($e.Script)" } else { "$e" })
            if ($k -and -not $seen.ContainsKey($k)) { [void]$entries.Add($e); $seen[$k] = $true }
        }
        $secs = @(Get-ParkedSection)
    }

    $added = @()
    $when  = (Get-Date).ToString('yyyy-MM-dd')
    foreach ($r in $Rows) {
        if ($seen.ContainsKey($r.Key)) { continue }
        [void]$entries.Add([PSCustomObject]@{
            Script = ".\$(Split-Path -Leaf $r.Script)"
            Reason = "$($r.Reason)"
            When   = $when
        })
        $seen[$r.Key] = $true
        $added += $r.Key
    }

    $path = $(if ($ForMe) {
        $layout = Get-UserLayoutOwn
        if (-not $layout) { $layout = [PSCustomObject]@{} }
        Add-Member -InputObject $layout -NotePropertyName 'Parked' -NotePropertyValue ([PSCustomObject][ordered]@{
            Scripts  = @($entries.ToArray())
            Sections = @($secs)
        }) -Force
        Save-UserLayout -Layout $layout
    } else {
        Save-ParkedScript -Entries @($entries.ToArray()) -Sections @($secs)
    })

    return [PSCustomObject]@{ Path = $path; Added = @($added); Total = $entries.Count }
}
#endregion

#region Form ---------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Park Scripts'
$form.ClientSize    = New-Object System.Drawing.Size(940, 600)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWhy = New-Object System.Windows.Forms.Label
$lblWhy.Location = New-Object System.Drawing.Point(12, 10)
$lblWhy.Size     = New-Object System.Drawing.Size(916, 32)
$lblWhy.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWhy.Text     = 'These tools have been absorbed by the editors. Ticked entries are the ones this build is sure about; untick anything you want to keep. Parking hides a button -- it does not delete, move or rename the file.'
$form.Controls.Add($lblWhy)

$grpScope = New-Object System.Windows.Forms.GroupBox
$grpScope.Text     = 'Park for'
$grpScope.Location = New-Object System.Drawing.Point(12, 46)
$grpScope.Size     = New-Object System.Drawing.Size(460, 50)
$form.Controls.Add($grpScope)
$rbMe = New-Object System.Windows.Forms.RadioButton
$rbMe.Text     = "Just me ($(Get-UserProfileKey))"
$rbMe.Location = New-Object System.Drawing.Point(12, 20)
$rbMe.Size     = New-Object System.Drawing.Size(240, 22)
$rbMe.Checked  = $true
$grpScope.Controls.Add($rbMe)
$rbTenant = New-Object System.Windows.Forms.RadioButton
$rbTenant.Text     = 'This tenant (everyone)'
$rbTenant.Location = New-Object System.Drawing.Point(262, 20)
$rbTenant.Size     = New-Object System.Drawing.Size(190, 22)
$grpScope.Controls.Add($rbTenant)

$lblScopeNote = New-Object System.Windows.Forms.Label
$lblScopeNote.Location  = New-Object System.Drawing.Point(486, 62)
$lblScopeNote.Size      = New-Object System.Drawing.Size(442, 32)
$lblScopeNote.Text      = "'Just me' affects nobody else. 'This tenant' hides them for every engineer here."
$form.Controls.Add($lblScopeNote)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(12, 104)
$grid.Size                = New-Object System.Drawing.Size(916, 390)
$grid.AllowUserToAddRows  = $false
$grid.SelectionMode       = 'FullRowSelect'
$grid.MultiSelect         = $false
$grid.RowHeadersVisible   = $false
$grid.AutoSizeColumnsMode = 'None'
$grid.Font                = New-Object System.Drawing.Font('Consolas', 9)
foreach ($c in @(@('Script', 210), @('State', 110), @('Reason', 556))) {
    $i = $grid.Columns.Add($c[0], $c[0]); $grid.Columns[$i].Width = $c[1]
}
$form.Controls.Add($grid)
Add-GridPickColumn -Grid $grid

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, 500)
$lblStatus.Size      = New-Object System.Drawing.Size(916, 34)
$lblStatus.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblStatus)

$btnAll  = New-Object System.Windows.Forms.Button
$btnAll.Text     = 'Tick all'
$btnAll.Location = New-Object System.Drawing.Point(12, 546)
$btnAll.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnAll)
$btnNone = New-Object System.Windows.Forms.Button
$btnNone.Text     = 'Untick all'
$btnNone.Location = New-Object System.Drawing.Point(110, 546)
$btnNone.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnNone)
$btnPark = New-Object System.Windows.Forms.Button
$btnPark.Text     = 'Park ticked'
$btnPark.Location = New-Object System.Drawing.Point(216, 546)
$btnPark.Size     = New-Object System.Drawing.Size(140, 28)
$btnPark.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnPark)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(828, 546)
$btnClose.Size     = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnClose)
#endregion

#region Wiring -------------------------------------------------------------------
function Update-ParkGrid {
    $grid.Rows.Clear()
    $rows = @()
    try { $rows = @(Get-ParkListEntry) } catch {
        $lblStatus.Text = "$($_.Exception.Message)"
        $lblStatus.ForeColor = (Get-ToolStatusColour -Kind Bad)
        return
    }
    if ($rows.Count -eq 0) {
        $lblStatus.Text = "No list found. Expected Config\ParkList.json beside the console."
        $lblStatus.ForeColor = (Get-ToolStatusColour -Kind Warn)
        $btnPark.Enabled = $false
        return
    }
    foreach ($r in $rows) {
        $i = $grid.Rows.Add($false, $r.Script, $r.State, $r.Reason)
        $grid.Rows[$i].Tag = $r
        # ONLY A PENDING ROW CAN BE TICKED. The others are shown so the window
        # accounts for every line in the file -- silence about a script you
        # expected to see is the thing that sends someone looking for a bug.
        if ($r.Actionable) {
            $grid.Rows[$i].Cells['Pick'].Value = $r.Recommend
        } else {
            $grid.Rows[$i].Cells['Pick'].ReadOnly = $true
            $grid.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Muted)
        }
    }
    $pending = @($rows | Where-Object { $_.Actionable }).Count
    $ticked  = @($rows | Where-Object { $_.Actionable -and $_.Recommend }).Count
    $btnPark.Enabled = ($pending -gt 0)
    $lblStatus.ForeColor = (Get-ToolStatusColour -Kind Muted)
    $lblStatus.Text = $(if ($pending -eq 0) {
        "Nothing to do -- every entry in the list is already parked, already actioned by you, or not installed here."
    } else {
        "$pending pending, $ticked ticked. Untick anything you want to keep, then press Park ticked."
    })
}

$btnAll.Add_Click({
    foreach ($row in $grid.Rows) { if ($row.Tag -and $row.Tag.Actionable) { $row.Cells['Pick'].Value = $true } }
})
$btnNone.Add_Click({
    foreach ($row in $grid.Rows) { if ($row.Tag -and $row.Tag.Actionable) { $row.Cells['Pick'].Value = $false } }
})
$btnClose.Add_Click({ $form.Close() })

$btnPark.Add_Click({
    $picked = @(Get-GridPicked -Grid $grid | Where-Object { $_.Actionable })
    if ($picked.Count -eq 0) { $lblStatus.Text = 'Nothing ticked.'; return }
    $forMe = [bool]$rbMe.Checked
    $who   = $(if ($forMe) { "for $(Get-UserProfileKey) only" } else { 'for EVERY engineer on this tenant' })

    $shown = @($picked | Select-Object -First 12 | ForEach-Object { "  $($_.Script)" }) -join "`r`n"
    if ($picked.Count -gt 12) { $shown += "`r`n  ... and $($picked.Count - 12) more" }
    $a = Show-ToolDialog(
        ("Park $($picked.Count) script(s) $who`?" + "`r`n`r`n" + $shown + "`r`n`r`n" +
         'Parking hides the button. The files stay where they are and anything that calls them still works.' + "`r`n" +
         'Close and reopen the console to see the change.'),
        'Park scripts', 'OKCancel', 'Question')
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        $res = Invoke-ParkListApply -Rows $picked -ForMe $forMe
        # THE APPLIED MARKER IS THIS PERSON'S EITHER WAY. A tenant park already
        # hides the tool from them, and recording it stops the row coming back
        # as pending if the tenant park is later lifted for some other reason.
        [void](Save-ParkListApplied -Keys @($res.Added))
        Write-CompletedLog -Action 'ParkList' -Target $res.Path -Result 'Success' `
            -Detail "$($res.Added.Count) script(s) parked $who; $($res.Total) now in the list." -Environment $Environment
        Update-ParkGrid
        $lblStatus.ForeColor = (Get-ToolStatusColour -Kind Good)
        $lblStatus.Text = "Parked $($res.Added.Count) script(s) $who. $($res.Total) script(s) now in the parked list. Reopen the console to see the buttons change."
        $null = Show-ToolDialog(
            ("$($res.Added.Count) script(s) parked $who." + "`r`n`r`n" + "Written to:" + "`r`n  $($res.Path)" + "`r`n`r`n" +
             'The console reads the parked list at startup -- close and reopen it.'),
            'Parked', 'OK', 'Information')
    } catch {
        $lblStatus.ForeColor = (Get-ToolStatusColour -Kind Bad)
        $lblStatus.Text = 'Park failed -- see the error dialog.'
        Show-HandlerError -ErrorRecord $_ -Context 'Parking scripts'
        Write-FailedLog -Action 'ParkList' -Target 'Parked list' -ErrorRecord $_ -Environment $Environment
    }
})

# THE GRID IS REBUILT WHEN THE SCOPE CHANGES. 'Already parked' means something
# different for the two scopes, and a stale State column would tick rows that
# the chosen destination has parked already.
$rbMe.Add_CheckedChanged({ Update-ParkGrid })

Update-ParkGrid
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Park Scripts' | Out-Null
[void]$form.ShowDialog()
#endregion
