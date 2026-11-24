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
# Lists every .ps1 beside the console and says whether anything uses it.
#
#   In Tools.json  -- a button in the console points at it.
#   Core           -- the console cannot run without it. Never movable.
#   Referenced     -- no button, but another script calls it. Moving it breaks
#                     that script, so the list names which one.
#   Unreferenced   -- nothing points at it. Safe to park.
#   Parked         -- already moved to DisabledTools.
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
$script:CoreFiles = @('Common.ps1', 'AdminConsole.ps1', 'ToolStatus.ps1')

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

        ScriptOnPrem is read as well as Script: a paired tool names two files
        and only one is in Script, so reading only Script would report every AD
        twin as having no button.
    #>
    $path = Get-ConfigFilePath -Name 'Tools.json'
    $result = [PSCustomObject]@{
        Path = $path; Map = @{}; Entries = 0; Error = $null
    }
    if (-not (Test-Path -LiteralPath $path)) {
        $result.Error = "Tools.json not found at '$path'."
        return $result
    }
    try {
        $tools = @(Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        $result.Error = "Tools.json could not be parsed: $($_.Exception.Message)"
        return $result
    }
    $map = @{}
    foreach ($t in $tools) {
        foreach ($k in @('Script', 'ScriptOnPrem')) {
            $v = "$($t.$k)".Trim()
            if (-not $v) { continue }
            $leaf = Split-Path -Leaf ($v -replace '/', '\\')
            if ($leaf) { $map[$leaf.ToLower()] = "$($t.Name)" }
        }
    }
    $result.Map = $map
    $result.Entries = @($tools).Count
    return $result
}

function Get-ScriptReference {
    <#
        Which other scripts mention this one by filename.

        Text search, deliberately: these are launched by path at run time --
        '& (Join-Path $PSScriptRoot ''X.ps1'')' -- so there is no call graph to
        follow, only the name appearing in the caller. It over-reports, and
        over-reporting is the right error: it warns about a move that might be
        safe rather than blessing one that is not.
    #>
    param([string[]]$AllFiles)
    $refs = @{}; $text = @{}
    foreach ($f in $AllFiles) {
        try { $text[$f] = Get-Content -LiteralPath (Join-Path $PSScriptRoot $f) -Raw -ErrorAction Stop }
        catch { $text[$f] = '' }
    }
    foreach ($target in $AllFiles) {
        $hits = @()
        foreach ($src in $AllFiles) {
            if ($src -ieq $target) { continue }
            if ($text[$src] -and $text[$src].IndexOf($target, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $hits += $src
            }
        }
        $refs[$target] = $hits
    }
    return $refs
}

function Get-ToolStatusRow {
    <#
        TWO INDEPENDENT FACTS, REPORTED SEPARATELY.

        The first version collapsed "has a button" and "is called by another
        script" into one Status column, with the button winning. That hid half
        the answer: a script with a button AND callers looked like it had no
        callers, and a script with callers but no button was labelled
        "Referenced", which reads as "in use" when the question asked was
        whether it is in Tools.json.

        Button and CalledBy are now separate columns and Verdict is derived
        from both, so the columns can be checked against each other.
    #>
    $listed = Get-ListedScript
    $root   = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
                Sort-Object Name)
    $names  = @($root | ForEach-Object { $_.Name })
    $refs   = Get-ScriptReference -AllFiles $names

    $rows = @()
    foreach ($f in $root) {
        $key    = $f.Name.ToLower()
        $button = if ($listed.Map.ContainsKey($key)) { $listed.Map[$key] } else { '' }
        $callers = @($refs[$f.Name])
        $isCore  = $script:CoreFiles -contains $f.Name

        $verdict =
            if     ($isCore)            { 'Core' }
            elseif ($button)            { 'In use' }
            elseif ($callers.Count -gt 0) { 'Helper only' }
            else                        { 'UNUSED' }

        $calledBy = if ($callers.Count -eq 0) { '' }
                    elseif ($callers.Count -le 3) { $callers -join ', ' }
                    else { "$(($callers | Select-Object -First 3) -join ', ') (+$($callers.Count - 3))" }

        $rows += [PSCustomObject]@{
            Name = $f.Name; Button = $button; CalledBy = $calledBy
            Verdict = $verdict; CallerCount = $callers.Count
            Modified = $f.LastWriteTime.ToString('yyyy-MM-dd')
            Parked = $false
        }
    }

    if (Test-Path -LiteralPath $script:DisabledDir) {
        foreach ($f in @(Get-ChildItem -LiteralPath $script:DisabledDir -Filter '*.ps1' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
            $rows += [PSCustomObject]@{
                Name = $f.Name; Button = ''; CalledBy = 'in DisabledTools'
                Verdict = 'Parked'; CallerCount = 0
                Modified = $f.LastWriteTime.ToString('yyyy-MM-dd')
                Parked = $true
            }
        }
    }
    return [PSCustomObject]@{ Rows = $rows; Listed = $listed }
}
#endregion

#region Form --------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Tool Status'
$form.Size = New-Object System.Drawing.Size(900, 640)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblTop = New-Object System.Windows.Forms.Label
$lblTop.Location = New-Object System.Drawing.Point(12, 10)
$lblTop.Size = New-Object System.Drawing.Size(860, 32)
$lblTop.Text = 'Parking MOVES a file to DisabledTools and deletes nothing. Core files cannot be parked. Re-run Publish-Toolkit afterwards -- the version manifest lists files by name.'
$lblTop.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblTop)

$lv = New-Object System.Windows.Forms.ListView
$lv.Location = New-Object System.Drawing.Point(12, 46)
$lv.Size = New-Object System.Drawing.Size(860, 470)
$lv.View = 'Details'
$lv.FullRowSelect = $true
$lv.CheckBoxes = $true
$lv.GridLines = $true
[void]$lv.Columns.Add('Script', 230)
[void]$lv.Columns.Add('Button in Tools.json', 190)
[void]$lv.Columns.Add('Called by', 290)
[void]$lv.Columns.Add('Verdict', 90)
[void]$lv.Columns.Add('Modified', 80)
$form.Controls.Add($lv)

# Proof the scan read its input. Without this the window looks identical
# whether Tools.json was parsed or silently failed -- which is exactly how the
# first version reported every script as having no button.
$lblSrc = New-Object System.Windows.Forms.Label
$lblSrc.Location = New-Object System.Drawing.Point(12, 520)
$lblSrc.Size = New-Object System.Drawing.Size(860, 18)
$lblSrc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblSrc)

$lblSum = New-Object System.Windows.Forms.Label
$lblSum.Location = New-Object System.Drawing.Point(12, 540)
$lblSum.Size = New-Object System.Drawing.Size(860, 20)
$lblSum.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblSum)

function Update-List {
    $lv.Items.Clear()
    $scan = Get-ToolStatusRow
    $rows = @($scan.Rows)

    if ($scan.Listed.Error) {
        $lblSrc.Text = "Tools.json COULD NOT BE READ -- $($scan.Listed.Error)  Every script below will show no button; that is this error, not the truth."
        $lblSrc.ForeColor = [System.Drawing.Color]::Firebrick
    } else {
        $lblSrc.Text = "Read $($scan.Listed.Entries) entry/entries from $($scan.Listed.Path)"
        $lblSrc.ForeColor = [System.Drawing.Color]::DimGray
    }

    foreach ($r in $rows) {
        $it = New-Object System.Windows.Forms.ListViewItem($r.Name)
        [void]$it.SubItems.Add($(if ($r.Button) { $r.Button } else { '--' }))
        [void]$it.SubItems.Add($(if ($r.CalledBy) { $r.CalledBy } else { '--' }))
        [void]$it.SubItems.Add($r.Verdict)
        [void]$it.SubItems.Add($r.Modified)
        $it.Tag = $r
        switch ($r.Verdict) {
            'Core'        { $it.ForeColor = [System.Drawing.Color]::DimGray }
            'In use'      { $it.ForeColor = [System.Drawing.Color]::FromArgb(0, 110, 60) }
            'Helper only' { $it.ForeColor = [System.Drawing.Color]::FromArgb(160, 90, 0) }
            'UNUSED'      { $it.ForeColor = [System.Drawing.Color]::FromArgb(160, 0, 0) }
            'Parked'      { $it.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 160) }
        }
        [void]$lv.Items.Add($it)
    }
    $c = @{}
    foreach ($r in $rows) { $c[$r.Verdict] = [int]$c[$r.Verdict] + 1 }
    $lblSum.Text = "In use: $([int]$c['In use'])   Core: $([int]$c['Core'])   Helper only: $([int]$c['Helper only'])   UNUSED: $([int]$c['UNUSED'])   Parked: $([int]$c['Parked'])"
}

# A Core row cannot be ticked. Refusing the tick is quieter than letting it be
# ticked and complaining at the point of action, and it makes the rule visible
# before anyone commits to it.
$lv.Add_ItemCheck({
    param($s, $e)
    $it = $lv.Items[$e.Index]
    if ($it.Tag -and $it.Tag.Verdict -eq 'Core' -and $e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
    }
})

function Get-Checked {
    $out = @()
    foreach ($it in $lv.Items) { if ($it.Checked -and $it.Tag) { $out += $it.Tag } }
    return $out
}

$y = 548
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
    if ($sel.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('Nothing ticked.', 'Tool Status'); return }

    # Referenced files are called out separately and by name. 'Are you sure?'
    # over a list of twenty is a question nobody can answer; 'these three are
    # called by something' is.
    $risky = @($sel | Where-Object { $_.CallerCount -gt 0 })
    $msg = "Move $($sel.Count) file(s) to DisabledTools:`r`n`r`n" +
           (($sel | ForEach-Object { "  $($_.Name)  [$($_.Verdict)]" }) -join "`r`n")
    if ($risky.Count -gt 0) {
        $msg += "`r`n`r`nWARNING -- $($risky.Count) of these are called by another script and will break it:`r`n" +
                (($risky | ForEach-Object { "  $($_.Name) -- called by $($_.CalledBy)" }) -join "`r`n")
    }
    $msg += "`r`n`r`nNothing is deleted. Restore puts them back.`r`n`r`nContinue?"
    $ans = [System.Windows.Forms.MessageBox]::Show($msg, 'Park tools',
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
    [void][System.Windows.Forms.MessageBox]::Show($out, 'Park tools')
    try { Write-CompletedLog -Action 'ToolStatus' -Target 'Park' -Result 'Success' -Detail "$done parked" } catch { }
})
$form.Controls.Add($btnPark)

$btnRestore = New-Object System.Windows.Forms.Button
$btnRestore.Text = 'Restore selected'
$btnRestore.Location = New-Object System.Drawing.Point(258, $y)
$btnRestore.Size = New-Object System.Drawing.Size(130, 30)
$btnRestore.Add_Click({
    $sel = @(Get-Checked | Where-Object { $_.Parked })
    if ($sel.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('Tick something in DisabledTools first.', 'Tool Status'); return }
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
    [void][System.Windows.Forms.MessageBox]::Show($out, 'Restore tools')
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
[void]$form.ShowDialog()
#endregion
