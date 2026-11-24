#Requires -Version 5.1
<#
    ConfigBackup.ps1
    Backs up and restores Config\*.json -- the per-deployment files that a
    folder copy destroys.

    Reasoning: ScriptNotes/ConfigBackup.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-08
#>
# TOOLKIT-HELP-BEGIN
# Copies Config\*.json somewhere safe, and puts them back.
#
# These are the files that are YOURS rather than the toolkit's: Tools.json,
# Settings.json, Sections.json, QuickLaunch.json, Themes.json, Tenant.json.
# Copying a new console folder over an old one replaces every one of them, and
# nothing warns you -- the loss shows up later as blank paths and missing
# buttons.
#
# RUN A BACKUP BEFORE COPYING A NEW CONSOLE OVER THIS ONE. That is the whole
# reason this exists.
#
# Backups are written OUTSIDE the console folder, beside it, so that replacing
# or renaming the console cannot take the backup with it. Each backup is a
# timestamped folder holding the files exactly as they were.
#
# Restore copies a backup back over Config\ -- and takes a fresh backup of what
# is there first, so restoring the wrong snapshot is not final either.
# TOOLKIT-HELP-END

param([string]$BackupRoot)

. (Join-Path $PSScriptRoot 'Common.ps1')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Locations ---------------------------------------------------------------
$script:ConfigDir = Join-Path $PSScriptRoot 'Config'

# OUTSIDE the console folder, deliberately, and this is the point of the tool.
#
# A backup inside Config\ or anywhere under the console is destroyed by the very
# operation it exists to protect against: copying a new console folder over the
# old one. Beside it survives that, survives a rename, and is still obvious to
# find.
#
# Keyed by console folder name so several consoles on one machine -- 'AdminConsole
# v3' beside 'MasterConsole' -- keep separate histories instead of overwriting
# each other's.
if (-not $BackupRoot) {
    $parent = Split-Path -Parent $PSScriptRoot
    $leaf   = Split-Path -Leaf $PSScriptRoot
    $BackupRoot = Join-Path (Join-Path $parent 'ConfigBackups') $leaf
}
$script:BackupRoot = $BackupRoot
#endregion

#region Backup / restore --------------------------------------------------------
function Get-ConfigFile {
    if (-not (Test-Path -LiteralPath $script:ConfigDir)) { return @() }
    return @(Get-ChildItem -LiteralPath $script:ConfigDir -Filter '*.json' -File -ErrorAction SilentlyContinue |
             Sort-Object Name)
}

function New-ConfigBackup {
    <#
        One timestamped folder per backup, plus a manifest.

        The manifest records name, size and hash. Without it a restore is a
        guess about whether anything actually differs -- and "restore" that
        silently replaces a file with an identical one is indistinguishable
        from one that replaces it with something wrong.
    #>
    param([string]$Reason = 'manual')

    $files = Get-ConfigFile
    if ($files.Count -eq 0) {
        return [PSCustomObject]@{ Ok = $false; Path = $null; Count = 0
                                  Message = "No .json files in $($script:ConfigDir)." }
    }
    $stamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $dest  = Join-Path $script:BackupRoot $stamp
    try {
        New-Item -Path $dest -ItemType Directory -Force -ErrorAction Stop | Out-Null
        $manifest = @()
        foreach ($f in $files) {
            Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $dest $f.Name) -Force -ErrorAction Stop
            $hash = ''
            try { $hash = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256 -ErrorAction Stop).Hash } catch { }
            $manifest += [PSCustomObject]@{
                Name = $f.Name; Bytes = $f.Length; Sha256 = $hash
                Modified = $f.LastWriteTime.ToString('s')
            }
        }
        [PSCustomObject]@{
            TakenUtc = (Get-Date).ToUniversalTime().ToString('s')
            Console  = $PSScriptRoot
            Reason   = $Reason
            Files    = $manifest
        } | ConvertTo-Json -Depth 4 |
            Set-Content -LiteralPath (Join-Path $dest 'backup-manifest.json') -Encoding UTF8

        return [PSCustomObject]@{ Ok = $true; Path = $dest; Count = $files.Count
                                  Message = "Backed up $($files.Count) file(s) to $dest" }
    } catch {
        return [PSCustomObject]@{ Ok = $false; Path = $dest; Count = 0
                                  Message = "Backup FAILED: $($_.Exception.Message)" }
    }
}

function Get-BackupSet {
    if (-not (Test-Path -LiteralPath $script:BackupRoot)) { return @() }
    $out = @()
    foreach ($d in @(Get-ChildItem -LiteralPath $script:BackupRoot -Directory -ErrorAction SilentlyContinue |
                     Sort-Object Name -Descending)) {
        $json = @(Get-ChildItem -LiteralPath $d.FullName -Filter '*.json' -File -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -ne 'backup-manifest.json' })
        $reason = ''
        $mf = Join-Path $d.FullName 'backup-manifest.json'
        if (Test-Path -LiteralPath $mf) {
            try { $reason = "$((Get-Content -LiteralPath $mf -Raw | ConvertFrom-Json).Reason)" } catch { }
        }
        $out += [PSCustomObject]@{
            Stamp = $d.Name; Path = $d.FullName; Count = $json.Count
            Files = (@($json | ForEach-Object { $_.Name }) -join ', ')
            Reason = $reason
        }
    }
    return $out
}

function Restore-ConfigBackup {
    param([string]$Path)

    $src = @(Get-ChildItem -LiteralPath $Path -Filter '*.json' -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -ne 'backup-manifest.json' })
    if ($src.Count -eq 0) { return "Nothing to restore -- that folder holds no config files." }

    # A backup of the CURRENT state before overwriting it. Restoring the wrong
    # snapshot is an easy mistake and this makes it a recoverable one; without
    # it, restore is the only destructive thing this tool does.
    $pre = New-ConfigBackup -Reason 'auto -- taken before a restore'

    $done = 0; $failed = @()
    foreach ($f in $src) {
        try {
            Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $script:ConfigDir $f.Name) -Force -ErrorAction Stop
            $done++
        } catch { $failed += "$($f.Name): $($_.Exception.Message)" }
    }
    $msg = "Restored $done file(s) from $(Split-Path -Leaf $Path)."
    if ($pre.Ok) { $msg += "`r`n`r`nWhat was there first is safe in: $(Split-Path -Leaf $pre.Path)" }
    else         { $msg += "`r`n`r`nWARNING -- the pre-restore backup failed: $($pre.Message)" }
    if ($failed.Count) { $msg += "`r`n`r`nFailed:`r`n" + ($failed -join "`r`n") }
    $msg += "`r`n`r`nClose and reopen any console window to pick the restored settings up."
    return $msg
}
#endregion

#region Form --------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Config Backup'
$form.Size = New-Object System.Drawing.Size(820, 560)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWhy = New-Object System.Windows.Forms.Label
$lblWhy.Location = New-Object System.Drawing.Point(12, 10)
$lblWhy.Size = New-Object System.Drawing.Size(780, 32)
$lblWhy.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWhy.Text = 'Back up before copying a new console folder over this one. A folder copy replaces every file below and warns about nothing.'
$form.Controls.Add($lblWhy)

$lblPaths = New-Object System.Windows.Forms.Label
$lblPaths.Location = New-Object System.Drawing.Point(12, 44)
$lblPaths.Size = New-Object System.Drawing.Size(780, 32)
$lblPaths.ForeColor = [System.Drawing.Color]::DimGray
$lblPaths.Text = "From: $($script:ConfigDir)" + [Environment]::NewLine +
                 "To  : $($script:BackupRoot)   (outside the console folder, so replacing it cannot take the backups too)"
$form.Controls.Add($lblPaths)

$grpNow = New-Object System.Windows.Forms.GroupBox
$grpNow.Text = 'Config files right now'
$grpNow.Location = New-Object System.Drawing.Point(12, 80)
$grpNow.Size = New-Object System.Drawing.Size(780, 120)
$form.Controls.Add($grpNow)

$lstNow = New-Object System.Windows.Forms.ListBox
$lstNow.Location = New-Object System.Drawing.Point(10, 20)
$lstNow.Size = New-Object System.Drawing.Size(760, 90)
$grpNow.Controls.Add($lstNow)

$grpOld = New-Object System.Windows.Forms.GroupBox
$grpOld.Text = 'Backups (newest first)'
$grpOld.Location = New-Object System.Drawing.Point(12, 208)
$grpOld.Size = New-Object System.Drawing.Size(780, 230)
$form.Controls.Add($grpOld)

$lvOld = New-Object System.Windows.Forms.ListView
$lvOld.Location = New-Object System.Drawing.Point(10, 20)
$lvOld.Size = New-Object System.Drawing.Size(760, 198)
$lvOld.View = 'Details'
$lvOld.FullRowSelect = $true
$lvOld.GridLines = $true
[void]$lvOld.Columns.Add('Taken', 150)
[void]$lvOld.Columns.Add('Files', 50)
[void]$lvOld.Columns.Add('Reason', 200)
[void]$lvOld.Columns.Add('Contents', 350)
$grpOld.Controls.Add($lvOld)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(12, 444)
$lblStatus.Size = New-Object System.Drawing.Size(780, 34)
$form.Controls.Add($lblStatus)

function Update-Lists {
    $lstNow.Items.Clear()
    foreach ($f in Get-ConfigFile) {
        [void]$lstNow.Items.Add(("{0,-28} {1,7:N0} bytes   modified {2}" -f $f.Name, $f.Length, $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm')))
    }
    if ($lstNow.Items.Count -eq 0) { [void]$lstNow.Items.Add('No .json files found in Config\.') }

    $lvOld.Items.Clear()
    foreach ($b in Get-BackupSet) {
        $it = New-Object System.Windows.Forms.ListViewItem($b.Stamp)
        [void]$it.SubItems.Add("$($b.Count)")
        [void]$it.SubItems.Add($b.Reason)
        [void]$it.SubItems.Add($b.Files)
        $it.Tag = $b
        [void]$lvOld.Items.Add($it)
    }
}

$y = 484
$btnBackup = New-Object System.Windows.Forms.Button
$btnBackup.Text = 'Back up now'
$btnBackup.Location = New-Object System.Drawing.Point(12, $y)
$btnBackup.Size = New-Object System.Drawing.Size(130, 30)
$btnBackup.Add_Click({
    $r = New-ConfigBackup -Reason 'manual'
    $lblStatus.Text = $r.Message
    $lblStatus.ForeColor = if ($r.Ok) { [System.Drawing.Color]::ForestGreen } else { [System.Drawing.Color]::Firebrick }
    Update-Lists
    try { Write-CompletedLog -Action 'ConfigBackup' -Target 'Backup' -Result $(if ($r.Ok) { 'Success' } else { 'Failed' }) -Detail $r.Message } catch { }
})
$form.Controls.Add($btnBackup)

$btnRestore = New-Object System.Windows.Forms.Button
$btnRestore.Text = 'Restore selected'
$btnRestore.Location = New-Object System.Drawing.Point(150, $y)
$btnRestore.Size = New-Object System.Drawing.Size(140, 30)
$btnRestore.Add_Click({
    if ($lvOld.SelectedItems.Count -ne 1) {
        $lblStatus.Text = 'Select exactly one backup to restore.'
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        return
    }
    $b = $lvOld.SelectedItems[0].Tag
    $ans = Show-ToolDialog(
        "Copy $($b.Count) file(s) from $($b.Stamp) over Config\ ?`r`n`r`n$($b.Files)`r`n`r`nWhat is there now is backed up first, so this can be undone.`r`n`r`nContinue?",
        'Restore config', [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $msg = Restore-ConfigBackup -Path $b.Path
    Update-Lists
    $null = Show-ToolDialog($msg, 'Restore config')
    $lblStatus.Text = ($msg -split "`r`n")[0]
    $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
    try { Write-CompletedLog -Action 'ConfigBackup' -Target 'Restore' -Result 'Success' -Detail "from $($b.Stamp)" } catch { }
})
$form.Controls.Add($btnRestore)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = 'Open backup folder'
$btnOpen.Location = New-Object System.Drawing.Point(298, $y)
$btnOpen.Size = New-Object System.Drawing.Size(150, 30)
$btnOpen.Add_Click({
    if (-not (Test-Path -LiteralPath $script:BackupRoot)) {
        New-Item -Path $script:BackupRoot -ItemType Directory -Force | Out-Null
    }
    Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$($script:BackupRoot)`""
})
$form.Controls.Add($btnOpen)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = 'Refresh'
$btnRefresh.Location = New-Object System.Drawing.Point(456, $y)
$btnRefresh.Size = New-Object System.Drawing.Size(100, 30)
$btnRefresh.Add_Click({ Update-Lists })
$form.Controls.Add($btnRefresh)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(692, $y)
$btnClose.Size = New-Object System.Drawing.Size(100, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

Update-Lists
try { Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath | Out-Null } catch { }
try { Set-ConsoleTheme -Form $form } catch { }
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
[void]$form.ShowDialog()
#endregion
