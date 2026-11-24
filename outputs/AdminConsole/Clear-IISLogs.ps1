#Requires -Version 5.1
<#
    Clear-IISLogs.ps1
    Delete IIS log files older than N days from the four hub servers.

    THE ONLY TOOL IN THIS TOOLKIT THAT DELETES ANYTHING. Everything about the
    flow below exists because of that: it scans first, it will not delete
    anything it has not just listed, and the confirmation names the count and
    the size rather than asking "are you sure?".

    -Environment IS ACCEPTED AND USED, and passed to every Write-*Log call.
#>
# TOOLKIT-HELP-RUNS: DELETES IIS log files on on-premises servers over an admin share. Destructive.
# TOOLKIT-VARIANT: OnPrem
# TOOLKIT-HELP-BEGIN
# Delete IIS log files older than a set age from the hub servers.
#
# THIS TOOL DELETES FILES. It is the only one here that does. Read the rest of
# this before running it.
#
# HOW IT WORKS. Press Scan first. Nothing is deleted by scanning -- it lists
# every file that WOULD go, with a total size, and only then does the Delete
# button become available. Delete acts on exactly that list and nothing else,
# so what you reviewed is what goes.
#
# WHAT IT WILL TOUCH, AND NOTHING ELSE:
#   *.log files only, in these two folders on each server:
#       C:\inetpub\logs\LogFiles\w3svc1
#       C:\inetpub\logs\LogFiles\w3svc2
#   Not subfolders. Not other extensions. Not any other path.
#
# WHY IT IS FUSSY ABOUT THE PATH. A typo in a log path that then gets deleted
# recursively is unrecoverable, so the folder is checked to be one of the two
# above before a single file is listed, and anything that is a junction or
# symlink is skipped rather than followed -- a reparse point in a log folder
# would take a delete somewhere nobody intended.
#
# THIRTY DAYS IS A DEFAULT, NOT A RULE. Check what your retention policy
# actually requires before lowering it. IIS logs are the evidence for "who
# connected and when", and once they are gone that question cannot be answered.
# If there is an incident open, do not run this.
#
# AGE IS BY LastWriteTime, not by the date in the filename. A log still being
# written to today keeps today's timestamp regardless of what it is called.
#
# It needs admin share access (\\server\C$) to each server.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

#region Load Common -------------------------------------------------------------
$root = $PSScriptRoot
while ($root -and -not (Test-Path -LiteralPath (Join-Path $root 'Common.ps1'))) {
    $parent = Split-Path -Parent $root
    if (-not $parent -or $parent -eq $root) { $root = ''; break }
    $root = $parent
}
if (-not $root) { throw "Common.ps1 not found above '$PSScriptRoot'." }
. (Join-Path $root 'Common.ps1')
#endregion

if (-not (Test-ConsoleLaunch)) { return }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# SERVERS COME FROM Tenant.json, NEVER FROM THIS FILE.
#
# This list used to be four hostnames written here as literals -- in this
# tool, in Check Queues and in Exchange Health Check, three copies of the same
# fact. That breaks the rule the whole toolkit rests on (Tenant.json and
# Settings.json are the only files allowed to name anything belonging to an
# organisation), and a toolkit handed to another tenant reached for servers
# that do not exist there, with an admin-share timeout as the only clue.
#
# Get-TenantExchangeServerList reads OnPremExchange.Servers from Tenant.json,
# falls back to asking the directory with Get-ExchangeServer when an
# on-premises session is open, and subtracts OnPremExchange.ExcludedServers
# either way. Empty is a real answer and the window says where to set it,
# because an empty list must not read as "nothing to do".
$script:Servers = @()
try { $script:Servers = @(Get-TenantExchangeServerList) } catch { $script:Servers = @() }

# AN EMPTY SERVER LIST IS AN ANSWER, NOT A QUIET NO-OP. With no Servers in
# Tenant.json and no on-premises session, every scan below would sweep zero
# servers, find zero files and report success -- which reads as "there is
# nothing to clean" rather than "I was never told where to look".
function Test-ServerListPresent {
    if (@($script:Servers).Count -gt 0) { return $true }
    $null = Show-ToolDialog(
        ("No on-premises Exchange servers are configured, so there is nothing to work against." + [Environment]::NewLine + [Environment]::NewLine +
         "Set OnPremExchange.Servers in Config\Tenant.json (Tenant Manager), or connect OnPrem and this tool will ask the directory itself." + [Environment]::NewLine + [Environment]::NewLine +
         "OnPremExchange.ExcludedServers is subtracted from either list."),
        'No servers configured', 'OK', 'Warning')
    return $false
}

# THE ONLY TWO FOLDERS THIS TOOL MAY TOUCH.
#
# A LIST, NOT A PARAMETER, and checked again at delete time. If a future change
# needs a third folder it should be added here deliberately, not typed into a
# box by somebody in a hurry at 4pm.
$script:LogFolders = @(
    'inetpub\logs\LogFiles\w3svc1'
    'inetpub\logs\LogFiles\w3svc2'
)

$script:Candidates = New-Object System.Collections.ArrayList
$script:Failed     = New-Object System.Collections.ArrayList
$script:ScanAge    = 0
$script:Scanned    = $false

#region Scan --------------------------------------------------------------------
function Test-AllowedLogPath {
    <#
        Is this folder one of the two this tool is allowed to delete from?

        CHECKED AT SCAN AND AGAIN AT DELETE. Checking once would be enough if
        nothing could change in between, but the whole point of the scan/delete
        split is that time passes between them. The second check costs nothing
        and closes the gap.
    #>
    param([string]$Path)

    $norm = "$Path".TrimEnd('\').ToLowerInvariant()
    foreach ($srv in $script:Servers) {
        foreach ($rel in $script:LogFolders) {
            $expect = ("\\$srv\C$\$rel").ToLowerInvariant()
            if ($norm -eq $expect) { return $true }
        }
    }
    return $false
}

function Get-OldLogFiles {
    param([Parameter(Mandatory)][string]$Server, [Parameter(Mandatory)][int]$Days)

    $r = [PSCustomObject]@{ Ok = $true; Rows = @(); Problems = @() }
    $cutoff = (Get-Date).AddDays(-1 * $Days)
    $rows   = New-Object System.Collections.ArrayList
    $probs  = New-Object System.Collections.ArrayList

    foreach ($rel in $script:LogFolders) {
        $path = "\\$Server\C$\$rel"

        if (-not (Test-AllowedLogPath -Path $path)) {
            [void]$probs.Add("$path is not an allowed log folder -- skipped.")
            continue
        }
        if (-not (Test-Path -LiteralPath $path)) {
            [void]$probs.Add("$path not found or not reachable.")
            continue
        }

        try {
            # -File and no -Recurse, on purpose. Subfolders are not in scope and
            # a recursive delete under a mistyped root is the failure mode this
            # tool must not have.
            $files = @(Get-ChildItem -LiteralPath $path -Filter '*.log' -File -ErrorAction Stop)
            foreach ($f in $files) {
                # SKIP REPARSE POINTS. A junction or symlink in a log folder
                # would take the delete somewhere nobody intended.
                if ($f.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                    [void]$probs.Add("$($f.FullName) is a reparse point -- skipped, not followed.")
                    continue
                }
                if ($f.LastWriteTime -ge $cutoff) { continue }

                [void]$rows.Add([PSCustomObject]@{
                    Server    = $Server
                    Folder    = $rel
                    Name      = $f.Name
                    FullName  = $f.FullName
                    SizeBytes = [long]$f.Length
                    Modified  = $f.LastWriteTime
                    AgeDays   = [int]((Get-Date) - $f.LastWriteTime).TotalDays
                })
            }
        } catch {
            $r.Ok = $false
            [void]$probs.Add("$path -- $($_.Exception.Message)")
        }
    }

    $r.Rows     = @($rows.ToArray())
    $r.Problems = @($probs.ToArray())
    return $r
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Clear IIS Logs'
$form.ClientSize      = New-Object System.Drawing.Size(1200, 900)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$lblAge          = New-Object System.Windows.Forms.Label
$lblAge.Text     = 'Delete logs older than'
$lblAge.Location = New-Object System.Drawing.Point(14, 19)
$lblAge.Size     = New-Object System.Drawing.Size(130, 20)
[void]$form.Controls.Add($lblAge)

$numDays               = New-Object System.Windows.Forms.NumericUpDown
$numDays.Location      = New-Object System.Drawing.Point(148, 16)
$numDays.Size          = New-Object System.Drawing.Size(70, 22)
$numDays.Minimum       = 7
$numDays.Maximum       = 3650
$numDays.Value         = 29
[void]$form.Controls.Add($numDays)

$lblDays          = New-Object System.Windows.Forms.Label
$lblDays.Text     = 'days'
$lblDays.Location = New-Object System.Drawing.Point(224, 19)
$lblDays.Size     = New-Object System.Drawing.Size(40, 20)
[void]$form.Controls.Add($lblDays)

$btnScan           = New-Object System.Windows.Forms.Button
$btnScan.Text      = 'Scan'
$btnScan.Location  = New-Object System.Drawing.Point(276, 14)
$btnScan.Size      = New-Object System.Drawing.Size(120, 28)
$btnScan.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnScan)

# DISABLED UNTIL A SCAN HAS RUN. The button cannot be pressed against a list
# that does not exist, which is the single most important thing on this form.
$btnDelete           = New-Object System.Windows.Forms.Button
$btnDelete.Text      = 'Delete listed files'
$btnDelete.Location  = New-Object System.Drawing.Point(408, 14)
$btnDelete.Size      = New-Object System.Drawing.Size(150, 28)
$btnDelete.FlatStyle = 'Flat'
$btnDelete.Enabled   = $false
$btnDelete.ForeColor = [System.Drawing.Color]::Firebrick
[void]$form.Controls.Add($btnDelete)

$lblProg           = New-Object System.Windows.Forms.Label
$lblProg.Location  = New-Object System.Drawing.Point(572, 19)
$lblProg.Size      = New-Object System.Drawing.Size(420, 20)
$lblProg.ForeColor = [System.Drawing.Color]::Gray
[void]$form.Controls.Add($lblProg)

$btnExport           = New-Object System.Windows.Forms.Button
$btnExport.Text      = 'Export CSV'
$btnExport.Location  = New-Object System.Drawing.Point(1000, 14)
$btnExport.Size      = New-Object System.Drawing.Size(120, 28)
$btnExport.FlatStyle = 'Flat'
$btnExport.Anchor    = 'Top,Right'
[void]$form.Controls.Add($btnExport)

$pb = Add-ReportProgressBar -Form $form -X 14 -Y 48 -Width 1172 -Height 12

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(14, 68)
$grid.Size                = New-Object System.Drawing.Size(1172, 380)
$grid.ReadOnly            = $true
$grid.AllowUserToAddRows  = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode       = 'FullRowSelect'
$grid.MultiSelect         = $false
$grid.Anchor              = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($grid)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(14, 458)
$txtOut.Size       = New-Object System.Drawing.Size(1172, 392)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Anchor     = 'Left,Right,Bottom'
[void]$form.Controls.Add($txtOut)

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(1066, 860)
$btnClose.Size      = New-Object System.Drawing.Size(120, 28)
$btnClose.FlatStyle = 'Flat'
$btnClose.Anchor    = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
[void]$form.Controls.Add($btnClose)
#endregion

#region Display -----------------------------------------------------------------
function Format-Size {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return "$Bytes B"
}

function Update-FileGrid {
    $t = New-Object System.Data.DataTable
    foreach ($c in @('Server', 'Folder', 'File')) { [void]$t.Columns.Add($c) }
    # Typed, so the grid sorts them as numbers rather than as text.
    [void]$t.Columns.Add('Size (bytes)', [long])
    [void]$t.Columns.Add('Age (days)', [int])
    [void]$t.Columns.Add('Modified')

    foreach ($r in @($script:Candidates | Sort-Object Server, Folder, Modified)) {
        [void]$t.Rows.Add($r.Server, $r.Folder, $r.Name, $r.SizeBytes, $r.AgeDays,
            $r.Modified.ToString('yyyy-MM-dd HH:mm:ss'))
    }
    $grid.DataSource = $t
    foreach ($c in @('Size (bytes)', 'Age (days)')) {
        if (-not $grid.Columns.Contains($c)) { continue }
        $grid.Columns[$c].DefaultCellStyle.Format    = 'N0'
        $grid.Columns[$c].DefaultCellStyle.Alignment = 'MiddleRight'
    }
}

function Invoke-Scan {
    $script:Candidates.Clear()
    $script:Failed.Clear()
    $script:Scanned = $false
    $btnDelete.Enabled = $false

    $days = [int]$numDays.Value
    $script:ScanAge = $days

    if (-not (Test-ServerListPresent)) { return }
    $pb.Maximum = [Math]::Max(1, @($script:Servers).Count)
    $pb.Value   = 0
    $pb.Visible = $true
    $btnScan.Enabled = $false

    try {
        foreach ($srv in $script:Servers) {
            $lblProg.Text = "Scanning $srv ..."
            $form.Refresh()

            $res = Get-OldLogFiles -Server $srv -Days $days
            foreach ($row in @($res.Rows)) { [void]$script:Candidates.Add($row) }
            foreach ($p in @($res.Problems)) { [void]$script:Failed.Add([PSCustomObject]@{ Server = $srv; Reason = $p }) }

            $pb.Value = [Math]::Min($pb.Maximum, $pb.Value + 1)
        }

        Update-FileGrid

        $total = (@($script:Candidates | Measure-Object SizeBytes -Sum).Sum)
        if (-not $total) { $total = 0 }

        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("SCAN ONLY -- nothing has been deleted.")
        [void]$L.Add('')
        [void]$L.Add("$(@($script:Candidates).Count) file(s) older than $days day(s), $(Format-Size -Bytes $total) in total.")
        [void]$L.Add('')
        foreach ($g in @($script:Candidates | Group-Object Server | Sort-Object Name)) {
            $sz = (@($g.Group | Measure-Object SizeBytes -Sum).Sum)
            if (-not $sz) { $sz = 0 }
            [void]$L.Add(("   {0,-16} {1,6} file(s)   {2}" -f $g.Name, @($g.Group).Count, (Format-Size -Bytes $sz)))
        }
        [void]$L.Add('')

        if (@($script:Candidates).Count -gt 0) {
            $oldest = @($script:Candidates | Sort-Object Modified)[0]
            $newest = @($script:Candidates | Sort-Object Modified -Descending)[0]
            [void]$L.Add("Oldest: $($oldest.Modified.ToString('yyyy-MM-dd'))   Newest to be deleted: $($newest.Modified.ToString('yyyy-MM-dd'))")
            [void]$L.Add('')
            [void]$L.Add('CHECK THE NEWEST DATE ABOVE before deleting. If it is more recent than you')
            [void]$L.Add('expected, the age is set too low -- and IIS logs are the only record of who')
            [void]$L.Add('connected and when. If an incident is open, stop here.')
            $btnDelete.Enabled = $true
            $script:Scanned = $true
        } else {
            [void]$L.Add('Nothing to delete at that age. Nothing further to do.')
        }

        if (@($script:Failed).Count -gt 0) {
            [void]$L.Add('')
            [void]$L.Add('THIS IS A PARTIAL ANSWER. The following were not read, so anything in them')
            [void]$L.Add('is NOT in the list above and will NOT be deleted:')
            foreach ($f in $script:Failed) { [void]$L.Add("   $($f.Server): $($f.Reason)") }
        }

        $txtOut.Text = ($L -join [Environment]::NewLine)
        $lblProg.Text = "$(@($script:Candidates).Count) file(s) listed."

        Write-CompletedLog -Action 'Report' -Target (Get-LogTarget 'Clear IIS Logs') `
            -Result $(if (@($script:Failed).Count -gt 0) { 'Partial' } else { 'Success' }) `
            -Detail "Scanned $(@($script:Servers).Count) server(s) for logs older than $days day(s). $(@($script:Candidates).Count) file(s), $total bytes. Nothing deleted." `
            -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Scan'
        Write-FailedLog -Action 'Report' -Target (Get-LogTarget 'Clear IIS Logs') `
            -Detail "$($_.Exception.Message)" -Environment $Environment
    } finally {
        $pb.Visible = $false
        $btnScan.Enabled = $true
    }
}

function Invoke-Delete {
    if (-not $script:Scanned -or @($script:Candidates).Count -eq 0) {
        $txtOut.Text = 'Nothing scanned. Press Scan first.'
        return
    }

    $total = (@($script:Candidates | Measure-Object SizeBytes -Sum).Sum)
    if (-not $total) { $total = 0 }
    $newest = @($script:Candidates | Sort-Object Modified -Descending)[0]

    # THE CONFIRMATION NAMES THE NUMBERS. "Are you sure?" is a button people
    # press without reading; a count, a size and the newest date being destroyed
    # are three things that make somebody stop if any of them is wrong.
    $msg = "Permanently delete $(@($script:Candidates).Count) file(s), $(Format-Size -Bytes $total)?" + [Environment]::NewLine + [Environment]::NewLine +
           "Servers      : $(($script:Candidates | ForEach-Object { $_.Server } | Sort-Object -Unique) -join ', ')" + [Environment]::NewLine +
           "Older than   : $script:ScanAge day(s)" + [Environment]::NewLine +
           "Newest to go : $($newest.Modified.ToString('yyyy-MM-dd HH:mm'))" + [Environment]::NewLine + [Environment]::NewLine +
           'This cannot be undone. There is no recycle bin over an admin share.'

    $ans = Show-ToolDialog($msg, 'Delete IIS logs',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
        $txtOut.Text = 'Cancelled. Nothing was deleted.'
        return
    }

    $pb.Maximum = [Math]::Max(1, @($script:Candidates).Count)
    $pb.Value   = 0
    $pb.Visible = $true
    $btnDelete.Enabled = $false
    $btnScan.Enabled   = $false

    $done = 0; $freed = [long]0
    $errs = New-Object System.Collections.ArrayList

    try {
        foreach ($f in @($script:Candidates)) {
            # THE PATH IS CHECKED AGAIN, per file, immediately before deletion.
            # The list was built earlier and this is the last point at which a
            # wrong path can be stopped.
            $dir = Split-Path -Parent $f.FullName
            if (-not (Test-AllowedLogPath -Path $dir)) {
                [void]$errs.Add("REFUSED (not an allowed folder): $($f.FullName)")
                continue
            }
            try {
                Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                $done++
                $freed += [long]$f.SizeBytes
            } catch {
                # A log still being written to is locked, and that is expected
                # rather than a fault -- it is reported and skipped.
                [void]$errs.Add("$($f.FullName) -- $($_.Exception.Message)")
            }
            $pb.Value = [Math]::Min($pb.Maximum, $pb.Value + 1)
            if (($done % 25) -eq 0) { $lblProg.Text = "Deleted $done ..."; $form.Refresh() }
        }

        $L = New-Object System.Collections.ArrayList
        [void]$L.Add("DELETED $done file(s), $(Format-Size -Bytes $freed) freed.")
        [void]$L.Add('')
        if (@($errs).Count -gt 0) {
            [void]$L.Add("$(@($errs).Count) file(s) were NOT deleted:")
            foreach ($e in $errs) { [void]$L.Add("   $e") }
            [void]$L.Add('')
            [void]$L.Add('A file in use by IIS is locked and cannot be removed while it is being')
            [void]$L.Add('written to. That is normal for the current log and is not a fault.')
            [void]$L.Add('')
        }
        [void]$L.Add('Press Scan again to see what remains.')
        $txtOut.Text = ($L -join [Environment]::NewLine)
        $lblProg.Text = "Deleted $done file(s)."

        Write-CompletedLog -Action 'Delete' -Target (Get-LogTarget 'Clear IIS Logs') `
            -Result $(if (@($errs).Count -gt 0) { 'Partial' } else { 'Success' }) `
            -Detail "Deleted $done IIS log file(s) older than $script:ScanAge day(s) across $(@($script:Servers).Count) server(s). $freed bytes freed. $(@($errs).Count) failure(s)." `
            -Environment $Environment
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Delete'
        Write-FailedLog -Action 'Delete' -Target (Get-LogTarget 'Clear IIS Logs') `
            -Detail "$($_.Exception.Message)" -Environment $Environment
    } finally {
        $pb.Visible = $false
        $btnScan.Enabled = $true
        # STAYS DISABLED. The list has just been acted on and is now stale --
        # pressing Delete twice must not be possible without a fresh scan.
        $btnDelete.Enabled = $false
        $script:Scanned = $false
    }
}

function Export-FileList {
    if (@($script:Candidates).Count -eq 0) { $txtOut.Text = 'Nothing to export. Press Scan first.'; return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV (*.csv)|*.csv'
    $dlg.FileName = "IISLogs-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    @($script:Candidates) | Select-Object Server, Folder, Name, FullName, SizeBytes, AgeDays, Modified |
        Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding UTF8
    $txtOut.Text = "Exported $(@($script:Candidates).Count) row(s) to $($dlg.FileName)"
}
#endregion

$btnScan.Add_Click({ try { Invoke-Scan } catch { Show-HandlerError -ErrorRecord $_ -Context 'Scan' } })
$btnDelete.Add_Click({ try { Invoke-Delete } catch { Show-HandlerError -ErrorRecord $_ -Context 'Delete' } })
$btnExport.Add_Click({ try { Export-FileList } catch { Show-HandlerError -ErrorRecord $_ -Context 'Export' } })
# CHANGING THE AGE INVALIDATES THE LIST. Without this, someone could scan at 90
# days, change the box to 7, and press Delete -- which would delete the 90-day
# list while the form said 7.
$numDays.Add_ValueChanged({
    try {
        $btnDelete.Enabled = $false
        $script:Scanned = $false
        if (@($script:Candidates).Count -gt 0) { $lblProg.Text = 'Age changed -- scan again before deleting.' }
    } catch { }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Clear IIS Logs' | Out-Null

$form.Add_Shown({
    try {
        $txtOut.Text = "THIS TOOL DELETES FILES. Press Scan first -- nothing is deleted by scanning." + [Environment]::NewLine +
            [Environment]::NewLine +
            "Servers : $($script:Servers -join ', ')" + [Environment]::NewLine +
            "Folders : $(($script:LogFolders | ForEach-Object { 'C:\' + $_ }) -join '  and  ')" + [Environment]::NewLine +
            "Files   : *.log only, not subfolders" + [Environment]::NewLine +
            "Environment: $(if ($Environment) { $Environment } else { '(none passed)' })" + [Environment]::NewLine +
            [Environment]::NewLine +
            'Delete becomes available only after a scan, and acts on exactly the list shown.'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Clear IIS Logs' }
})

[void]$form.ShowDialog()
