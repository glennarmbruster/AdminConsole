#Requires -Version 5.1
<#
    Exchange-HealthCheck.ps1
    Services, queues, database copies, IIS logging and disk space on the four
    hub servers, on five tabs.

    REFACTORED FROM A CONSOLE SCRIPT. The original wrote colour-coded tables to
    the host and optionally an HTML file. The collection logic is kept almost
    verbatim -- it was sound -- and only the presentation changed: one remote
    snapshot per server, then five tabs instead of a scrolling wall.

    -Environment IS ACCEPTED AND USED, and passed to every Write-*Log call.
#>
# TOOLKIT-HELP-RUNS: Reads services, queues, database copies and disks from on-premises servers. Read-only.
# TOOLKIT-VARIANT: OnPrem
# TOOLKIT-HELP-BEGIN
# The health of the four hub servers, on five tabs. Press Run.
#
# WHAT EACH TAB IS FOR:
#
#   Summary & Server Overview   Start here. Every problem found across all the
#                               other tabs, listed once, plus uptime and who
#                               last rebooted each server.
#   Services / IIS / Logging    Exchange services that are not running when they
#                               should be, IIS state, and whether IIS logging is
#                               actually writing.
#   Database Copy Status        Mounted, healthy, and how far behind each copy
#                               is. Copy and replay queues matter more than the
#                               status word.
#   Transport Queues            Queue lengths and anything retrying.
#   Drive Space                 Free space per drive, worst first.
#
# AUTOMOUNT AND CONTENTINDEXSTATE ARE NOT NOISE. A copy that is Healthy with a
# failed content index cannot serve search, and a database that is mounted but
# whose copies are all behind is a failover that will lose mail.
#
# A SERVER THAT DID NOT ANSWER IS NAMED, not skipped. An empty row for a server
# that is down looks exactly like a server with nothing wrong.
#
# It needs WinRM to each server. Queue and database data additionally need the
# Exchange management tools.
#
# Read-only. It reports and changes nothing.
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

# Thresholds, from the original script's parameters.
$script:QueueWarning        = 50
$script:QueueCritical       = 500
$script:DiskWarningPercent  = 15
$script:DiskCriticalPercent = 10

# Services that matter most on an Exchange mailbox server. Kept as-is.
$script:ImportantServices = @(
    'MSExchangeADTopology', 'MSExchangeIS', 'MSExchangeRepl', 'MSExchangeRPC'
    'MSExchangeTransport', 'MSExchangeMailboxTransportSubmission'
    'MSExchangeMailboxTransportDelivery', 'MSExchangeFrontEndTransport'
    'MSExchangeTransportLogSearch', 'MSExchangeMailboxAssistants'
    'MSExchangeDelivery', 'MSExchangeSubmission', 'MSExchangeServiceHost'
    'MSExchangeDagMgmt', 'MSExchangeHM', 'MSExchangeAntispamUpdate'
    'MSExchangeCompliance', 'MSExchangeDiagnostics', 'MSExchangeEdgeSync'
    'MSExchangeFastSearch', 'HostControllerService', 'MSExchangeThrottling'
    'MSExchangeUM', 'MSExchangeUMCR', 'W3SVC', 'WAS', 'IISADMIN', 'RemoteRegistry'
)

$script:Snapshots = @{}
$script:Problems  = New-Object System.Collections.ArrayList
$script:Failed    = New-Object System.Collections.ArrayList
$script:Queues    = New-Object System.Collections.ArrayList
$script:Copies    = New-Object System.Collections.ArrayList
$script:LastRun   = $null

#region Collection --------------------------------------------------------------
function Get-ServerSnapshot {
    <#
        One round trip per server.

        EVERYTHING IN ONE Invoke-Command ON PURPOSE. Services, uptime, the
        reboot event, IIS state and disks are five questions, and asking them
        separately is five round trips per server -- twenty across the estate,
        most of the runtime spent on connection setup rather than work.
    #>
    param([Parameter(Mandatory)][string]$ComputerName)

    Invoke-Command -ComputerName $ComputerName -ArgumentList (, $script:ImportantServices) -ErrorAction Stop -ScriptBlock {
        param([string[]]$ServiceName)

        $os = Get-CimInstance -ClassName Win32_OperatingSystem

        $services = foreach ($name in $ServiceName) {
            $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
            if ($svc) {
                $wmi = Get-CimInstance -ClassName Win32_Service -Filter "Name='$($svc.Name)'" -ErrorAction SilentlyContinue
                [pscustomobject]@{
                    Name        = $svc.Name
                    DisplayName = $svc.DisplayName
                    Status      = [string]$svc.Status
                    StartMode   = if ($wmi) { $wmi.StartMode } else { 'Unknown' }
                }
            }
        }

        # 1074 = clean shutdown/restart request (carries the user),
        # 6008 = previous shutdown was unexpected.
        $rebootEvent = $null
        try {
            $rebootEvent = Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = 1074, 6006, 6008 } -MaxEvents 40 -ErrorAction Stop |
                Where-Object { $_.TimeCreated -le $os.LastBootUpTime.AddMinutes(5) } |
                Select-Object -First 1
        } catch { }

        $rebootUser = 'Unknown'; $rebootReason = 'Not recorded'
        if ($rebootEvent) {
            if ($rebootEvent.Id -eq 1074) {
                $props = @($rebootEvent.Properties | ForEach-Object { [string]$_.Value })
                if ($props.Count -ge 7 -and $props[6]) { $rebootUser = $props[6] }
                $processName = if ($props.Count -ge 1) { Split-Path -Leaf $props[0] } else { 'unknown process' }
                $comment = if ($props.Count -ge 5 -and $props[4]) { $props[4] } else { 'no comment' }
                $rebootReason = "$processName / $comment"
            } elseif ($rebootEvent.Id -eq 6008) {
                $rebootUser = 'N/A (unexpected shutdown)'
                $rebootReason = 'Previous shutdown was unexpected'
            } else {
                $rebootReason = 'Event log service stopped (6006)'
            }
        }

        $iis = [pscustomobject]@{
            ServiceStatus = 'Not installed'; LoggingState = 'Unknown'; LogPath = 'n/a'
            LastWrite = $null; LogSizeGB = 0; FileCount = 0; AppPoolStopped = @()
        }
        $w3svc = Get-Service -Name W3SVC -ErrorAction SilentlyContinue
        if ($w3svc) {
            $iis.ServiceStatus = [string]$w3svc.Status
            $logRoot = $null
            try {
                Import-Module WebAdministration -ErrorAction Stop
                $logRoot = (Get-WebConfigurationProperty -Filter '/system.applicationHost/sites/siteDefaults/logFile' -Name directory).Value
                $logRoot = [Environment]::ExpandEnvironmentVariables($logRoot)
                $iis.LoggingState = if ((Get-WebConfigurationProperty -Filter '/system.applicationHost/sites/siteDefaults/logFile' -Name enabled).Value) { 'Enabled' } else { 'Disabled' }
                $iis.AppPoolStopped = @(Get-ChildItem IIS:\AppPools -ErrorAction SilentlyContinue |
                    Where-Object { $_.State -ne 'Started' } | Select-Object -ExpandProperty Name)
            } catch {
                $logRoot = Join-Path $env:SystemDrive 'inetpub\logs\LogFiles'
                $iis.LoggingState = 'Unverified (WebAdministration unavailable)'
            }
            if ($logRoot -and (Test-Path $logRoot)) {
                $iis.LogPath = $logRoot
                $files = @(Get-ChildItem -Path $logRoot -Filter *.log -Recurse -ErrorAction SilentlyContinue)
                $iis.FileCount = $files.Count
                if ($files.Count -gt 0) {
                    $iis.LogSizeGB = [Math]::Round((($files | Measure-Object -Property Length -Sum).Sum / 1GB), 2)
                    $iis.LastWrite = ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
                }
            } else {
                $iis.LogPath = if ($logRoot) { "$logRoot (missing)" } else { 'n/a' }
            }
        }

        $disks = foreach ($disk in Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3') {
            [pscustomobject]@{
                Drive       = $disk.DeviceID
                Label       = $disk.VolumeName
                SizeGB      = [Math]::Round($disk.Size / 1GB, 1)
                FreeGB      = [Math]::Round($disk.FreeSpace / 1GB, 1)
                FreePercent = if ($disk.Size -gt 0) { [Math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1) } else { 0 }
            }
        }

        [pscustomobject]@{
            ComputerName = $env:COMPUTERNAME
            OS           = $os.Caption
            LastBootTime = $os.LastBootUpTime
            UptimeHours  = [Math]::Round(((Get-Date) - $os.LastBootUpTime).TotalHours, 1)
            RebootUser   = $rebootUser
            RebootReason = $rebootReason
            Services     = @($services)
            Iis          = $iis
            Disks        = @($disks)
        }
    }
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Exchange Health Check'
$form.ClientSize      = New-Object System.Drawing.Size(1280, 860)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'

$btnRun           = New-Object System.Windows.Forms.Button
$btnRun.Text      = 'Run'
$btnRun.Location  = New-Object System.Drawing.Point(14, 14)
$btnRun.Size      = New-Object System.Drawing.Size(120, 30)
$btnRun.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnRun)

$lblProg           = New-Object System.Windows.Forms.Label
$lblProg.Location  = New-Object System.Drawing.Point(148, 20)
$lblProg.Size      = New-Object System.Drawing.Size(700, 20)
$lblProg.ForeColor = [System.Drawing.Color]::Gray
[void]$form.Controls.Add($lblProg)

$btnExport           = New-Object System.Windows.Forms.Button
$btnExport.Text      = 'Export CSV'
$btnExport.Location  = New-Object System.Drawing.Point(1000, 14)
$btnExport.Size      = New-Object System.Drawing.Size(120, 30)
$btnExport.FlatStyle = 'Flat'
$btnExport.Anchor    = 'Top,Right'
[void]$form.Controls.Add($btnExport)

$pb = Add-ReportProgressBar -Form $form -X 14 -Y 50 -Width 1252 -Height 12

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(14, 70)
$tabs.Size     = New-Object System.Drawing.Size(1252, 740)
$tabs.Anchor   = 'Top,Left,Right,Bottom'
[void]$form.Controls.Add($tabs)

function New-GridTab {
    param([string]$Title)
    $tp      = New-Object System.Windows.Forms.TabPage
    $tp.Text = $Title
    $g                     = New-Object System.Windows.Forms.DataGridView
    $g.Dock                = 'Fill'
    $g.ReadOnly            = $true
    $g.AllowUserToAddRows  = $false
    $g.AutoSizeColumnsMode = 'Fill'
    $g.SelectionMode       = 'FullRowSelect'
    $g.MultiSelect         = $false
    [void]$tp.Controls.Add($g)
    [void]$tabs.TabPages.Add($tp)
    return $g
}

# TAB 1 IS SPLIT: the summary reads as prose and the overview as a table, and
# forcing either into the other's shape makes both worse.
$tpSummary      = New-Object System.Windows.Forms.TabPage
$tpSummary.Text = 'Summary and Server Overview'
[void]$tabs.TabPages.Add($tpSummary)

# A SPLITCONTAINER, NOT TWO DOCKED CONTROLS.
#
# It was Dock=Top on the text and Dock=Fill on the grid, ordered with
# BringToFront. That is z-order-dependent: WinForms docks the BACK of the
# collection first, so the Fill control claimed the whole page and the overview
# grid ended up UNDERNEATH the summary rather than below it. The server list was
# built and populated the whole time and simply could not be seen.
#
# A split is deterministic, and it lets the divider be dragged when the problem
# list runs long.
$split               = New-Object System.Windows.Forms.SplitContainer
$split.Dock          = 'Fill'
$split.Orientation   = 'Horizontal'
$split.SplitterWidth = 6
[void]$tpSummary.Controls.Add($split)

$txtSummary            = New-Object System.Windows.Forms.TextBox
$txtSummary.Multiline  = $true
$txtSummary.ReadOnly   = $true
$txtSummary.ScrollBars = 'Both'
$txtSummary.WordWrap   = $false
$txtSummary.Dock       = 'Fill'
$txtSummary.Font       = New-Object System.Drawing.Font('Consolas', 9)
[void]$split.Panel1.Controls.Add($txtSummary)

$gridOverview                     = New-Object System.Windows.Forms.DataGridView
$gridOverview.Dock                = 'Fill'
$gridOverview.ReadOnly            = $true
$gridOverview.AllowUserToAddRows  = $false
$gridOverview.AutoSizeColumnsMode = 'Fill'
$gridOverview.SelectionMode       = 'FullRowSelect'
[void]$split.Panel2.Controls.Add($gridOverview)

$gridServices = New-GridTab -Title 'Services and IIS'
# LOGGING GETS ITS OWN TAB. It was one squashed row per server on the services
# grid, sharing columns named for services -- so the log path, size, file count
# and last-write time were all crammed into a Verdict cell and truncated.
$gridLogging  = New-GridTab -Title 'Logging'
$gridCopies   = New-GridTab -Title 'Mailbox Database Copy Status'
$gridQueues   = New-GridTab -Title 'Transport Queues'
$gridDisks    = New-GridTab -Title 'Drive Space'

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(1146, 820)
$btnClose.Size      = New-Object System.Drawing.Size(120, 28)
$btnClose.FlatStyle = 'Flat'
$btnClose.Anchor    = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
[void]$form.Controls.Add($btnClose)
#endregion

#region Render ------------------------------------------------------------------
function Set-GridRows {
    param($Grid, [string[]]$Columns, $Rows, [hashtable]$NumericColumns = @{})

    $t = New-Object System.Data.DataTable
    foreach ($c in $Columns) {
        if ($NumericColumns.ContainsKey($c)) { [void]$t.Columns.Add($c, $NumericColumns[$c]) }
        else { [void]$t.Columns.Add($c) }
    }
    foreach ($r in @($Rows)) {
        $vals = foreach ($c in $Columns) {
            $v = $r.$c
            if ($null -eq $v) { [System.DBNull]::Value } else { $v }
        }
        [void]$t.Rows.Add($vals)
    }
    $Grid.DataSource = $t
    foreach ($c in $NumericColumns.Keys) {
        if ($Grid.Columns.Contains($c)) { $Grid.Columns[$c].DefaultCellStyle.Alignment = 'MiddleRight' }
    }
}

function Set-RowColour {
    param($Grid, [string]$Column, [scriptblock]$IsBad, [scriptblock]$IsWarn)
    foreach ($gr in $Grid.Rows) {
        $v = $gr.Cells[$Column].Value
        if ($IsBad.InvokeReturnAsIs($v))       { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
        elseif ($IsWarn.InvokeReturnAsIs($v))  { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGoldenrod }
    }
}

function Invoke-HealthCheck {
    $script:Snapshots = @{}
    $script:Problems.Clear(); $script:Failed.Clear()
    $script:Queues.Clear();   $script:Copies.Clear()

    if (-not (Test-ServerListPresent)) { return }
    $pb.Maximum = [Math]::Max(1, @($script:Servers).Count + 2)
    $pb.Value   = 0
    $pb.Visible = $true
    $btnRun.Enabled = $false

    try {
        foreach ($srv in $script:Servers) {
            $lblProg.Text = "Collecting from $srv ..."
            $form.Refresh()
            try {
                $script:Snapshots[$srv] = Get-ServerSnapshot -ComputerName $srv
            } catch {
                [void]$script:Failed.Add([PSCustomObject]@{ Server = $srv; Reason = "$($_.Exception.Message)" })
            }
            $pb.Value = [Math]::Min($pb.Maximum, $pb.Value + 1)
        }

        # ---- database copies -------------------------------------------------
        $lblProg.Text = 'Reading database copy status ...'; $form.Refresh()
        foreach ($srv in $script:Servers) {
            try {
                # Invoke-OnPremCommand, not Invoke-EmsCommand. The two do the
                # same job -- resolve an on-premises cmdlet explicitly rather
                # than trusting import order -- and are named differently in
                # the two consoles: the second tenant's Common.ps1 calls it
                # Invoke-EmsCommand below its tenant marker. Repointed on the
                # way across 2026-09-12 rather than carrying the alias, so
                # this file reads like every other tool in THIS toolkit.
                foreach ($c in @(Invoke-OnPremCommand -Name 'Get-MailboxDatabaseCopyStatus' -Parameters @{
                        Server = $srv; ErrorAction = 'Stop' })) {
                    [void]$script:Copies.Add([PSCustomObject]@{
                        Server        = $srv
                        Database      = "$($c.DatabaseName)"
                        Status        = "$($c.Status)"
                        CopyQueue     = [long]0 + $c.CopyQueueLength
                        ReplayQueue   = [long]0 + $c.ReplayQueueLength
                        ContentIndex  = "$($c.ContentIndexState)"
                        Activation    = "$($c.ActivationPreference)"
                    })
                }
            } catch {
                [void]$script:Failed.Add([PSCustomObject]@{ Server = $srv; Reason = "Get-MailboxDatabaseCopyStatus -- $($_.Exception.Message)" })
            }
        }
        $pb.Value = [Math]::Min($pb.Maximum, $pb.Value + 1)

        # ---- transport queues ------------------------------------------------
        $lblProg.Text = 'Reading transport queues ...'; $form.Refresh()
        foreach ($srv in $script:Servers) {
            try {
                foreach ($q in @(Invoke-OnPremCommand -Name 'Get-Queue' -Parameters @{
                        Server = $srv; ErrorAction = 'Stop' })) {
                    [void]$script:Queues.Add([PSCustomObject]@{
                        Server    = $srv
                        Queue     = "$($q.Identity)"
                        Type      = "$($q.DeliveryType)"
                        Status    = "$($q.Status)"
                        Messages  = [int]0 + $q.MessageCount
                        NextHop   = "$($q.NextHopDomain)"
                        LastError = "$($q.LastError)"
                    })
                }
            } catch {
                [void]$script:Failed.Add([PSCustomObject]@{ Server = $srv; Reason = "Get-Queue -- $($_.Exception.Message)" })
            }
        }
        $pb.Value = $pb.Maximum

        Update-AllTabs
        $script:LastRun = Get-Date
        $lblProg.Text = "Completed at $($script:LastRun.ToString('HH:mm:ss'))"

        Write-CompletedLog -Action 'Report' -Target (Get-LogTarget 'Exchange Health Check') `
            -Result $(if (@($script:Failed).Count -gt 0) { 'Partial' } else { 'Success' }) `
            -Detail "Checked $(@($script:Servers).Count) server(s). $(@($script:Problems).Count) problem(s), $(@($script:Failed).Count) collection failure(s)." `
            -Environment $Environment
    } catch {
        $txtSummary.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Health check'
        Write-FailedLog -Action 'Report' -Target (Get-LogTarget 'Exchange Health Check') `
            -Detail "$($_.Exception.Message)" -Environment $Environment
    } finally {
        $pb.Visible = $false
        $btnRun.Enabled = $true
    }
}

function Update-AllTabs {
    # ---- services --------------------------------------------------------
    $svcRows = New-Object System.Collections.ArrayList
    foreach ($srv in $script:Servers) {
        if (-not $script:Snapshots.ContainsKey($srv)) { continue }
        $snap = $script:Snapshots[$srv]
        foreach ($s in @($snap.Services)) {
            # AUTOMATIC AND STOPPED IS THE PROBLEM. A stopped service set to
            # Manual or Disabled is configuration, not a fault, and flagging it
            # buries the ones that matter.
            $bad = ($s.StartMode -eq 'Auto' -and $s.Status -ne 'Running')
            [void]$svcRows.Add([PSCustomObject]@{
                Server = $srv; Service = $s.Name; 'Display name' = $s.DisplayName
                Status = $s.Status; 'Start mode' = $s.StartMode
                Verdict = $(if ($bad) { 'STOPPED - should be running' } else { 'ok' })
            })
            if ($bad) { [void]$script:Problems.Add("$srv : service $($s.Name) is $($s.Status) but set to Automatic") }
        }
        $iis = $snap.Iis
        [void]$svcRows.Add([PSCustomObject]@{
            Server = $srv; Service = 'IIS (W3SVC)'; 'Display name' = 'Internet Information Services'
            Status = $iis.ServiceStatus; 'Start mode' = ''
            # LOGGING DETAIL MOVED TO ITS OWN TAB. It used to be crammed into
            # this Verdict cell and truncated.
            Verdict = $(if (@($iis.AppPoolStopped).Count -gt 0) { "STOPPED - app pool(s): $(@($iis.AppPoolStopped) -join ', ')" } else { 'ok' })
        })
        if (@($iis.AppPoolStopped).Count -gt 0) { [void]$script:Problems.Add("$srv : app pool(s) stopped -- $(@($iis.AppPoolStopped) -join ', ')") }
    }
    Set-GridRows -Grid $gridServices -Columns @('Server', 'Service', 'Display name', 'Status', 'Start mode', 'Verdict') -Rows $svcRows
    Set-RowColour -Grid $gridServices -Column 'Verdict' -IsBad { param($v) "$v" -like 'STOPPED*' } -IsWarn { param($v) $false }

    # ---- logging ---------------------------------------------------------
    #
    # A LOG PATH THAT EXISTS AND IS NOT BEING WRITTEN TO IS THE INTERESTING
    # CASE, and it is invisible if you only look at whether logging is Enabled.
    # IIS can be configured to log, have a valid directory, and still not be
    # writing -- and the first anybody knows is when the logs are needed.
    $logRows = New-Object System.Collections.ArrayList
    foreach ($srv in $script:Servers) {
        if (-not $script:Snapshots.ContainsKey($srv)) { continue }
        $iis = $script:Snapshots[$srv].Iis

        $lastWrite = 'never'
        $ageHours  = $null
        if ($iis.LastWrite) {
            $lastWrite = ([datetime]$iis.LastWrite).ToString('yyyy-MM-dd HH:mm')
            $ageHours  = [Math]::Round(((Get-Date) - [datetime]$iis.LastWrite).TotalHours, 1)
        }

        $verdict = 'ok'
        if ("$($iis.LoggingState)" -eq 'Disabled') {
            $verdict = 'LOGGING DISABLED'
            [void]$script:Problems.Add("$srv : IIS logging is DISABLED")
        } elseif ("$($iis.LogPath)" -like '*(missing)*') {
            $verdict = 'LOG PATH MISSING'
            [void]$script:Problems.Add("$srv : IIS log path is missing -- $($iis.LogPath)")
        } elseif ($null -eq $ageHours) {
            $verdict = 'NO LOG FILES'
            [void]$script:Problems.Add("$srv : IIS logging is $($iis.LoggingState) but there are no log files")
        } elseif ($ageHours -gt 2) {
            $verdict = "STALE - nothing written for $ageHours hour(s)"
            [void]$script:Problems.Add("$srv : IIS logs have not been written for $ageHours hour(s)")
        } elseif ("$($iis.LoggingState)" -like 'Unverified*') {
            $verdict = 'unverified - WebAdministration unavailable'
        }

        [void]$logRows.Add([PSCustomObject]@{
            Server = $srv; 'IIS service' = $iis.ServiceStatus; 'Logging' = $iis.LoggingState
            'Log path' = $iis.LogPath; 'Files' = [int]$iis.FileCount; 'Size GB' = [double]$iis.LogSizeGB
            'Last write' = $lastWrite; 'Age hrs' = $ageHours; 'Verdict' = $verdict
        })
    }
    Set-GridRows -Grid $gridLogging -Columns @('Server', 'IIS service', 'Logging', 'Log path', 'Files', 'Size GB', 'Last write', 'Age hrs', 'Verdict') `
        -Rows $logRows -NumericColumns @{ 'Files' = [int]; 'Size GB' = [double]; 'Age hrs' = [double] }
    Set-RowColour -Grid $gridLogging -Column 'Verdict' `
        -IsBad  { param($v) "$v" -match 'DISABLED|MISSING|NO LOG FILES' } `
        -IsWarn { param($v) "$v" -match 'STALE|unverified' }

    # ---- database copies -------------------------------------------------
    foreach ($c in $script:Copies) {
        if ("$($c.Status)" -notmatch '(?i)mounted|healthy') { [void]$script:Problems.Add("$($c.Server) : $($c.Database) copy status is $($c.Status)") }
        if ($c.CopyQueue -gt 10)   { [void]$script:Problems.Add("$($c.Server) : $($c.Database) copy queue is $($c.CopyQueue)") }
        if ($c.ReplayQueue -gt 10) { [void]$script:Problems.Add("$($c.Server) : $($c.Database) replay queue is $($c.ReplayQueue)") }
        if ("$($c.ContentIndex)" -notmatch '(?i)healthy') { [void]$script:Problems.Add("$($c.Server) : $($c.Database) content index is $($c.ContentIndex)") }
    }
    Set-GridRows -Grid $gridCopies -Columns @('Server', 'Database', 'Status', 'CopyQueue', 'ReplayQueue', 'ContentIndex', 'Activation') `
        -Rows @($script:Copies | Sort-Object Server, Database) -NumericColumns @{ CopyQueue = [long]; ReplayQueue = [long] }
    Set-RowColour -Grid $gridCopies -Column 'Status' -IsBad { param($v) "$v" -notmatch '(?i)mounted|healthy' } -IsWarn { param($v) $false }

    # ---- queues ----------------------------------------------------------
    foreach ($q in $script:Queues) {
        if ("$($q.Status)" -match '(?i)retry') { [void]$script:Problems.Add("$($q.Server) : queue $($q.Queue) is retrying ($($q.Messages) msg) -- $($q.LastError)") }
        elseif ($q.Messages -ge $script:QueueCritical) { [void]$script:Problems.Add("$($q.Server) : queue $($q.Queue) holds $($q.Messages) message(s)") }
    }
    Set-GridRows -Grid $gridQueues -Columns @('Server', 'Queue', 'Type', 'Status', 'Messages', 'NextHop', 'LastError') `
        -Rows @($script:Queues | Sort-Object Messages -Descending) -NumericColumns @{ Messages = [int] }
    Set-RowColour -Grid $gridQueues -Column 'Messages' `
        -IsBad  { param($v) try { [int]$v -ge $script:QueueCritical } catch { $false } } `
        -IsWarn { param($v) try { [int]$v -ge $script:QueueWarning } catch { $false } }

    # ---- disks -----------------------------------------------------------
    $diskRows = New-Object System.Collections.ArrayList
    foreach ($srv in $script:Servers) {
        if (-not $script:Snapshots.ContainsKey($srv)) { continue }
        foreach ($d in @($script:Snapshots[$srv].Disks)) {
            [void]$diskRows.Add([PSCustomObject]@{
                Server = $srv; Drive = $d.Drive; Label = $d.Label
                'Size GB' = [double]$d.SizeGB; 'Free GB' = [double]$d.FreeGB; 'Free %' = [double]$d.FreePercent
            })
            if ($d.FreePercent -le $script:DiskCriticalPercent) { [void]$script:Problems.Add("$srv : drive $($d.Drive) has $($d.FreePercent)% free ($($d.FreeGB) GB)") }
        }
    }
    Set-GridRows -Grid $gridDisks -Columns @('Server', 'Drive', 'Label', 'Size GB', 'Free GB', 'Free %') `
        -Rows @($diskRows | Sort-Object 'Free %') -NumericColumns @{ 'Size GB' = [double]; 'Free GB' = [double]; 'Free %' = [double] }
    Set-RowColour -Grid $gridDisks -Column 'Free %' `
        -IsBad  { param($v) try { [double]$v -le $script:DiskCriticalPercent } catch { $false } } `
        -IsWarn { param($v) try { [double]$v -le $script:DiskWarningPercent } catch { $false } }

    # ---- overview + summary ----------------------------------------------
    $ovRows = New-Object System.Collections.ArrayList
    foreach ($srv in $script:Servers) {
        if ($script:Snapshots.ContainsKey($srv)) {
            $s = $script:Snapshots[$srv]
            # STATE SUMMARISES THE SERVER IN ONE WORD, so the overview answers
            # "is anything wrong here" without reading five other columns. It
            # counts this server's own problems rather than restating them.
            $mine = @($script:Problems | Where-Object { "$_" -like "$srv :*" })
            $state = if (@($mine).Count -eq 0) { 'OK' } else { "$(@($mine).Count) issue(s)" }

            # A SERVER UP FOR UNDER AN HOUR IS WORTH NOTICING. It has either
            # just been patched or just crashed, and either way the rest of this
            # report describes a server that has not settled yet.
            if ([double]$s.UptimeHours -lt 1) { $state = "$state - rebooted recently" }

            [void]$ovRows.Add([PSCustomObject]@{
                Server = $srv; State = $state; OS = $s.OS
                'Last boot' = ([datetime]$s.LastBootTime).ToString('yyyy-MM-dd HH:mm')
                'Uptime hrs' = [double]$s.UptimeHours
                'Rebooted by' = $s.RebootUser; 'Reason' = $s.RebootReason
            })
        } else {
            [void]$ovRows.Add([PSCustomObject]@{
                Server = $srv; State = 'DID NOT RESPOND'; OS = ''; 'Last boot' = ''
                'Uptime hrs' = $null; 'Rebooted by' = ''; 'Reason' = ''
            })
        }
    }
    Set-GridRows -Grid $gridOverview -Columns @('Server', 'State', 'OS', 'Last boot', 'Uptime hrs', 'Rebooted by', 'Reason') `
        -Rows $ovRows -NumericColumns @{ 'Uptime hrs' = [double] }
    Set-RowColour -Grid $gridOverview -Column 'State' `
        -IsBad  { param($v) "$v" -eq 'DID NOT RESPOND' } `
        -IsWarn { param($v) "$v" -match 'issue|rebooted' }

    $L = New-Object System.Collections.ArrayList
    [void]$L.Add("Exchange health check  --  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$L.Add("Servers: $($script:Servers -join ', ')")
    [void]$L.Add('')

    if (@($script:Problems).Count -eq 0) {
        [void]$L.Add('NO PROBLEMS FOUND.')
        [void]$L.Add('')
        [void]$L.Add('Every Automatic service is running, no database copy is unhealthy or')
        [void]$L.Add('behind, no queue is retrying, and no drive is low. The other tabs hold the')
        [void]$L.Add('detail behind that.')
    } else {
        [void]$L.Add("$(@($script:Problems).Count) PROBLEM(S) FOUND -- every one is also on its own tab:")
        [void]$L.Add('')
        foreach ($p in $script:Problems) { [void]$L.Add("   $p") }
    }

    if (@($script:Failed).Count -gt 0) {
        [void]$L.Add('')
        [void]$L.Add('THIS IS A PARTIAL ANSWER. The following did not respond, so anything wrong')
        [void]$L.Add('on them is NOT listed above -- an empty row for a server that is down looks')
        [void]$L.Add('exactly like a server with nothing wrong:')
        foreach ($f in $script:Failed) { [void]$L.Add("   $($f.Server): $($f.Reason)") }
    }

    $txtSummary.Text = ($L -join [Environment]::NewLine)
}

function Export-HealthCsv {
    if (@($script:Snapshots.Keys).Count -eq 0 -and @($script:Queues).Count -eq 0) {
        $txtSummary.Text = 'Nothing to export. Press Run first.'; return
    }
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Choose a folder for the health check CSV files'
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    @($script:Queues) | Export-Csv -LiteralPath (Join-Path $dlg.SelectedPath "Health-Queues-$stamp.csv") -NoTypeInformation -Encoding UTF8
    @($script:Copies) | Export-Csv -LiteralPath (Join-Path $dlg.SelectedPath "Health-Copies-$stamp.csv") -NoTypeInformation -Encoding UTF8
    @($script:Problems | ForEach-Object { [PSCustomObject]@{ Problem = $_ } }) |
        Export-Csv -LiteralPath (Join-Path $dlg.SelectedPath "Health-Problems-$stamp.csv") -NoTypeInformation -Encoding UTF8
    $txtSummary.Text = "Exported to $($dlg.SelectedPath)" + [Environment]::NewLine + [Environment]::NewLine + $txtSummary.Text
}
#endregion

$btnRun.Add_Click({ try { Invoke-HealthCheck } catch { Show-HandlerError -ErrorRecord $_ -Context 'Run' } })
$btnExport.Add_Click({ try { Export-HealthCsv } catch { Show-HandlerError -ErrorRecord $_ -Context 'Export' } })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Exchange Health Check' | Out-Null

$form.Add_Shown({
    try {
        $txtSummary.Text = "Press Run." + [Environment]::NewLine +
            [Environment]::NewLine +
            "Servers: $($script:Servers -join ', ')" + [Environment]::NewLine +
            "Environment: $(if ($Environment) { $Environment } else { '(none passed)' })" + [Environment]::NewLine +
            [Environment]::NewLine +
            'One remote call per server for services, uptime, IIS and disks, then the' + [Environment]::NewLine +
            'queues and database copies. Expect it to take a few seconds per server.'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Opening Exchange Health Check' }
})

[void]$form.ShowDialog()
