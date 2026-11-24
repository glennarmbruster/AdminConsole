#Requires -Version 5.1
<#
    ServerHealth.ps1
    Disk, patch and update state across the servers named in Tenant.json
    ServerHealth.Servers, on four tabs.

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    2026-09-16  Dashboard button: one-page executive HTML beside the CSVs.
    2026-09-16  Pending reboot no longer counts PendingFileRenameOperations,
                which is present on healthy servers and was flagging the whole
                fleet; the dashboard drops the Reboot column and its count
                entirely (still on the Overview tab, with the reason).
    2026-09-16  Missing updates carry Category, Released date and AgeDays on
                the Updates tab and CSV; the dashboard lists every missing
                update by name, severity, release date and age, worst first.
                Defender definition updates and the KBs in
                $script:DashboardExcludeKB are left off the dashboard and its
                counts; the Updates tab marks them 'Excluded'.

    SIMILAR TO Exchange-HealthCheck.ps1, NOT THE SAME TOOL. That one asks
    Exchange questions -- queues, database copies, transport services -- of the
    Exchange servers. This one asks operating-system questions -- disk, patches,
    reboots -- of a WIDER list that mostly is not Exchange. Same shape, same
    tabs, same one-round-trip-per-server collection; different question.
#>
# TOOLKIT-HELP-RUNS: Reads disks, installed patches, pending updates and the Exchange build from the configured servers. Read-only.
# TOOLKIT-VARIANT: OnPrem
# TOOLKIT-HELP-BEGIN
# Disk space, patch level and pending updates for the servers listed in
# Tenant.json under ServerHealth.Servers. Press Run.
#
# WHAT EACH TAB IS FOR:
#
#   Overview      One row per server. Start here: OS build, worst drive, last
#                 patch date, how many updates are pending, Exchange build, and
#                 whether the server needs a reboot.
#   Drives        Free space per drive, worst first.
#   Updates       Installed patch history and, per server, the updates Windows
#                 Update says are still missing -- KB, severity and title.
#   Exchange      Build number and the CU/SU that build corresponds to, for the
#                 servers that have Exchange on them.
#
# "MISSING UPDATES" MEANS WHAT THAT SERVER'S OWN WINDOWS UPDATE AGENT SAYS.
# The check runs a Windows Update search on each server, so the answer is
# whatever that server would install if you clicked Check for updates -- which
# on a WSUS-managed estate means approved updates, not everything Microsoft has
# ever shipped. A server whose update agent cannot answer is REPORTED AS SUCH,
# with the error code, rather than shown as having nothing pending.
#
# IT IS THE SLOW PART. A Windows Update search takes anything from ten seconds
# to a couple of minutes per server, and the servers are done one at a time.
# Untick "Check for missing updates" for a fast disk-and-patch-date pass.
#
# EXCHANGE BUILD IS REPORTED, CURRENCY IS NOT. The build number is read from
# ExSetup.exe and named from a table in this script -- "2019 CU15 Sep25HU" and
# so on. The script does NOT say whether that is the newest available, because
# nothing on the server knows: it would mean shipping a list of every future
# Microsoft release. The table's own as-of date is shown so you can see how old
# the naming is. Compare against learn.microsoft.com for currency.
#
# DASHBOARD writes a one-page executive summary as a single HTML file beside
# the CSVs and opens it: fleet health at a glance, the servers that need
# attention and why, and one row per server with uptime, last patch, missing
# updates, worst drive and Exchange release, then every missing update by
# name with its severity, release date and age (worst and oldest first) so
# a reader can tell a ninety-day-old Critical from a Defender signature that
# shipped this morning. Self-contained -- it renders
# from a share, an attachment or an offline laptop -- and prints to one
# landscape page. Same data and same colours as the Overview tab.
#
# It needs WinRM to each server, and administrator rights on them.
#
# Read-only. It reports and changes nothing -- no updates are downloaded or
# installed, and nothing is rebooted.
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

$ReportDir = Get-ReportPath

# Thresholds. Percentage AND an absolute floor, because neither alone is right:
# 10% of a 4TB volume is 400GB and fine, while 10% of a 60GB system drive is
# 6GB and about to stop the server patching itself.
$script:DiskWarnPercent = 15
$script:DiskCritPercent = 10
$script:DiskWarnGB      = 20
$script:DiskCritGB      = 10

# A server that has not taken a patch in this long is worth a look. Not an
# error -- a change freeze is a legitimate reason -- so it colours amber and
# says how many days, rather than claiming something is wrong.
$script:PatchStaleDays = 45

# NOT ON THE EXECUTIVE DASHBOARD. Defender definition updates arrive daily
# and are always 'missing' somewhere for an hour; and a KB an executive has
# already accepted as deferred is noise on every page after the first. These
# still appear on the Updates tab and in the CSV -- marked 'Excluded' in the
# Dashboard column -- so nothing is hidden from the engineer; they simply do
# not reach the one-page summary or its counts. Match on KB (case-insensitive,
# with or without the KB prefix), on the Windows Update category, or on the
# title.
$script:DashboardExcludeKB       = @('KB5067398', 'KB5121364')
$script:DashboardExcludeCategory = '^(?i)Definition Updates$'
$script:DashboardExcludeTitle    = '(?i)Security Intelligence Update|Defender Antivirus|Windows Defender|Antimalware Platform'

function Test-DashboardExcluded {
    param([Parameter(Mandatory)]$Update)
    $kbs = @("$($Update.KB)" -split '\s*,\s*' | Where-Object { $_ })
    foreach ($k in $kbs) {
        $n = ($k -replace '(?i)^KB', '')
        foreach ($x in $script:DashboardExcludeKB) { if ($n -eq ($x -replace '(?i)^KB', '')) { return $true } }
    }
    if ("$($Update.Category)" -match $script:DashboardExcludeCategory) { return $true }
    if ("$($Update.Title)" -match $script:DashboardExcludeTitle) { return $true }
    return $false
}

#region Exchange build table ----------------------------------------------------
# BUILD NUMBER -> RELEASE NAME. Transcribed from
# learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates
#
# THE TABLE'S OWN DATE IS SHOWN IN THE REPORT, and that is the honest part of
# this. A build-to-name table is a snapshot of what Microsoft had shipped on the
# day it was written; it can name what is installed, and it can NEVER say
# whether something newer exists. So this tool reports "2019 CU15 Sep25HU" and
# stops. It does not compute "behind by N", because the only way to be right
# about that is to know the future, and the failure mode of guessing is a report
# that says CURRENT next to a server three CUs behind.
#
# Refresh it by re-reading that page and adding rows. Nothing breaks if it goes
# stale -- an unrecognised revision falls back to naming the CU and printing the
# raw build, which is still more use than the number alone.
$script:ExBuildTableAsOf = '2026-09-13'

# Major.Minor.Build -> the CU (or RTM) that base build is.
$script:ExCuByBase = @{
    '15.2.2562' = 'SE RTM'
    '15.2.1748' = '2019 CU15'; '15.2.1544' = '2019 CU14'; '15.2.1258' = '2019 CU13'
    '15.2.1118' = '2019 CU12'; '15.2.986'  = '2019 CU11'; '15.2.922'  = '2019 CU10'
    '15.2.858'  = '2019 CU9';  '15.2.792'  = '2019 CU8';  '15.2.721'  = '2019 CU7'
    '15.2.659'  = '2019 CU6';  '15.2.595'  = '2019 CU5';  '15.2.529'  = '2019 CU4'
    '15.2.464'  = '2019 CU3';  '15.2.397'  = '2019 CU2';  '15.2.330'  = '2019 CU1'
    '15.2.221'  = '2019 RTM'
    '15.1.2507' = '2016 CU23'; '15.1.2375' = '2016 CU22'; '15.1.2308' = '2016 CU21'
    '15.1.2242' = '2016 CU20'; '15.1.2176' = '2016 CU19'; '15.1.2106' = '2016 CU18'
    '15.1.2044' = '2016 CU17'; '15.1.1979' = '2016 CU16'; '15.1.1913' = '2016 CU15'
    '15.1.1847' = '2016 CU14'; '15.1.1779' = '2016 CU13'; '15.1.1713' = '2016 CU12'
    '15.1.1591' = '2016 CU11'; '15.1.1531' = '2016 CU10'; '15.1.1466' = '2016 CU9'
    '15.1.1415' = '2016 CU8'
}

# Full four-part build -> the exact SU/HU on top of that CU, and its date.
#
# ONLY THE CUs STILL IN SERVICE ARE ENUMERATED HERE -- SE RTM, 2019 CU15 and
# CU14, 2016 CU23. Older CUs get their CU name and the raw build, which is the
# right amount of detail for a server that is already years behind: the SU level
# is not the finding, the CU is.
$script:ExSuByBuild = @{
    '15.2.2562.49' = 'SE RTM Sep26SU (2026-09-08)'
    '15.2.2562.46' = 'SE RTM Aug26SU (2026-08-11)'
    '15.2.2562.45' = 'SE RTM Jul26SU (2026-07-14)'
    '15.2.2562.43' = 'SE RTM Jun26SU (2026-06-09)'
    '15.2.2562.41' = 'SE RTM May26HU (2026-05-07)'
    '15.2.2562.37' = 'SE RTM Feb26SU (2026-02-10)'
    '15.2.2562.35' = 'SE RTM Dec25SU (2025-12-09)'
    '15.2.2562.29' = 'SE RTM Oct25SU (2025-10-14)'
    '15.2.2562.27' = 'SE RTM Sep25HU (2025-09-08)'
    '15.2.2562.20' = 'SE RTM Aug25SU (2025-08-12)'
    '15.2.2562.17' = 'SE RTM (2025-07-01)'

    '15.2.1748.51' = '2019 CU15 Sep26SU (2026-09-08)'
    '15.2.1748.49' = '2019 CU15 Aug26SU (2026-08-11)'
    '15.2.1748.48' = '2019 CU15 Jul26SU (2026-07-14)'
    '15.2.1748.46' = '2019 CU15 Jun26SU (2026-06-09)'
    '15.2.1748.43' = '2019 CU15 Feb26SU (2026-02-10)'
    '15.2.1748.42' = '2019 CU15 Dec25SU (2025-12-09)'
    '15.2.1748.39' = '2019 CU15 Oct25SU (2025-10-14)'
    '15.2.1748.37' = '2019 CU15 Sep25HU (2025-09-08)'
    '15.2.1748.36' = '2019 CU15 Aug25SU (2025-08-12)'
    '15.2.1748.26' = '2019 CU15 May25HU (2025-05-29)'
    '15.2.1748.24' = '2019 CU15 Apr25HU (2025-04-18)'
    '15.2.1748.10' = '2019 CU15 (2025-02-10)'

    '15.2.1544.46' = '2019 CU14 Sep26SU (2026-09-08)'
    '15.2.1544.44' = '2019 CU14 Aug26SU (2026-08-11)'
    '15.2.1544.43' = '2019 CU14 Jul26SU (2026-07-14)'
    '15.2.1544.41' = '2019 CU14 Jun26SU (2026-06-09)'
    '15.2.1544.39' = '2019 CU14 Feb26SU (2026-02-10)'
    '15.2.1544.37' = '2019 CU14 Dec25SU (2025-12-09)'
    '15.2.1544.36' = '2019 CU14 Oct25SU (2025-10-14)'
    '15.2.1544.34' = '2019 CU14 Sep25HU (2025-09-08)'
    '15.2.1544.33' = '2019 CU14 Aug25SU (2025-08-12)'
    '15.2.1544.27' = '2019 CU14 May25HU (2025-05-29)'
    '15.2.1544.25' = '2019 CU14 Apr25HU (2025-04-18)'
    '15.2.1544.14' = '2019 CU14 Nov24SUv2 (2024-11-27)'
    '15.2.1544.13' = '2019 CU14 Nov24SU (2024-11-12)'
    '15.2.1544.11' = '2019 CU14 Apr24HU (2024-04-23)'
    '15.2.1544.9'  = '2019 CU14 Mar24SU (2024-03-12)'
    '15.2.1544.4'  = '2019 CU14 (2024-02-13)'

    '15.1.2507.73' = '2016 CU23 Sep26SU (2026-09-08)'
    '15.1.2507.72' = '2016 CU23 Aug26SU (2026-08-11)'
    '15.1.2507.71' = '2016 CU23 Jul26SU (2026-07-14)'
    '15.1.2507.69' = '2016 CU23 Jun26SU (2026-06-09)'
    '15.1.2507.66' = '2016 CU23 Feb26SU (2026-02-10)'
    '15.1.2507.63' = '2016 CU23 Dec25SU (2025-12-09)'
    '15.1.2507.61' = '2016 CU23 Oct25SU (2025-10-14)'
    '15.1.2507.59' = '2016 CU23 Sep25HU (2025-09-08)'
    '15.1.2507.58' = '2016 CU23 Aug25SU (2025-08-12)'
    '15.1.2507.57' = '2016 CU23 May25HU (2025-05-29)'
    '15.1.2507.55' = '2016 CU23 Apr25HU (2025-04-18)'
    '15.1.2507.44' = '2016 CU23 Nov24SUv2 (2024-11-27)'
    '15.1.2507.43' = '2016 CU23 Nov24SU (2024-11-12)'
    '15.1.2507.39' = '2016 CU23 Apr24HU (2024-04-23)'
    '15.1.2507.37' = '2016 CU23 Mar24SU (2024-03-12)'
    '15.1.2507.6'  = '2016 CU23 (2022-04-20)'
}

function Resolve-ExchangeRelease {
    <#
        Name a build. Never claims currency -- see the table comment.

        Three answers, in descending confidence:
          exact build in the SU table  -> the release name and its date
          base build in the CU table   -> the CU, plus the raw build, and a note
                                          that the SU level is not named here
          neither                      -> 'Unrecognised build', and the number

        The third is a real answer and is printed as one. A build this table has
        never heard of is either newer than the table or a version nobody
        expected to find, and both are worth seeing rather than blanking.
    #>
    param([string]$Build)

    $b = "$Build".Trim()
    if (-not $b) { return 'Not installed' }

    # 15.02.1748.037 and 15.2.1748.37 are the same build written two ways --
    # ExSetup reports the padded form, the Learn table lists both. Strip the
    # padding so one lookup key serves.
    $parts = @($b -split '\.')
    if ($parts.Count -ge 4) {
        $norm = (($parts[0..3] | ForEach-Object { [string][int]$_ }) -join '.')
    } else {
        $norm = $b
    }

    if ($script:ExSuByBuild.ContainsKey($norm)) {
        return $script:ExSuByBuild[$norm]
    }

    if ($parts.Count -ge 3) {
        $base = (($parts[0..2] | ForEach-Object { [string][int]$_ }) -join '.')
        if ($script:ExCuByBase.ContainsKey($base)) {
            return "$($script:ExCuByBase[$base]) + SU not in table ($norm)"
        }
    }

    return "Unrecognised build ($norm)"
}
#endregion

#region Servers -----------------------------------------------------------------
$script:Servers = @()
try { $script:Servers = @(Get-TenantHealthServerList) } catch { $script:Servers = @() }

function Test-ServerListPresent {
    <#
        AN EMPTY LIST IS AN ANSWER, NOT A QUIET NO-OP. With no ServerHealth
        block the run would sweep zero servers, find zero problems and report
        success -- which reads as "the estate is healthy" rather than "I was
        never told where to look". Same rule as Exchange Health Check.
    #>
    if (@($script:Servers).Count -gt 0) { return $true }

    # NAME THE FILE AND THE KEY. "No servers configured" on its own is true of
    # the object in memory and may be false of the file on disk, and it sends
    # somebody to edit a Tenant.json that is already correct.
    $where = ''
    try { $where = Get-TenantHealthServerSource } catch { }

    $null = Show-ToolDialog(
        ("No servers are configured for the health check, so there is nothing to check." + [Environment]::NewLine + [Environment]::NewLine +
         "What was read:" + [Environment]::NewLine + "  $where" + [Environment]::NewLine + [Environment]::NewLine +
         "Set it in Tenant Manager -- the 'Server Health list:' box -- or add this to Config\Tenant.json by hand:" + [Environment]::NewLine + [Environment]::NewLine +
         '  "ServerHealth": { "Servers": [ { "Name": "SRV01", "Exchange": true }, { "Name": "SRV02" } ] }' + [Environment]::NewLine + [Environment]::NewLine +
         "A plain list of names works too. With no ServerHealth block at all, the servers from OnPremExchange.Servers are used instead."),
        'No servers configured', 'OK', 'Warning')
    return $false
}
#endregion

#region Collection --------------------------------------------------------------
function Get-HealthSnapshot {
    <#
        One round trip per server.

        EVERYTHING IN ONE Invoke-Command, for the reason Exchange Health Check
        gives: disks, OS build, patch history, pending reboot, the Exchange
        build and the update search are six questions, and asking them
        separately is six connection setups per server.

        THE UPDATE SEARCH RUNS ON THE TARGET, AND HAS TO.

        IUpdateSearcher::Search IS supported for remote callers with
        administrator rights (learn.microsoft.com "Using WUA From a Remote
        Computer"), but BeginSearch/EndSearch -- the async pair -- are not. So
        this uses the synchronous Search, inside the remote session, where the
        call is local to the machine whose updates are being counted.

        It is also the slow part, which is why -IncludeUpdates exists: the
        search contacts WSUS or Windows Update and can take minutes. When it
        fails it fails LOUDLY into WuStatus rather than returning an empty
        collection, because "no updates pending" and "could not ask" must not
        look the same.
    #>
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [bool]$IncludeUpdates = $true
    )

    Invoke-Command -ComputerName $ComputerName -ArgumentList $IncludeUpdates -ErrorAction Stop -ScriptBlock {
        param([bool]$DoUpdates)

        $os = Get-CimInstance -ClassName Win32_OperatingSystem

        # ---- disks ----------------------------------------------------------
        $disks = foreach ($d in Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3') {
            [pscustomobject]@{
                Drive       = $d.DeviceID
                Label       = $d.VolumeName
                SizeGB      = [Math]::Round($d.Size / 1GB, 1)
                FreeGB      = [Math]::Round($d.FreeSpace / 1GB, 1)
                FreePercent = if ($d.Size -gt 0) { [Math]::Round(($d.FreeSpace / $d.Size) * 100, 1) } else { 0 }
            }
        }

        # ---- OS build -------------------------------------------------------
        #
        # UBR IS THE PART THAT MOVES. Win32_OperatingSystem.Version stops at
        # 10.0.17763 for every Server 2019 ever built; the revision that says
        # WHICH patch level lives in the registry as UBR. Without it two servers
        # a year apart in patching report an identical OS version.
        $ubr = $null
        try {
            $ubr = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR -ErrorAction Stop).UBR
        } catch { }
        $displayVersion = ''
        try {
            $cv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
            if ($cv.PSObject.Properties['DisplayVersion']) { $displayVersion = "$($cv.DisplayVersion)" }
            elseif ($cv.PSObject.Properties['ReleaseId'])  { $displayVersion = "$($cv.ReleaseId)" }
        } catch { }

        $osBuild = "$($os.Version)"
        if ($null -ne $ubr) { $osBuild = "$osBuild.$ubr" }

        # ---- installed patches ----------------------------------------------
        #
        # Get-HotFix, not the update history, because history records ATTEMPTS
        # -- including failures and superseded entries -- while Get-HotFix lists
        # what is actually on the box. InstalledOn is genuinely null on some
        # rows (Microsoft never populated it for certain packages), so it is
        # filtered before Sort rather than sorted with nulls in the set.
        $hotfixes = @()
        try {
            $hotfixes = @(Get-HotFix -ErrorAction Stop |
                Where-Object { $_.InstalledOn } |
                Sort-Object InstalledOn -Descending |
                Select-Object -First 25 |
                ForEach-Object {
                    [pscustomobject]@{
                        HotFixID    = "$($_.HotFixID)"
                        Description = "$($_.Description)"
                        InstalledOn = $_.InstalledOn
                        InstalledBy = "$($_.InstalledBy)"
                    }
                })
        } catch { }
        $lastPatch = $null
        if ($hotfixes.Count -gt 0) { $lastPatch = $hotfixes[0].InstalledOn }

        # ---- pending reboot --------------------------------------------------
        #
        # TWO signals, not three. Servicing and Auto Update are key EXISTENCE
        # checks and each genuinely means 'reboot to finish an install'.
        #
        # PendingFileRenameOperations is NOT a third signal and used to be read
        # as one: the value exists on a healthy server essentially all the time
        # -- antivirus, temp cleanup and any installer that touches a locked
        # file leave entries behind, and the array routinely holds empty
        # strings -- so counting it flagged every server in the fleet as
        # 'reboot pending', which is worse than not reporting it at all. It is
        # collected as a separate note for the engineer on the Updates side and
        # deliberately does not make a server 'needs reboot'.
        $reboot = @()
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
            $reboot += 'Servicing'
        }
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
            $reboot += 'Windows Update'
        }
        $fileRename = 0
        try {
            $pfro = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' `
                        -Name PendingFileRenameOperations -ErrorAction Stop
            if ($pfro) {
                # Non-empty entries only. The raw array is padded with blanks.
                $fileRename = @(@($pfro.PendingFileRenameOperations) | Where-Object { "$_".Trim() }).Count
            }
        } catch { }

        # ---- Exchange --------------------------------------------------------
        #
        # ExSetup.exe's file version, which is the build Microsoft's own
        # documentation tells you to read (option 2 on the build-numbers page)
        # and the only one that reflects an SU. Get-ExchangeServer's
        # AdminDisplayVersion shows the CU and NOT the SU on top of it, so a
        # fully patched server and an unpatched one report the same string.
        #
        # The install path comes from the registry rather than being assumed to
        # be Program Files: it is a setup choice, and several estates have it
        # on a data volume.
        $exBuild = ''
        $exPath  = ''
        try {
            $key = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup' -ErrorAction Stop
            if ($key.PSObject.Properties['MsiInstallPath']) { $exPath = "$($key.MsiInstallPath)" }
        } catch { }
        if ($exPath) {
            $setupExe = Join-Path $exPath 'bin\ExSetup.exe'
            if (Test-Path -LiteralPath $setupExe) {
                try { $exBuild = "$((Get-Item -LiteralPath $setupExe).VersionInfo.FileVersion)".Trim() } catch { }
            }
        }

        # ---- missing updates -------------------------------------------------
        $missing  = @()
        $wuStatus = 'Not checked'
        if ($DoUpdates) {
            try {
                $session  = New-Object -ComObject Microsoft.Update.Session
                $searcher = $session.CreateUpdateSearcher()
                # The documented criteria string. IsHidden=0 leaves out updates
                # an administrator has deliberately declined on that server --
                # they are hidden because somebody decided so, and listing them
                # as missing every month is how a report trains people to
                # ignore it.
                $result = $searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
                foreach ($u in $result.Updates) {
                    $kb = ''
                    try { if (@($u.KBArticleIDs).Count -gt 0) { $kb = 'KB' + (@($u.KBArticleIDs) -join ', KB') } } catch { }
                    # WHEN IT CAME OUT. LastDeploymentChangeTime is the date this
                    # revision of the update was published to Windows Update --
                    # for practical purposes the release date -- and its age is
                    # the number an executive reads: a critical update that has
                    # been available for ninety days is a different conversation
                    # from one that shipped on Tuesday. Category names the kind
                    # (Security Updates, Critical Updates, Updates, Definition
                    # Updates), which is what tells a security patch apart from
                    # a Defender signature when the severity field is blank.
                    $released = $null
                    try {
                        $d = $u.LastDeploymentChangeTime
                        if ($d -and ([DateTime]$d).Year -gt 2000) { $released = [DateTime]$d }
                    } catch { }
                    $cat = ''
                    try {
                        $names = @()
                        foreach ($c in @($u.Categories)) { if ("$($c.Name)") { $names += "$($c.Name)" } }
                        # The most specific first: 'Security Updates' beats 'Windows Server 2019'.
                        $pref = @($names | Where-Object { $_ -match '(?i)^(Security|Critical|Definition|Update Rollups|Updates|Feature Packs|Service Packs)' })
                        $cat = $(if ($pref.Count) { $pref[0] } elseif ($names.Count) { $names[0] } else { '' })
                    } catch { }
                    $missing += [pscustomobject]@{
                        KB       = $kb
                        Severity = "$($u.MsrcSeverity)"
                        Title    = "$($u.Title)"
                        Released = $released
                        Category = $cat
                    }
                }
                $wuStatus = 'OK'
            } catch {
                # The HRESULT is the whole value of this branch. 0x8024402C is a
                # WSUS it cannot reach; 0x80072EE2 is a timeout; 0x80240044 is
                # the call being refused. Each has a different fix, and
                # "failed" alone sends somebody to the wrong one.
                $code = ''
                try { $code = '0x{0:X8}' -f $_.Exception.HResult } catch { }
                $wuStatus = "Failed $code -- $($_.Exception.Message)".Trim()
            }
        }

        [pscustomobject]@{
            ComputerName   = $env:COMPUTERNAME
            OS             = "$($os.Caption)"
            OSDisplay      = $displayVersion
            OSBuild        = $osBuild
            LastBootTime   = $os.LastBootUpTime
            UptimeDays     = [Math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1)
            Disks          = @($disks)
            HotFixes       = @($hotfixes)
            LastPatch      = $lastPatch
            PendingReboot  = @($reboot)
            FileRenamePending = $fileRename
            ExchangeBuild  = $exBuild
            ExchangePath   = $exPath
            MissingUpdates = @($missing)
            WuStatus       = $wuStatus
        }
    }
}
#endregion

#region Results -----------------------------------------------------------------
$script:Snapshots = @{}
$script:Failed    = New-Object System.Collections.ArrayList
$script:RunAt     = $null

function Get-WorstDrive {
    <#
        The drive most likely to cause a problem, and WHY.

        WORST BY STATE, THEN BY PERCENTAGE -- not simply the emptiest. A 4TB
        volume at 12% has 480GB free and needs no attention this year; a 60GB
        system drive at 12% has 7GB and will fail the next cumulative update.
        Ranking on percentage alone puts the first one at the top of the report
        and the second one out of sight.
    #>
    param($Disks)

    $worst = $null; $worstRank = -1
    foreach ($d in @($Disks)) {
        $state = Get-DriveState -Disk $d
        $rank = switch ($state) { 'Critical' { 2 } 'Warning' { 1 } default { 0 } }
        # Percentage breaks ties WITHIN a state, so the fullest critical drive
        # is the one named.
        if ($rank -gt $worstRank -or ($rank -eq $worstRank -and $worst -and $d.FreePercent -lt $worst.FreePercent)) {
            $worst = $d; $worstRank = $rank
        }
    }
    return $worst
}

function Get-DriveState {
    <#
        PERCENTAGE **AND** AN ABSOLUTE FLOOR, either of which can trip it.

        Percentage alone is wrong on very large volumes (10% of 4TB is fine) and
        an absolute floor alone is wrong on small ones (20GB free on a 60GB
        drive is a third of it). Requiring both to be breached would miss the
        case each is there for, so this is an OR.
    #>
    param($Disk)
    if (-not $Disk) { return 'Unknown' }
    if ($Disk.FreePercent -le $script:DiskCritPercent -or $Disk.FreeGB -le $script:DiskCritGB) { return 'Critical' }
    if ($Disk.FreePercent -le $script:DiskWarnPercent -or $Disk.FreeGB -le $script:DiskWarnGB)  { return 'Warning' }
    return 'OK'
}
#endregion

#region Form --------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Server Health'
$form.ClientSize      = New-Object System.Drawing.Size(1280, 860)
$form.FormBorderStyle = 'Sizable'
$form.StartPosition   = 'Manual'
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$btnRun           = New-Object System.Windows.Forms.Button
$btnRun.Text      = 'Run'
$btnRun.Location  = New-Object System.Drawing.Point(14, 14)
$btnRun.Size      = New-Object System.Drawing.Size(120, 30)
$btnRun.FlatStyle = 'Flat'
[void]$form.Controls.Add($btnRun)

# THE SLOW PART, MADE OPTIONAL AND TICKED BY DEFAULT.
#
# The update search is most of the runtime -- ten seconds to two minutes per
# server -- and it is also the answer the tool exists for, so it is on unless
# somebody turns it off. Untick it for a disk-and-patch-date pass that finishes
# in seconds, which is what you want when chasing a full drive.
$chkUpdates          = New-Object System.Windows.Forms.CheckBox
$chkUpdates.Text     = 'Check for missing updates (slow)'
$chkUpdates.Location = New-Object System.Drawing.Point(146, 20)
$chkUpdates.Size     = New-Object System.Drawing.Size(230, 20)
$chkUpdates.Checked  = $true
[void]$form.Controls.Add($chkUpdates)

$lblProg          = New-Object System.Windows.Forms.Label
$lblProg.Text     = 'Idle.'
$lblProg.Location = New-Object System.Drawing.Point(390, 22)
$lblProg.Size     = New-Object System.Drawing.Size(470, 18)
[void]$form.Controls.Add($lblProg)

$pb           = New-Object System.Windows.Forms.ProgressBar
$pb.Location  = New-Object System.Drawing.Point(390, 42)
$pb.Size      = New-Object System.Drawing.Size(470, 8)
$pb.Visible   = $false
[void]$form.Controls.Add($pb)

$btnExport           = New-Object System.Windows.Forms.Button
$btnExport.Text      = 'Export CSV'
$btnExport.Location  = New-Object System.Drawing.Point(1000, 14)
$btnExport.Size      = New-Object System.Drawing.Size(120, 30)
$btnExport.FlatStyle = 'Flat'
$btnExport.Anchor    = 'Top,Right'
$btnExport.Enabled   = $false
[void]$form.Controls.Add($btnExport)

# THE EXECUTIVE DASHBOARD. One page, large type, colour where it matters:
# how many servers are healthy, which ones are not and why, when each was last
# patched, how long each has been up, what is missing. Written as a single
# self-contained HTML file beside the CSVs -- no external scripts, no fonts to
# fetch, nothing that needs a network to render -- and opened in the browser.
# It prints to one landscape page. Asked for 2026-09-16.
$btnDash             = New-Object System.Windows.Forms.Button
$btnDash.Text        = 'Dashboard'
$btnDash.Location    = New-Object System.Drawing.Point(872, 14)
$btnDash.Size        = New-Object System.Drawing.Size(120, 30)
$btnDash.FlatStyle   = 'Flat'
$btnDash.Anchor      = 'Top,Right'
$btnDash.Enabled     = $false
[void]$form.Controls.Add($btnDash)

$btnFolder           = New-Object System.Windows.Forms.Button
$btnFolder.Text      = 'Reports'
$btnFolder.Location  = New-Object System.Drawing.Point(1128, 14)
$btnFolder.Size      = New-Object System.Drawing.Size(100, 30)
$btnFolder.FlatStyle = 'Flat'
$btnFolder.Anchor    = 'Top,Right'
$btnFolder.Add_Click({ Open-ReportFolder -Path $ReportDir })
[void]$form.Controls.Add($btnFolder)

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(14, 60)
$tabs.Size     = New-Object System.Drawing.Size(1252, 744)
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

# TAB 1 IS SPLIT, the same shape Exchange Health Check settled on: the summary
# reads as prose and the overview as a table, and forcing either into the
# other's shape makes both worse. A SplitContainer rather than two docked
# controls, because docking is z-order dependent and put one under the other.
$tpOverview      = New-Object System.Windows.Forms.TabPage
$tpOverview.Text = 'Overview'
[void]$tabs.TabPages.Add($tpOverview)

$split               = New-Object System.Windows.Forms.SplitContainer
$split.Dock          = 'Fill'
$split.Orientation   = 'Horizontal'
$split.SplitterWidth = 6
[void]$tpOverview.Controls.Add($split)

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

$gridDrives   = New-GridTab -Title 'Drives'
$gridUpdates  = New-GridTab -Title 'Updates'
# A one-line note above the grid saying what the dashboard leaves out, so the
# engineer reading this tab is never surprised by a count that differs.
$lblUpdNote           = New-Object System.Windows.Forms.Label
$lblUpdNote.Dock      = 'Top'
$lblUpdNote.AutoSize  = $false
$lblUpdNote.Height    = 34
$lblUpdNote.Padding   = New-Object System.Windows.Forms.Padding(6, 4, 6, 0)
$lblUpdNote.Text      = "Dashboard column = 'Excluded': not shown on the executive dashboard or in its counts (Defender definition updates and $($script:DashboardExcludeKB -join ', ')). Everything is still listed here and in the CSV."
[void]$gridUpdates.Parent.Controls.Add($lblUpdNote)
$lblUpdNote.BringToFront()
$gridExchange = New-GridTab -Title 'Exchange'

$btnClose           = New-Object System.Windows.Forms.Button
$btnClose.Text      = 'Close'
$btnClose.Location  = New-Object System.Drawing.Point(1146, 816)
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

function Set-RowColourByState {
    <#
        Colour from a STATE COLUMN, not by re-deriving the rule.

        Exchange Health Check passes two scriptblocks that re-test the value.
        That means the threshold lives in two places -- the collector and the
        colouring -- and they drift. Here the collector writes 'Critical' /
        'Warning' / 'OK' into a column the operator can also read, and the
        colour just follows the word. One rule, visible in the report.
    #>
    param($Grid, [string]$Column)
    if (-not $Grid.Columns.Contains($Column)) { return }
    foreach ($gr in $Grid.Rows) {
        $v = "$($gr.Cells[$Column].Value)"
        if     ($v -eq 'Critical') { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
        elseif ($v -eq 'Warning')  { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGoldenrod }
        elseif ($v -eq 'Unknown')  { $gr.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray }
    }
}

function Get-OverviewRows {
    $rows = New-Object System.Collections.ArrayList
    foreach ($srv in $script:Servers) {
        $name = $srv.Name
        $snap = $script:Snapshots[$name]

        if (-not $snap) {
            # A SERVER THAT DID NOT ANSWER IS A ROW, NOT A GAP. An absent row
            # for an unreachable server looks exactly like a server with
            # nothing wrong, which is the failure this whole report exists to
            # avoid.
            $why = 'Did not answer'
            $f = @($script:Failed | Where-Object { $_.Server -eq $name })
            if ($f.Count -gt 0) { $why = "$($f[0].Reason)" }
            [void]$rows.Add([pscustomobject]@{
                Server = $name; State = 'Critical'; OS = ''; OSBuild = ''
                WorstDrive = ''; FreeGB = $null; FreePct = $null
                LastPatch = ''; PatchAgeDays = $null; Missing = ''
                Exchange = ''; Reboot = ''; Note = $why
            })
            continue
        }

        $worst = Get-WorstDrive -Disks $snap.Disks
        $diskState = Get-DriveState -Disk $worst

        $patchAge = $null
        if ($snap.LastPatch) { $patchAge = [int]((Get-Date) - $snap.LastPatch).TotalDays }

        # WORST OF THE THREE SIGNALS, so the Overview can be read down one
        # column. Disk state, patch staleness and a failed update search each
        # push the row up; nothing pushes it down.
        $state = $diskState
        if ($state -ne 'Critical') {
            if ($null -ne $patchAge -and $patchAge -gt $script:PatchStaleDays) { $state = 'Warning' }
            if ("$($snap.WuStatus)" -like 'Failed*') { $state = 'Warning' }
            if ($srv.Exchange -and -not $snap.ExchangeBuild) { $state = 'Warning' }
        }

        $missing = switch -Wildcard ("$($snap.WuStatus)") {
            'OK'          { "$(@($snap.MissingUpdates).Count)" }
            'Not checked' { 'not checked' }
            default       { 'unknown' }
        }

        $ex = ''
        if ($snap.ExchangeBuild)  { $ex = Resolve-ExchangeRelease -Build $snap.ExchangeBuild }
        elseif ($srv.Exchange)    { $ex = 'EXPECTED BUT NOT FOUND' }

        [void]$rows.Add([pscustomobject]@{
            Server       = $name
            State        = $state
            OS           = "$($snap.OS)"
            OSBuild      = $(if ($snap.OSDisplay) { "$($snap.OSBuild) ($($snap.OSDisplay))" } else { "$($snap.OSBuild)" })
            WorstDrive   = $(if ($worst) { "$($worst.Drive)" } else { '' })
            FreeGB       = $(if ($worst) { [double]$worst.FreeGB } else { $null })
            FreePct      = $(if ($worst) { [double]$worst.FreePercent } else { $null })
            LastPatch    = $(if ($snap.LastPatch) { $snap.LastPatch.ToString('yyyy-MM-dd') } else { 'unknown' })
            PatchAgeDays = $patchAge
            Missing      = $missing
            Exchange     = $ex
            Reboot       = $(if (@($snap.PendingReboot).Count -gt 0) { (@($snap.PendingReboot) -join ', ') } else { '' })
            Note         = $(
                $n = @()
                if ("$($snap.WuStatus)" -like 'Failed*') { $n += "$($snap.WuStatus)" }
                # Informational only. A queued file rename is NOT a pending
                # reboot -- see the collection block -- but saying nothing
                # about it invites the question of why the Reboot column
                # used to say yes on a server that was fine.
                if ([int]"$($snap.FileRenamePending)" -gt 0) { $n += "$([int]$snap.FileRenamePending) file rename(s) queued (not a reboot)" }
                ($n -join '; ')
            )
        })
    }
    return @($rows)
}

function Get-DriveRows {
    $rows = New-Object System.Collections.ArrayList
    foreach ($srv in $script:Servers) {
        $snap = $script:Snapshots[$srv.Name]
        if (-not $snap) { continue }
        foreach ($d in @($snap.Disks)) {
            [void]$rows.Add([pscustomobject]@{
                Server  = $srv.Name
                State   = (Get-DriveState -Disk $d)
                Drive   = "$($d.Drive)"
                Label   = "$($d.Label)"
                SizeGB  = [double]$d.SizeGB
                FreeGB  = [double]$d.FreeGB
                FreePct = [double]$d.FreePercent
            })
        }
    }
    # Worst first: critical before warning before OK, then emptiest.
    $rank = @{ 'Critical' = 0; 'Warning' = 1; 'OK' = 2; 'Unknown' = 3 }
    return @($rows | Sort-Object @{ Expression = { $rank["$($_.State)"] } }, FreePct)
}

function Get-UpdateRows {
    <#
        One row per MISSING update, plus one summary row per server.

        The summary row is what makes an empty list readable: a server with
        nothing missing produces no update rows at all, and without a line
        saying so it is indistinguishable from a server that was skipped.
    #>
    $rows = New-Object System.Collections.ArrayList
    foreach ($srv in $script:Servers) {
        $snap = $script:Snapshots[$srv.Name]
        if (-not $snap) { continue }

        $status = "$($snap.WuStatus)"
        if ($status -ne 'OK') {
            [void]$rows.Add([pscustomobject]@{
                Server = $srv.Name; State = $(if ($status -eq 'Not checked') { 'Unknown' } else { 'Warning' })
                KB = ''; Severity = ''; Category = ''; Title = $status; Released = ''; AgeDays = $null; Dashboard = ''
                LastPatch = $(if ($snap.LastPatch) { $snap.LastPatch.ToString('yyyy-MM-dd') } else { 'unknown' })
            })
            continue
        }

        $missing = @($snap.MissingUpdates)
        if ($missing.Count -eq 0) {
            [void]$rows.Add([pscustomobject]@{
                Server = $srv.Name; State = 'OK'; KB = ''; Severity = ''; Category = ''
                Title = 'Nothing pending'; Released = ''; AgeDays = $null; Dashboard = ''
                LastPatch = $(if ($snap.LastPatch) { $snap.LastPatch.ToString('yyyy-MM-dd') } else { 'unknown' })
            })
            continue
        }

        foreach ($u in $missing) {
            # Severity is blank on plenty of legitimate updates -- Microsoft
            # only rates security ones. Blank is shown as blank rather than
            # invented as 'Low', which would be a claim nobody made.
            $sev = "$($u.Severity)"
            $rel = $null; $age = $null
            if ($u.PSObject.Properties['Released'] -and $u.Released) {
                try { $rel = [DateTime]$u.Released; $age = [int]((Get-Date) - $rel).TotalDays } catch { }
            }
            [void]$rows.Add([pscustomobject]@{
                Server = $srv.Name
                State  = $(if ($sev -match '(?i)^(critical|important)$') { 'Critical' } else { 'Warning' })
                KB       = "$($u.KB)"
                Severity = $sev
                Category = $(if ($u.PSObject.Properties['Category']) { "$($u.Category)" } else { '' })
                Title    = "$($u.Title)"
                Released = $(if ($rel) { $rel.ToString('yyyy-MM-dd') } else { '' })
                AgeDays  = $age
                Dashboard = $(if (Test-DashboardExcluded -Update $u) { 'Excluded' } else { '' })
                LastPatch = $(if ($snap.LastPatch) { $snap.LastPatch.ToString('yyyy-MM-dd') } else { 'unknown' })
            })
        }
    }
    return @($rows)
}

function Get-ExchangeRows {
    $rows = New-Object System.Collections.ArrayList
    foreach ($srv in $script:Servers) {
        $snap = $script:Snapshots[$srv.Name]
        if (-not $snap) { continue }

        # LISTED WHEN EXCHANGE IS EXPECTED **OR** FOUND, which is not the same
        # as "when the flag is set". A server marked Exchange=false that turns
        # out to have Exchange on it belongs on this tab -- that is a config
        # error worth seeing, and hiding it would make the flag authoritative
        # over the machine.
        if (-not $srv.Exchange -and -not $snap.ExchangeBuild) { continue }

        $state = 'OK'
        $note  = ''
        if (-not $snap.ExchangeBuild) {
            $state = 'Warning'; $note = 'Marked as an Exchange server, but no Exchange install was found'
        } elseif (-not $srv.Exchange) {
            $state = 'Warning'; $note = 'Exchange found on a server not marked as one -- check Tenant.json'
        }

        [void]$rows.Add([pscustomobject]@{
            Server   = $srv.Name
            State    = $state
            Build    = "$($snap.ExchangeBuild)"
            Release  = $(if ($snap.ExchangeBuild) { Resolve-ExchangeRelease -Build $snap.ExchangeBuild } else { 'Not installed' })
            Path     = "$($snap.ExchangePath)"
            Note     = $note
        })
    }
    return @($rows)
}
#endregion

#region Run ---------------------------------------------------------------------
function Update-AllTabs {
    $ov = Get-OverviewRows
    Set-GridRows -Grid $gridOverview `
        -Columns @('Server','State','OS','OSBuild','WorstDrive','FreeGB','FreePct','LastPatch','PatchAgeDays','Missing','Exchange','Reboot','Note') `
        -Rows $ov `
        -NumericColumns @{ FreeGB = [double]; FreePct = [double]; PatchAgeDays = [int] }
    Set-RowColourByState -Grid $gridOverview -Column 'State'

    $dr = Get-DriveRows
    Set-GridRows -Grid $gridDrives `
        -Columns @('Server','State','Drive','Label','SizeGB','FreeGB','FreePct') `
        -Rows $dr `
        -NumericColumns @{ SizeGB = [double]; FreeGB = [double]; FreePct = [double] }
    Set-RowColourByState -Grid $gridDrives -Column 'State'

    $up = Get-UpdateRows
    Set-GridRows -Grid $gridUpdates `
        -Columns @('Server','State','KB','Severity','Category','Title','Released','AgeDays','Dashboard','LastPatch') -Rows $up `
        -NumericColumns @{ AgeDays = [int] }
    Set-RowColourByState -Grid $gridUpdates -Column 'State'
    $exclN = @($up | Where-Object { $_.Dashboard -eq 'Excluded' }).Count
    $lblUpdNote.Text = "Dashboard column = 'Excluded': not shown on the executive dashboard or in its counts (Defender definition updates and $($script:DashboardExcludeKB -join ', ')). This run: $exclN excluded row(s). Everything is still listed here and in the CSV."

    $ex = Get-ExchangeRows
    Set-GridRows -Grid $gridExchange `
        -Columns @('Server','State','Build','Release','Path','Note') -Rows $ex
    Set-RowColourByState -Grid $gridExchange -Column 'State'

    # ---- summary ----------------------------------------------------------
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Run at $($script:RunAt.ToString('yyyy-MM-dd HH:mm:ss'))   --   $(@($script:Servers).Count) server(s) configured")
    [void]$sb.AppendLine("Exchange build names are from a table dated $($script:ExBuildTableAsOf). It names what is installed; it does not say whether anything newer exists.")
    [void]$sb.AppendLine('')

    $problems = New-Object System.Collections.ArrayList
    foreach ($r in $ov) {
        if ($r.State -eq 'OK') { continue }
        $bits = New-Object System.Collections.ArrayList
        if ($r.Note) { [void]$bits.Add("$($r.Note)") }
        if ($r.WorstDrive -and (Get-DriveState -Disk ([pscustomobject]@{ FreeGB = $r.FreeGB; FreePercent = $r.FreePct })) -ne 'OK') {
            [void]$bits.Add("$($r.WorstDrive) has $($r.FreeGB) GB free ($($r.FreePct)%)")
        }
        if ($null -ne $r.PatchAgeDays -and $r.PatchAgeDays -gt $script:PatchStaleDays) {
            [void]$bits.Add("last patched $($r.PatchAgeDays) days ago")
        }
        if ($r.Exchange -eq 'EXPECTED BUT NOT FOUND') { [void]$bits.Add('no Exchange install found') }
        if ($r.Reboot) { [void]$bits.Add("reboot pending ($($r.Reboot))") }
        if ($bits.Count -eq 0) { [void]$bits.Add('see the tabs') }
        [void]$problems.Add(("  {0,-18} {1}" -f $r.Server, (@($bits) -join '; ')))
    }

    if ($problems.Count -eq 0) {
        [void]$sb.AppendLine('Nothing flagged. Every server answered, every drive is above the thresholds, and no patch is overdue.')
    } else {
        [void]$sb.AppendLine("FLAGGED ($($problems.Count)):")
        foreach ($p in $problems) { [void]$sb.AppendLine($p) }
    }

    $totalMissing = 0
    $unknown = 0
    foreach ($srv in $script:Servers) {
        $snap = $script:Snapshots[$srv.Name]
        if (-not $snap) { continue }
        if ("$($snap.WuStatus)" -eq 'OK') { $totalMissing += @($snap.MissingUpdates).Count }
        elseif ("$($snap.WuStatus)" -ne 'Not checked') { $unknown++ }
    }
    [void]$sb.AppendLine('')
    if ($chkUpdates.Checked) {
        [void]$sb.AppendLine("Missing updates: $totalMissing across the servers that answered.")
        if ($unknown -gt 0) {
            [void]$sb.AppendLine("  $unknown server(s) could not be asked -- see the Updates tab for the error on each. NOT counted as zero.")
        }
    } else {
        [void]$sb.AppendLine('Missing updates were NOT checked on this run.')
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Thresholds: drive warning at $($script:DiskWarnPercent)% or $($script:DiskWarnGB) GB, critical at $($script:DiskCritPercent)% or $($script:DiskCritGB) GB; patch stale after $($script:PatchStaleDays) days.")
    $txtSummary.Text = $sb.ToString()
}

function Invoke-ServerHealth {
    $script:Snapshots = @{}
    $script:Failed.Clear()

    if (-not (Test-ServerListPresent)) { return }

    $doUpdates = [bool]$chkUpdates.Checked
    $script:RunAt = Get-Date

    $pb.Maximum = [Math]::Max(1, @($script:Servers).Count)
    $pb.Value   = 0
    $pb.Visible = $true
    $btnRun.Enabled = $false
    $btnExport.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        $i = 0
        foreach ($srv in $script:Servers) {
            $i++
            $lblProg.Text = "Collecting from $($srv.Name) ($i of $(@($script:Servers).Count)) ..."
            # Refresh, not DoEvents. DoEvents pumps the whole message queue and
            # would let Run be pressed again mid-collection; Refresh just
            # repaints, which is all that is wanted here.
            $form.Refresh()
            try {
                $script:Snapshots[$srv.Name] = Get-HealthSnapshot -ComputerName $srv.Name -IncludeUpdates $doUpdates
            } catch {
                [void]$script:Failed.Add([pscustomobject]@{ Server = $srv.Name; Reason = "$($_.Exception.Message)" })
            }
            $pb.Value = [Math]::Min($pb.Maximum, $pb.Value + 1)
        }

        Update-AllTabs
        $answered = @($script:Snapshots.Keys).Count
        $lblProg.Text = "Done. $answered of $(@($script:Servers).Count) server(s) answered."
        $btnExport.Enabled = $true
        $btnDash.Enabled   = $true

        Write-CompletedLog -Action 'Report' -Target 'Server Health' -Result 'Success' `
            -Detail ("$answered of $(@($script:Servers).Count) server(s) answered; updates checked: $doUpdates.") `
            -Environment $Environment
    } catch {
        $lblProg.Text = 'Failed.'
        $txtSummary.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Server health'
        Write-FailedLog -Action 'Report' -Target 'Server Health' -ErrorRecord $_ -Environment $Environment
    } finally {
        $pb.Visible = $false
        $form.Cursor = $prev
        $btnRun.Enabled = $true
    }
}

function Export-ServerHealthCsv {
    <#
        FOUR FILES, ONE PER TAB, INTO THE REPORTS FOLDER.

        Not one wide file: the four tabs have genuinely different shapes -- one
        row per server, per drive, per missing update, per Exchange install --
        and flattening them into a single sheet produces a table that is mostly
        blank cells. Four narrow files each open in Excel and mean something.

        All four carry the SAME timestamp, because they are four views of one
        run and dating them separately would invite comparing a drive list
        against an update list from a different afternoon.
    #>
    if (-not $script:RunAt) { return }
    if (-not (Test-Path -LiteralPath $ReportDir)) {
        New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
    }
    $stamp = $script:RunAt.ToString('yyyy-MM-dd_HHmm')
    $written = New-Object System.Collections.ArrayList

    $sets = @(
        @{ Name = 'ServerHealth-Overview'; Rows = (Get-OverviewRows) }
        @{ Name = 'ServerHealth-Drives';   Rows = (Get-DriveRows) }
        @{ Name = 'ServerHealth-Updates';  Rows = (Get-UpdateRows) }
        @{ Name = 'ServerHealth-Exchange'; Rows = (Get-ExchangeRows) }
    )
    foreach ($s in $sets) {
        $rows = @($s.Rows)
        # NO ZERO-ROW FILES. An empty CSV and a failed run look identical in a
        # folder listing a week later -- the same rule the other reports keep.
        if ($rows.Count -eq 0) { continue }
        $path = Join-Path $ReportDir "$($s.Name)_$stamp.csv"
        $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
        [void]$written.Add($path)
    }

    if ($written.Count -eq 0) {
        $lblProg.Text = 'Nothing to export -- run the check first.'
        return
    }
    $lblProg.Text = "Exported $($written.Count) file(s) to the reports folder."
    Write-CompletedLog -Action 'Report' -Target ($written[0]) -Result 'Success' `
        -Detail ("Server Health export: $($written.Count) file(s), stamp $stamp.") -Environment $Environment
}

function ConvertTo-HtmlText {
    param([string]$Text)
    return [System.Net.WebUtility]::HtmlEncode("$Text")
}

function New-ServerHealthDashboard {
    <#
        ONE PAGE FOR SOMEBODY WHO WILL NOT OPEN THE CSV.

        Top to bottom: the four numbers that answer 'are we all right'; a
        donut of healthy / warning / critical; the servers that need attention
        and the one line that says why; then every server on one row each --
        status, uptime, last patch and its age, missing updates, worst drive
        as a bar and Exchange release. Reboot state is deliberately not here
        -- see the Overview tab.

        SELF-CONTAINED ON PURPOSE. Inline CSS, inline SVG, no script, no web
        fonts, no CDN. It has to render from a network share, from an email
        attachment, and from a laptop with no internet, and it has to look
        the same in a year. The print stylesheet drops the page chrome and
        fits a landscape sheet.

        Colours are the same three the grids use (green / amber / red), read
        off the same Get-OverviewRows, so this page and the Overview tab can
        never disagree about a server.
    #>
    if (-not $script:RunAt) { $lblProg.Text = 'Run the check first.'; return }
    if (-not (Test-Path -LiteralPath $ReportDir)) { New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null }

    $rows  = @(Get-OverviewRows)
    $stamp = $script:RunAt.ToString('yyyy-MM-dd_HHmm')
    $path  = Join-Path $ReportDir "ServerHealth-Dashboard_$stamp.html"

    $tenant = ''
    try { $tenant = "$((Get-TenantConfig).TenantName)".Trim() } catch { }
    if (-not $tenant) { try { $tenant = "$(Get-ActiveTenantLabel)" } catch { } }

    $ok   = @($rows | Where-Object { $_.State -eq 'OK' }).Count
    $warn = @($rows | Where-Object { $_.State -eq 'Warning' }).Count
    $crit = @($rows | Where-Object { $_.State -eq 'Critical' }).Count
    $total = $rows.Count

    # EVERY MISSING UPDATE, BY NAME, WORST FIRST. Severity outranks age:
    # a Critical from last week sits above an Important from last quarter,
    # and within a severity the oldest is first because it has been ignored
    # longest. Blank severity (non-security updates) goes last.
    $upd = @(Get-UpdateRows | Where-Object { $_.KB -or ($_.Title -and $_.State -ne 'OK' -and $_.Title -notmatch '^(Not checked|Failed)') })
    $upd = @($upd | Where-Object { $_.Title -ne 'Nothing pending' -and $_.Dashboard -ne 'Excluded' })
    # The per-server count and the fleet total are recomputed from what is
    # left, so the dashboard never shows a '3' with two lines under it.
    $updCount = @{}
    foreach ($u in $upd) { $updCount["$($u.Server)"] = 1 + $(if ($updCount.ContainsKey("$($u.Server)")) { $updCount["$($u.Server)"] } else { 0 }) }
    $missingTotal = 0
    foreach ($r in $rows) {
        $n = 0
        if ([int]::TryParse("$($r.Missing)", [ref]$n)) {
            $n = $(if ($updCount.ContainsKey("$($r.Server)")) { [int]$updCount["$($r.Server)"] } else { 0 })
            $r.Missing = "$n"
            $missingTotal += $n
        }
    }
    $sevRank = { param($x) switch -Regex ("$($x.Severity)") { '^(?i)critical$' { 0 } '^(?i)important$' { 1 } '^(?i)moderate$' { 2 } '^(?i)low$' { 3 } default { 4 } } }
    $upd = @($upd | Sort-Object @{ e = { & $sevRank $_ } }, @{ e = { if ($null -ne $_.AgeDays) { -[int]$_.AgeDays } else { 0 } } }, Server)
    $critUpd = @($upd | Where-Object { $_.Severity -match '(?i)^critical$' }).Count
    $impUpd  = @($upd | Where-Object { $_.Severity -match '(?i)^important$' }).Count
    $oldestUpd = ($upd | Where-Object { $null -ne $_.AgeDays } | Measure-Object -Property AgeDays -Maximum).Maximum
    $oldestUpdTxt = $(if ($null -ne $oldestUpd) { "$([int]$oldestUpd) d" } else { 'n/a' })
    $oldestPatch = ($rows | Where-Object { $null -ne $_.PatchAgeDays } | Measure-Object -Property PatchAgeDays -Maximum).Maximum
    # NO REBOOT ON THE DASHBOARD. It is on the Overview tab, with the reason,
    # where an engineer can act on it. On a one-page executive summary a
    # reboot count invites the wrong question -- it reads as an outage waiting
    # to happen when it usually means 'a patch finishes at the next reboot
    # window' -- and the underlying signal was noisy enough to flag a whole
    # healthy fleet. Dropped rather than caveated.

    $upDays = @()
    foreach ($srv in $script:Servers) {
        $snap = $script:Snapshots[$srv.Name]
        if ($snap -and $null -ne $snap.UptimeDays) { $upDays += [double]$snap.UptimeDays }
    }
    $longestUp = $(if ($upDays.Count) { ($upDays | Measure-Object -Maximum).Maximum } else { $null })

    # ---- donut: three arcs on a 100-unit circumference ----
    $r0 = 15.9155   # circumference == 100, so a dasharray is a percentage
    $pOk   = $(if ($total) { [Math]::Round(100.0 * $ok   / $total, 2) } else { 0 })
    $pWarn = $(if ($total) { [Math]::Round(100.0 * $warn / $total, 2) } else { 0 })
    $pCrit = $(if ($total) { [Math]::Round(100.0 * $crit / $total, 2) } else { 0 })
    $donut = @"
<svg viewBox="0 0 42 42" class="donut" role="img" aria-label="Fleet health">
  <circle cx="21" cy="21" r="$r0" fill="none" stroke="#e6e9ee" stroke-width="6"></circle>
  <circle cx="21" cy="21" r="$r0" fill="none" stroke="#2e9e5b" stroke-width="6" stroke-dasharray="$pOk $(100-$pOk)" stroke-dashoffset="25"></circle>
  <circle cx="21" cy="21" r="$r0" fill="none" stroke="#e0a100" stroke-width="6" stroke-dasharray="$pWarn $(100-$pWarn)" stroke-dashoffset="$(25-$pOk)"></circle>
  <circle cx="21" cy="21" r="$r0" fill="none" stroke="#d64545" stroke-width="6" stroke-dasharray="$pCrit $(100-$pCrit)" stroke-dashoffset="$(25-$pOk-$pWarn)"></circle>
  <text x="21" y="20.5" class="donut-n">$ok/$total</text>
  <text x="21" y="26" class="donut-l">healthy</text>
</svg>
"@

    # ---- attention list ----
    $attn = New-Object System.Text.StringBuilder
    foreach ($r in @($rows | Where-Object { $_.State -ne 'OK' } | Sort-Object @{ e = { if ($_.State -eq 'Critical') { 0 } else { 1 } } }, Server)) {
        $why = @()
        if ($r.Note) { $why += "$($r.Note)" }
        if ($r.WorstDrive -and $null -ne $r.FreePct) {
            $ds = Get-DriveState -Disk ([pscustomobject]@{ FreeGB = $r.FreeGB; FreePercent = $r.FreePct })
            if ($ds -ne 'OK') { $why += "$($r.WorstDrive) has $($r.FreeGB) GB free ($($r.FreePct)%)" }
        }
        if ($null -ne $r.PatchAgeDays -and $r.PatchAgeDays -gt $script:PatchStaleDays) { $why += "last patched $($r.PatchAgeDays) days ago" }
        $mn = 0
        if ([int]::TryParse("$($r.Missing)", [ref]$mn) -and $mn -gt 0) {
            $mine = @($upd | Where-Object { $_.Server -eq $r.Server })
            $mc = @($mine | Where-Object { $_.Severity -match '(?i)^critical$' }).Count
            $mi = @($mine | Where-Object { $_.Severity -match '(?i)^important$' }).Count
            $mo = ($mine | Where-Object { $null -ne $_.AgeDays } | Measure-Object -Property AgeDays -Maximum).Maximum
            $txt = "$mn update(s) missing"
            $bits = @(); if ($mc) { $bits += "$mc critical" }; if ($mi) { $bits += "$mi important" }
            if ($bits.Count) { $txt += " ($($bits -join ', '))" }
            if ($null -ne $mo) { $txt += ", oldest $([int]$mo) days" }
            $why += $txt
        }
        if ($r.Exchange -eq 'EXPECTED BUT NOT FOUND') { $why += 'Exchange expected but not found' }
        if ($why.Count -eq 0) { $why += 'see the Overview tab' }
        $cls = $(if ($r.State -eq 'Critical') { 'crit' } else { 'warn' })
        [void]$attn.AppendLine("<li class=`"$cls`"><b>$(ConvertTo-HtmlText $r.Server)</b> &mdash; $(ConvertTo-HtmlText ($why -join '; '))</li>")
    }
    if ($attn.Length -eq 0) { [void]$attn.AppendLine('<li class="ok"><b>Every server is healthy.</b> Nothing needs attention.</li>') }

    # ---- table ----
    $tbl = New-Object System.Text.StringBuilder
    foreach ($r in @($rows | Sort-Object Server)) {
        $snap = $script:Snapshots[$r.Server]
        $up = $(if ($snap -and $null -ne $snap.UptimeDays) { "$([Math]::Round([double]$snap.UptimeDays, 0)) d" } else { '&ndash;' })
        $patch = $(if ($r.LastPatch -and $r.LastPatch -ne 'unknown') { "$($r.LastPatch)" } else { 'unknown' })
        if ($null -ne $r.PatchAgeDays) { $patch += " <span class=`"muted`">($($r.PatchAgeDays) d)</span>" }
        $patchCls = $(if ($null -ne $r.PatchAgeDays -and $r.PatchAgeDays -gt $script:PatchStaleDays) { 'warn-t' } else { '' })
        $bar = '&ndash;'
        if ($r.WorstDrive -and $null -ne $r.FreePct) {
            $ds  = Get-DriveState -Disk ([pscustomobject]@{ FreeGB = $r.FreeGB; FreePercent = $r.FreePct })
            $bc  = $(if ($ds -eq 'Critical') { '#d64545' } elseif ($ds -eq 'Warning') { '#e0a100' } else { '#2e9e5b' })
            $pct = [Math]::Max(0, [Math]::Min(100, [double]$r.FreePct))
            $bar = "<div class=`"bar`"><div class=`"fill`" style=`"width:$pct%;background:$bc`"></div></div><span class=`"bar-l`">$(ConvertTo-HtmlText $r.WorstDrive) $($r.FreeGB) GB free ($($r.FreePct)%)</span>"
        }
        $mn = 0; $miss = "$($r.Missing)"
        $missCls = $(if ([int]::TryParse($miss, [ref]$mn) -and $mn -gt 0) { 'warn-t' } elseif ($miss -eq '0') { 'ok-t' } else { 'muted' })
        $pill = "<span class=`"pill $($r.State.ToLower())`">$($r.State)</span>"
        $ex = $(if ($r.Exchange) { ConvertTo-HtmlText $r.Exchange } else { '&ndash;' })
        [void]$tbl.AppendLine("<tr><td class=`"srv`">$(ConvertTo-HtmlText $r.Server)</td><td>$pill</td><td class=`"muted`">$(ConvertTo-HtmlText $r.OS)</td><td>$up</td><td class=`"$patchCls`">$patch</td><td class=`"$missCls`">$(ConvertTo-HtmlText $miss)</td><td class=`"barcell`">$bar</td><td>$ex</td></tr>")
    }

    $updTbl = New-Object System.Text.StringBuilder
    foreach ($u in $upd) {
        $sevTxt = $(if ("$($u.Severity)") { "$($u.Severity)" } else { $(if ("$($u.Category)") { "$($u.Category)" } else { '&ndash;' }) })
        $sevCls = switch -Regex ("$($u.Severity)") { '^(?i)critical$' { 'critical' } '^(?i)important$' { 'warning' } default { 'muted-pill' } }
        $ageTxt = $(if ($null -ne $u.AgeDays) { "$($u.AgeDays) d" } else { '&ndash;' })
        $ageCls = $(if ($null -ne $u.AgeDays -and [int]$u.AgeDays -gt $script:PatchStaleDays) { 'warn-t' } else { '' })
        [void]$updTbl.AppendLine("<tr><td class=`"srv`">$(ConvertTo-HtmlText $u.Server)</td><td><span class=`"pill $sevCls`">$(ConvertTo-HtmlText $sevTxt)</span></td><td>$(ConvertTo-HtmlText $u.KB)</td><td>$(ConvertTo-HtmlText $u.Title)</td><td>$(ConvertTo-HtmlText $u.Released)</td><td class=`"$ageCls`">$ageTxt</td></tr>")
    }
    $updSection = ''
    if ($upd.Count -gt 0) {
        $updSection = @"
<h2>Missing updates &middot; $($upd.Count) across the fleet &middot; $critUpd critical, $impUpd important &middot; oldest $oldestUpdTxt</h2>
<table class="upd">
<thead><tr><th>Server</th><th>Severity</th><th>KB</th><th>Update</th><th>Released</th><th>Age</th></tr></thead>
<tbody>
$($updTbl.ToString())</tbody>
</table>
"@
    }

    $oldestTxt = $(if ($null -ne $oldestPatch) { "$([int]$oldestPatch) days" } else { 'n/a' })
    $upTxt     = $(if ($null -ne $longestUp)   { "$([int]$longestUp) days" }   else { 'n/a' })
    $crClass   = $(if ($crit -gt 0) { 'crit' } elseif ($warn -gt 0) { 'warn' } else { 'ok' })
    $headline  = $(if ($crit -gt 0) { "$crit server(s) need action now" } elseif ($warn -gt 0) { "$warn server(s) need attention" } else { 'All servers healthy' })
    $updated   = $script:RunAt.ToString('dddd d MMMM yyyy, HH:mm')

    $html = @"
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>Server Health - $(ConvertTo-HtmlText $tenant) - $stamp</title>
<style>
  :root { --ok:#2e9e5b; --warn:#e0a100; --crit:#d64545; --ink:#1f2933; --muted:#6b7280; --line:#e6e9ee; --card:#ffffff; --bg:#f4f6f9; }
  * { box-sizing:border-box; }
  body { margin:0; padding:28px 36px; background:var(--bg); color:var(--ink); font:16px/1.4 "Segoe UI", Arial, sans-serif; }
  h1 { margin:0; font-size:34px; font-weight:600; letter-spacing:-0.2px; }
  .sub { color:var(--muted); font-size:16px; margin-top:4px; }
  .head { display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:22px; }
  .headline { font-size:22px; font-weight:600; padding:10px 16px; border-radius:8px; color:#fff; }
  .headline.ok { background:var(--ok); } .headline.warn { background:var(--warn); } .headline.crit { background:var(--crit); }
  .grid { display:grid; grid-template-columns: 220px repeat(4, 1fr); gap:16px; margin-bottom:22px; }
  .card { background:var(--card); border:1px solid var(--line); border-radius:10px; padding:16px 18px; }
  .kpi .n { font-size:40px; font-weight:700; line-height:1; }
  .kpi .l { color:var(--muted); font-size:15px; margin-top:6px; }
  .kpi .l2 { display:block; margin-top:3px; }
  .kpi.ok .n { color:var(--ok); } .kpi.warn .n { color:var(--warn); } .kpi.crit .n { color:var(--crit); }
  .donut { width:170px; height:170px; display:block; margin:0 auto; }
  .donut-n { font-size:7.5px; font-weight:700; text-anchor:middle; fill:var(--ink); }
  .donut-l { font-size:3.2px; text-anchor:middle; fill:var(--muted); }
  h2 { font-size:19px; margin:0 0 10px; font-weight:600; }
  ul.attn { list-style:none; margin:0 0 22px; padding:0; }
  ul.attn li { padding:10px 14px; border-left:6px solid var(--line); background:var(--card); border-radius:6px; margin-bottom:6px; font-size:16px; }
  ul.attn li.crit { border-left-color:var(--crit); } ul.attn li.warn { border-left-color:var(--warn); } ul.attn li.ok { border-left-color:var(--ok); }
  table { width:100%; border-collapse:collapse; background:var(--card); border:1px solid var(--line); border-radius:10px; overflow:hidden; font-size:15px; }
  th { text-align:left; font-weight:600; color:var(--muted); font-size:13px; text-transform:uppercase; letter-spacing:.04em; padding:10px 12px; border-bottom:1px solid var(--line); background:#fafbfc; }
  td { padding:9px 12px; border-bottom:1px solid var(--line); vertical-align:middle; }
  tr:last-child td { border-bottom:0; }
  td.srv { font-weight:600; }
  .pill { display:inline-block; padding:3px 10px; border-radius:999px; color:#fff; font-size:13px; font-weight:600; }
  .pill.ok { background:var(--ok); } .pill.warning { background:var(--warn); } .pill.critical { background:var(--crit); } .pill.muted-pill { background:#8a94a6; }
  table.upd { margin-bottom:22px; } table.upd td:nth-child(4) { max-width:520px; }
  .muted { color:var(--muted); } .warn-t { color:#b07800; font-weight:600; } .ok-t { color:var(--ok); font-weight:600; }
  .barcell { min-width:220px; }
  .bar { height:10px; background:var(--line); border-radius:5px; overflow:hidden; }
  .fill { height:100%; }
  .bar-l { display:block; font-size:12.5px; color:var(--muted); margin-top:3px; }
  .foot { color:var(--muted); font-size:13px; margin-top:16px; }
  @media print {
    @page { size: landscape; margin: 10mm; }
    body { padding:0; background:#fff; font-size:13px; }
    h1 { font-size:26px; } .headline { font-size:17px; } .kpi .n { font-size:30px; }
    .grid { gap:10px; margin-bottom:12px; } .card { padding:10px 12px; }
    .donut { width:120px; height:120px; }
    table { font-size:12px; } td, th { padding:5px 8px; }
    ul.attn li { padding:6px 10px; font-size:13px; }
    .foot { margin-top:8px; }
    h2 { page-break-after:avoid; } table { page-break-inside:auto; } tr { page-break-inside:avoid; }
  }
</style></head>
<body>
<div class="head">
  <div><h1>Server Health</h1><div class="sub">$(ConvertTo-HtmlText $tenant) &middot; $total server(s) &middot; checked $updated</div></div>
  <div class="headline $crClass">$headline</div>
</div>
<div class="grid">
  <div class="card">$donut</div>
  <div class="card kpi $(if ($crit) { 'crit' } else { 'ok' })"><div class="n">$crit</div><div class="l">Critical</div></div>
  <div class="card kpi $(if ($warn) { 'warn' } else { 'ok' })"><div class="n">$warn</div><div class="l">Warning</div></div>
  <div class="card kpi $(if ($critUpd) { 'crit' } elseif ($missingTotal) { 'warn' } else { 'ok' })"><div class="n">$missingTotal</div><div class="l">Updates missing &middot; $critUpd critical, $impUpd important &middot; oldest $oldestUpdTxt</div></div>
  <div class="card kpi $(if ($null -ne $oldestPatch -and $oldestPatch -gt $script:PatchStaleDays) { 'warn' } else { 'ok' })"><div class="n">$oldestTxt</div><div class="l">Oldest patch age<span class="l2">Longest uptime $upTxt</span></div></div>
</div>
<h2>Needs attention</h2>
<ul class="attn">
$($attn.ToString())</ul>
$updSection
<h2>Every server</h2>
<table>
<thead><tr><th>Server</th><th>Status</th><th>OS</th><th>Uptime</th><th>Last patch</th><th>Missing</th><th>Worst drive (free)</th><th>Exchange</th></tr></thead>
<tbody>
$($tbl.ToString())</tbody>
</table>
<div class="foot">Thresholds: drive warning at $($script:DiskWarnPercent)% or $($script:DiskWarnGB) GB free, critical at $($script:DiskCritPercent)% or $($script:DiskCritGB) GB; patches stale after $($script:PatchStaleDays) days. Generated by the Admin Console Server Health tool. Four CSVs with the full detail carry the same timestamp in the reports folder.</div>
</body></html>
"@

    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($path, ($html -replace "`r?`n", "`r`n"), $enc)
    $lblProg.Text = "Dashboard written: $(Split-Path -Leaf $path)"
    Write-CompletedLog -Action 'Report' -Target $path -Result 'Success' `
        -Detail "Server Health executive dashboard: $total server(s), $crit critical, $warn warning." -Environment $Environment
    try { Start-Process -FilePath $path } catch { $lblProg.Text += ' (could not open the browser; the file is in the reports folder)' }
}

$btnRun.Add_Click({
    try { Invoke-ServerHealth }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Run the server health check' }
})
$btnExport.Add_Click({
    try { Export-ServerHealthCsv }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Export the server health CSVs' }
})
$btnDash.Add_Click({
    try { New-ServerHealthDashboard }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Build the executive dashboard' }
})
#endregion

#region Show --------------------------------------------------------------------
$lblProg.Text = "$(@($script:Servers).Count) server(s) configured. Press Run."
if (@($script:Servers).Count -eq 0) {
    $lblProg.Text = 'No servers configured -- set ServerHealth.Servers in Tenant.json.'
}

Set-FormScreenCorner -Form $form -Corner 'TopLeft'
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Server Health' | Out-Null

[void]$form.ShowDialog()
#endregion
