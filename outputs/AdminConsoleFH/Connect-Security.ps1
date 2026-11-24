#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    Connect-Security.ps1
    Connects this PowerShell process to Security and Compliance (IPPSSession).

    Same shape as Connect-SPO.ps1 and Connect-ExO.ps1.
#>
# TOOLKIT-HELP-BEGIN
# Connects this console to Security and Compliance (IPPSSession).
#
# This is the endpoint behind content search, eDiscovery, retention labels and
# DLP policies.
#
# The console closes, the sign-in happens, and the console reopens with the
# connection live. That is deliberate -- a sign-in window cannot accept typing
# while the console's own window is in front of it.
#
# IT IS NOT THE SAME AS EXCHANGE ONLINE, even though both come from the
# ExchangeOnlineManagement module and both may use the same account. They are
# separate sessions to separate endpoints, and the header reports them
# separately. Being connected to ExO tells you nothing about this one.
#
# BOTH SESSIONS IN ONE PROCESS SHARE CMDLET NAMES. Where a name exists in both
# endpoints, the one that answers is whichever module state won last. If a
# compliance cmdlet starts returning Exchange-shaped results, that is why --
# reconnect this session.
# TOOLKIT-HELP-END
# TOOLKIT-REQUIRES: ConsoleClosed
#
# B2: this script raises a sign-in window, and a modal ShowDialog() disables
# every other window created on the same thread -- the prompt would draw and
# then ignore all input. The console reads this marker, closes its form, runs
# this script with the message loop gone, and re-opens itself afterwards.

param(
    [switch]$Force
)

. (Join-Path $PSScriptRoot 'Common.ps1')

# FIRST THING AFTER Common LOADS.
if (-not (Test-ConsoleLaunch)) { return }

# No WinForms. The sign-in prompt is the window.

function Write-Line {
    param([string]$Text, [string]$Colour)
    if ($Colour) { Write-Host $Text -ForegroundColor $Colour } else { Write-Host $Text }
}

# Test-IppsConnectionState is NOT defined here -- Common.ps1 has it. A local
# copy would win for this file only, so a fix in Common would never reach it.

$before = Test-IppsConnectionState
if ($before.Connected -and -not $Force) {
    Write-Line "Already connected. $($before.Detail)" 'Green'
    Write-Line 'Nothing to do. Re-run with -Force to reconnect.'
    return
}

if (-not (Get-Command -Name 'Connect-IPPSSession' -ErrorAction SilentlyContinue)) {
    # A MISSING MODULE AND A MISSING SIGN-IN ARE DIFFERENT PROBLEMS, and
    # reporting the first as the second sends the operator to click Connect
    # forever.
    Write-Line 'The ExchangeOnlineManagement module is NOT INSTALLED in this window.' 'Red'
    Write-Line 'This is not a sign-in problem and clicking Connect again will not fix it. Install it with:'
    Write-Line '    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser'
    Write-FailedLog -Action 'Connect' -Target 'Security and Compliance' `
        -Detail 'Refused: ExchangeOnlineManagement (Connect-IPPSSession) is not available in this process.'
    return
}

Write-Line 'A sign-in window will appear. Complete it and the console comes back on its own.'
Write-Line ''

# THE eDISCOVERY CMDLETS NEED -EnableSearchOnlySession, AND THIS SCRIPT DID NOT
# PASS IT.
#
# Microsoft made the switch mandatory (MC1131771, enforced from late September
# 2025) for the *-ComplianceSearch family and New-ComplianceSearchAction, with
# module version 3.9.0 as the floor. What the switch changes is the AUDIENCE OF
# THE ACCESS TOKEN, so it is decided when the session is created and cannot be
# added afterwards.
#
# THAT MATTERS MORE THAN IT LOOKS, because a tool launched from this console
# runs IN THIS PROCESS and uses the session this button made. Connecting with
# the switch in a separate PowerShell window does nothing for Content Search
# and Purge -- it never sees that session.
#
# Enforcement is not uniform: Get-ComplianceSearch keeps answering, so the
# Security indicator goes green and searches create and run, and then
# New-ComplianceSearchAction refuses. A preview that fails while every
# indicator is green is the shape this produces.
#
# PASSED ONLY IF THE MODULE HAS THE PARAMETER. Below 3.9.0 the switch does not
# exist and naming it fails the call outright -- which would turn a working
# connection into no connection at all for anyone who has not updated.
$searchOnly = @{}
$ippsCmd = Get-Command -Name 'Connect-IPPSSession' -ErrorAction SilentlyContinue
$modVer = ''
if ($ippsCmd) {
    $modVer = "$($ippsCmd.Module.Version)"
    if ($ippsCmd.Parameters.ContainsKey('EnableSearchOnlySession')) {
        $searchOnly['EnableSearchOnlySession'] = $true
        Write-Line "ExchangeOnlineManagement $modVer -- connecting with -EnableSearchOnlySession, which the eDiscovery cmdlets require."
    } else {
        Write-Line "ExchangeOnlineManagement $modVer does NOT support -EnableSearchOnlySession." 'Yellow'
        Write-Line 'Signing in will still work, and so will most of this endpoint, but the eDiscovery cmdlets -- Content Search preview, export and purge -- will be refused, usually as a 403 with nothing to explain it. Version 3.9.0 or later is the floor:' 'Yellow'
        Write-Line '    Update-Module -Name ExchangeOnlineManagement'
        Write-Line ''
    }
}

try {
    Connect-IPPSSession @searchOnly -ErrorAction Stop | Out-Null

    # READ BACK, and read back with a cmdlet that exists ONLY in this endpoint.
    # Probing with a name Exchange Online also provides would report success
    # whenever ExO happened to be connected in the same process.
    $after = Test-IppsConnectionState
    if ($after.Connected) {
        Write-Line ''
        Write-Line "Connected. $($after.Detail)" 'Green'
        Write-CompletedLog -Action 'Connect' -Target 'Security and Compliance' -Result 'Success' `
            -Detail 'Connected to Security and Compliance (IPPSSession).'
    } else {
        Write-Line ''
        Write-Line 'Connect-IPPSSession returned without error, but no session answered afterwards.' 'Yellow'
        Write-Line $after.Detail
        Write-CompletedLog -Action 'Connect' -Target 'Security and Compliance' -Result 'Partial' `
            -Detail 'Connect-IPPSSession did not throw but Get-ComplianceSearch found no session.'
    }
} catch {
    Write-Line ''
    Write-Line (Format-ErrorDetail -ErrorRecord $_ -Context 'Connect-IPPSSession') 'Red'
    Write-Line ''
    Write-Line 'If the sign-in completed but this still failed, there are two usual causes:'
    Write-Line '  - the account lacks an eDiscovery or Compliance role, so the session opens and then cannot read anything; or'
    Write-Line '  - the rights are PIM-eligible and were activated AFTER connecting, which leaves this token without them. Activate first, then connect again.'
    Write-Line 'App-only and certificate-based authentication are NOT supported for the eDiscovery cmdlets at all.'
    Write-FailedLog -Action 'Connect' -Target 'Security and Compliance' `
        -Detail 'Connect-IPPSSession failed.' -ErrorRecord $_
}
