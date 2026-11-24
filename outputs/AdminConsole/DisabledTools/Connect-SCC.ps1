#Requires -Version 5.1
<#
    Connect-SCC.ps1
    Connects this PowerShell process to Security and Compliance.

    Modelled on Connect-ExO.ps1, deliberately and closely. Everything that
    script learned about interactive sign-in from a console button applies
    here unchanged -- the sign-in window does not care which service is on the
    other end of it.

    Author   : Glenn Armbruster
    Revised  : 2026-08-10
#>
# TOOLKIT-HELP-BEGIN
# Connects this console to Security and Compliance.
#
# The console closes, the sign-in happens, and the console reopens with the
# connection live. That is deliberate -- a sign-in window cannot accept typing
# while the console's own window is open in front of it, so the console gets out
# of the way.
#
# Only Search and Purge needs this. It is a separate connection from Exchange
# Online and is not made automatically, because most days nothing wants it.
#
# Complete the sign-in and MFA prompt when it appears; the console comes back on
# its own.
# TOOLKIT-HELP-END

param(
    [string]$UserPrincipalName,
    [switch]$Force
)

. (Join-Path $PSScriptRoot 'Common.ps1')

# No WinForms. Same rule as Connect-ExO.ps1: the sign-in prompt IS the window,
# and anything of ours competing with it is the problem this script exists to
# avoid.

function Write-Line {
    param([string]$Text, [string]$Colour)
    if ($Colour) { Write-Host $Text -ForegroundColor $Colour } else { Write-Host $Text }
}

Write-Line ''
Write-Line '=== Connect to Security and Compliance ===' 'Cyan'

if (-not (Get-Command -Name 'Test-SccConnectionState' -ErrorAction SilentlyContinue)) {
    Write-Line 'This script needs a newer Common.ps1 than the one in this folder.' 'Red'
    Write-Line 'Missing: Test-SccConnectionState' 'Red'
    Write-Line ''
    return
}

$before = Test-SccConnectionState
if ($before.Connected -and -not $Force) {
    Write-Line "Already connected. $($before.Detail)" 'Green'
    Write-Line 'Nothing to do. Re-run with -Force to reconnect.'
    Write-Line ''
    return
}
if (-not $before.Connected) { Write-Line $before.Detail 'DarkGray' }

if (-not (Get-Command -Name 'Connect-IPPSSession' -ErrorAction SilentlyContinue)) {
    Write-Line 'ExchangeOnlineManagement is not installed in this process.' 'Red'
    Write-Line 'Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    Write-Line ''
    Write-CompletedLog -Action 'Connect' -Target 'Security and Compliance' -Result 'Failed' `
        -Detail 'ExchangeOnlineManagement module not available'
    return
}

# The version floor is specific to purge, and it is checked HERE rather than in
# the tool because this is where it can still be acted on. Finding out at purge
# time means the search has already run.
#
# -EnableSearchOnlySession arrived in 3.9.0-Preview1 and purge requires 3.9.0 or
# later. An older module connects perfectly well and then refuses the purge.
$minVersion = [version]'3.9.0'
$installed  = $null
try {
    $m = Get-Module -Name 'ExchangeOnlineManagement' -ListAvailable -ErrorAction Stop |
         Sort-Object Version -Descending | Select-Object -First 1
    if ($m) { $installed = [version]$m.Version }
} catch { }

$params = @{ ShowBanner = $false; ErrorAction = 'Stop' }

if ((Get-Command Connect-IPPSSession).Parameters.ContainsKey('EnableSearchOnlySession')) {
    $params['EnableSearchOnlySession'] = $true
    Write-Line 'Using -EnableSearchOnlySession (required for purge).' 'DarkGray'
} else {
    Write-Line ''
    Write-Line "This module version has no -EnableSearchOnlySession. Searching will work; PURGE WILL NOT." 'Yellow'
    if ($installed) { Write-Line "Installed: $installed. Purge needs $minVersion or later." 'Yellow' }
    Write-Line 'Update-Module ExchangeOnlineManagement' 'Yellow'
    Write-Line ''
}

# Same reasoning as Connect-ExO.ps1: from 3.7.0 the module signs in through Web
# Account Manager, whose broker window is disabled by a modal message loop on
# the same thread. Without WAM the sign-in goes to a browser, a separate process
# our message loop cannot interfere with.
if ((Get-Command Connect-IPPSSession).Parameters.ContainsKey('DisableWAM')) {
    $params['DisableWAM'] = $true
    Write-Line 'Using -DisableWAM (sign-in goes to the browser, not the WAM broker).' 'DarkGray'
}

if ($UserPrincipalName) { $params['UserPrincipalName'] = $UserPrincipalName }

# The sign-in window is an ActiveX WebBrowser control and can only be created on
# an STA thread that is not carrying somebody else's synchronization context.
# ShowDialog installs one and does not remove it when the form closes, so the
# leftover context is cleared before the sign-in rather than after it fails.
$apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
$staleCtx  = [System.Threading.SynchronizationContext]::Current
if ($staleCtx) {
    Write-Line "Clearing a leftover $($staleCtx.GetType().Name) before sign-in." 'DarkGray'
    [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
}
if ($apartment -ne [System.Threading.ApartmentState]::STA) {
    Write-Line ''
    Write-Line "This PowerShell thread is $apartment, not STA. The sign-in window cannot open here." 'Red'
    Write-Line 'Start the console from powershell.exe (STA by default) rather than with -MTA, or from pwsh with -STA.' 'Yellow'
    Write-CompletedLog -Action 'Connect' -Target 'Security and Compliance' -Result 'Failed' `
        -Detail "Thread apartment is $apartment; interactive sign-in requires STA."
    Write-Line ''
    return
}

Write-Line ''
Write-Line 'A sign-in window will open. Complete MFA there, then return here.' 'Yellow'
Write-Line ''

# ErrorAction is already in $params. Passing it again as well is a binding
# error that fails before the module is called -- the trap that broke every
# attempt in Connect-ExO.ps1 until it was found.
#
# 2>$null suppresses the error STREAM, not the exception: the module writes
# pages of MSAL plumbing to it before the real terminating error arrives.
try {
    Connect-IPPSSession @params 2>$null
} catch {
    $msg = "$($_.Exception.Message)"

    if ($msg -match 'single-threaded apartment|8856f961') {
        Write-Line ''
        Write-Line 'The sign-in window could not be created on this thread. Retrying once with a clean context.' 'Yellow'
        [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
        try {
            Connect-IPPSSession @params 2>$null
            $msg = $null
        } catch {
            $msg = "$($_.Exception.Message)"
        }
    }

    if ($msg) {
        Write-Line ''
        Write-Line "Connection failed: $msg" 'Red'
        Write-CompletedLog -Action 'Connect' -Target 'Security and Compliance' -Result 'Failed' -Detail $msg
        Write-Line ''
        return
    }
}

# The module cache is keyed on what is loaded, and connecting is exactly the
# event that changes it. Refresh rather than trust a lookup made before there
# was anything to find.
$after = Test-SccConnectionState -Refresh
Write-Line ''
if ($after.Connected) {
    Write-Line "Connected. $($after.Detail)" 'Green'
    Write-CompletedLog -Action 'Connect' -Target 'Security and Compliance' -Result 'Success' `
        -Detail "$($after.Detail)"
} else {
    Write-Line "Sign-in completed but no usable session is present: $($after.Detail)" 'Red'
    Write-CompletedLog -Action 'Connect' -Target 'Security and Compliance' -Result 'Failed' `
        -Detail "$($after.Detail)"
}
Write-Line ''
