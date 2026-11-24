#Requires -Version 5.1
<#
    Connect-ExO.ps1
    Connects this PowerShell process to Exchange Online.

    Reasoning: ScriptNotes/Connect-Scripts.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Connects this console to Exchange Online.
#
# The console closes, the sign-in happens, and the console reopens with the
# connection live. That is deliberate -- a sign-in window cannot accept typing
# while the console's own window is open in front of it, so the console gets out
# of the way.
#
# Complete the sign-in and MFA prompt when it appears; the console comes back on
# its own.
# TOOLKIT-HELP-END

param(
    [string]$UserPrincipalName,
    [switch]$Force
)

. (Join-Path $PSScriptRoot 'Common.ps1')

# No WinForms. This script must NOT open a window of its own -- the sign-in
# prompt is the window, and anything of ours competing with it is the problem
# this script exists to avoid.

function Write-Line {
    param([string]$Text, [string]$Colour)
    if ($Colour) { Write-Host $Text -ForegroundColor $Colour } else { Write-Host $Text }
}

Write-Line ''
Write-Line '=== Connect to Exchange Online ===' 'Cyan'

$before = Test-ExoConnectionState
if ($before.Connected -and -not $Force) {
    Write-Line "Already connected. $($before.Detail)" 'Green'
    Write-Line 'Nothing to do. Re-run with -Force to reconnect.'
    Write-Line ''
    return
}
if (-not $before.Connected) { Write-Line $before.Detail 'DarkGray' }

if (-not (Get-Command -Name 'Connect-ExchangeOnline' -ErrorAction SilentlyContinue)) {
    Write-Line 'ExchangeOnlineManagement is not installed in this process.' 'Red'
    Write-Line 'Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    Write-Line ''
    Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' `
        -Detail 'ExchangeOnlineManagement module not available'
    return
}

$params = @{ ShowBanner = $false; ErrorAction = 'Stop' }

# -DisableWAM is the point of this script.
#
# From version 3.7.0 the module signs in through Web Account Manager, which
# raises its own broker window. When the Admin Console launched a connect
# script from a button, that window appeared while $form.ShowDialog() held the
# thread in a modal message loop -- and a modal form disables other windows on
# the same thread. The prompt drew and then never accepted input, which reads
# as "the MFA page never finished loading".
#
# Microsoft document the switch for exactly this class of problem: "If you
# encounter WAM-related issues during sign in, you can use this switch to
# disable WAM." Without WAM the sign-in goes to a browser instead, which is a
# separate process and unaffected by our message loop.
#
# Version-guarded because the parameter only exists from 3.7.2-Preview1. Passing
# it to an older module is a binding error, and "your connect script is broken"
# is a worse outcome than a prompt that needs a second click.
if ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) {
    $params['DisableWAM'] = $true
    Write-Line 'Using -DisableWAM (sign-in goes to the browser, not the WAM broker).' 'DarkGray'
} else {
    Write-Line 'This module version has no -DisableWAM; using the default sign-in path.' 'DarkGray'
}

# Skips the username page. Not cosmetic on a shared jump box, where the broker
# may remember somebody else's account.
if ($UserPrincipalName) { $params['UserPrincipalName'] = $UserPrincipalName }

Write-Line ''
Write-Line 'A sign-in window will open. Complete MFA there, then return here.' 'Yellow'
Write-Line ''

try {
    Connect-ExchangeOnline @params
} catch {
    Write-Line ''
    Write-Line "Connection failed: $($_.Exception.Message)" 'Red'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' `
        -Detail "$($_.Exception.Message)"
    Write-Line ''
    return
}

# Prove it rather than assume it. Connect-ExchangeOnline returning without an
# error is not the same as a usable session, and the console's buttons are
# about to be re-coloured from this same probe -- so the message here and the
# button state cannot disagree.
$after = Test-ExoConnectionState
Write-Line ''
if ($after.Connected) {
    Write-Line "Connected. $($after.Detail)" 'Green'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Success' `
        -Detail "$($after.Detail)"
} else {
    Write-Line "Sign-in completed but no session is present: $($after.Detail)" 'Red'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' `
        -Detail "$($after.Detail)"
}

# The tenant is cached from the first thing that asked. Connecting is exactly
# the event that can change the answer, so the cache is dropped here rather
# than left to go stale behind a correct-looking label.
Clear-ConnectedTenantDomainCache
$mismatch = Test-TenantDomainMismatch
if ($mismatch) {
    Write-Line ''
    Write-Line $mismatch 'Yellow'
}
Write-Line ''
