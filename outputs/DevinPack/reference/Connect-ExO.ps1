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

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

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

# MODULE FIRST, SIGN-IN SECOND.
#
# This used to print the Install-Module line and stop, which is correct and
# useless: the operators these guardrails exist for read "not installed" and
# stop too. Three states, three different answers -- and only one of them
# involves downloading anything.
$mod = Test-ToolkitModule -Id 'ExO'

if ($mod.State -eq 'Installed') {
    # On disk, just not in this process. Free to fix and not worth asking about.
    Write-Line $mod.Detail 'DarkGray'
    $imp = Import-ToolkitModule -Id 'ExO'
    Write-Line $imp.Detail $(if ($imp.Ok) { 'Green' } else { 'Red' })
    if (-not $imp.Ok) {
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' -Detail $imp.Detail
        return
    }
    $mod = Test-ToolkitModule -Id 'ExO'
}

if ($mod.State -ne 'Imported') {
    Write-Line "$($mod.Name) is not installed for your account." 'Yellow'

    # ASKED, NEVER ASSUMED. A tool that quietly installs software is a
    # different thing from a tool that offers to.
    if (-not (Request-ToolkitModuleInstall -Id 'ExO')) {
        Write-Line 'Not installed -- you said no. Nothing was changed.' 'DarkGray'
        Write-Line "To do it yourself later:  Install-Module $($mod.Name) -Scope CurrentUser"
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' `
            -Detail "$($mod.Name) not installed; operator declined the install."
        return
    }

    Write-Line 'Installing. This can take a minute or two the first time.' 'Cyan'
    $ins = Install-ToolkitModule -Id 'ExO'
    foreach ($s in @($ins.Steps)) { Write-Line "  $s" 'DarkGray' }

    if (-not $ins.Ok) {
        Write-Line $ins.Detail 'Red'
        Write-Line ''
        Write-Line 'If this machine cannot reach the PowerShell Gallery, your desktop team' 'Yellow'
        Write-Line 'will need to install it for you. The command is:' 'Yellow'
        Write-Line "    Install-Module $($mod.Name) -Scope CurrentUser"
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' -Detail $ins.Detail
        return
    }

    Write-Line $ins.Detail 'Green'
    # The dashboard's cached view is now wrong -- it was built before the
    # install. Clearing it is cheaper than letting it report Missing for a
    # module that is now loaded.
    $global:ToolkitModuleStateMap = $null
} else {
    Write-Line $mod.Detail 'DarkGray'
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

# The sign-in window is an ActiveX WebBrowser control, and it can only be
# created on an STA thread that is not carrying somebody else's synchronization
# context.
#
# This script is started from AdminConsole immediately after ShowDialog()
# returns. ShowDialog installs a WindowsFormsSynchronizationContext on the
# thread and does NOT remove it when the form closes, so MSAL captures a context
# whose message pump has already gone. That is the reported failure:
#
#   ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2' cannot be
#   instantiated because the current thread is not in a single-threaded
#   apartment
#
# 8856f961- is the IE WebBrowser control, and the trace ends in
# WinFormsLegacyWebUi.OnAuthenticate -- the embedded browser that -DisableWAM
# selects.
#
# Clearing the context is cheap and safe: nothing after this point needs it, and
# the console is rebuilt from scratch when it reopens.
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
    Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' `
        -Detail "Thread apartment is $apartment; interactive sign-in requires STA."
    Write-Line ''
    return
}

Write-Line ''
Write-Line 'A sign-in window will open. Complete MFA there, then return here.' 'Yellow'
Write-Line ''

# ErrorAction is ALREADY in $params (line 60). Do not add it here as well --
# splatting a key and passing the same parameter explicitly is a binding error,
# "parameter 'ErrorAction' is specified more than once", and it fails before the
# module is even called. That broke every connect attempt until it was spotted.
#
# 2>$null suppresses the error STREAM, not the exception. The module writes
# forty lines of MSAL async plumbing to it before the terminating error
# surfaces, and none of those lines names a cause -- only the last one does, and
# the catch below prints it.
try {
    Connect-ExchangeOnline @params 2>$null
} catch {
    $msg = "$($_.Exception.Message)"

    # One targeted retry, and ONLY for this error. The context is already
    # cleared above, so a second attempt starts clean. Retrying anything else
    # would just hide it -- which is why this checks the message rather than
    # wrapping every failure in a loop.
    if ($msg -match 'single-threaded apartment|8856f961') {
        Write-Line ''
        Write-Line 'The sign-in window could not be created on this thread. Retrying once with a clean context.' 'Yellow'
        [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
        try {
            Connect-ExchangeOnline @params 2>$null
            $msg = $null
        } catch {
            $msg = "$($_.Exception.Message)"
        }
    }

    if ($msg) {
        Write-Line ''
        Write-Line "Connection failed: $msg" 'Red'
        Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' -Detail $msg
        Write-Line ''
        return
    }
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
