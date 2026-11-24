#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    Connect-Teams.ps1
    Connects this PowerShell process to Microsoft Teams.

    Same shape as Connect-SPO.ps1 and Connect-ExO.ps1.
#>
# TOOLKIT-HELP-BEGIN
# Connects this console to Microsoft Teams.
#
# IT USES DEVICE SIGN-IN. No window pops up. A code and a URL are printed in
# the console window -- open the URL anywhere, type the code, and this finishes
# on its own. If you are waiting for a prompt to appear you are waiting for
# something that is not coming; read the code instead.
#
# The console closes first so there IS a console to print the code into, then
# reopens with the connection live.
#
# ONLY THE TEAMS TOOLS NEED THIS. The Exchange, SharePoint and AD tools do not
# use it and are not greyed when it is absent.
#
# CONNECTING TO TEAMS DOES NOT CONNECT YOU TO ANYTHING ELSE. It is a separate
# sign-in from Exchange Online even when it is the same account, and the header
# reports the two independently.
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

# FIRST THING AFTER Common LOADS. A guard that runs after work has started is a
# guard that already lost.
if (-not (Test-ConsoleLaunch)) { return }

# No WinForms. This script must NOT open a window of its own -- the sign-in
# prompt is the window, and anything of ours competing with it is the problem
# this script exists to avoid.

function Write-Line {
    param([string]$Text, [string]$Colour)
    if ($Colour) { Write-Host $Text -ForegroundColor $Colour } else { Write-Host $Text }
}

# Test-TeamsConnectionState is NOT defined here. Common.ps1 already has it, and
# a local copy would win over the shared one for this file only -- so a fix made
# in Common would never reach this tool. FUN001 exists for exactly that.

$before = Test-TeamsConnectionState
if ($before.Connected -and -not $Force) {
    Write-Line "Already connected. $($before.Detail)" 'Green'
    Write-Line 'Nothing to do. Re-run with -Force to reconnect.'
    return
}

if (-not (Get-Command -Name 'Connect-MicrosoftTeams' -ErrorAction SilentlyContinue)) {
    # THE TWO REASONS A SERVICE IS UNAVAILABLE ARE NOT THE SAME PROBLEM, and
    # saying "not connected" when the module is missing sends the operator to
    # click Connect forever.
    Write-Line 'The MicrosoftTeams module is NOT INSTALLED in this window.' 'Red'
    Write-Line 'This is not a sign-in problem and clicking Connect again will not fix it. Install it with:'
    Write-Line '    Install-Module -Name MicrosoftTeams -Scope CurrentUser'
    Write-FailedLog -Action 'Connect' -Target 'Microsoft Teams' `
        -Detail 'Refused: MicrosoftTeams is not available in this process.'
    return
}

# DEVICE AUTHENTICATION, NOT THE INTERACTIVE BROWSER PROMPT.
#
# Plain Connect-MicrosoftTeams does not work here -- it raises an embedded
# browser sign-in that fails in this environment. -UseDeviceAuthentication
# prints a code and a URL and waits instead, which works precisely because it
# needs no browser of its own.
#
# THE CODE APPEARS IN THIS WINDOW, so it has to be read here. That is also why
# the console closes first: there has to be somewhere to print it.
# THE PAGE IS OPENED FOR YOU, BEFORE the connect call rather than after.
#
# The device-login URL is fixed -- https://microsoft.com/devicelogin -- so it
# can be opened without knowing the code. The code itself cannot be handled the
# same way: it is generated INSIDE Connect-MicrosoftTeams and printed as the
# cmdlet blocks waiting for the sign-in, so there is no point at which this
# script can read it and paste it in. Opening the page first is the whole of
# what is automatable here.
#
# A FAILURE TO OPEN IS NOT FATAL. If there is no default browser, or the shell
# refuses, the sign-in still works by hand -- so this warns and carries on
# rather than stopping a connection that was going to succeed.
$deviceUrl = 'https://microsoft.com/devicelogin'
try {
    Start-Process $deviceUrl -ErrorAction Stop
    Write-Line "Opened $deviceUrl in your browser." 'Cyan'
} catch {
    Write-Line "Could not open a browser automatically. Go to $deviceUrl yourself." 'Yellow'
}
Write-Line ''
Write-Line 'A CODE will be printed below in a moment. Type it into that page.' 'Cyan'
Write-Line 'This window waits until you finish -- nothing else will pop up.'
Write-Line ''

try {
    Connect-MicrosoftTeams -UseDeviceAuthentication -ErrorAction Stop | Out-Null

    # READ BACK. An absent exception is not evidence a session exists -- the
    # same reason every write in this toolkit re-reads.
    $after = Test-TeamsConnectionState
    if ($after.Connected) {
        Write-Line ''
        Write-Line "Connected. $($after.Detail)" 'Green'
        Write-CompletedLog -Action 'Connect' -Target 'Microsoft Teams' -Result 'Success' `
            -Detail 'Connected to Microsoft Teams via device authentication.'
    } else {
        Write-Line ''
        Write-Line 'Connect-MicrosoftTeams returned without error, but no session answered afterwards.' 'Yellow'
        Write-Line $after.Detail
        Write-CompletedLog -Action 'Connect' -Target 'Microsoft Teams' -Result 'Partial' `
            -Detail 'Connect-MicrosoftTeams did not throw but Get-CsTenant found no session.'
    }
} catch {
    Write-Line ''
    Write-Line (Format-ErrorDetail -ErrorRecord $_ -Context 'Connect-MicrosoftTeams') 'Red'
    Write-Line ''
    Write-Line 'If the code expired before it was entered, just run this again -- codes are short-lived and that is the commonest cause.'
    Write-Line 'If the sign-in completed but this still failed, the account may lack a Teams administrator role -- the module connects and then cannot read the tenant.'
    Write-FailedLog -Action 'Connect' -Target 'Microsoft Teams' `
        -Detail 'Connect-MicrosoftTeams -UseDeviceAuthentication failed.' -ErrorRecord $_
}
