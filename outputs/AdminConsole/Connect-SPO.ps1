#Requires -Version 5.1
<#
    Connect-SPO.ps1
    Connects this PowerShell process to SharePoint Online.

    Reasoning: ScriptNotes/Connect-Scripts.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Connects this console to SharePoint Online.
#
# Same pattern as the Exchange Online button: the console closes, you sign in,
# and it reopens with the connection live.
#
# Needed by the OneDrive and SharePoint reports. The Exchange tools do not use
# it.
# TOOLKIT-HELP-END

param(
    [string]$Url,
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

function Write-Line {
    param([string]$Text, [string]$Colour)
    if ($Colour) { Write-Host $Text -ForegroundColor $Colour } else { Write-Host $Text }
}

Write-Line ''
Write-Line '=== Connect to SharePoint Online ===' 'Cyan'

$before = Test-SpoConnectionState
if ($before.Connected -and -not $Force) {
    Write-Line "Already connected. $($before.Detail)" 'Green'
    Write-Line 'Nothing to do. Re-run with -Force to reconnect.'
    Write-Line ''
    return
}
if (-not $before.Connected) { Write-Line $before.Detail 'DarkGray' }

# MODULE FIRST, SIGN-IN SECOND. Same three-state handling as Connect-ExO.ps1;
# the reasoning is written out there.
$mod = Test-ToolkitModule -Id 'SPO'

if ($mod.State -eq 'Installed') {
    Write-Line $mod.Detail 'DarkGray'
    $imp = Import-ToolkitModule -Id 'SPO'
    Write-Line $imp.Detail $(if ($imp.Ok) { 'Green' } else { 'Red' })
    if (-not $imp.Ok) {
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'SharePoint Online' -Result 'Failed' -Detail $imp.Detail
        return
    }
    $mod = Test-ToolkitModule -Id 'SPO'
}

if ($mod.State -ne 'Imported') {
    Write-Line "$($mod.Name) is not installed for your account." 'Yellow'

    if (-not (Request-ToolkitModuleInstall -Id 'SPO')) {
        Write-Line 'Not installed -- you said no. Nothing was changed.' 'DarkGray'
        Write-Line "To do it yourself later:  Install-Module $($mod.Name) -Scope CurrentUser"
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'SharePoint Online' -Result 'Failed' `
            -Detail "$($mod.Name) not installed; operator declined the install."
        return
    }

    Write-Line 'Installing. This can take a minute or two the first time.' 'Cyan'
    $ins = Install-ToolkitModule -Id 'SPO'
    foreach ($s in @($ins.Steps)) { Write-Line "  $s" 'DarkGray' }

    if (-not $ins.Ok) {
        Write-Line $ins.Detail 'Red'
        Write-Line ''
        Write-Line 'If this machine cannot reach the PowerShell Gallery, your desktop team' 'Yellow'
        Write-Line 'will need to install it for you. The command is:' 'Yellow'
        Write-Line "    Install-Module $($mod.Name) -Scope CurrentUser"
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'SharePoint Online' -Result 'Failed' -Detail $ins.Detail
        return
    }

    Write-Line $ins.Detail 'Green'
    $global:ToolkitModuleStateMap = $null
} else {
    Write-Line $mod.Detail 'DarkGray'
}

if (-not $Url) { $Url = Get-SpoAdminUrl }
if (-not $Url) {
    Write-Line 'No SharePoint admin URL, and none could be derived.' 'Red'
    Write-Line 'Set it in ToolManager -> Settings..., or pass -Url.'
    Write-Line ''
    return
}
Write-Line "Admin centre: $Url" 'DarkGray'

$params = @{ Url = $Url; ErrorAction = 'Stop' }

# -UseSystemBrowser is the point of this script.
#
# Without it the module raises its own sign-in dialog. Launched from an Admin
# Console button, that dialog was created while $form.ShowDialog() held the
# thread in a modal message loop -- and a modal form disables other windows on
# the same thread, so the prompt drew and never accepted input.
#
# The system browser is a separate process. Nothing this console does to its
# own message loop can reach it, which is what makes this the fix rather than a
# workaround. Microsoft document it as the MSAL sign-in path.
#
# Version-guarded: the parameter is not present in every build of the module,
# and passing an unknown parameter fails the whole call.
if ((Get-Command Connect-SPOService).Parameters.ContainsKey('UseSystemBrowser')) {
    $params['UseSystemBrowser'] = $true
    Write-Line 'Using -UseSystemBrowser (sign-in goes to your default browser).' 'DarkGray'
} else {
    Write-Line 'This module version has no -UseSystemBrowser; using the built-in prompt.' 'DarkGray'
}

Write-Line ''
Write-Line 'A sign-in window will open. Complete MFA there, then return here.' 'Yellow'
Write-Line ''

try {
    Connect-SPOService @params
} catch {
    Write-Line ''
    Write-Line "Connection failed: $($_.Exception.Message)" 'Red'
    # The single most common cause is the wrong host: the admin centre is
    # <tenant>-admin.sharepoint.com, not the tenant's ordinary SharePoint URL.
    # Saying so here saves reading an MSAL error that does not mention it.
    Write-Line ''
    Write-Line "If that mentions the URL, check the admin centre host. Expected the -admin form; got: $Url" 'DarkGray'
    Write-CompletedLog -Action 'Connect' -Target 'SharePoint Online' -Result 'Failed' `
        -Detail "$Url -- $($_.Exception.Message)"
    Write-Line ''
    return
}

$after = Test-SpoConnectionState
Write-Line ''
if ($after.Connected) {
    Write-Line "Connected. $($after.Detail)" 'Green'
    Write-CompletedLog -Action 'Connect' -Target 'SharePoint Online' -Result 'Success' -Detail $Url
} else {
    Write-Line "Sign-in completed but no session is present: $($after.Detail)" 'Red'
    Write-CompletedLog -Action 'Connect' -Target 'SharePoint Online' -Result 'Failed' `
        -Detail "$Url -- $($after.Detail)"
}

# One SPO connection is supported per PowerShell session, so a second connect
# replaces the first rather than adding to it. Worth saying out loud on a jump
# box where two people share a window.
Write-Line ''
