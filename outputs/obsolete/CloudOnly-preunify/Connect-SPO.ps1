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

if (-not (Get-Command -Name 'Connect-SPOService' -ErrorAction SilentlyContinue)) {
    Write-Line 'Microsoft.Online.SharePoint.PowerShell is not installed in this process.' 'Red'
    Write-Line 'Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser'
    Write-Line ''
    Write-CompletedLog -Action 'Connect' -Target 'SharePoint Online' -Result 'Failed' `
        -Detail 'Microsoft.Online.SharePoint.PowerShell module not available'
    return
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
