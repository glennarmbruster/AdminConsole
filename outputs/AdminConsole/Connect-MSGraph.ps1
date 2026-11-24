#Requires -Version 5.1
<#
    Connect-MSGraph.ps1
    Connects this PowerShell process to Microsoft Graph through the tenant's app.

    Reasoning: ScriptNotes/Connect-Scripts.notes.md, and the Microsoft Graph
    region of Common.ps1.

    Author   : Glenn Armbruster
    Created  : 2026-09-17
#>
# TOOLKIT-HELP-BEGIN
# Connects this console to Microsoft Graph -- the API behind Entra ID, licensing,
# sign-in logs, Identity Protection and service health.
#
# Same pattern as the Exchange Online button: the console closes, you sign in
# in the browser, and it reopens with the connection live.
#
# WHAT IT ASKS FOR. Each Graph tool declares the permissions it needs in its
# own header. This script reads every one of those and signs in ONCE with the
# whole set, so one prompt covers every Graph tool on the console. It asks for
# nothing a tool does not use. The list is printed before the sign-in.
#
# If the sign-in refuses a permission, the tenant's Graph app has not been
# granted it yet -- that is an app-registration change, not something this
# console can do.
#
# Needed by: Sign-in Activity, Risky Users, Auth Methods, Licence Assignment,
# Service Health. The Exchange tools do not use it.
# TOOLKIT-HELP-END

param(
    # Extra scopes on top of what the tools declare. Rarely needed.
    [string[]]$Scopes = @(),
    [switch]$Force
)

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

function Write-Line {
    param([string]$Text, [string]$Colour)
    if ($Colour) { Write-Host $Text -ForegroundColor $Colour } else { Write-Host $Text }
}

Write-Line ''
Write-Line '=== Connect to Microsoft Graph ===' 'Cyan'

# TENANT FIRST. Nothing below makes sense without an app to sign in through,
# and the message for a missing one has to name the file, not the module.
$cfg = Get-TenantGraphConfig
if (-not $cfg.Configured) {
    Write-Line 'Microsoft Graph is not configured for this tenant.' 'Red'
    Write-Line 'Tenant.json needs Capabilities.MicrosoftGraph = true and a Graph.ClientId (the'
    Write-Line 'Application (client) ID of the tenant''s Graph app). Tenant Manager edits both.'
    Write-Line ''
    return
}

# WHAT THE TOOLS NEED. Read from the scripts, not typed here: a list in this
# file would be right on the day it was written and wrong the first time a
# tool was added without someone remembering to come back.
$catalogue = @(Get-ToolkitGraphScopeCatalogue -Root $PSScriptRoot)
$want = @($catalogue | ForEach-Object { $_.Scope })
foreach ($s in $Scopes) { if ($s -and $want -notcontains $s) { $want += $s } }
if ($want.Count -eq 0) { $want = @('User.Read') }

$before = Test-GraphConnectionState
if ($before.Connected -and -not $before.WrongApp -and -not $Force) {
    $missing = @($want | Where-Object { $before.Scopes -notcontains $_ })
    if ($missing.Count -eq 0) {
        Write-Line "Already connected. $($before.Detail)" 'Green'
        Write-Line 'Every scope the tools declare is held. Nothing to do; re-run with -Force to reconnect.'
        Write-Line ''
        return
    }
    Write-Line "Connected as $($before.Account), but missing $($missing.Count) scope(s) the tools declare:" 'Yellow'
    foreach ($m in $missing) { Write-Line "    $m" 'Yellow' }
    Write-Line 'Reconnecting to add them.' 'DarkGray'
} elseif ($before.Connected -and $before.WrongApp) {
    Write-Line $before.Detail 'Yellow'
    Write-Line 'Replacing that sign-in with one through the tenant''s app.' 'DarkGray'
} elseif (-not $before.Connected) {
    Write-Line $before.Detail 'DarkGray'
}

# MODULE FIRST, SIGN-IN SECOND. Same three-state handling as Connect-ExO.ps1;
# the reasoning is written out there. Only Microsoft.Graph.Authentication --
# the full SDK is forty modules the tools do not use.
$mod = Test-ToolkitModule -Id 'Graph'

if ($mod.State -eq 'Installed') {
    Write-Line $mod.Detail 'DarkGray'
    $imp = Import-ToolkitModule -Id 'Graph'
    Write-Line $imp.Detail $(if ($imp.Ok) { 'Green' } else { 'Red' })
    if (-not $imp.Ok) {
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'Microsoft Graph' -Result 'Failed' -Detail $imp.Detail
        return
    }
    $mod = Test-ToolkitModule -Id 'Graph'
}

if ($mod.State -ne 'Imported') {
    Write-Line "$($mod.Name) is not installed for your account." 'Yellow'

    if (-not (Request-ToolkitModuleInstall -Id 'Graph')) {
        Write-Line 'Not installed -- you said no. Nothing was changed.' 'DarkGray'
        Write-Line "To do it yourself later:  Install-Module $($mod.Name) -Scope CurrentUser"
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'Microsoft Graph' -Result 'Failed' `
            -Detail "$($mod.Name) not installed; operator declined the install."
        return
    }

    Write-Line 'Installing. This can take a minute the first time.' 'Cyan'
    $ins = Install-ToolkitModule -Id 'Graph'
    foreach ($s in @($ins.Steps)) { Write-Line "  $s" 'DarkGray' }

    if (-not $ins.Ok) {
        Write-Line $ins.Detail 'Red'
        Write-Line ''
        Write-Line 'If this machine cannot reach the PowerShell Gallery, your desktop team' 'Yellow'
        Write-Line 'will need to install it for you. The command is:' 'Yellow'
        Write-Line "    Install-Module $($mod.Name) -Scope CurrentUser"
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'Microsoft Graph' -Result 'Failed' -Detail $ins.Detail
        return
    }

    Write-Line $ins.Detail 'Green'
    $global:ToolkitModuleStateMap = $null
} else {
    Write-Line $mod.Detail 'DarkGray'
}

Write-Line ''
Write-Line 'Signing in with these delegated scopes (declared by the tools that use them):' 'Cyan'
foreach ($row in $catalogue) {
    Write-Line ("    {0,-46} {1}" -f $row.Scope, ($row.Scripts -join ', ')) 'DarkGray'
}
foreach ($s in $Scopes) { if ($s) { Write-Line ("    {0,-46} (asked for on the command line)" -f $s) 'DarkGray' } }

# The sign-in is drawn by MSAL after the console's window has closed, but
# ShowDialog leaves a WindowsFormsSynchronizationContext on this thread whose
# pump is gone. Cleared here as well as in the console, because this script
# can be run from a bare prompt. Same fix as Connect-ExO.ps1.
$staleCtx = [System.Threading.SynchronizationContext]::Current
if ($staleCtx) {
    [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
}

Write-Line ''
Write-Line 'A sign-in window will open in your browser. Complete MFA there, then return here.' 'Yellow'
Write-Line ''

try {
    Connect-ToolkitGraph -Scopes $want -Force:$Force | Out-Null
} catch {
    Write-Line ''
    Write-Line "Connection failed: $($_.Exception.Message)" 'Red'
    Write-Line ''
    Write-Line 'If the message says a permission was not consented or is invalid, the tenant''s' 'DarkGray'
    Write-Line 'Graph app has not been granted one of the scopes above. That is an' 'DarkGray'
    Write-Line 'app-registration change for whoever owns the app -- the list above is what to send.' 'DarkGray'
    Write-CompletedLog -Action 'Connect' -Target 'Microsoft Graph' -Result 'Failed' `
        -Detail "$($cfg.ClientId) -- $($_.Exception.Message)"
    Write-Line ''
    return
}

$after = Test-GraphConnectionState
Write-Line ''
if ($after.Connected -and -not $after.WrongApp) {
    Write-Line "Connected. $($after.Detail)" 'Green'
    $held    = @($after.Scopes)
    $missing = @($want | Where-Object { $held -notcontains $_ })
    if ($missing.Count -gt 0) {
        # Consent can succeed for SOME scopes and silently drop others. A tool
        # would find out at its first call; better to say it here.
        Write-Line "Granted, but $($missing.Count) requested scope(s) are NOT on the token:" 'Yellow'
        foreach ($m in $missing) { Write-Line "    $m" 'Yellow' }
        Write-Line 'The tools that declare those will say so when opened. Ask the app owner to add them.' 'Yellow'
    }
    Write-CompletedLog -Action 'Connect' -Target 'Microsoft Graph' -Result $(if ($missing.Count) { 'Partial' } else { 'Success' }) `
        -Detail "$($after.Account) via $($cfg.ClientId); $($held.Count) scope(s)$(if ($missing.Count) { '; missing: ' + ($missing -join ', ') })"
} else {
    Write-Line "Sign-in completed but no usable session is present: $($after.Detail)" 'Red'
    Write-CompletedLog -Action 'Connect' -Target 'Microsoft Graph' -Result 'Failed' -Detail "$($after.Detail)"
}
Write-Line ''
