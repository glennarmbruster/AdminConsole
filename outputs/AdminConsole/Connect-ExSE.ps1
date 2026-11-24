#Requires -Version 5.1
<#
    Connect-ExSE.ps1
    Connects this PowerShell process to on-premises Exchange (Server SE)
    by importing a remote PowerShell session from the hybrid server.

    Reasoning: ScriptNotes/Connect-Scripts.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Connects this console to on-premises Exchange.
#
# Uses your current Windows sign-in -- no password prompt -- by importing a
# session from the hybrid server.
#
# Needed by anything that touches the on-premises directory: distribution list
# management, new starters, terminations and the on-premises reports.
# TOOLKIT-HELP-END

param(
    [string]$ConnectionUri,
    [switch]$PromptForCredential,
    [switch]$Force
)

. (Join-Path $PSScriptRoot 'Common.ps1')

# THE AUDIT LOG'S Environment COLUMN COMES FROM CONFIG, NOT FROM A LITERAL.
#
# It was a hardcoded literal -- one organisation's name for one of its directories,
# in a script that ships to tenants that have never heard of it. Every row this
# tool wrote claimed to be from a directory that does not exist here.
#
# Blank is a fine answer: Write-CompletedLog falls back to Get-LogEnvironmentLabel
# when Environment is empty, which is what it is there for.
$script:LogDirectoryId = ''
try { $script:LogDirectoryId = "$((Get-DefaultTenantDirectory).Id)" } catch { }

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
Write-Line '=== Connect to on-premises Exchange (Server SE) ===' 'Cyan'

$before = Test-EmsAvailable
if ($before.Available -and -not $Force) {
    Write-Line "Already connected. $($before.Detail)" 'Green'
    Write-Line 'Nothing to do. Re-run with -Force to reconnect.'
    Write-Line ''
    return
}

# Importing on-prem Exchange into a window that already holds Exchange Online
# is the collision the whole toolkit is careful about (ARCHITECTURE section 2).
# Get-Mailbox, Set-Mailbox and Get-DistributionGroup exist in BOTH, and after
# -AllowClobber the on-prem copies win by load order -- silently, at every call
# site, with nothing on screen to say which world answered.
#
# Not blocked, because a hybrid admin legitimately needs both and knows it. But
# it is said out loud, once, at the moment it becomes true.
$exo = Test-ExoConnectionState
if ($exo.Connected) {
    Write-Line ''
    Write-Line 'WARNING: Exchange Online is already connected in this window.' 'Yellow'
    Write-Line '  Get-Mailbox, Set-Mailbox and Get-DistributionGroup exist in both.' 'Yellow'
    Write-Line '  After this import the ON-PREMISES versions will answer.' 'Yellow'
    Write-Line '  For cloud work, use a separate PowerShell window.' 'Yellow'
    Write-Line ''
}

if (-not $ConnectionUri) { $ConnectionUri = Get-ExchangeServerUri }
if (-not $ConnectionUri) {
    Write-Line 'No Exchange server configured. Set it in ToolManager -> Settings..., or pass -ConnectionUri.' 'Red'
    Write-Line ''
    return
}
Write-Line "Endpoint: $ConnectionUri" 'DarkGray'

$sessionParams = @{
    ConfigurationName = 'Microsoft.Exchange'
    ConnectionUri     = $ConnectionUri
    Authentication    = 'Kerberos'
    ErrorAction       = 'Stop'
}

# No Get-Credential by default, and that is the important difference from the
# script this replaces.
#
# Two reasons. First, -Credential is not mandatory on New-PSSession -- the
# documented default is "Current user" -- so with Kerberos the logged-on
# ticket is used, which is the normal single-sign-on case on a jump box and
# needs no prompt at all.
#
# Second, a prompt is the thing that breaks here. Get-Credential raises a
# Windows credential dialog, and when this runs from a console button the
# console is inside $form.ShowDialog() -- a modal loop that disables other
# windows on the same thread. That is exactly what made the Exchange Online
# and SharePoint sign-in prompts hang. Not prompting sidesteps it entirely.
#
# -PromptForCredential is there for the case that actually needs it: signing in
# as a different admin account. If the dialog does not appear, run this script
# directly from the prompt rather than from the console button.
if ($PromptForCredential) {
    Write-Line ''
    Write-Line 'A credential prompt will open. If it does not appear, close the console' 'Yellow'
    Write-Line 'and run .\Connect-ExSE.ps1 -PromptForCredential from the prompt instead.' 'Yellow'
    $cred = Get-Credential -Message "Exchange on-premises ($ConnectionUri)"
    if (-not $cred) {
        Write-Line 'Cancelled -- no credentials entered.' 'DarkGray'
        Write-Line ''
        return
    }
    $sessionParams['Credential'] = $cred
} else {
    Write-Line "Using the logged-on account ($env:USERDOMAIN\$env:USERNAME) via Kerberos." 'DarkGray'
    Write-Line 'Use -PromptForCredential to sign in as a different account.' 'DarkGray'
}

Write-Line ''
Write-Line 'Opening the remote session...' 'DarkGray'

$session = $null
try {
    $session = New-PSSession @sessionParams
} catch {
    Write-Line ''
    Write-Line "Could not open the session: $($_.Exception.Message)" 'Red'
    Write-Line ''
    Write-Line 'Common causes:' 'DarkGray'
    Write-Line '  - The logged-on account has no Exchange role. Retry with -PromptForCredential.' 'DarkGray'
    Write-Line '  - The server name is wrong. Check ToolManager -> Settings...' 'DarkGray'
    Write-Line '  - Kerberos needs the FQDN. A short host name will not authenticate.' 'DarkGray'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange on-premises' -Result 'Failed' `
        -Detail "$ConnectionUri -- $($_.Exception.Message)" -Environment $script:LogDirectoryId
    Write-Line ''
    return
}

try {
    # THIS is where the button version of the old script would have failed
    # silently.
    #
    # Import-PSSession has no -Global parameter and no documented scope
    # behaviour. It builds a temporary module and returns the PSModuleInfo. Run
    # interactively, its commands land where you want them. Run inside a .ps1
    # launched with '&' -- which is how every console button works -- there is
    # no guarantee they outlive the script, and in practice they do not. The
    # script would report success and leave the window with no Exchange
    # cmdlets in it.
    #
    # Import-Module -Global is the documented fix: "imports modules into the
    # global session state so they're available to all commands in the
    # session." Capture the module, then re-import it globally.
    #
    # -AllowClobber because Exchange Online may already own Get-Mailbox and
    # friends; without it those commands are skipped and the import looks
    # partial for no visible reason. See the warning above -- this is the
    # deliberate half of that collision.
    $module = Import-PSSession -Session $session -DisableNameChecking -AllowClobber -ErrorAction Stop
    Import-Module -ModuleInfo $module -Global -DisableNameChecking -ErrorAction Stop
} catch {
    Write-Line ''
    Write-Line "Session opened but the import failed: $($_.Exception.Message)" 'Red'
    try { Remove-PSSession -Session $session -ErrorAction SilentlyContinue } catch { }
    Write-CompletedLog -Action 'Connect' -Target 'Exchange on-premises' -Result 'Failed' `
        -Detail "$ConnectionUri -- import failed: $($_.Exception.Message)" -Environment $script:LogDirectoryId
    Write-Line ''
    return
}

# Held globally so it survives this script's scope and can be found later.
# Microsoft: "you must keep the connection to the other session open. If you
# close it, the imported commands are no longer available." A session nobody
# holds a reference to is a session nobody can deliberately close.
$global:ExchangeOnPremSession = $session

# The forest-wide view is the reason this console can look up two directories
# without being told which. Without it, recipient lookups silently return only
# the local domain -- an empty result that looks like "no such mailbox".
try {
    Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
    Write-Line 'View entire forest: on.' 'DarkGray'
} catch {
    Write-Line "Connected, but Set-ADServerSettings failed: $($_.Exception.Message)" 'Yellow'
    Write-Line 'Lookups may only see the local domain.' 'Yellow'
}

$after = Test-EmsAvailable
Write-Line ''
if ($after.Available) {
    Write-Line "Connected. $($after.Detail)" 'Green'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange on-premises' -Result 'Success' `
        -Detail $ConnectionUri -Environment $script:LogDirectoryId
} else {
    # The import reported success and the probe disagrees. Almost always the
    # scope problem above, so name it rather than leaving a puzzle.
    Write-Line "Import completed but the cmdlets are not present: $($after.Detail)" 'Red'
    Write-Line 'If this persists, run this script directly rather than from the console button.' 'Yellow'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange on-premises' -Result 'Failed' `
        -Detail "$ConnectionUri -- imported but cmdlets absent" -Environment $script:LogDirectoryId
}
Write-Line ''
