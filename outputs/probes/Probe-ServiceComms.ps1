#Requires -Version 5.1
# ---------------------------------------------------------------------------
<#
    Probe-ServiceComms.ps1

    Can this account read Service Health and Message Center through Graph?

    A PROBE, NOT A TOOL. Deliberately different from everything in AdminConsole:

      - Console output, no WinForms. You read it once and act on it. A form
        would be more code between you and the answer.
      - It does NOT dot-source Common.ps1. A probe that depends on the toolkit
        can fail because of the toolkit, and then you have learned nothing
        about Graph. Self-contained means every failure here is a Graph
        failure.
      - It reads and writes nothing. No log entry, no report file.

    Run it, read the verdict at the bottom, delete it. If the answer is yes,
    the tool gets built against exactly the endpoints proved here.

    Author   : Glenn Armbruster
    Revised  : 2026-08-12
#>

$ErrorActionPreference = 'Continue'

# The two scopes the service communications API documents. Both are
# admin-consent-only, so a tenant that has never consented them to the
# Microsoft Graph PowerShell app will raise a consent prompt on first run --
# that is a real possible outcome of this probe, not a fault.
$Scopes = @('ServiceHealth.Read.All', 'ServiceMessage.Read.All')

$Root = 'https://graph.microsoft.com/v1.0/admin/serviceAnnouncement'

function Write-Head {
    param([string]$Text)
    Write-Host ''
    Write-Host $Text -ForegroundColor Cyan
    Write-Host ('-' * 72) -ForegroundColor DarkGray
}

function Write-Good { param([string]$T) Write-Host "  $T" -ForegroundColor Green }
function Write-Bad  { param([string]$T) Write-Host "  $T" -ForegroundColor Red }
function Write-Warn { param([string]$T) Write-Host "  $T" -ForegroundColor Yellow }
function Write-Info { param([string]$T) Write-Host "  $T" }

# ---------------------------------------------------------------------------
Write-Head 'STEP 1  Graph module'

Write-Info "PowerShell $($PSVersionTable.PSVersion)  ($($PSVersionTable.PSEdition))"

# Get-Command only finds a cmdlet in an unloaded module because AUTO-LOADING
# goes looking for it. Set this preference to None -- which some managed
# profiles and group policies do -- and every Get-Command for an unimported
# module returns nothing, with no error and no clue. An installed, healthy,
# correctly-pathed module then looks exactly like a missing one.
$autoLoad = try { "$PSModuleAutoloadingPreference" } catch { '' }
if ($autoLoad -and $autoLoad -ne 'All') {
    Write-Warn "PSModuleAutoloadingPreference = '$autoLoad'. Modules will not load on demand;"
    Write-Warn 'anything not already imported will look absent. Importing explicitly below.'
}

$connect = Get-Command Connect-MgGraph -ErrorAction SilentlyContinue

# IMPORT BY EXPLICIT VERSION, NOT BY NAME.
#
# Corrected 2026-08-12 after Test-GraphAuthVersion.ps1 disproved the first
# diagnosis. A bare 'Import-Module Microsoft.Graph.Authentication' failed here
# with "Method 'GetTokenAsync' ... does not have an implementation", and the
# obvious conclusion -- six installed versions, so the install is broken -- was
# WRONG. Every one of those six imports cleanly in a fresh -NoProfile process,
# including the newest.
#
# So the versions on disk are fine and the resolution is not. Naming the
# version removes the ambiguity that a bare import leaves PowerShell to settle
# across two scopes and six candidates.
#
# If this still fails, the poison is already in the process before the probe
# starts -- something the profile loaded -- and the answer is -NoProfile, which
# the verdict below says.
if (-not $connect) {
    $newest = Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication' -ErrorAction SilentlyContinue |
              Sort-Object Version -Descending | Select-Object -First 1
    if ($newest) {
        try {
            Import-Module Microsoft.Graph.Authentication -RequiredVersion $newest.Version -ErrorAction Stop
        } catch {
            Write-Warn "Pinned import of $($newest.Version) failed: $($_.Exception.Message)"
            Write-Warn 'That version loads fine in a clean process, so something in THIS session'
            Write-Warn 'loaded a conflicting Graph assembly first. Retry with:'
            Write-Warn "  powershell.exe -NoProfile -File `"$PSCommandPath`""
        }
    }
    $connect = Get-Command Connect-MgGraph -ErrorAction SilentlyContinue
    if ($connect) {
        Write-Warn "Loaded only after pinning version $($newest.Version) -- auto-loading did not work here."
    }
}

if (-not $connect) {
    # "IT IS INSTALLED" AND "THIS SHELL CAN SEE IT" ARE DIFFERENT STATEMENTS,
    # and the gap between them is almost always which PowerShell you are in.
    # Windows PowerShell 5.1 loads user modules from
    #   Documents\WindowsPowerShell\Modules
    # and PowerShell 7 from
    #   Documents\PowerShell\Modules
    # Install-Module -Scope CurrentUser writes to whichever one you ran it in,
    # so a module installed from 7 is genuinely absent from 5.1 and vice versa.
    # Telling you to install it again, in the shell that already cannot see it,
    # would just make a second copy in the wrong place.
    #
    # So: look before advising.
    Write-Bad 'Connect-MgGraph is not available IN THIS SHELL.'

    $found = @()
    try {
        $found = @(Get-Module -ListAvailable -Name 'Microsoft.Graph*' -ErrorAction SilentlyContinue)
    } catch { }

    if ($found.Count -gt 0) {
        Write-Warn 'But Graph modules ARE on this machine, visible to this shell:'
        foreach ($m in ($found | Sort-Object Name, Version | Select-Object -First 12)) {
            Write-Info ("    {0,-42} {1,-10} {2}" -f $m.Name, $m.Version, $m.ModuleBase)
        }
        if ($found.Count -gt 12) { Write-Info "    ... and $($found.Count - 12) more" }

        # No second import attempt here. The pinned one above already ran and
        # already printed its error; repeating it would print the same failure
        # twice and read like two separate problems.
        $auth = @($found | Where-Object { $_.Name -eq 'Microsoft.Graph.Authentication' })
        if ($auth.Count -eq 0) {
            Write-Warn 'Microsoft.Graph.Authentication specifically is NOT among them.'
        }
    } else {
        Write-Info 'No Microsoft.Graph* modules are visible to this shell at all.'
    }

    if (-not $connect) {
        Write-Info ''
        Write-Info 'Module search paths this shell uses:'
        foreach ($p in ($env:PSModulePath -split ';' | Where-Object { $_ })) { Write-Info "    $p" }
        Write-Info ''
        Write-Info 'TRY THESE IN ORDER -- cheapest first:'
        Write-Info ''
        Write-Info '  1. Close this window and open a new one. An assembly loaded earlier in a'
        Write-Info '     session cannot be unloaded, so a conflicting Graph DLL pulled in by'
        Write-Info '     anything you ran before makes every later import fail. A restart is'
        Write-Info '     the only way to clear it, and it costs nothing.'
        Write-Info ''
        Write-Info '  2. Skip your profile, in case that is what loads it:'
        Write-Info "       powershell.exe -NoProfile -File `"$PSCommandPath`""
        Write-Info ''
        Write-Info '  3. Run Test-GraphAuthVersion.ps1. It tests each installed version in its'
        Write-Info '     own fresh process, so it can tell a genuinely broken module apart from'
        Write-Info '     a poisoned session -- they look identical from in here.'
        Write-Info ''
        Write-Info 'Only if ALL versions fail there is the install itself the problem.'
        return
    }
}

$modVer = try { (Get-Module Microsoft.Graph.Authentication -ListAvailable |
                 Sort-Object Version -Descending | Select-Object -First 1).Version } catch { $null }
Write-Good "Connect-MgGraph found.  Microsoft.Graph.Authentication $(if ($modVer) { "v$modVer" } else { '(version unknown)' })"

# ---------------------------------------------------------------------------
Write-Head 'STEP 2  Connect'

$ctx = $null
try { $ctx = Get-MgContext } catch { $ctx = $null }

$needConnect = $true
if ($ctx -and $ctx.Scopes) {
    $missing = @($Scopes | Where-Object { $ctx.Scopes -notcontains $_ })
    if ($missing.Count -eq 0) {
        Write-Good 'Already connected with both scopes -- reusing the session.'
        $needConnect = $false
    } else {
        Write-Info "Connected, but missing: $($missing -join ', ')  -- reconnecting."
    }
}

if ($needConnect) {
    Write-Info 'A sign-in window will open. If it asks an admin to consent, that is'
    Write-Info 'the answer to half this probe -- note what it asks for.'

    # -NoWelcome exists only in Graph SDK v2, and on v1 an unknown parameter is
    # a hard binding error. Splatted conditionally rather than assumed, which
    # is the same guard PimActivation.ps1 uses.
    $p = @{ Scopes = $Scopes; ErrorAction = 'Stop' }
    if ($connect.Parameters.ContainsKey('NoWelcome')) { $p.NoWelcome = $true }

    try {
        Connect-MgGraph @p | Out-Null
        $ctx = Get-MgContext
    } catch {
        Write-Bad "Connect failed: $($_.Exception.Message)"
        Write-Info 'If this mentions consent, an admin must approve the two scopes for the'
        Write-Info 'Microsoft Graph PowerShell application before any of this can work.'
        return
    }
}

if (-not $ctx) { Write-Bad 'No Graph context after connecting.'; return }

Write-Info ''
Write-Info "Account  : $($ctx.Account)"
Write-Info "Tenant   : $($ctx.TenantId)"
Write-Info "AuthType : $($ctx.AuthType)"

# GRANTED, not requested. Asking for a scope and receiving it are different
# events, and the difference is the entire point of connecting before probing.
foreach ($s in $Scopes) {
    if ($ctx.Scopes -contains $s) { Write-Good "granted: $s" }
    else                          { Write-Bad  "NOT granted: $s" }
}

# ---------------------------------------------------------------------------
Write-Head 'STEP 3  Who am I, and do I hold an admin role'

# Delegated access to this API applies RBAC ON TOP OF the scope: the signed-in
# user must hold at least one admin role. A correctly granted scope with no
# active role still returns 403, and that 403 reads like a permissions bug
# rather than "you have not activated anything today". Worth knowing which of
# the two you are looking at before building on it.
try {
    $me = Invoke-MgGraphRequest -Method GET -ErrorAction Stop `
              -Uri 'https://graph.microsoft.com/v1.0/me?$select=displayName,userPrincipalName'
    Write-Info "Signed in as: $($me.displayName)  <$($me.userPrincipalName)>"
} catch {
    Write-Warn "Could not read /me: $($_.Exception.Message)"
}

try {
    $roles = Invoke-MgGraphRequest -Method GET -ErrorAction Stop `
                 -Uri 'https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole?$select=displayName'
    $names = @(@($roles.value) | ForEach-Object { "$($_.displayName)" } | Where-Object { $_ } | Sort-Object)
    if ($names.Count -gt 0) {
        Write-Good "Active directory roles ($($names.Count)):"
        foreach ($n in $names) { Write-Info "    $n" }
    } else {
        Write-Warn 'No ACTIVE directory roles. If your roles are PIM-eligible, activate first --'
        Write-Warn 'eligible is not active, and this API only sees active.'
    }
} catch {
    Write-Warn "Could not list roles: $($_.Exception.Message)"
    Write-Info 'Not fatal. The endpoint tests below are the ones that matter.'
}

# ---------------------------------------------------------------------------
Write-Head 'STEP 4  The endpoints a tool would actually call'

function Test-Endpoint {
    <#
        One endpoint, one verdict line, plus a sample so the shape of the data
        is visible. The sample matters as much as the pass/fail: a tool built
        against guessed field names is a tool that has to be rewritten.
    #>
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][string]$Uri,
        [string[]]$Show = @()
    )

    Write-Host ''
    Write-Host "  $Label" -ForegroundColor White
    Write-Host "    $Uri" -ForegroundColor DarkGray

    try {
        $r = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop

        # Invoke-MgGraphRequest returns a Hashtable on 5.1 by default and a
        # PSObject if asked. Reading 'value' through the indexer covers the
        # hashtable case; the property access covers the other. Assuming one
        # shape is how a working call reports zero rows.
        $items = @()
        if ($r -is [System.Collections.IDictionary]) { $items = @($r['value']) }
        elseif ($r.PSObject.Properties['value'])     { $items = @($r.value) }

        Write-Good "OK -- $($items.Count) item(s) returned."

        if ($items.Count -gt 0 -and $Show.Count -gt 0) {
            $first = $items[0]
            Write-Host '    sample:' -ForegroundColor DarkGray
            foreach ($f in $Show) {
                $v = if ($first -is [System.Collections.IDictionary]) { $first[$f] } else { $first.$f }
                $t = "$v"
                if ($t.Length -gt 90) { $t = $t.Substring(0, 90) + '...' }
                Write-Host ("      {0,-22} {1}" -f $f, $t) -ForegroundColor DarkGray
            }
        }
        return $true
    } catch {
        $code = ''
        try { $code = "$($_.Exception.Response.StatusCode.value__)" } catch { }
        Write-Bad ("FAILED{0}: {1}" -f $(if ($code) { " (HTTP $code)" } else { '' }), $_.Exception.Message)
        if ($code -eq '403') {
            Write-Warn 'A 403 here is usually the ROLE, not the scope -- see step 3.'
        }
        return $false
    }
}

$okHealth = Test-Endpoint -Label 'Service Health -- per-service status' `
    -Uri "$Root/healthOverviews" `
    -Show @('id', 'service', 'status')

$okIssues = Test-Endpoint -Label 'Service Health -- current issues' `
    -Uri "$Root/issues?`$top=5" `
    -Show @('id', 'title', 'classification', 'status', 'service', 'startDateTime')

$okMsgs = Test-Endpoint -Label 'Message Center -- posts' `
    -Uri "$Root/messages?`$top=5" `
    -Show @('id', 'title', 'category', 'severity', 'actionRequiredByDateTime', 'lastModifiedDateTime')

# ---------------------------------------------------------------------------
Write-Head 'VERDICT'

if ($okHealth -and $okIssues) {
    Write-Good 'Service Health is readable. A tool can be built.'
} else {
    Write-Bad  'Service Health is NOT readable by this account right now.'
}

if ($okMsgs) {
    Write-Good 'Message Center is readable. A tool can be built.'
} else {
    Write-Bad  'Message Center is NOT readable by this account right now.'
}

Write-Host ''
Write-Info 'If something failed, the useful distinction is:'
Write-Info '  consent prompt at sign-in  -> tenant has not approved the scope (admin action)'
Write-Info '  scope granted but HTTP 403 -> you hold no ACTIVE admin role (PIM, probably)'
Write-Info '  scope NOT granted          -> the request was refused, not the data'
Write-Host ''
Write-Info 'Disconnect-MgGraph when finished if you would rather not leave the session open.'
Write-Host ''
