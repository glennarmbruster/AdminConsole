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

# Cheap and harmless when auto-loading works, decisive when it does not. This
# runs BEFORE the "not available" branch so a suppressed auto-load never gets
# reported as a missing module.
if (-not $connect) {
    try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop } catch { }
    $connect = Get-Command Connect-MgGraph -ErrorAction SilentlyContinue
    if ($connect) { Write-Warn 'Found only after an explicit Import-Module -- auto-loading is not working here.' }
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

        $auth = @($found | Where-Object { $_.Name -eq 'Microsoft.Graph.Authentication' })
        if ($auth.Count -gt 0) {
            # Present but not auto-loading. Usually a half-upgraded SDK: several
            # versions side by side, and the assembly load fails silently.
            # Importing explicitly turns a silent miss into a real error message.
            Write-Info ''
            Write-Info 'Microsoft.Graph.Authentication is present but did not auto-load. Trying explicitly...'
            try {
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                $connect = Get-Command Connect-MgGraph -ErrorAction SilentlyContinue
                if ($connect) { Write-Good 'Imported. Continuing.' }
            } catch {
                Write-Bad "Import failed: $($_.Exception.Message)"
                if ($auth.Count -gt 1) {
                    Write-Warn ("$($auth.Count) versions of Microsoft.Graph.Authentication are installed. " +
                                'Mixed versions are the usual cause -- remove the old ones.')
                }
            }
        } else {
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
        Write-Info 'If Graph works in your OTHER PowerShell, that is the answer -- run this there.'
        Write-Info 'PimActivation.ps1 uses the same module, so whichever shell that works in is the'
        Write-Info 'one to use here.'
        Write-Info ''
        Write-Info 'Otherwise install just the piece needed (not the 40-module meta-package):'
        Write-Info '  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
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
