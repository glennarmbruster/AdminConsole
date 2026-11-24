#Requires -Version 5.1
<#
.SYNOPSIS
    M365Logic -- all cloud-side (Exchange Online / SharePoint Online)
    administration LOGIC for the M365_ toolkit, with no UI of any kind.
.DESCRIPTION
    ADDED 2026-07-24 -- first module in the "logic module + UI layer" split
    the operator specified after working through the browser-host
    architecture:

        "tools should gradually move from a 'standalone script' model to a
         'logic module + UI layer' model. The business logic should live in
         reusable functions or modules that can be called by WinForms, the
         browser UI, scheduled jobs, or the ad-hoc console."

    This file is the logic half. It contains ZERO UI code -- no
    System.Windows.Forms, no System.Drawing, no HTML, no HttpListener, no
    Write-Host, no MessageBox, no prompts, no Read-Host. Every function here
    takes parameters and returns objects or strings. That's what lets the
    same function serve a WinForms window, a browser endpoint, a scheduled
    job, or a bare console one-liner without modification.

    WHERE THE UI LIVES INSTEAD (all three consume this module):
      - M365Common.ps1        -- WinForms UI layer. Keeps only
                                 Show-TargetPromptDialog / Resolve-ToolTarget
                                 (the modal "enter a target" dialog), which
                                 are genuinely WinForms and genuinely UI.
      - M365WebCommon.ps1     -- browser UI layer. HttpListener hosting,
                                 routing, shared CSS, the connection gate.
      - M365_*.ps1 tools      -- thin presentation: build a window or a
                                 page, call functions from this module,
                                 render what comes back.

    WHY A .psm1 AND NOT ANOTHER DOT-SOURCED .ps1: functions in a module get
    their own scope. A dot-sourced .ps1 dumps everything into the caller's
    scope, where a variable the caller happens to also use can silently
    collide -- which is exactly the class of bug that bit this codebase
    before (a shared $script:DefaultDC in Common.ps1 broke
    disable-intake.ps1's own DC handling on 2026-07-21). Module scope makes
    that structurally impossible: $script:M365DefaultDomainCache below is
    private to this module and cannot be seen, read, or clobbered by any
    consumer. Import-Module also gives explicit control over what's public
    (see the manifest's FunctionsToExport) instead of exporting whatever
    happens to be defined.

    PS 5.1 FLOOR, DELIBERATELY -- this module must load in BOTH Windows
    PowerShell 5.1 (what the WinForms suite runs on) and PowerShell 7 (what
    the browser host runs on). So nothing here uses PS7-only features:
    no $IsWindows/$IsMacOS/$IsLinux, no ternary operator, no ?? operator,
    no -Parallel. Those all stay in the PS7-only web layer
    (M365WebCommon.ps1) where they're safe.

    CLOUD-ONLY, DELIBERATELY -- nothing in this module touches on-prem
    Exchange or the ActiveDirectory module. That's not an oversight; it's
    the module boundary. The on-prem tools (DLManager.ps1, DL_SendToAs.ps1,
    GroupEditor.ps1's on-prem branch, MailAttributes/AllAttributes,
    DLMemberships, and AdminConsole.ps1's HQ/Franchise auto-detect) depend
    on the on-prem Exchange Management Shell and RSAT, which are Windows
    PowerShell 5.1-only and cannot follow into a PS7 host. Keeping them out
    of this module means this module always fully loads in either host --
    if on-prem logic were mixed in here, importing it under PS7 would
    half-work at best. On-prem logic belongs in a separate sibling module
    (not built yet) if/when it's worth extracting.

    IMPORTANT -- THIS MODULE NEVER CONFIRMS ANYTHING. Destructive functions
    here (Disable-M365User, Revoke-*, Invoke-QuarantineRelease,
    Set-M365OneDriveSiteAdmin) do exactly what they're told, immediately,
    with no prompt. Confirmation is the UI layer's job, and every existing
    caller already does it (WinForms MessageBox, browser confirm()). This is
    the correct split -- a scheduled job must not block on a dialog -- but
    it means any NEW caller of these functions is responsible for its own
    confirmation step. Do not "helpfully" add prompts here later; that
    would break the non-interactive callers this split exists to enable.

    PROVENANCE: every function below is moved, logic-unchanged, from the
    tool script that previously owned it --
      - Get-EXOCommand / Invoke-EXOCommand / Format-ErrorDetail /
        Get-M365DefaultDomain / Get-TargetSmtp / Write-CompletedLog /
        Resolve-Target / Get-InboxRulesForMailbox  <- M365Common.ps1
      - Get-MailboxHealthSnapshot / Format-MailboxHealthReport
                                                   <- M365_MailboxHealth.ps1
                                                      (+ its _Web sibling)
      - Get-MailboxState / Format-MailboxState / Grant-* / Revoke-*
                                        <- M365_MailboxPermissions_Web.ps1
      - Invoke-MessageTrace / Invoke-MessageTraceDetail /
        Find-QuarantineMatch / Invoke-QuarantineRelease /
        Get-TraceTimeZone / ConvertTo-TraceDisplayTime / Format-TraceDetail
                                             <- M365_MessageTrace_Web.ps1
      - ConvertTo-NormalizedUserList / Get-OneDrivePersonalSiteUrl /
        Disable-M365User                          <- M365_Disable_Web.ps1
    Verification notes and flagged assumptions from those files are carried
    forward verbatim into the relevant function comments below rather than
    dropped -- they're still true, and still worth reading before trusting
    a field.

    NEW IN THIS MODULE (not moved from anywhere):
      Test-M365ExoConnection -- see its own comment. Added because a
      long-running host will lose its EXO session to token expiry or idle
      teardown, and the failure mode looks like the confusing
      "no command named Get-Mailbox exists in this session" error the
      operator already hit once for a different reason.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24
    Requires: ExchangeOnlineManagement (all EXO functions).
              Microsoft.Online.SharePoint.PowerShell only for
              Set-M365OneDriveSiteAdmin / the OneDrive step of
              Disable-M365User -- every other function here is EXO-only, so
              the module loads and works fine with no SPO module present as
              long as those two aren't called.
#>

#region Module-private state -----------------------------------------------------
# Module-scoped, therefore genuinely private -- no consumer can see or
# clobber this (see header note on why that matters here specifically).
$script:M365DefaultDomainCache = $null
#endregion

#region EXO cmdlet resolution ----------------------------------------------------
function Get-EXOCommand {
    <#
    .SYNOPSIS
        Resolves a cmdlet name to the Exchange Online-sourced command
        specifically, not an on-prem copy that may also be loaded.
    .DESCRIPTION
        Kept from the original Common.ps1 lineage. Get-Mailbox,
        Get-DistributionGroup, Get-AcceptedDomain and friends all exist in
        BOTH on-prem Exchange and Exchange Online Management. If an
        operator's session also has an on-prem Exchange module/snap-in
        loaded (e.g. from working the on-prem AdminConsole suite earlier in
        the same session), a bare call could silently resolve against the
        wrong environment -- acting on the wrong tenant with no error. Every
        EXO-flavored call in this module goes through Invoke-EXOCommand
        below rather than calling a cmdlet name directly.

        Connect-ExchangeOnline generates a temporary module per session,
        named tmpEXO_<random>; that's the preferred match. The
        'ExchangeOnlineManagement' fallback covers the module's own
        statically-shipped cmdlets.
    #>
    param([Parameter(Mandatory)][string]$Name)
    $all = @(Get-Command -Name $Name -All -ErrorAction SilentlyContinue)
    if ($all.Count -eq 0) {
        throw "No command named '$Name' exists in this session. Is Exchange Online connected (Connect-ExchangeOnline) to the target tenant? Note that Connect-ExchangeOnline is per-PROCESS -- connecting in a different window/process does not carry over to this one."
    }
    $cmd = $all | Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' } | Select-Object -First 1
    if (-not $cmd) {
        $cmd = $all | Where-Object { $_.ModuleName -eq 'ExchangeOnlineManagement' -or $_.Source -eq 'ExchangeOnlineManagement' } | Select-Object -First 1
    }
    if (-not $cmd) {
        $sources = ($all | ForEach-Object { "$($_.Name) [ModuleName=$($_.ModuleName); Source=$($_.Source)]" }) -join '  |  '
        throw "Found '$Name' but none of the loaded copies look like Exchange Online: $sources -- run 'Get-Command $Name -All | Select Name,ModuleName,Source' to confirm which one this should be; an on-prem Exchange module/snap-in may be shadowing it."
    }
    return $cmd
}

function Invoke-EXOCommand {
    <#
    .SYNOPSIS
        Splat-friendly wrapper around Get-EXOCommand.
    .EXAMPLE
        Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $cmd = Get-EXOCommand -Name $Name
    & $cmd @Parameters
}

function Test-M365ExoConnection {
    <#
    .SYNOPSIS
        Reports whether this process currently has a usable Exchange Online
        connection. Returns an object, never throws, never prompts.
    .DESCRIPTION
        ADDED 2026-07-24 (not moved from any existing script) -- built for
        the long-running-host concern: a host process that stays alive all
        day WILL eventually lose its EXO session to access-token expiry or
        idle session teardown, and when it does, the next cmdlet call fails
        with "no command named Get-Mailbox exists in this session," which
        reads like "you never connected" rather than "your connection
        aged out." A caller can use this to detect that state and either
        reconnect or surface an honest message.

        Two-stage check, cheapest first:
          1. Are EXO cmdlets even present in this session? (Same
             Get-Command probe Get-EXOCommand itself does.) If not, nothing
             is connected -- no point asking anything else.
          2. If ExchangeOnlineManagement exposes Get-ConnectionInformation
             (Exchange Online Management v3+), call it for live per-
             connection state, which is what actually distinguishes
             "connected" from "module loaded but session dead."

        UNVERIFIED, deliberately handled defensively: the exact property
        names and shape of Get-ConnectionInformation's output have NOT been
        confirmed against a live tenant here, and they have changed across
        module versions. So this reads State/TokenStatus only if present,
        treats their ABSENCE as "can't tell -- assume usable," and never
        throws. Practical consequence: a false "connected" is possible if
        the shape differs from expectation, in which case the underlying
        cmdlet call still fails with its own clear error, exactly as it did
        before this function existed. A false "disconnected" is not
        possible, which is the safer direction for a caller that might
        reconnect on the strength of it.
    .OUTPUTS
        PSCustomObject with:
          Connected     [bool]   -- best available answer
          CmdletsLoaded [bool]   -- stage 1 result
          Detail        [string] -- human-readable explanation
          Organization  [string] -- tenant, when discoverable ($null otherwise)
    #>
    param()
    $result = [ordered]@{
        Connected     = $false
        CmdletsLoaded = $false
        Detail        = ''
        Organization  = $null
    }

    $probe = @(Get-Command -Name 'Get-Mailbox' -All -ErrorAction SilentlyContinue |
        Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' -or
                       $_.ModuleName -eq 'ExchangeOnlineManagement' -or $_.Source -eq 'ExchangeOnlineManagement' })
    if ($probe.Count -eq 0) {
        $result.Detail = 'No Exchange Online cmdlets are present in this PowerShell process. Run Connect-ExchangeOnline in THIS process (it is per-process -- connecting elsewhere does not carry over).'
        return [PSCustomObject]$result
    }
    $result.CmdletsLoaded = $true

    $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
    if (-not $gci) {
        # Older module without Get-ConnectionInformation -- cmdlet presence
        # is the best signal available. See UNVERIFIED note above.
        $result.Connected = $true
        $result.Detail    = 'Exchange Online cmdlets are loaded. This module version has no Get-ConnectionInformation, so live session state could not be verified -- assuming usable.'
        return [PSCustomObject]$result
    }

    try {
        $conns = @(& $gci -ErrorAction Stop)
        if ($conns.Count -eq 0) {
            $result.Detail = 'Exchange Online cmdlets are loaded, but there is no active connection (the session may have expired or been disconnected). Run Connect-ExchangeOnline again in this process.'
            return [PSCustomObject]$result
        }
        # Prefer a connection that positively reports itself healthy; if no
        # connection object exposes State/TokenStatus at all, fall back to
        # "a connection object exists" rather than guessing it's dead.
        $healthy = $conns | Where-Object {
            (-not $_.PSObject.Properties['State']       -or "$($_.State)"       -eq 'Connected') -and
            (-not $_.PSObject.Properties['TokenStatus'] -or "$($_.TokenStatus)" -ne 'Expired')
        }
        $chosen = $conns | Select-Object -First 1
        if ($chosen.PSObject.Properties['Organization']) { $result.Organization = "$($chosen.Organization)" }
        if (@($healthy).Count -gt 0) {
            $result.Connected = $true
            $orgNote = ''
            if ($result.Organization) { $orgNote = " (organization: $($result.Organization))" }
            $result.Detail = "Connected to Exchange Online$orgNote."
        } else {
            $result.Detail = 'Exchange Online cmdlets are loaded, but the existing connection reports itself expired/disconnected. Run Connect-ExchangeOnline again in this process.'
        }
    } catch {
        # Get-ConnectionInformation itself failed -- don't claim
        # disconnected on the strength of that, see UNVERIFIED note.
        $result.Connected = $true
        $result.Detail    = "Exchange Online cmdlets are loaded. Get-ConnectionInformation could not be read ($($_.Exception.Message)) -- assuming usable."
    }
    return [PSCustomObject]$result
}
#endregion

#region SPO / Entra connection probes ---------------------------------------------
# ADDED 2026-07-24 -- the operator's actual daily workflow is three connects,
# every day, in Windows PowerShell 5.x:
#     Connect-ExchangeOnline
#     Connect-SPOService
#     Connect-AzureAD
# so the host shows one pill per service rather than a single "M365
# connected" light. Test-M365ExoConnection above covers the first; these two
# cover the others.
#
# A CORRECTION WORTH RECORDING, since earlier notes in this codebase
# (SRGDisable/ANZDisable headers) overstate it: the AzureAD module is
# DEPRECATED (Microsoft announced deprecation 2024-03-30), and specific
# cmdlets that call the retired Azure AD Graph endpoint do fail -- which is
# exactly what broke Get-AzureADGroup during a real Disable run. But
# Connect-AzureAD itself still authenticates, which the operator confirmed
# from daily use. "Some cmdlets are dead" is not "the module is dead," and
# the earlier note conflated them.
#
# POWERSHELL VERSION MATTERS HERE, a lot: the AzureAD module is Windows
# PowerShell 5.1 ONLY (not supported on PowerShell 7 -- confirmed against
# Microsoft's own module-compatibility docs), and
# Microsoft.Online.SharePoint.PowerShell does not work natively in PS7
# either (it needs Import-Module -UseWindowsPowerShell compatibility mode).
# Only ExchangeOnlineManagement is genuinely PS7-native. That's why the web
# host now targets 5.1+ instead of 7.0: run it in Windows PowerShell 5.1 and
# all three of these probes can actually report Connected. Run it under PS7
# (e.g. on macOS) and ExO works while these two will read NotConnected
# because their modules can't load there -- which is correct reporting, not
# a bug.

function Test-M365SpoConnection {
    <#
    .SYNOPSIS
        Reports whether SharePoint Online is connected. Never throws.
    .DESCRIPTION
        COSTS AN API CALL. Unlike Exchange Online, the SPO module exposes no
        local connection-state cmdlet (there is no SPO equivalent of
        Get-ConnectionInformation), so the only reliable probe is to make a
        cheap real call and see whether it succeeds. Get-SPOTenant is used
        because it's tenant-scoped, returns quickly, and needs no arguments.

        Because it's a real round trip, callers should NOT poll this on a
        timer -- check it on page load and on an explicit refresh. The host's
        connection route treats it as an "expensive" probe for that reason.
    .OUTPUTS
        PSCustomObject: Connected [bool], CmdletsLoaded [bool], Detail [string]
    #>
    param()
    $r = [ordered]@{ Connected = $false; CmdletsLoaded = $false; Detail = '' }

    if (-not (Get-Command -Name 'Get-SPOTenant' -ErrorAction SilentlyContinue)) {
        $r.Detail = 'SharePoint Online cmdlets are not loaded in this process. Import Microsoft.Online.SharePoint.PowerShell and run Connect-SPOService. Note this module is not natively supported on PowerShell 7 -- in PS7 it needs Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell.'
        return [PSCustomObject]$r
    }
    $r.CmdletsLoaded = $true
    try {
        $null = Get-SPOTenant -ErrorAction Stop
        $r.Connected = $true
        $r.Detail    = 'Connected to SharePoint Online.'
    } catch {
        $r.Detail = "SharePoint Online cmdlets are loaded but not connected (or the connection has expired): $($_.Exception.Message)"
    }
    return [PSCustomObject]$r
}

function Test-M365AadConnection {
    <#
    .SYNOPSIS
        Reports whether a directory connection (AzureAD, or Graph/Entra)
        exists. Never throws.
    .DESCRIPTION
        Checks in the order that matches the operator's real workflow:

        1. AzureAD -- Get-AzureADCurrentSessionInfo. This is the documented
           way to read the current AzureAD session; it throws when nothing is
           connected, so it doubles as the probe. Cheap-ish but does touch
           the service.
        2. Microsoft Graph -- Get-MgContext. Purely local (reads in-process
           state, no API call), so it's free. Checked second only because
           AzureAD is what the operator connects today; whichever answers
           first wins and .Provider says which.
        3. Microsoft.Entra -- Get-EntraContext, same idea, for the module
           that replaced AzureAD.

        .Provider lets the UI show WHICH directory module is actually
        connected rather than a vague green light -- useful precisely because
        AzureAD is deprecated and this will eventually be Graph instead.
    .OUTPUTS
        PSCustomObject: Connected [bool], Provider [string], Detail [string],
                        Account [string]
    #>
    param()
    $r = [ordered]@{ Connected = $false; Provider = $null; Detail = ''; Account = $null }
    $tried = @()

    # 1. AzureAD (the operator's current daily connect)
    if (Get-Command -Name 'Get-AzureADCurrentSessionInfo' -ErrorAction SilentlyContinue) {
        try {
            $s = Get-AzureADCurrentSessionInfo -ErrorAction Stop
            $r.Connected = $true
            $r.Provider  = 'AzureAD'
            if ($s.Account)    { $r.Account = "$($s.Account)" }
            elseif ($s.TenantId) { $r.Account = "tenant $($s.TenantId)" }
            $r.Detail = "Connected via the AzureAD module$(if ($r.Account) { " as $($r.Account)" }). Note: AzureAD is deprecated -- some cmdlets that call the retired Azure AD Graph endpoint will fail even though the connection itself is fine."
            return [PSCustomObject]$r
        } catch {
            $tried += 'AzureAD (cmdlets present, no active session)'
        }
    } else {
        $tried += 'AzureAD (module not loaded -- it is Windows PowerShell 5.1 only, so it cannot load under PowerShell 7)'
    }

    # 2. Microsoft Graph -- local check, no API call
    if (Get-Command -Name 'Get-MgContext' -ErrorAction SilentlyContinue) {
        try {
            $ctx = Get-MgContext -ErrorAction Stop
            if ($ctx) {
                $r.Connected = $true
                $r.Provider  = 'Graph'
                if ($ctx.Account) { $r.Account = "$($ctx.Account)" }
                $r.Detail = "Connected via Microsoft Graph$(if ($r.Account) { " as $($r.Account)" })."
                return [PSCustomObject]$r
            }
            $tried += 'Graph (module present, no context)'
        } catch {
            $tried += 'Graph (context read failed)'
        }
    } else {
        $tried += 'Graph (module not loaded)'
    }

    # 3. Microsoft.Entra
    if (Get-Command -Name 'Get-EntraContext' -ErrorAction SilentlyContinue) {
        try {
            $ctx = Get-EntraContext -ErrorAction Stop
            if ($ctx) {
                $r.Connected = $true
                $r.Provider  = 'Entra'
                if ($ctx.Account) { $r.Account = "$($ctx.Account)" }
                $r.Detail = "Connected via the Microsoft.Entra module$(if ($r.Account) { " as $($r.Account)" })."
                return [PSCustomObject]$r
            }
            $tried += 'Entra (module present, no context)'
        } catch {
            $tried += 'Entra (context read failed)'
        }
    } else {
        $tried += 'Entra (module not loaded)'
    }

    $r.Detail = "No directory connection found. Tried: $($tried -join '; ')."
    return [PSCustomObject]$r
}

function Get-M365ConnectionSummary {
    <#
    .SYNOPSIS
        One object describing all three service connections, for the host's
        pill row.
    .PARAMETER IncludeExpensive
        SPO's probe (and to a lesser degree the directory one) makes a real
        service call. Without this switch only the free/local checks run and
        the expensive ones report Skipped, so the host can poll cheaply on a
        timer and do the full check only on page load or an explicit refresh.
    .PARAMETER IncludeDirectory
        REVISED 2026-07-24 -- the directory check (AzureAD / Graph / Entra) is
        now OFF unless asked for.

        Reason: an audit of every script in the toolkit found that NOTHING
        calls a directory cmdlet for any functional purpose. The only such
        calls anywhere were the three read-only probes inside
        Test-M365AadConnection itself -- i.e. the pill existed solely to
        report on a connection nothing used. The operator confirmed they had
        stopped using it and only need Exchange Online and SharePoint Online.

        Skipping it by default also removes a pointless service call from
        every full refresh. The probe function is deliberately KEPT rather
        than deleted, so switching the pill back on is one flag rather than a
        rebuild, should Graph work start up later.
    #>
    param(
        [switch]$IncludeExpensive,
        [switch]$IncludeDirectory
    )

    $exo = Test-M365ExoConnection

    if ($IncludeExpensive) {
        $spo = Test-M365SpoConnection
    } else {
        $spo = [PSCustomObject]@{ Connected = $false; CmdletsLoaded = $false; Detail = 'Not checked on this pass (checking SharePoint costs a service call -- use Refresh).'; Skipped = $true }
    }

    if (-not $IncludeDirectory) {
        # Not monitored at all -- distinct from "checked and disconnected",
        # and distinct from "skipped this pass". The caller should hide the
        # pill entirely rather than show a state.
        $aad = [PSCustomObject]@{ Connected = $false; Provider = $null; Detail = 'Directory connection is not monitored (nothing in this toolkit uses it).'; Account = $null; Skipped = $true; NotMonitored = $true }
    } elseif ($IncludeExpensive) {
        $aad = Test-M365AadConnection
    } else {
        $aad = [PSCustomObject]@{ Connected = $false; Provider = $null; Detail = 'Not checked on this pass (use Refresh).'; Account = $null; Skipped = $true }
    }

    return [PSCustomObject]@{
        Exo = $exo
        Spo = $spo
        Aad = $aad
    }
}
#endregion

#region Error formatting ---------------------------------------------------------
function Format-ErrorDetail {
    <#
    .SYNOPSIS
        Uniform, detailed error text for every catch block in the toolkit.
    .DESCRIPTION
        Same level of detail everywhere: cmdlet, script line, exception
        type, message, category, inner exception if present. Returns a
        string -- the caller decides whether it lands in a TextBox, an HTML
        block, a log file, or the console.
    #>
    param(
        [Parameter(Mandatory)] $ErrorRecord,
        [string]$Context = ''
    )
    $ex   = $ErrorRecord.Exception
    $cmd  = if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.MyCommand) {
                $ErrorRecord.InvocationInfo.MyCommand.Name
            } else { '(cmdlet unknown)' }
    $line = if ($ErrorRecord.InvocationInfo) { $ErrorRecord.InvocationInfo.ScriptLineNumber } else { '?' }
    $sb = New-Object System.Text.StringBuilder
    $prefix = if ($Context) { "ERROR [$Context]" } else { 'ERROR' }
    [void]$sb.AppendLine("$prefix -- $cmd (script line $line)")
    [void]$sb.AppendLine("    Message  : $($ex.Message)")
    [void]$sb.AppendLine("    Type     : $($ex.GetType().FullName)")
    if ($ErrorRecord.CategoryInfo) {
        [void]$sb.AppendLine("    Category : $($ErrorRecord.CategoryInfo.Category) ($($ErrorRecord.CategoryInfo.Reason))")
    }
    if ($ex.InnerException) {
        [void]$sb.AppendLine("    Inner    : $($ex.InnerException.Message)")
    }
    return $sb.ToString().TrimEnd()
}
#endregion

#region Address parsing ----------------------------------------------------------
function Get-M365DefaultDomain {
    <#
    .SYNOPSIS
        Returns the CONNECTED tenant's default accepted domain, or $null.
    .DESCRIPTION
        This is the piece that makes the whole M365_ toolkit
        tenant-agnostic: the older ANZ/SRG suites each hardcoded their own
        mail-domain constant, so pointing them at a different tenant meant
        editing a file. This asks the tenant instead.

        Returns $null (never throws) if nothing is connected yet or the
        lookup fails for any other reason -- callers treat $null as "the
        operator must supply a full email address," not as fatal, because
        this runs silently every time a bare name is typed. Result is
        cached module-privately after the first success so it isn't
        re-queried per keystroke.
    #>
    if ($script:M365DefaultDomainCache) { return $script:M365DefaultDomainCache }
    try {
        $accepted = @(Invoke-EXOCommand -Name 'Get-AcceptedDomain' -Parameters @{ ErrorAction = 'Stop' })
        $default = $accepted | Where-Object { $_.Default -eq $true } | Select-Object -First 1
        if ($default) {
            $script:M365DefaultDomainCache = $default.DomainName
            return $script:M365DefaultDomainCache
        }
    } catch {
        # Not connected yet, or the lookup failed -- not fatal, and
        # deliberately not surfaced as an error. See DESCRIPTION.
    }
    return $null
}

function Clear-M365DefaultDomainCache {
    <#
    .SYNOPSIS
        Forgets the cached default domain so the next lookup re-queries.
    .DESCRIPTION
        Needed by any caller that can switch tenants within one process
        life -- most obviously a long-running browser host where the
        operator disconnects and reconnects to a different tenant. Without
        this, the cache from tenant A would silently keep appending A's
        domain to bare names typed against tenant B. Short-lived WinForms
        tools never need it.
    #>
    $script:M365DefaultDomainCache = $null
}

function Get-TargetSmtp {
    <#
    .SYNOPSIS
        Normalizes a typed name into a full SMTP address.
    .DESCRIPTION
        "bob.smith" and "bob.smith@<tenant default domain>" resolve the
        same. A domain is appended only if the operator didn't type one AND
        a default domain could actually be resolved from the connected
        tenant; otherwise returns $null so the caller can ask for a full
        address rather than guessing.
    #>
    param([string]$Text)
    if (-not $Text) { return $null }
    $t = $Text.Trim()
    if ($t -eq '') { return $null }
    if ($t -like '*@*') { return $t }
    $domain = Get-M365DefaultDomain
    if (-not $domain) { return $null }
    return "$t@$domain"
}

function ConvertTo-NormalizedUserList {
    <#
    .SYNOPSIS
        Splits a free-text multi-user box into a deduplicated list.
    .DESCRIPTION
        Tolerates newline, comma, and semicolon separators; trims; drops
        blanks; dedupes case-insensitively while preserving the order the
        operator typed. No domain-append happens here -- pipe each result
        through Get-TargetSmtp if addresses are wanted.
    #>
    param([string]$Text)
    if (-not $Text) { return @() }
    $raw = @($Text -split '[\r\n,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $result = @()
    foreach ($u in $raw) {
        if ($seen.Add($u.ToLowerInvariant())) { $result += $u }
    }
    return $result
}
#endregion

#region Object-type resolution (EXO only -- no on-prem branch) -------------------
function Resolve-Target {
    <#
    .SYNOPSIS
        Detects whether an address is a Mailbox, UnifiedGroup, or
        DistributionGroup. Cloud-only.
    .DESCRIPTION
        Tries each type in turn via EXO, collecting (not throwing) the
        failures so a caller can show why nothing matched. Type is
        'NotFound' if none of the three resolve.

        There is no on-prem 'OnpremGroup' type here, unlike the original
        on-prem-aware suite -- a cloud-only module has no synced on-prem AD
        groups to find. See the module header on the on-prem boundary.
    #>
    param([string]$Smtp)
    $result = [ordered]@{
        Type             = 'NotFound'
        Smtp             = $Smtp
        ResolutionErrors = @()
    }
    try {
        Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'Mailbox'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (EXO) resolution attempt')
    }
    try {
        Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'UnifiedGroup'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroup (EXO) resolution attempt')
    }
    try {
        Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'DistributionGroup'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup (EXO) resolution attempt')
    }
    return [PSCustomObject]$result
}
#endregion

#region Admin portal deep links ---------------------------------------------------
# ADDED 2026-07-24 -- one place for every admin-portal URL the toolkit links
# to. Centralized deliberately: portal URLs are Microsoft's to change, and
# when one breaks, fixing it should be a one-line edit here rather than a
# hunt through view files.
#
# VERIFICATION STATUS as of 2026-07-24 (checked against Microsoft Learn /
# current docs before shipping -- re-check if a link ever 404s):
#   VERIFIED  Defender Threat Explorer  https://security.microsoft.com/threatexplorerv3
#             (documented as the direct link; note the "v3" suffix -- the
#             bare /threatexplorer path is NOT the documented one)
#   VERIFIED  Defender Quarantine, Email tab
#             https://security.microsoft.com/quarantine?viewid=Email
#   PARTIAL   Entra user blade pattern
#             #view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/<section>/userId/<guid>
#             The blade name, the ~/<section>/ segment, and the userId
#             segment are confirmed in Microsoft's own docs and answers, but
#             not every section slug is. So only 'overview' is used below --
#             it's the safest, and it's also the page where BOTH follow-up
#             actions are reachable (the Block sign-in button sits on its
#             toolbar, and Groups is in its left nav). Deliberately not
#             guessing at a deeper slug like ~/Groups/ that might 404.
#   VERIFIED  Entra all-users list  https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/AllUsers
#             -- used as the fallback whenever an object GUID isn't available.
#
# The Entra links need the user's Entra object GUID, which this toolkit gets
# for free: Get-Mailbox's ExternalDirectoryObjectId IS that GUID (confirmed
# -- "ExternalDirectoryObjectId is the Microsoft Entra ID object GUID for the
# mailbox"). That matters a lot here, because it means precise per-user deep
# links work with NOTHING but the Exchange connection -- no Microsoft Graph,
# no Microsoft.Entra module, and no tenant-admin consent, which is exactly
# what wasn't available when license removal and sign-in blocking had to be
# dropped from the automated path in the first place.
# URL TEMPLATES -- operator-overridable, see Set-M365PortalUrlTemplate below.
#
# WHY TEMPLATES INSTEAD OF FIXED STRINGS: the operator asked whether the
# Defender pages could open with a filter already applied -- Quarantine
# pre-filtered to a sender, Threat Explorer pre-filled with the address
# being investigated. Both portals DO have those filters in their UI
# (Quarantine filters on Message ID / sender / recipient; Explorer on
# sender / subject / recipient / NetworkMessageID), but Microsoft publishes
# NO query-string syntax for pre-filling them, and inventing parameter names
# would produce links that quietly land on an unfiltered page -- worse than
# no link, because it looks like it worked.
#
# So: these are templates. Any template may contain {sender}, {recipient},
# or {messageId}. Get-M365PortalUrl substitutes (URL-encoded) and reports
# whether substitution actually happened, so the UI can label the link
# honestly either way. Templates ship WITHOUT placeholders -- i.e. the plain,
# documented, known-good base URLs.
#
# TO TURN ON PREPOPULATION once you've discovered the real syntax (filter in
# the portal, then read your address bar):
#     Set-M365PortalUrlTemplate -Kind Quarantine `
#         -Template 'https://security.microsoft.com/quarantine?viewid=Email&<theparam>={sender}'
# Put that call in the host script after the module import and every link in
# the toolkit starts prepopulating. No code edit, nothing to hunt for.
$script:M365PortalUrlTemplates = @{
    EntraUser      = 'https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/{objectId}'
    EntraUsers     = 'https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/AllUsers'
    ThreatExplorer = 'https://security.microsoft.com/threatexplorerv3'
    Quarantine     = 'https://security.microsoft.com/quarantine?viewid=Email'
    ExchangeAdmin  = 'https://admin.exchange.microsoft.com/'
}

function Set-M365PortalUrlTemplate {
    <#
    .SYNOPSIS
        Overrides one portal URL template at runtime.
    .DESCRIPTION
        Lets the operator enable filter prepopulation (or fix a URL Microsoft
        has moved) without editing this module. See the region comment above
        for the discovery recipe and an example.

        Placeholders understood by Get-M365PortalUrl: {sender},
        {recipient}, {messageId}, {objectId}. Values are URL-encoded on
        substitution, so don't pre-encode them in the template.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('EntraUser', 'EntraUsers', 'ThreatExplorer', 'Quarantine', 'ExchangeAdmin')]
        [string]$Kind,
        [Parameter(Mandatory)][string]$Template
    )
    $script:M365PortalUrlTemplates[$Kind] = $Template
}

function Get-M365PortalUrl {
    <#
    .SYNOPSIS
        Builds an admin-portal URL from its template. Returns a string;
        never throws.
    .PARAMETER Kind
        EntraUser        -- a specific user's Entra overview page. Needs
                            -ObjectId; falls back to EntraUsers without one.
        EntraUsers       -- the Entra all-users list.
        ThreatExplorer   -- Defender Threat Explorer.
        Quarantine       -- Defender Quarantine, Email tab.
        ExchangeAdmin    -- Exchange admin center home.
    .PARAMETER ObjectId
        Entra object GUID. For a mailbox this is Get-Mailbox's
        ExternalDirectoryObjectId.
    .PARAMETER Sender / -Recipient / -MessageId
        Substituted into {sender} / {recipient} / {messageId} IF the active
        template contains those placeholders. Ignored otherwise -- passing
        them is always safe.
    .PARAMETER Detailed
        Return an object instead of a string:
            Url            [string] the built URL
            Prepopulated   [bool]   did any placeholder actually substitute?
        UI code uses this to label a link honestly ("Open Quarantine" vs
        "Open Quarantine for this sender") rather than promising filtering
        that isn't configured.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('EntraUser', 'EntraUsers', 'ThreatExplorer', 'Quarantine', 'ExchangeAdmin')]
        [string]$Kind,
        [string]$ObjectId,
        [string]$Sender,
        [string]$Recipient,
        [string]$MessageId,
        [switch]$Detailed
    )

    # No GUID for a user-specific link -- fall back to the list rather than
    # emitting a deep link with an empty id in it.
    if ($Kind -eq 'EntraUser' -and [string]::IsNullOrWhiteSpace($ObjectId)) {
        $Kind = 'EntraUsers'
    }

    $template = $script:M365PortalUrlTemplates[$Kind]
    $url = $template
    $didSubstitute = $false

    $subs = @(
        @{ Token = '{objectId}';  Value = $ObjectId }
        @{ Token = '{sender}';    Value = $Sender }
        @{ Token = '{recipient}'; Value = $Recipient }
        @{ Token = '{messageId}'; Value = $MessageId }
    )
    foreach ($s in $subs) {
        if ($url -like "*$($s.Token)*") {
            $val = if ([string]::IsNullOrWhiteSpace($s.Value)) { '' } else { [uri]::EscapeDataString($s.Value) }
            $url = $url.Replace($s.Token, $val)
            # {objectId} is plumbing, not a user-visible filter -- don't let
            # it count as "prepopulated" or every Entra link would claim to
            # be filtered.
            if ($s.Token -ne '{objectId}' -and -not [string]::IsNullOrWhiteSpace($s.Value)) {
                $didSubstitute = $true
            }
        }
    }

    if ($Detailed) {
        return [PSCustomObject]@{ Url = $url; Prepopulated = $didSubstitute }
    }
    return $url
}

function Get-M365EntraObjectId {
    <#
    .SYNOPSIS
        Returns a mailbox's Entra object GUID, or $null.
    .DESCRIPTION
        Get-Mailbox's ExternalDirectoryObjectId is the Entra object GUID
        (see the region comment above). Returns $null rather than throwing if
        the mailbox can't be read or the property is empty -- callers use it
        to build an optional deep link, so a failure here should degrade the
        link to the users list, not break the operation that needed it.
    #>
    param([Parameter(Mandatory)][string]$Smtp)
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        if ($mbx.ExternalDirectoryObjectId) { return "$($mbx.ExternalDirectoryObjectId)" }
    } catch {
        # Deliberately swallowed -- see DESCRIPTION.
    }
    return $null
}
#endregion

#region Suite identity + tenant guard ---------------------------------------------
# ADDED 2026-07-24, when ANZW/SRGW web hosts were added alongside the M365 one.
#
# WHY THIS EXISTS: this module is tenant-agnostic by design -- Get-TargetSmtp
# asks the CONNECTED tenant for its default domain rather than reading a
# hardcoded constant. That's what lets one engine serve every tenant. But it
# also means three hosts running the same engine are indistinguishable to the
# engine itself: start the SRG host, connect to ANZ by mistake, and every tool
# will happily operate on ANZ while the window says SRG.
#
# So a host may declare which tenant it EXPECTS. Test-M365ExpectedTenant
# compares that against the connected tenant's real default domain, and the
# host surfaces a mismatch loudly. This is a guard, not a lock -- it warns and
# labels rather than blocking, because a legitimate mismatch is possible (a
# tenant with several accepted domains, for instance) and silently refusing to
# work would be worse than saying plainly what it sees.
$script:M365SuiteName    = 'M365'
$script:M365LogFileName  = 'M365-CompletedTasks.csv'
$script:M365ExpectedDomain = $null

function Set-M365SuiteIdentity {
    <#
    .SYNOPSIS
        Names the running suite and, optionally, the tenant it expects.
    .DESCRIPTION
        Called once at host startup. -LogFileName keeps each suite's completed
        task history in its own CSV, so ANZ work never interleaves with SRG's
        (the same separation the WinForms suites already maintain).
        -ExpectedDomain arms the tenant guard; omit it to leave the host
        genuinely tenant-agnostic.
    #>
    param(
        [Parameter(Mandatory)][string]$SuiteName,
        [string]$LogFileName,
        [string]$ExpectedDomain
    )
    $script:M365SuiteName = $SuiteName
    if ($LogFileName)   { $script:M365LogFileName = $LogFileName }
    if ($ExpectedDomain) { $script:M365ExpectedDomain = $ExpectedDomain }
    # A previous host in this process may have cached a different tenant's
    # domain; drop it so the guard and Get-TargetSmtp both re-resolve.
    Clear-M365DefaultDomainCache
}

function Get-M365SuiteName { return $script:M365SuiteName }

function Test-M365ExpectedTenant {
    <#
    .SYNOPSIS
        Is the connected tenant the one this suite expects? Never throws.
    .OUTPUTS
        PSCustomObject:
          Armed      [bool]   -- did the host declare an expected domain?
          Match      [bool]   -- true when it matches, or when not armed
          Expected   [string]
          Actual     [string] -- the connected tenant's default domain, or $null
          Detail     [string]
    #>
    param()
    $r = [ordered]@{ Armed = $false; Match = $true; Expected = $script:M365ExpectedDomain; Actual = $null; Detail = '' }

    if (-not $script:M365ExpectedDomain) {
        $r.Detail = 'No expected tenant declared -- this host is tenant-agnostic and will work against whatever is connected.'
        return [PSCustomObject]$r
    }
    $r.Armed = $true

    $actual = Get-M365DefaultDomain
    $r.Actual = $actual
    if (-not $actual) {
        # Can't tell yet (probably not connected). Not a mismatch -- reporting
        # one here would cry wolf before the operator has even signed in.
        $r.Detail = "Expecting $($script:M365ExpectedDomain), but the connected tenant is not known yet (connect Exchange Online first)."
        return [PSCustomObject]$r
    }
    if ($actual -ieq $script:M365ExpectedDomain) {
        $r.Detail = "Connected tenant matches this host: $actual."
    } else {
        $r.Match  = $false
        $r.Detail = "TENANT MISMATCH -- this host is for $($script:M365ExpectedDomain), but the connected tenant's default domain is $actual. Every tool here will act on $actual. Close this host, or reconnect to the right tenant."
    }
    return [PSCustomObject]$r
}
#endregion

#region Completed-tasks logging ---------------------------------------------------
function Write-CompletedLog {
    <#
    .SYNOPSIS
        Appends one completed-task row to M365-CompletedTasks.csv.
    .DESCRIPTION
        Writes to Logs\Completed\ relative to THIS MODULE's own folder --
        deliberately not the older suites' hardcoded
        E:\Scripts\M365\Logs\Completed, which is a real path on Express
        Personnel's server that won't exist on an arbitrary machine (a
        laptop pointed at a personal tenant, for instance). Note that
        $PSScriptRoot inside a module resolves to the module's directory,
        not the calling script's, so every consumer -- WinForms, web,
        scheduled job -- lands in the same log without having to agree on
        anything.
    #>
    param([string]$Task)
    $folder = Join-Path $PSScriptRoot 'Logs\Completed'
    # REVISED 2026-07-24 -- filename comes from Set-M365SuiteIdentity so each
    # suite (M365 / ANZW / SRGW) keeps its own history rather than three
    # tenants' work interleaving in one CSV. Defaults to M365-CompletedTasks.csv
    # when no host has set it.
    $file   = Join-Path $folder $script:M365LogFileName
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $file)) {
        "Task,DateTime" | Out-File -FilePath $file -Encoding utf8
    }
    $dt = Get-Date -Format "yyyy-MM-dd HHmm"
    if ($Task -and $Task.Trim() -ne "") {
        Add-Content -Path $file -Value ('"{0}","{1}"' -f $Task.Trim(), $dt)
    }
}
#endregion

#region Inbox rules --------------------------------------------------------------
function Get-InboxRulesForMailbox {
    <#
    .SYNOPSIS
        Returns a mailbox's inbox rules, or the error that prevented it.
    .DESCRIPTION
        Returns @{ Rules = @(...); Error = $null|<string> } rather than
        throwing, so a composite caller (Get-MailboxHealthSnapshot) can
        record a per-field failure and still return everything else.
    #>
    param([string]$Smtp)
    try {
        $rules = @(Invoke-EXOCommand -Name 'Get-InboxRule' -Parameters @{ Mailbox = $Smtp; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ Rules = $rules; Error = $null }
    } catch {
        return [PSCustomObject]@{ Rules = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule') }
    }
}
#endregion

#region Mailbox audit (Search-UnifiedAuditLog) ------------------------------------
# MOVED 2026-07-24 from MailboxAudit.ps1's "#region Constants" and
# "#region Data layer". Search logic and JSON parsing are unchanged; the ONLY
# behavioural change is how cancellation works -- see Invoke-M365MailboxAudit.
#
# The operation list was verified against Microsoft Learn's "Audit log
# activities" (Exchange mailbox activities table) when the WinForms tool was
# built, and is carried over verbatim in the same order as the source table.
$script:M365MailboxOperations = [ordered]@{
    'Accessed mailbox attachments'                                  = 'AttachmentAccess'
    'Accessed Mailbox folder'                                       = 'FolderBind'
    'Accessed mailbox items'                                        = 'MailItemsAccessed'
    'Accessed message'                                              = 'MessageBind'
    'Added delegate mailbox permissions'                            = 'Add-MailboxPermission'
    'Added or removed user with delegate access to calendar folder' = 'UpdateCalendarDelegation'
    'Added permissions to folder'                                   = 'AddFolderPermissions'
    'Copied messages to another folder'                             = 'Copy'
    'Created mailbox item'                                          = 'Create'
    'Created new inbox rule in Outlook web app'                     = 'New-InboxRule'
    'Delete priority cleanup label item'                            = 'PriorityCleanupDelete'
    'Deleted messages from Deleted Items folder'                    = 'SoftDelete'
    'Labeled item as a record'                                      = 'ApplyRecord'
    'Labeled message as a record'                                   = 'ApplyRecordLabel'
    'Modified folder permission'                                    = 'ModifyFolderPermissions'
    'Modified inbox rule from Outlook web app'                      = 'Set-InboxRule'
    'Moved messages to another folder'                              = 'Move'
    'Moved messages to Deleted Items folder'                        = 'MoveToDeletedItems'
    'Preserve Mailbox item'                                         = 'PreservedMailItemProactively'
    'Priority cleanup label applied'                                = 'ApplyPriorityCleanup'
    'Purged messages from the mailbox'                              = 'HardDelete'
    'Removed delegate mailbox permissions'                          = 'Remove-MailboxPermission'
    'Removed permissions from folder'                               = 'RemoveFolderPermissions'
    'Search items in a mailbox'                                     = 'SearchQueryInitiated'
    'Sent message'                                                  = 'Send'
    'Sent message using Send As permissions'                        = 'SendAs'
    'Sent message using Send On Behalf permissions'                 = 'SendOnBehalf'
    'Update item label'                                             = 'UpdateComplianceTag'
    'Updated inbox rules from Outlook client'                       = 'UpdateInboxRules'
    'Updated message'                                               = 'Update'
    'User signed in to mailbox'                                     = 'MailboxLogin'
}
# Reverse lookup (Operation value -> FriendlyName) for display.
$script:M365OperationFriendlyNames = @{}
foreach ($key in $script:M365MailboxOperations.Keys) {
    $script:M365OperationFriendlyNames[$script:M365MailboxOperations[$key]] = $key
}
# Documented ReturnLargeSet ceiling (Search-UnifiedAuditLog reference) -- used to
# detect and warn on a possibly-truncated result set rather than presenting it
# silently as complete.
$script:M365AuditResultCeiling = 50000
$script:M365AuditPageSize      = 5000

function Get-M365MailboxAuditOperations {
    <#
    .SYNOPSIS
        The FriendlyName -> Operation map, for building a picker.
    #>
    return $script:M365MailboxOperations
}

function Invoke-M365MailboxAudit {
    <#
    .SYNOPSIS
        Pages Search-UnifiedAuditLog for one mailbox. Read-only.
    .DESCRIPTION
        Pages until the service returns no further records, a short page
        arrives, or the documented 50,000-record ReturnLargeSet ceiling is hit.
        .HitCeiling tells the caller the result set may be TRUNCATED so it can
        say so, rather than presenting a capped result as complete.

        CHANGED FROM THE WINFORMS TOOL -- cancellation. The original checked a
        script-scope $script:AuditCancelRequested flag that its Cancel button
        set. That works for a modal WinForms dialog but is wrong here: the web
        host is single-threaded and processes one request at a time, so no
        second request could ever run to flip such a flag mid-search, and a
        module-level mutable flag shared across callers would be a genuine
        hazard (a stale $true from an earlier run would silently truncate the
        next one). Instead the loop guard and the ceiling govern completion, and
        -MaxPages lets a caller bound the work up front. Nothing else about the
        search changed.

        The 15-page default is inherited from the original as a defensive guard
        (15 x 5,000 = 75,000, comfortably past the 50,000 ceiling that normally
        stops it first).
    .OUTPUTS
        PSCustomObject: Records [array] (raw rows, AuditData still JSON),
                        HitCeiling [bool], Pages [int]
    #>
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][datetime]$StartUtc,
        [Parameter(Mandatory)][datetime]$EndUtc,
        [string[]]$Operations,
        [int]$MaxPages = 15
    )
    $sessionId  = "M365Audit_$([Guid]::NewGuid().ToString())"
    $all        = New-Object System.Collections.Generic.List[object]
    $hitCeiling = $false
    $iteration  = 0

    while ($iteration -lt $MaxPages) {
        $iteration++
        $params = @{
            StartDate      = $StartUtc
            EndDate        = $EndUtc
            UserIds        = $Mailbox
            SessionId      = $sessionId
            SessionCommand = 'ReturnLargeSet'
            ResultSize     = $script:M365AuditPageSize
            ErrorAction    = 'Stop'
        }
        if ($Operations -and $Operations.Count -gt 0) { $params['Operations'] = $Operations }
        $page = @(Invoke-EXOCommand -Name 'Search-UnifiedAuditLog' -Parameters $params)
        if ($page.Count -eq 0) { break }
        foreach ($row in $page) { $all.Add($row) }
        if ($all.Count -ge $script:M365AuditResultCeiling) { $hitCeiling = $true; break }
        if ($page.Count -lt $script:M365AuditPageSize) { break }   # short page -- no more data
    }
    return [PSCustomObject]@{
        Records    = $all.ToArray()
        HitCeiling = $hitCeiling
        Pages      = $iteration
    }
}

function ConvertTo-M365AuditRecord {
    <#
    .SYNOPSIS
        Flattens one raw Search-UnifiedAuditLog row into display fields.
    .DESCRIPTION
        Parses the AuditData JSON. Field names were verified against
        Microsoft's "Learn about properties in Exchange audit records" when the
        WinForms tool was built; carried over unchanged.

        Every field is read defensively: audit records vary by operation type,
        and a property missing on one operation must not blank the whole row.
        If the JSON fails to parse at all, it falls back to the top-level row
        properties rather than dropping the record.
    #>
    param($RawRecord)
    $parsed = $null
    try { $parsed = $RawRecord.AuditData | ConvertFrom-Json -ErrorAction Stop } catch { $parsed = $null }
    $opValue  = if ($parsed -and $parsed.Operation) { "$($parsed.Operation)" } else { "$($RawRecord.Operations)" }
    $friendly = if ($script:M365OperationFriendlyNames.ContainsKey($opValue)) { $script:M365OperationFriendlyNames[$opValue] } else { $opValue }
    $creationUtc = if ($parsed -and $parsed.CreationTime) {
        try {
            [datetime]::Parse($parsed.CreationTime, [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)
        } catch { $RawRecord.CreationDate }
    } else { $RawRecord.CreationDate }

    [PSCustomObject]@{
        # UTC ISO-8601 rather than the original's .ToLocalTime(): a browser
        # renders the timestamp in the viewer's own zone, so the server should
        # hand over an unambiguous instant instead of pre-converting to the
        # SERVER's local time (which would be wrong if they ever differ).
        TimeUtc           = if ($creationUtc) { ([datetime]$creationUtc).ToUniversalTime().ToString('o') } else { $null }
        Operation         = $opValue
        OperationFriendly = $friendly
        Actor             = if ($parsed) { "$($parsed.UserId)" } else { "$($RawRecord.UserIds)" }
        MailboxOwnerUPN   = if ($parsed) { "$($parsed.MailboxOwnerUPN)" } else { '' }
        ClientIP          = if ($parsed) { "$($parsed.ClientIP)" } else { '' }
        ClientInfo        = if ($parsed) { "$($parsed.ClientInfoString)" } else { '' }
        Subject           = if ($parsed -and $parsed.Item -and $parsed.Item.Subject) { "$($parsed.Item.Subject)" } else { '' }
        Folder            = if ($parsed -and $parsed.Item -and $parsed.Item.ParentFolder -and $parsed.Item.ParentFolder.Name) { "$($parsed.Item.ParentFolder.Name)" } else { '' }
        ResultStatus      = if ($parsed) { "$($parsed.ResultStatus)" } else { '' }
        RawJson           = if ($parsed) { ($parsed | ConvertTo-Json -Depth 8) } else { "$($RawRecord.AuditData)" }
    }
}

function Get-M365FullAccessHolders {
    <#
    .SYNOPSIS
        Who currently holds Full Access on a mailbox (excluding SELF/inherited).
    .DESCRIPTION
        MOVED 2026-07-24 from SharedMailboxAudit.ps1's Get-FullAccessHolders.
        Used to cross-reference audit activity against the current delegate
        list, which is the whole point of the shared-mailbox audit: "who did
        this, and are they still supposed to be able to?"

        Returns an object rather than throwing so the audit can still run and
        report the holder-list failure separately.
    #>
    param([Parameter(Mandatory)][string]$Smtp)
    try {
        $fa = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.User -notlike 'NT AUTHORITY\SELF' -and
                $_.IsInherited -eq $false -and
                $_.Deny -eq $false -and
                ($_.AccessRights -contains 'FullAccess')
            })
        return [PSCustomObject]@{ Holders = @($fa | ForEach-Object { "$($_.User)" }); Error = $null }
    } catch {
        return [PSCustomObject]@{ Holders = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access holders)') }
    }
}
#endregion

#region Inbox rule actions --------------------------------------------------------
# MOVED 2026-07-24 from InboxRules.ps1's "#region Actions - inbox rules",
# logic unchanged (unchanged there in turn from the original monolith's lines
# 807-982). Get-InboxRulesForMailbox, which reads the rules, is above.
#
# CARRYING FORWARD A REAL BUGFIX -- do not "tidy" this by adding -Mailbox back.
# All three functions accept -Smtp but deliberately do NOT pass it to the
# cmdlet. $RuleIdentity comes from Get-InboxRule -Mailbox <smtp>, and Exchange
# returns that Identity ALREADY SCOPED to the mailbox it was queried against.
# Passing -Mailbox alongside an already-scoped -Identity throws:
#     "You can't use both the Identity and Mailbox parameters to specify a
#      mailbox"
# -- confirmed against Microsoft Learn, where -Mailbox exists only to
# disambiguate a bare rule NAME. The operator hit this for real on
# Enable/Disable (fixed 2026-07-09); Remove had the identical latent bug.
# -Smtp is retained because callers use it for the completed-tasks log entry.
function Enable-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Enable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Disable-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-TargetInboxRule {
    # DESTRUCTIVE and irreversible -- a deleted inbox rule cannot be restored.
    # NO CONFIRMATION here; the caller must confirm.
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Remove-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}

function Get-InboxRuleSummaries {
    <#
    .SYNOPSIS
        Inbox rules for a mailbox, flattened for display, with the suspicious
        ones flagged.
    .DESCRIPTION
        Wraps Get-InboxRulesForMailbox and reduces each rule to the fields a UI
        actually shows, plus a Suspicious flag using the SAME criteria
        Get-MailboxHealthSnapshot uses (an ENABLED rule that forwards,
        redirects, or deletes). Keeping that rule-of-thumb in one shape means
        Mailbox Health and this view can't disagree about which rules are worth
        a second look.

        Suspicious is a prompt for human review, NOT a judgment that a rule is
        malicious -- plenty of legitimate rules forward or delete.
    .OUTPUTS
        PSCustomObject: Rules [array], Error [string|$null]
    #>
    param([string]$Smtp)
    $res = Get-InboxRulesForMailbox -Smtp $Smtp
    if ($res.Error) { return [PSCustomObject]@{ Rules = @(); Error = $res.Error } }

    $rules = @($res.Rules | ForEach-Object {
        $actions = @()
        if ($_.ForwardTo -or $_.ForwardAsAttachmentTo) { $actions += 'Forward' }
        if ($_.RedirectTo) { $actions += 'Redirect' }
        if ($_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage) { $actions += 'Delete' }
        $suspicious = ([bool]$_.Enabled) -and ($actions.Count -gt 0)

        [PSCustomObject]@{
            Name        = "$($_.Name)"
            Identity    = "$($_.Identity)"
            Enabled     = [bool]$_.Enabled
            Priority    = $_.Priority
            Description = if ($_.Description) { ($_.Description -replace '\s+', ' ').Trim() } else { '' }
            FullDescription = "$($_.Description)"
            Actions     = ($actions -join ', ')
            Suspicious  = $suspicious
        }
    })
    return [PSCustomObject]@{ Rules = $rules; Error = $null }
}
#endregion

#region Mailbox health -----------------------------------------------------------
function Get-MailboxHealthSnapshot {
    <#
    .SYNOPSIS
        Read-only composite health view of one mailbox.
    .DESCRIPTION
        Every field is fetched in its own try/catch, so one failure never
        costs the rest of the report -- partial data plus a per-field error
        list is more useful than an all-or-nothing throw. Errors accumulate
        in .Errors.

        Interpretation / verification notes, carried forward from the
        original tool unchanged because they're still true:
          - "Protocol usage" is Get-CASMailbox protocol ENABLEMENT
            (OWA/ActiveSync/POP/IMAP/MAPI/EWS), NOT traffic volume.
          - "Delegates" lists Full Access and Send As grantees by name, each
            from its own cmdlet. Does NOT include Outlook's separate
            calendar/folder delegate list.
          - DL memberships are NOT part of this snapshot -- no DL
            Memberships tool exists in the cloud toolkit yet.
          - Quota, Forwarding, Litigation Hold, and Retention Hold are all
            confirmed Get-Mailbox output properties per Microsoft Learn's
            "Property sets in Exchange Online PowerShell module cmdlets."
          - AccountDisabled: Microsoft Learn does NOT confirm Get-Mailbox
            returns this as an output property (only as a Set-Mailbox
            parameter). Still flagged UNVERIFIED -- treat a blank value as
            "unknown," NOT as "account is enabled."
          - ArchiveEnabled is derived from ArchiveGuid, not ArchiveStatus:
            per Microsoft KB4486848 ArchiveStatus can read "None" even when
            an archive is active.
          - "Suspicious Inbox Rules" flags ENABLED rules with a forwarding,
            redirect, or delete action. That's a prompt for human review,
            not a judgment that a rule is malicious.
    #>
    param([string]$Smtp)
    $health = [ordered]@{
        LastLogonTime              = $null
        ItemCount                  = $null
        MailboxSize                = $null
        ProhibitSendQuota          = $null
        ProhibitSendReceiveQuota   = $null
        ArchiveEnabled             = $null
        ArchiveSize                = $null
        AccountDisabled            = $null
        ForwardingSmtpAddress      = $null
        ForwardingAddress          = $null
        DeliverToMailboxAndForward = $null
        LitigationHoldEnabled      = $null
        RetentionHoldEnabled       = $null
        InPlaceHoldsCount          = $null
        FullAccessDelegates        = @()
        SendAsDelegates            = @()
        RuleCount                  = $null
        SuspiciousRules            = @()
        OWAEnabled                 = $null
        ActiveSyncEnabled          = $null
        PopEnabled                 = $null
        ImapEnabled                = $null
        MAPIEnabled                = $null
        EwsEnabled                 = $null
        Errors                     = @()
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.LastLogonTime = $stats.LastLogonTime
        $health.ItemCount     = $stats.ItemCount
        $health.MailboxSize   = $stats.TotalItemSize
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (health)')
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.ArchiveEnabled             = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $health.AccountDisabled            = $mbx.AccountDisabled
        $health.ProhibitSendQuota          = $mbx.ProhibitSendQuota
        $health.ProhibitSendReceiveQuota   = $mbx.ProhibitSendReceiveQuota
        $health.ForwardingSmtpAddress      = $mbx.ForwardingSmtpAddress
        $health.ForwardingAddress          = $mbx.ForwardingAddress
        $health.DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
        $health.LitigationHoldEnabled      = $mbx.LitigationHoldEnabled
        $health.RetentionHoldEnabled       = $mbx.RetentionHoldEnabled
        $health.InPlaceHoldsCount          = if ($mbx.InPlaceHolds) { @($mbx.InPlaceHolds).Count } else { 0 }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (quota/forwarding/hold/account status)')
    }
    if ($health.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $health.ArchiveSize = $astats.TotalItemSize
        } catch {
            $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (health)')
        }
    }
    try {
        $fa = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.User -notlike 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false -and $_.Deny -eq $false -and ($_.AccessRights -contains 'FullAccess') })
        $health.FullAccessDelegates = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access delegates)')
    }
    try {
        $sa = @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.Trustee -notlike 'NT AUTHORITY\SELF' -and $_.AccessControlType -eq 'Allow' })
        $health.SendAsDelegates = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As delegates)')
    }
    try {
        $rules = Get-InboxRulesForMailbox -Smtp $Smtp
        if ($rules.Error) {
            $health.Errors += $rules.Error
        } else {
            $health.RuleCount = $rules.Rules.Count
            $health.SuspiciousRules = @(
                $rules.Rules | Where-Object {
                    $_.Enabled -and (
                        $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or
                        $_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage
                    )
                } | ForEach-Object {
                    $actions = @()
                    if ($_.ForwardTo -or $_.ForwardAsAttachmentTo) { $actions += 'Forward' }
                    if ($_.RedirectTo) { $actions += 'Redirect' }
                    if ($_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage) { $actions += 'Delete' }
                    "$($_.Name)  [$($actions -join ', ')]"
                }
            )
        }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule (rule count / suspicious rule scan)')
    }
    try {
        $cas = Invoke-EXOCommand -Name 'Get-CASMailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.OWAEnabled        = $cas.OWAEnabled
        $health.ActiveSyncEnabled = $cas.ActiveSyncEnabled
        $health.PopEnabled        = $cas.PopEnabled
        $health.ImapEnabled       = $cas.ImapEnabled
        $health.MAPIEnabled       = $cas.MAPIEnabled
        $health.EwsEnabled        = $cas.EwsEnabled
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-CASMailbox (protocol access)')
    }
    return [PSCustomObject]$health
}

function Format-MailboxHealthReport {
    <#
    .SYNOPSIS
        Renders a Get-MailboxHealthSnapshot object as the plain-text report.
    .DESCRIPTION
        Pure formatting, no queries. Kept separate from the snapshot on
        purpose: a scheduled job may want the OBJECT (to filter, compare, or
        export) while an interactive UI wants this TEXT. Both UI layers use
        this identical function, so a mailbox looked up through WinForms and
        the same mailbox through the browser read exactly the same.
    #>
    param([string]$Smtp, $H)
    $sb = New-Object System.Text.StringBuilder
    $issues = @()
    if ($H.AccountDisabled -eq $true) { $issues += "Account sign-in is disabled (AccountDisabled=True) -- unverified as a Get-Mailbox field, confirm before relying on it." }
    if ($H.ForwardingSmtpAddress -or $H.ForwardingAddress) {
        $fwdTarget = if ($H.ForwardingSmtpAddress) { $H.ForwardingSmtpAddress } else { $H.ForwardingAddress }
        $fwdMode = if ($H.DeliverToMailboxAndForward) { "copy kept in mailbox AND forwarded" } else { "forwarded ONLY -- nothing kept in this mailbox" }
        $issues += "Mail forwarding is configured to '$fwdTarget' ($fwdMode)."
    }
    if ($H.LitigationHoldEnabled -eq $true) { $issues += "Litigation Hold is enabled." }
    if ($H.RetentionHoldEnabled -eq $true) { $issues += "Retention Hold is enabled." }
    if ($H.InPlaceHoldsCount -gt 0) { $issues += "$($H.InPlaceHoldsCount) In-Place/eDiscovery hold(s) applied." }
    if ($H.SuspiciousRules.Count -gt 0) { $issues += "$($H.SuspiciousRules.Count) enabled Inbox rule(s) forward, redirect, or delete mail -- see 'Suspicious Inbox Rules' below." }
    [void]$sb.AppendLine("Mailbox                : $Smtp")
    [void]$sb.AppendLine("=== Potential Issues (review before escalating) ===")
    if ($issues.Count -eq 0) { [void]$sb.AppendLine("  None of the checks below flagged anything.") }
    else { $issues | ForEach-Object { [void]$sb.AppendLine("  - $_") } }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Basics ---")
    [void]$sb.AppendLine("Last Logon             : $($H.LastLogonTime)")
    [void]$sb.AppendLine("Item Count              : $($H.ItemCount)")
    [void]$sb.AppendLine("Mailbox Size            : $($H.MailboxSize)")
    [void]$sb.AppendLine("ProhibitSendQuota       : $($H.ProhibitSendQuota)  (compare to Mailbox Size above)")
    [void]$sb.AppendLine("ProhibitSendReceiveQuota: $($H.ProhibitSendReceiveQuota)")
    [void]$sb.AppendLine("Archive Enabled         : $($H.ArchiveEnabled)")
    if ($H.ArchiveEnabled) { [void]$sb.AppendLine("Archive Size            : $($H.ArchiveSize)") }
    [void]$sb.AppendLine("Account Disabled        : $($H.AccountDisabled)  (unverified Get-Mailbox field -- see Get-MailboxHealthSnapshot help)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Forwarding & Holds ---")
    [void]$sb.AppendLine("Forwarding SMTP Address : $($H.ForwardingSmtpAddress)")
    [void]$sb.AppendLine("Forwarding Address      : $($H.ForwardingAddress)")
    [void]$sb.AppendLine("Deliver And Forward     : $($H.DeliverToMailboxAndForward)")
    [void]$sb.AppendLine("Litigation Hold         : $($H.LitigationHoldEnabled)")
    [void]$sb.AppendLine("Retention Hold          : $($H.RetentionHoldEnabled)")
    [void]$sb.AppendLine("In-Place Holds          : $($H.InPlaceHoldsCount)")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Delegates ---")
    [void]$sb.AppendLine("Full Access grantees ($($H.FullAccessDelegates.Count)):")
    if ($H.FullAccessDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
    else { $H.FullAccessDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
    [void]$sb.AppendLine("Send As grantees ($($H.SendAsDelegates.Count)):")
    if ($H.SendAsDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
    else { $H.SendAsDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Inbox Rules ---")
    [void]$sb.AppendLine("Inbox Rule Count        : $($H.RuleCount)")
    [void]$sb.AppendLine("Suspicious Inbox Rules (enabled, forward/redirect/delete):")
    if ($H.SuspiciousRules.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
    else { $H.SuspiciousRules | ForEach-Object { [void]$sb.AppendLine("  $_") } }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Protocol Access (enabled/disabled -- not traffic volume) ---")
    [void]$sb.AppendLine("  OWA        : $($H.OWAEnabled)")
    [void]$sb.AppendLine("  ActiveSync : $($H.ActiveSyncEnabled)")
    [void]$sb.AppendLine("  POP        : $($H.PopEnabled)")
    [void]$sb.AppendLine("  IMAP       : $($H.ImapEnabled)")
    [void]$sb.AppendLine("  MAPI       : $($H.MAPIEnabled)")
    [void]$sb.AppendLine("  EWS        : $($H.EwsEnabled)")
    if ($H.Errors.Count -gt 0) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("--- Field errors (partial data above is still valid) ---")
        $H.Errors | ForEach-Object { [void]$sb.AppendLine($_) }
    }
    return $sb.ToString()
}
#endregion

#region Mailbox permissions ------------------------------------------------------
function Get-MailboxState {
    <#
    .SYNOPSIS
        Core properties plus Full Access / Send As grantees for one mailbox.
    .DESCRIPTION
        Per-field try/catch, same partial-data philosophy as
        Get-MailboxHealthSnapshot. ArchiveEnabled is derived from
        ArchiveGuid rather than ArchiveStatus (Microsoft KB4486848 --
        ArchiveStatus can read "None" while an archive is active), and
        ArchiveStatus is reported alongside it for reference.
    #>
    param([string]$Smtp)
    $state = [ordered]@{
        DisplayName          = $null
        PrimarySmtpAddress   = $null
        Alias                = $null
        RecipientTypeDetails = $null
        Hidden               = $null
        LitigationHold       = $null
        ArchiveEnabled       = $null
        ArchiveStatus        = $null
        MailboxSize          = $null
        MailboxItemCount     = $null
        ArchiveSize          = $null
        ArchiveItemCount     = $null
        FullAccess           = @()
        SendAs               = @()
        Errors               = @()
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.DisplayName          = $mbx.DisplayName
        $state.PrimarySmtpAddress   = $mbx.PrimarySmtpAddress
        $state.Alias                = $mbx.Alias
        $state.RecipientTypeDetails = $mbx.RecipientTypeDetails
        $state.Hidden               = $mbx.HiddenFromAddressListsEnabled
        $state.LitigationHold       = $mbx.LitigationHoldEnabled
        $state.ArchiveEnabled       = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $state.ArchiveStatus        = $mbx.ArchiveStatus
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (core properties)')
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.MailboxSize      = $stats.TotalItemSize
        $state.MailboxItemCount = $stats.ItemCount
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (mailbox size)')
    }
    if ($state.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $state.ArchiveSize      = $astats.TotalItemSize
            $state.ArchiveItemCount = $astats.ItemCount
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (archive size)')
        }
    }
    try {
        $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.User -notlike 'NT AUTHORITY\SELF' -and
                $_.IsInherited -eq $false -and
                $_.Deny -eq $false -and
                ($_.AccessRights -contains 'FullAccess')
            }
        $state.FullAccess = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access)')
    }
    try {
        $sa = Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.Trustee -notlike 'NT AUTHORITY\SELF' -and
                $_.AccessControlType -eq 'Allow'
            }
        $state.SendAs = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As)')
    }
    return [PSCustomObject]$state
}

function Format-MailboxState {
    <#
    .SYNOPSIS
        Renders a Get-MailboxState object as plain text under a heading.
    .DESCRIPTION
        -Heading is what makes BEFORE/AFTER pairs readable; callers pass
        'CURRENT STATE', 'BEFORE', 'AFTER'.
    #>
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading (Mailbox) ---")
    [void]$sb.AppendLine("  Display Name        : $($State.DisplayName)")
    [void]$sb.AppendLine("  Primary SMTP        : $($State.PrimarySmtpAddress)")
    [void]$sb.AppendLine("  Alias               : $($State.Alias)")
    [void]$sb.AppendLine("  Mailbox Type        : $($State.RecipientTypeDetails)")
    [void]$sb.AppendLine("  Hidden From GAL     : $($State.Hidden)")
    [void]$sb.AppendLine("  Litigation Hold     : $($State.LitigationHold)")
    [void]$sb.AppendLine("  Mailbox Size        : $($State.MailboxSize)  ($($State.MailboxItemCount) items)")
    [void]$sb.AppendLine("  Archive Enabled     : $($State.ArchiveEnabled)  (ArchiveStatus: $($State.ArchiveStatus))")
    if ($State.ArchiveEnabled) {
        [void]$sb.AppendLine("  Archive Size        : $($State.ArchiveSize)  ($($State.ArchiveItemCount) items)")
    }
    [void]$sb.AppendLine("  Full Access grantees:")
    if ($State.FullAccess.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.FullAccess | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    [void]$sb.AppendLine("  Send As grantees    :")
    if ($State.SendAs.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.SendAs | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    if ($State.Errors.Count -gt 0) {
        [void]$sb.AppendLine("  --- Field errors (partial data above is still valid) ---")
        $State.Errors | ForEach-Object { [void]$sb.AppendLine("  $_") }
    }
    return $sb.ToString()
}

# The four grant/revoke primitives. NO CONFIRMATION -- see module header.
# AutoMapping is left at the cmdlet default here (unlike
# MailboxOneDriveAccess.ps1, which explicitly forces -AutoMapping:$true);
# that difference is inherited from the tool each came from and is
# deliberately not "harmonized" without the operator deciding which
# behavior should win.
function Grant-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Grant-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}

function Invoke-MailboxPermissionEdit {
    <#
    .SYNOPSIS
        Runs one grant/revoke with a BEFORE/AFTER snapshot pair and logs it.
    .DESCRIPTION
        The orchestration both UI layers were duplicating: snapshot, act,
        snapshot, log, return the combined text. NO CONFIRMATION -- the
        caller must have already confirmed (see module header).

        -Action is a scriptblock so this stays agnostic about which of the
        four primitives is being run; -ActionLabel is what appears in the
        returned text, -Verb is what goes in the completed-tasks log.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$Grantee,
        [Parameter(Mandatory)][string]$ActionLabel,
        [Parameter(Mandatory)][string]$Verb,
        [Parameter(Mandatory)][scriptblock]$Action
    )
    $out = New-Object System.Text.StringBuilder
    $before = Get-MailboxState -Smtp $Smtp
    [void]$out.Append((Format-MailboxState -Heading 'BEFORE' -State $before))
    & $Action
    [void]$out.AppendLine("$ActionLabel : SUCCESS ($Grantee)")
    $after = Get-MailboxState -Smtp $Smtp
    [void]$out.Append((Format-MailboxState -Heading 'AFTER' -State $after))
    Write-CompletedLog -Task "$Verb $Smtp (grantee: $Grantee)"
    return $out.ToString()
}
#endregion

#region Message trace ------------------------------------------------------------
function Invoke-MessageTrace {
    <#
    .SYNOPSIS
        Message trace over a date range. Sender and recipient are both
        independently optional.
    .DESCRIPTION
        Leaving both -SenderAddress and -RecipientAddress blank traces every
        message in the range, which is the behavior the tools settled on
        after dropping the old "sent from / received to THIS mailbox"
        direction radio -- a trace is often about a sender, a recipient,
        both, or neither.

        Get-MessageTraceV2 is cloud-only per Microsoft Learn, so there was
        never an on-prem concern here. ResultSize 5000 is inherited from the
        original tools.
    #>
    param(
        [string]$SenderAddress,
        [string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [string]$Subject,
        [string]$Status
    )
    $params = @{
        StartDate   = $StartDate
        EndDate     = $EndDate
        ResultSize  = 5000
        ErrorAction = 'Stop'
    }
    if ($SenderAddress)    { $params.SenderAddress    = $SenderAddress }
    if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }
    if ($Subject) {
        $params.Subject           = $Subject
        $params.SubjectFilterType = 'Contains'
    }
    if ($Status) { $params.Status = $Status }
    Invoke-EXOCommand -Name 'Get-MessageTraceV2' -Parameters $params
}

function Invoke-MessageTraceDetail {
    <#
    .SYNOPSIS
        Per-message delivery detail for one trace result.
    #>
    param(
        [Parameter(Mandatory)][guid]$MessageTraceId,
        [Parameter(Mandatory)][string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )
    Invoke-EXOCommand -Name 'Get-MessageTraceDetailV2' -Parameters @{
        MessageTraceId   = $MessageTraceId
        RecipientAddress = $RecipientAddress
        StartDate        = $StartDate
        EndDate          = $EndDate
        ErrorAction      = 'Stop'
    }
}

function Find-QuarantineMatch {
    <#
    .SYNOPSIS
        Finds quarantine entries matching a trace result's Message-ID.
    .DESCRIPTION
        Throws (rather than returning empty) when there's no Message-ID to
        match on at all -- that's a different situation from "searched and
        found nothing," and the caller should say so differently.
    #>
    param(
        [string]$MessageId,
        [string]$RecipientAddress,
        [datetime]$StartReceivedDate,
        [datetime]$EndReceivedDate
    )
    if (-not $MessageId) {
        throw "This trace result has no Message-ID value to match against quarantine -- cannot look this message up automatically."
    }
    @(Invoke-EXOCommand -Name 'Get-QuarantineMessage' -Parameters @{
        MessageId         = $MessageId
        RecipientAddress  = $RecipientAddress
        StartReceivedDate = $StartReceivedDate
        EndReceivedDate   = $EndReceivedDate
        ErrorAction       = 'Stop'
    })
}

function Get-QuarantineMessageDetail {
    <#
    .SYNOPSIS
        Full detail for one quarantined message, by its quarantine Identity.
    .DESCRIPTION
        ADDED 2026-07-24. Rationale: the operator wanted a "view in
        quarantine" option alongside Release, and the honest limitation is
        that Defender has no documented per-message deep link -- a portal
        link can only land on the quarantine LIST. But the toolkit already
        holds the quarantine Identity, so the better answer is to show the
        detail HERE rather than punt to a page the operator then has to
        search. That's usually what "view it first" actually means: enough
        context to decide whether to release.

        -Identity is the quarantine identifier in Microsoft's documented
        GUID1\GUID2 form (as returned by Get-QuarantineMessage / this
        toolkit's Find-QuarantineMatch).

        Returns a flat object of the fields worth reading before a release
        decision. Every field is read defensively -- Get-QuarantineMessage's
        exact property set varies with the quarantine type (spam vs malware
        vs phish vs mail-flow-rule), so a property missing on one type must
        not blank the whole detail view.
    .OUTPUTS
        PSCustomObject; .Error is set (and the rest left empty) on failure.
    #>
    param([Parameter(Mandatory)][string]$Identity)
    $d = [ordered]@{
        Identity         = $Identity
        Subject          = $null
        SenderAddress    = $null
        RecipientAddress = $null
        ReceivedTime     = $null
        Expires          = $null
        Type             = $null
        QuarantineTypes  = $null
        PolicyName       = $null
        PolicyType       = $null
        Direction        = $null
        ReleaseStatus    = $null
        Released         = $null
        ReleasedUsers    = $null
        MessageId        = $null
        Size             = $null
        SpamConfidence   = $null
        Error            = $null
    }
    try {
        $q = Invoke-EXOCommand -Name 'Get-QuarantineMessage' -Parameters @{ Identity = $Identity; ErrorAction = 'Stop' }
        if (-not $q) { $d.Error = "No quarantined message found with identity '$Identity' (it may have been released, deleted, or expired)."; return [PSCustomObject]$d }

        # Read each property only if present -- see DESCRIPTION on why.
        $get = {
            param($obj, $name)
            if ($obj.PSObject.Properties[$name]) { return $obj.$name }
            return $null
        }
        $d.Subject          = & $get $q 'Subject'
        $d.SenderAddress    = & $get $q 'SenderAddress'
        $d.RecipientAddress = @(& $get $q 'RecipientAddress') -join ', '
        $d.ReceivedTime     = & $get $q 'ReceivedTime'
        $d.Expires          = & $get $q 'Expires'
        $d.Type             = & $get $q 'Type'
        $d.QuarantineTypes  = & $get $q 'QuarantineTypes'
        $d.PolicyName       = & $get $q 'PolicyName'
        $d.PolicyType       = & $get $q 'PolicyType'
        $d.Direction        = & $get $q 'Direction'
        $d.ReleaseStatus    = & $get $q 'ReleaseStatus'
        $d.Released         = & $get $q 'Released'
        $d.ReleasedUsers    = @(& $get $q 'ReleasedUsers') -join ', '
        $d.MessageId        = & $get $q 'MessageId'
        $d.Size             = & $get $q 'Size'
        $d.SpamConfidence   = & $get $q 'SpamConfidenceLevel'
    } catch {
        $d.Error = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-QuarantineMessage (detail)'
    }
    return [PSCustomObject]$d
}

function Format-QuarantineMessageDetail {
    <#
    .SYNOPSIS
        Renders Get-QuarantineMessageDetail as plain text.
    #>
    param($Detail)
    if ($Detail.Error) { return $Detail.Error }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Subject          : $($Detail.Subject)")
    [void]$sb.AppendLine("From             : $($Detail.SenderAddress)")
    [void]$sb.AppendLine("To               : $($Detail.RecipientAddress)")
    [void]$sb.AppendLine("Received         : $($Detail.ReceivedTime)")
    [void]$sb.AppendLine("Expires from qtn : $($Detail.Expires)")
    [void]$sb.AppendLine("Direction        : $($Detail.Direction)")
    [void]$sb.AppendLine("Quarantine type  : $($Detail.Type)  $(if ($Detail.QuarantineTypes) { "($($Detail.QuarantineTypes))" })")
    [void]$sb.AppendLine("Policy           : $($Detail.PolicyName)  $(if ($Detail.PolicyType) { "[$($Detail.PolicyType)]" })")
    [void]$sb.AppendLine("Spam confidence  : $($Detail.SpamConfidence)")
    [void]$sb.AppendLine("Size             : $($Detail.Size)")
    [void]$sb.AppendLine("Release status   : $($Detail.ReleaseStatus)   Released: $($Detail.Released)")
    if ($Detail.ReleasedUsers) { [void]$sb.AppendLine("Released to      : $($Detail.ReleasedUsers)") }
    [void]$sb.AppendLine("Message-ID       : $($Detail.MessageId)")
    [void]$sb.AppendLine("Quarantine ID    : $($Detail.Identity)")
    return $sb.ToString()
}

function Get-QuarantineMessagePreview {
    <#
    .SYNOPSIS
        Previews a quarantined message's contents.
    .DESCRIPTION
        Wraps Preview-QuarantineMessage (verified against Microsoft Learn --
        signature is -Identity <GUID1\GUID2> with optional -RecipientAddress
        and -EntityType).

        Per Microsoft's own documentation this does NOT work on files
        quarantined by Safe Attachments for SharePoint/OneDrive/Teams, so a
        failure here is not necessarily a bug -- the returned .Error says so
        rather than pretending the message is unreadable.

        READ-ONLY. Previewing does not release, deliver, or alter the message
        in any way.
    #>
    param(
        [Parameter(Mandatory)][string]$Identity,
        [string]$RecipientAddress
    )
    $p = @{ Identity = $Identity; ErrorAction = 'Stop' }
    if ($RecipientAddress) { $p.RecipientAddress = $RecipientAddress }
    try {
        $result = Invoke-EXOCommand -Name 'Preview-QuarantineMessage' -Parameters $p
        return [PSCustomObject]@{ Preview = $result; Error = $null }
    } catch {
        return [PSCustomObject]@{
            Preview = $null
            Error   = (Format-ErrorDetail -ErrorRecord $_ -Context 'Preview-QuarantineMessage (note: previewing is not supported for Safe Attachments files from SharePoint/OneDrive/Teams)')
        }
    }
}

function Invoke-QuarantineRelease {
    <#
    .SYNOPSIS
        Releases one quarantined message to ALL original recipients.
    .DESCRIPTION
        NO CONFIRMATION -- see module header. -ReleaseToAll is inherited
        from the original tools; this releases to every original recipient,
        not just the one whose trace row was selected.
    #>
    param([Parameter(Mandatory)]$QuarantineIdentity)
    Invoke-EXOCommand -Name 'Release-QuarantineMessage' -Parameters @{
        Identity     = $QuarantineIdentity
        ReleaseToAll = $true
        Confirm      = $false
        ErrorAction  = 'Stop'
    } | Out-Null
}

function Get-TraceTimeZone {
    <#
    .SYNOPSIS
        Maps a friendly label to a TimeZoneInfo, falling back to UTC.
    .DESCRIPTION
        Windows time-zone IDs, so this is for the WinForms/5.1 side. (The
        browser layer converts client-side with IANA names instead --
        America/Chicago rather than 'Central Standard Time' -- because
        Windows IDs aren't available in a browser.) Never throws; an
        unknown label or a missing zone yields UTC.
    #>
    param([string]$Label)
    $ids = @{
        'Eastern'  = 'Eastern Standard Time'
        'Central'  = 'Central Standard Time'
        'Mountain' = 'Mountain Standard Time'
        'Pacific'  = 'Pacific Standard Time'
        'UTC'      = 'UTC'
    }
    $id = $ids[$Label]
    if (-not $id) { $id = 'UTC' }
    try {
        return [System.TimeZoneInfo]::FindSystemTimeZoneById($id)
    } catch {
        return [System.TimeZoneInfo]::Utc
    }
}

function ConvertTo-TraceDisplayTime {
    <#
    .SYNOPSIS
        Converts a UTC trace timestamp into a target zone for display.
    .DESCRIPTION
        Trace timestamps come back UTC but not always Kind-tagged as such,
        so this forces DateTimeKind.Utc before converting -- otherwise the
        framework may treat an untagged value as local and double-shift it.
        Returns the input unchanged if conversion fails.
    #>
    param($UtcValue, [System.TimeZoneInfo]$TimeZone)
    if (-not $UtcValue) { return $UtcValue }
    try {
        $forced = [DateTime]::SpecifyKind([DateTime]$UtcValue, [DateTimeKind]::Utc)
        return [System.TimeZoneInfo]::ConvertTimeFromUtc($forced, $TimeZone)
    } catch {
        return $UtcValue
    }
}

function Format-TraceDetail {
    <#
    .SYNOPSIS
        Renders Invoke-MessageTraceDetail output as plain text.
    #>
    param($Details, [System.TimeZoneInfo]$TimeZone)
    $sb = New-Object System.Text.StringBuilder
    $rows = @($Details)
    if ($rows.Count -eq 0) {
        [void]$sb.AppendLine("(no detail records returned)")
        return $sb.ToString()
    }
    foreach ($d in $rows) {
        $displayDate = if ($TimeZone) { ConvertTo-TraceDisplayTime -UtcValue $d.Date -TimeZone $TimeZone } else { $d.Date }
        [void]$sb.AppendLine("$displayDate  [$($d.Event)]  $($d.Action)")
        if ($d.Detail) { [void]$sb.AppendLine("    $($d.Detail)") }
    }
    return $sb.ToString()
}
#endregion

#region Calendar permissions ------------------------------------------------------
# MOVED 2026-07-24 from CalendarPermissions.ps1's "#region Actions - calendar
# permissions", logic unchanged. That file in turn had them unchanged from the
# original Mailbox-Group-Inspector.ps1 monolith (lines 1272-1455), so this is
# the same code that's been in production -- only its home has changed.
# EXO-only; no on-prem dependency.
function Get-CalendarFolderIdentity {
    <#
    .SYNOPSIS
        Resolves a mailbox's actual calendar folder identity.
    .DESCRIPTION
        Resolves the real calendar folder rather than hardcoding ":\Calendar",
        since the folder can be renamed or (rarely) localized. Falls back to
        the first Calendar-scope folder if none matches the default root path.
    #>
    param([string]$Smtp)
    $folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{ Identity = $Smtp; FolderScope = 'Calendar'; ErrorAction = 'Stop' })
    if ($folders.Count -eq 0) { throw "No Calendar-scope folder found for $Smtp." }
    $primary = $folders | Where-Object { $_.FolderPath -eq '/Calendar' } | Select-Object -First 1
    if (-not $primary) { $primary = $folders | Select-Object -First 1 }
    return "$($Smtp):\$($primary.Name)"
}

function Get-CalendarPermissions {
    <#
    .SYNOPSIS
        Current calendar folder permissions, or the error that prevented it.
    .DESCRIPTION
        Returns an object rather than throwing so a UI can show the failure
        alongside the rest of its state.
    #>
    param([string]$Smtp)
    try {
        $calId = Get-CalendarFolderIdentity -Smtp $Smtp
        $perms = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $calId; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ CalendarIdentity = $calId; Permissions = $perms; Error = $null }
    } catch {
        return [PSCustomObject]@{ CalendarIdentity = $null; Permissions = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxFolderPermission (Calendar)') }
    }
}

function Set-CalendarPermission {
    <#
    .SYNOPSIS
        Grants or updates one user's calendar role. NO CONFIRMATION.
    .DESCRIPTION
        Add-MailboxFolderPermission fails if the user already has an entry, so
        this checks first (best-effort) rather than parsing the resulting error
        text. If the check itself errors for any reason, it falls through to
        Add-; a genuine problem will surface clearly from that call instead.

        Returns 'Added' or 'Updated' so the caller can report which happened.
    #>
    param([string]$CalendarIdentity, [string]$Grantee, [string]$Role)
    $existing = @()
    try {
        $existing = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $CalendarIdentity; User = $Grantee; ErrorAction = 'Stop' })
    } catch {
        $existing = @()
    }
    if ($existing.Count -gt 0) {
        Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Updated'
    } else {
        Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Added'
    }
}

function Remove-CalendarPermission {
    <#
    .SYNOPSIS
        Removes a user's calendar permissions entirely. NO CONFIRMATION.
    .DESCRIPTION
        Remove-MailboxFolderPermission removes ALL permissions for this user on
        the folder -- there is no selective/partial removal, per Microsoft
        Learn. The caller must confirm before calling this.
    #>
    param([string]$CalendarIdentity, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxFolderPermission' -Parameters @{
        Identity = $CalendarIdentity; User = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}

# The documented MailboxFolderPermission roles, for populating a UI dropdown
# in one place instead of each front end keeping its own list.
function Get-CalendarPermissionRoles {
    return @('Owner', 'PublishingEditor', 'Editor', 'PublishingAuthor', 'Author',
             'NonEditingAuthor', 'Reviewer', 'Contributor', 'AvailabilityOnly',
             'LimitedDetails', 'None')
}
#endregion

#region Out of office -------------------------------------------------------------
# MOVED 2026-07-24 from OOFEditor.ps1's "#region Actions - out of office",
# logic unchanged (and unchanged there from the original monolith's lines
# 1108-1269). EXO-only.
function Get-OutOfOfficeConfig {
    <#
    .SYNOPSIS
        Reads a mailbox's auto-reply (Out of Office) configuration.
    #>
    param([string]$Smtp)
    Invoke-EXOCommand -Name 'Get-MailboxAutoReplyConfiguration' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
}

function Set-OutOfOfficeConfig {
    <#
    .SYNOPSIS
        Writes a mailbox's auto-reply configuration. NO CONFIRMATION.
    .DESCRIPTION
        StartTime/EndTime are only sent when the state is 'Scheduled' --
        passing them otherwise is meaningless and Exchange may reject it.
        The caller must confirm before calling this.
    #>
    param(
        [string]$Smtp,
        [string]$AutoReplyState,
        [string]$ExternalAudience,
        [string]$InternalMessage,
        [string]$ExternalMessage,
        [Nullable[DateTime]]$StartTime,
        [Nullable[DateTime]]$EndTime
    )
    $params = @{
        Identity         = $Smtp
        AutoReplyState   = $AutoReplyState
        ExternalAudience = $ExternalAudience
        InternalMessage  = $InternalMessage
        ExternalMessage  = $ExternalMessage
        Confirm          = $false
        ErrorAction      = 'Stop'
    }
    if ($AutoReplyState -eq 'Scheduled') {
        if ($StartTime) { $params.StartTime = $StartTime }
        if ($EndTime)   { $params.EndTime   = $EndTime }
    }
    Invoke-EXOCommand -Name 'Set-MailboxAutoReplyConfiguration' -Parameters $params | Out-Null
}

# Valid dropdown values, kept here so both front ends agree.
function Get-OutOfOfficeStates    { return @('Disabled', 'Enabled', 'Scheduled') }
function Get-OutOfOfficeAudiences { return @('None', 'Known', 'All') }
#endregion

#region OneDrive / SharePoint Online ---------------------------------------------
function Get-OneDrivePersonalSiteUrl {
    <#
    .SYNOPSIS
        Builds a user's OneDrive personal-site URL from their address.
    .DESCRIPTION
        UNVERIFIED / FLAGGED ASSUMPTION, carried forward from the original
        tools: this URL is ALGORITHMICALLY CONSTRUCTED (every '.' and '@'
        replaced with '_'), not looked up via any API. It matches
        Microsoft's standard personal-site pattern and matched the
        operator's real example exactly, but if a user's OneDrive was
        provisioned under a different alias than their current primary SMTP
        (e.g. after a mailbox rename), the constructed URL can be wrong.
        Set-SPOUser then fails with a clear "site not found"-style error
        rather than silently acting on the wrong site -- but callers should
        surface the constructed URL so the operator can eyeball it.

        -SpoTenantPrefix is REQUIRED and has no default on purpose. It is
        genuinely tenant-specific and NOT reliably derivable from the mail
        domain -- the original ANZ tool hardcoded 'expressprosau' precisely
        because stripping the TLD from 'expresspros.com.au' would have
        produced the wrong value. Guessing here would point a
        site-collection-admin grant at the wrong tenant's SharePoint, so the
        caller must supply it.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$SpoTenantPrefix
    )
    $slug = ($Smtp -replace '\.', '_') -replace '@', '_'
    return "https://$SpoTenantPrefix-my.sharepoint.com/personal/$slug"
}

function Set-M365OneDriveSiteAdmin {
    <#
    .SYNOPSIS
        Grants or removes site-collection-admin on a OneDrive personal site.
    .DESCRIPTION
        NO CONFIRMATION -- see module header. Requires
        Microsoft.Online.SharePoint.PowerShell (Connect-SPOService) in this
        process; that module is NOT needed by any other function here.

        Set-SPOUser is called by bare name rather than through
        Invoke-EXOCommand -- it's a SharePoint cmdlet with no Exchange
        Online equivalent to collide with, so the EXO collision guard
        doesn't apply.

        UNVERIFIED: whether Microsoft.Online.SharePoint.PowerShell loads at
        all under PowerShell 7 cross-platform. It has historically been a
        Windows PowerShell 5.1-oriented module. If it doesn't, this specific
        function is the piece that would need rewriting against
        PnP.PowerShell (different cmdlets and auth flow -- a real rewrite,
        not a substitution). Everything else in this module is EXO-only and
        unaffected either way.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$LoginName,
        [Parameter(Mandatory)][bool]$IsSiteCollectionAdmin
    )
    Set-SPOUser -Site $SiteUrl -LoginName $LoginName -IsSiteCollectionAdmin $IsSiteCollectionAdmin -ErrorAction Stop | Out-Null
}
#endregion

#region Mailbox + OneDrive access (bulk) ------------------------------------------
# MOVED 2026-07-24 from MailboxOneDriveAccess.ps1 (the merged
# Set-MailboxAccess + onedriveaccess tool), logic unchanged.
#
# SHAPE: many-to-many. Every Mailbox in one list gets the selected access
# granted to every User/Manager in the other -- a full cross product, which is
# what both original tools did.
#
# The five Access Types are five COMPLETE, mutually exclusive end states, not
# independent toggles: Full, SendAs, OneDrive, FullSendAs, FullOneDrive.
# Send-As + OneDrive is deliberately absent per the operator's rule ("there
# should never be a Send-As, OneDrive need") -- it isn't validated against,
# it simply isn't one of the five, so it can't be requested.
function Get-M365AccessFlags {
    <#
    .SYNOPSIS
        Turns an Access Type name into three booleans.
    .DESCRIPTION
        Single source of truth for what each Access Type does. Both the
        pre-flight gate and the action loop read these same flags, so there's
        no second list anywhere that could drift.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Full', 'SendAs', 'OneDrive', 'FullSendAs', 'FullOneDrive')]
        [string]$AccessType
    )
    switch ($AccessType) {
        'Full'         { return [PSCustomObject]@{ Full = $true;  SendAs = $false; OneDrive = $false } }
        'SendAs'       { return [PSCustomObject]@{ Full = $false; SendAs = $true;  OneDrive = $false } }
        'OneDrive'     { return [PSCustomObject]@{ Full = $false; SendAs = $false; OneDrive = $true  } }
        'FullSendAs'   { return [PSCustomObject]@{ Full = $true;  SendAs = $true;  OneDrive = $false } }
        'FullOneDrive' { return [PSCustomObject]@{ Full = $true;  SendAs = $false; OneDrive = $true  } }
    }
}

function Get-M365AccessTypeText {
    param([Parameter(Mandatory)][string]$AccessType)
    switch ($AccessType) {
        'Full'         { return 'Full Access' }
        'SendAs'       { return 'Send-As' }
        'OneDrive'     { return 'OneDrive' }
        'FullSendAs'   { return 'Full Access & Send-As' }
        'FullOneDrive' { return 'Full Access & OneDrive' }
        default        { return $AccessType }
    }
}

function Invoke-M365MailboxOneDriveAccess {
    <#
    .SYNOPSIS
        Grants or removes Full Access / Send-As / OneDrive admin across a
        Mailboxes x Users cross product. DESTRUCTIVE on Remove. NO
        CONFIRMATION.
    .DESCRIPTION
        Behaviour preserved from the merged WinForms tool:

        PRE-FLIGHT runs only when the Access Type actually touches Exchange
        (Full and/or Send-As). Every Mailbox AND every User must resolve via
        Get-Mailbox first; if any fails, NOTHING is changed at all -- not
        Exchange, not OneDrive. A OneDrive-only run skips pre-flight entirely,
        since it needs no Exchange connection.

        MessageCopyForSentAs is set once per mailbox whenever Send-As is being
        GRANTED (standalone or combined) so Send-As mail lands in Sent Items.
        It is NEVER cleared on Remove, in any combination -- that was explicit
        in the original and is preserved.

        Already-satisfied operations are reported as SKIP, not failures
        (removing access that isn't there is not an error).

        -SpoTenantPrefix is required only when the Access Type includes
        OneDrive. It cannot be derived from the mail domain -- see
        Get-OneDrivePersonalSiteUrl -- so it must be supplied.
    .OUTPUTS
        PSCustomObject:
          Text                 [string] the full run log
          OneDriveGrantedFor   [string[]] mailboxes where at least one OneDrive
                               grant succeeded -- lets a UI surface the
                               OneDrive URL only for access that really exists
          Succeeded / Failed / Skipped [int] counts
    #>
    param(
        [Parameter(Mandatory)][string[]]$Mailboxes,
        [Parameter(Mandatory)][string[]]$Users,
        [Parameter(Mandatory)]
        [ValidateSet('Full', 'SendAs', 'OneDrive', 'FullSendAs', 'FullOneDrive')]
        [string]$AccessType,
        [Parameter(Mandatory)][ValidateSet('Grant', 'Remove')][string]$Mode,
        [string]$SpoTenantPrefix
    )

    $flags   = Get-M365AccessFlags -AccessType $AccessType
    $isGrant = ($Mode -eq 'Grant')
    $verb    = if ($isGrant) { 'Granted' } else { 'Removed' }
    $actionText = "$Mode $(Get-M365AccessTypeText -AccessType $AccessType)"

    $log = New-Object System.Text.StringBuilder
    [void]$log.AppendLine("=== Mailbox & OneDrive Access ===")
    [void]$log.AppendLine("Action      : $actionText")
    [void]$log.AppendLine("Mailbox(es) : $($Mailboxes -join ', ')")
    [void]$log.AppendLine("User(s)     : $($Users -join ', ')")
    [void]$log.AppendLine("")

    if ($flags.OneDrive -and [string]::IsNullOrWhiteSpace($SpoTenantPrefix)) {
        [void]$log.AppendLine("STOPPED -- this Access Type includes OneDrive but no SharePoint tenant prefix was supplied, and it cannot be guessed safely. Nothing was changed.")
        return [PSCustomObject]@{ Text = $log.ToString(); OneDriveGrantedFor = @(); Succeeded = 0; Failed = 0; Skipped = 0 }
    }

    # --- Pre-flight, Exchange-touching runs only.
    if ($flags.Full -or $flags.SendAs) {
        $notFound = @()
        foreach ($m in $Mailboxes) {
            try { Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $m; ErrorAction = 'Stop' } | Out-Null }
            catch { $notFound += $m }
        }
        foreach ($u in $Users) {
            try { Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $u; ErrorAction = 'Stop' } | Out-Null }
            catch { $notFound += $u }
        }
        if ($notFound.Count -gt 0) {
            $list = (($notFound | Sort-Object -Unique) -join "`r`n    ")
            [void]$log.AppendLine("PRE-FLIGHT FAILED -- these could not be resolved:")
            [void]$log.AppendLine("    $list")
            [void]$log.AppendLine("")
            [void]$log.AppendLine("No changes were made. Correct the above and try again.")
            return [PSCustomObject]@{ Text = $log.ToString(); OneDriveGrantedFor = @(); Succeeded = 0; Failed = 0; Skipped = 0 }
        }
        [void]$log.AppendLine("Pre-flight  : all identities resolved OK.")
        [void]$log.AppendLine("")
    }

    $succeeded = 0; $failed = 0; $skipped = 0
    $oneDriveOk = New-Object System.Collections.Generic.List[string]

    foreach ($m in $Mailboxes) {

        if ($flags.SendAs -and $isGrant) {
            try {
                Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{ Identity = $m; MessageCopyForSentAs = $true; ErrorAction = 'Stop' } | Out-Null
                [void]$log.AppendLine("  [OK]   MessageCopyForSentAs = true on $m")
            } catch {
                [void]$log.AppendLine("  [WARN] Could not set MessageCopyForSentAs on $m : $($_.Exception.Message)")
            }
        }

        $siteUrl = $null
        if ($flags.OneDrive) {
            $siteUrl = Get-OneDrivePersonalSiteUrl -Smtp $m -SpoTenantPrefix $SpoTenantPrefix
            [void]$log.AppendLine("  OneDrive site for $m : $siteUrl")
        }

        foreach ($u in $Users) {

            if ($flags.Full) {
                if ($isGrant) {
                    try {
                        Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
                            Identity = $m; User = $u; AccessRights = 'FullAccess'; AutoMapping = $true; ErrorAction = 'Stop'
                        } | Out-Null
                        [void]$log.AppendLine("  [OK]   Granted Full Access : $u -> $m")
                        $succeeded++
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Full Access : $u -> $m : $($_.Exception.Message)")
                        $failed++
                    }
                } else {
                    try {
                        $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $m; User = $u; ErrorAction = 'SilentlyContinue' } |
                              Where-Object { $_.AccessRights -contains 'FullAccess' }
                        if ($fa) {
                            Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
                                Identity = $m; User = $u; AccessRights = 'FullAccess'; InheritanceType = 'All'; Confirm = $false; ErrorAction = 'Stop'
                            } | Out-Null
                            [void]$log.AppendLine("  [OK]   Removed Full Access : $u -> $m")
                            $succeeded++
                        } else {
                            [void]$log.AppendLine("  [SKIP] No Full Access found : $u -> $m")
                            $skipped++
                        }
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Remove Full Access : $u -> $m : $($_.Exception.Message)")
                        $failed++
                    }
                }
            }

            if ($flags.SendAs) {
                if ($isGrant) {
                    try {
                        Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
                            Identity = $m; Trustee = $u; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
                        } | Out-Null
                        [void]$log.AppendLine("  [OK]   Granted Send-As : $u -> $m")
                        $succeeded++
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Send-As : $u -> $m : $($_.Exception.Message)")
                        $failed++
                    }
                } else {
                    try {
                        $sa = Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $m; Trustee = $u; ErrorAction = 'SilentlyContinue' } |
                              Where-Object { $_.AccessRights -contains 'SendAs' }
                        if ($sa) {
                            Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
                                Identity = $m; Trustee = $u; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
                            } | Out-Null
                            [void]$log.AppendLine("  [OK]   Removed Send-As : $u -> $m")
                            $succeeded++
                        } else {
                            [void]$log.AppendLine("  [SKIP] No Send-As found : $u -> $m")
                            $skipped++
                        }
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Remove Send-As : $u -> $m : $($_.Exception.Message)")
                        $failed++
                    }
                }
            }

            if ($flags.OneDrive) {
                try {
                    Set-M365OneDriveSiteAdmin -SiteUrl $siteUrl -LoginName $u -IsSiteCollectionAdmin $isGrant
                    [void]$log.AppendLine("  [OK]   $verb OneDrive admin : $u -> $m")
                    $succeeded++
                    if ($isGrant -and -not $oneDriveOk.Contains($m)) { $oneDriveOk.Add($m) }
                } catch {
                    [void]$log.AppendLine("  [FAIL] OneDrive admin : $u -> $m : $($_.Exception.Message)")
                    $failed++
                }
            }
        }
    }

    [void]$log.AppendLine("")
    [void]$log.AppendLine("Summary : $succeeded succeeded, $failed failed, $skipped with nothing to do.")

    if ($succeeded -gt 0) {
        $userList = ($Users -join '+')
        foreach ($m in $Mailboxes) {
            Write-CompletedLog -Task "$verb MailboxOneDriveAccess [$actionText] $m (users: $userList)"
        }
    }

    return [PSCustomObject]@{
        Text               = $log.ToString()
        OneDriveGrantedFor = @($oneDriveOk)
        Succeeded          = $succeeded
        Failed             = $failed
        Skipped            = $skipped
    }
}
#endregion

#region Disable / termination ----------------------------------------------------
function Disable-M365User {
    <#
    .SYNOPSIS
        Runs the termination steps for ONE user. Destructive. No
        confirmation.
    .DESCRIPTION
        NO CONFIRMATION -- see module header. The caller must confirm before
        calling this, and every existing caller does.

        FOUR STEPS, EACH ATTEMPTED INDEPENDENTLY -- a failure in one step
        does not skip the others for that user, same partial-success
        philosophy as the snapshot functions:
          1+2. Hide from address lists AND convert to a shared mailbox, in
               one Set-Mailbox call.
          3.   If a manager was supplied and -GrantMailboxAccess: mailbox
               FullAccess only. This NEVER grants Send-As -- per the
               operator's explicit instruction there is no
               Add-RecipientPermission call in this function at all.
          4.   If a manager was supplied and -GrantOneDriveAccess:
               site-collection-admin on the user's OneDrive so the manager
               can browse/copy files. Requires -SpoTenantPrefix.

        NOT DONE HERE, deliberately: license-group removal and sign-in
        block. Both were originally automated via the AzureAD module, which
        is permanently dead (Microsoft retired the underlying Azure AD Graph
        API on 2025-08-31), then briefly via Microsoft.Entra, which needed
        one-time tenant-admin consent that wasn't available. The function
        reports both as NOT ATTEMPTED in its output and log line so the
        caller can remind the operator to handle them manually. If a tenant
        ever has that consent available, re-adding them is a scope decision
        for the operator, not something to quietly switch on.

        CMDLET VERIFICATION carried forward: Set-Mailbox -Type Shared and
        -HiddenFromAddressListsEnabled are both confirmed Set-Mailbox
        parameters, and per Microsoft's "Convert a mailbox" documentation
        existing mail, folders, and permissions are retained across the
        Shared conversion. Add-MailboxPermission -AccessRights FullAccess
        -InheritanceType All is the same call this codebase already proved
        in its mailbox-permissions tools. Set-SPOUser
        -IsSiteCollectionAdmin was confirmed against Microsoft Learn.
    .OUTPUTS
        REVISED 2026-07-24 -- now returns a PSCustomObject, not a bare
        string:
            Smtp      [string] -- the user processed
            ObjectId  [string] -- their Entra object GUID, or $null
            Text      [string] -- the same per-user report text as before
            Statuses  [hashtable] -- HideShared / FullAccess / OneDrive
        The string is still there as .Text, so the only change a caller needs
        is reading .Text instead of using the return value directly.

        WHY: the two steps this function deliberately does NOT automate
        (license-group removal and sign-in blocking) have to be finished by
        hand in the Entra portal, and the operator asked for the tool to link
        them per user rather than just naming them. A precise Entra deep link
        needs the object GUID, so it's captured here -- once, alongside work
        already being done -- and handed back for the UI to turn into a link.
        Also writes one completed-tasks log row, unchanged.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [string]$ManagerSmtp,
        [bool]$GrantMailboxAccess = $false,
        [bool]$GrantOneDriveAccess = $false,
        [string]$SpoTenantPrefix
    )
    $out = New-Object System.Text.StringBuilder
    [void]$out.AppendLine("=== $Smtp ===")

    # Grab the Entra object GUID BEFORE the mailbox is altered, so the
    # follow-up deep link is available even if a later step fails. Non-fatal
    # by design -- a missing GUID just means the UI links to the users list.
    $objectId = Get-M365EntraObjectId -Smtp $Smtp

    # Steps 1+2: hide + convert to shared, in one call.
    $statusHideShared = 'FAILED'
    try {
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $Smtp; HiddenFromAddressListsEnabled = $true; Type = 'Shared'; ErrorAction = 'Stop'
        } | Out-Null
        $statusHideShared = 'OK'
        [void]$out.AppendLine("  Hide + convert to shared mailbox: OK")
    } catch {
        [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context 'Hide + convert to shared'))
    }

    # Step 3: manager Full Access (never Send-As).
    $statusFullAccess = 'SKIPPED'
    if ($ManagerSmtp -and $GrantMailboxAccess) {
        try {
            Grant-FullAccessPermission -Smtp $Smtp -Grantee $ManagerSmtp
            $statusFullAccess = 'OK'
            [void]$out.AppendLine("  Manager Full Access ($ManagerSmtp): OK")
        } catch {
            $statusFullAccess = 'FAILED'
            [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context 'Manager Full Access'))
        }
    } else {
        $reason = if (-not $ManagerSmtp) { 'no manager supplied' } else { 'mailbox access not requested' }
        [void]$out.AppendLine("  Manager Full Access: SKIPPED ($reason)")
    }

    # Step 4: manager OneDrive site-collection admin.
    $statusOneDrive = 'SKIPPED'
    if ($ManagerSmtp -and $GrantOneDriveAccess) {
        if (-not $SpoTenantPrefix) {
            $statusOneDrive = 'FAILED'
            [void]$out.AppendLine("  Manager OneDrive access: FAILED -- OneDrive access was requested but no SharePoint tenant prefix was supplied, and it cannot be guessed safely (see Get-OneDrivePersonalSiteUrl help).")
        } else {
            $odUrl = Get-OneDrivePersonalSiteUrl -Smtp $Smtp -SpoTenantPrefix $SpoTenantPrefix
            try {
                Set-M365OneDriveSiteAdmin -SiteUrl $odUrl -LoginName $ManagerSmtp -IsSiteCollectionAdmin $true
                $statusOneDrive = 'OK'
                [void]$out.AppendLine("  Manager OneDrive access ($ManagerSmtp -> $odUrl): OK")
            } catch {
                $statusOneDrive = 'FAILED'
                [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Manager OneDrive access ($odUrl)"))
            }
        }
    } else {
        $reason = if (-not $ManagerSmtp) { 'no manager supplied' } else { 'OneDrive access not requested' }
        [void]$out.AppendLine("  Manager OneDrive access: SKIPPED ($reason)")
    }

    [void]$out.AppendLine("  License removal: NOT ATTEMPTED -- handle manually.")
    [void]$out.AppendLine("  Sign-in block: NOT ATTEMPTED -- handle manually.")

    $managerLabel = if ($ManagerSmtp) { $ManagerSmtp } else { '(none)' }
    Write-CompletedLog -Task "Disabled $Smtp (manager: $managerLabel; HideShared=$statusHideShared; FullAccess=$statusFullAccess; OneDrive=$statusOneDrive; LicenseRemoved=NOT ATTEMPTED; SignInBlocked=NOT ATTEMPTED)"

    return [PSCustomObject]@{
        Smtp     = $Smtp
        ObjectId = $objectId
        Text     = $out.ToString()
        Statuses = @{
            HideShared = $statusHideShared
            FullAccess = $statusFullAccess
            OneDrive   = $statusOneDrive
        }
    }
}
#endregion
