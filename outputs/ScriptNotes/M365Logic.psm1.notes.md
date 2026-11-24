# M365Logic.psm1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
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
```


## region Module-private state

- **L114-115** Module-scoped, therefore genuinely private -- no consumer can see or clobber this (see header note on why that matters here specifically).

## function Get-EXOCommand

<!-- line 121 -->

```text
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
```


## function Invoke-EXOCommand

<!-- line 158 -->

```text
.SYNOPSIS
        Splat-friendly wrapper around Get-EXOCommand.
    .EXAMPLE
        Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
```


## function Test-M365ExoConnection

<!-- line 173 -->

```text
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
```

- **L233-234** Older module without Get-ConnectionInformation -- cmdlet presence is the best signal available. See UNVERIFIED note above.
- **L246-248** Prefer a connection that positively reports itself healthy; if no connection object exposes State/TokenStatus at all, fall back to "a connection object exists" rather than guessing it's dead.
- **L264-265** Get-ConnectionInformation itself failed -- don't claim disconnected on the strength of that, see UNVERIFIED note.

## region SPO / Entra connection probes

- **L274-302** ADDED 2026-07-24 -- the operator's actual daily workflow is three connects, every day, in Windows PowerShell 5.x: Connect-ExchangeOnline Connect-SPOService Connect-AzureAD so the host shows one pill per service rather than a single "M365 connected" light. Test-M365ExoConnection above covers the first; these two cover the others. A CORRECTION WORTH RECORDING, since earlier notes in this codebase (SRGDisable/ANZDisable headers) overstate it: the AzureAD module is DEPRECATED (Microsoft announced deprecation 2024-03-30), and specific cmdlets that call the retired Azure AD Graph endpoint do fail -- which is exactly what broke Get-AzureADGroup during a real Disable run. But Connect-AzureAD itself still authenticates, which the operator confirmed from daily use. "Some cmdlets are dead" is not "the module is dead," and the earlier note conflated them. POWERSHELL VERSION MATTERS HERE, a lot: the AzureAD module is Windows PowerShell 5.1 ONLY (not supported on PowerShell 7 -- confirmed against Microsoft's own module-compatibility docs), and Microsoft.Online.SharePoint.PowerShell does not work natively in PS7 either (it needs Import-Module -UseWindowsPowerShell compatibility mode). Only ExchangeOnlineManagement is genuinely PS7-native. That's why the web host now targets 5.1+ instead of 7.0: run it in Windows PowerShell 5.1 and all three of these probes can actually report Connected. Run it under PS7 (e.g. on macOS) and ExO works while these two will read NotConnected because their modules can't load there -- which is correct reporting, not a bug.

## function Test-M365SpoConnection

<!-- line 305 -->

```text
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
```


## function Test-M365AadConnection

<!-- line 340 -->

```text
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
```

- **L369** 1. AzureAD (the operator's current daily connect)
- **L386** 2. Microsoft Graph -- local check, no API call
- **L405** 3. Microsoft.Entra

## function Get-M365ConnectionSummary

<!-- line 429 -->

```text
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
```

- **L468-470** Not monitored at all -- distinct from "checked and disconnected", and distinct from "skipped this pass". The caller should hide the pill entirely rather than show a state.

## function Format-ErrorDetail

<!-- line 488 -->

```text
.SYNOPSIS
        Uniform, detailed error text for every catch block in the toolkit.
    .DESCRIPTION
        Same level of detail everywhere: cmdlet, script line, exception
        type, message, category, inner exception if present. Returns a
        string -- the caller decides whether it lands in a TextBox, an HTML
        block, a log file, or the console.
```


## function Get-M365DefaultDomain

<!-- line 523 -->

```text
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
```

- **L548-549** Not connected yet, or the lookup failed -- not fatal, and deliberately not surfaced as an error. See DESCRIPTION.

## function Clear-M365DefaultDomainCache

<!-- line 555 -->

```text
.SYNOPSIS
        Forgets the cached default domain so the next lookup re-queries.
    .DESCRIPTION
        Needed by any caller that can switch tenants within one process
        life -- most obviously a long-running browser host where the
        operator disconnects and reconnects to a different tenant. Without
        this, the cache from tenant A would silently keep appending A's
        domain to bare names typed against tenant B. Short-lived WinForms
        tools never need it.
```


## function Get-TargetSmtp

<!-- line 570 -->

```text
.SYNOPSIS
        Normalizes a typed name into a full SMTP address.
    .DESCRIPTION
        "bob.smith" and "bob.smith@<tenant default domain>" resolve the
        same. A domain is appended only if the operator didn't type one AND
        a default domain could actually be resolved from the connected
        tenant; otherwise returns $null so the caller can ask for a full
        address rather than guessing.
```


## function ConvertTo-NormalizedUserList

<!-- line 591 -->

```text
.SYNOPSIS
        Splits a free-text multi-user box into a deduplicated list.
    .DESCRIPTION
        Tolerates newline, comma, and semicolon separators; trims; drops
        blanks; dedupes case-insensitively while preserving the order the
        operator typed. No domain-append happens here -- pipe each result
        through Get-TargetSmtp if addresses are wanted.
```


## function Resolve-Target

<!-- line 614 -->

```text
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
```


## region Admin portal deep links

- **L659-714** ADDED 2026-07-24 -- one place for every admin-portal URL the toolkit links to. Centralized deliberately: portal URLs are Microsoft's to change, and when one breaks, fixing it should be a one-line edit here rather than a hunt through view files. VERIFICATION STATUS as of 2026-07-24 (checked against Microsoft Learn / current docs before shipping -- re-check if a link ever 404s): VERIFIED Defender Threat Explorer https://security.microsoft.com/threatexplorerv3 (documented as the direct link; note the "v3" suffix -- the bare /threatexplorer path is NOT the documented one) VERIFIED Defender Quarantine, Email tab https://security.microsoft.com/quarantine?viewid=Email PARTIAL Entra user blade pattern #view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/<section>/userId/<guid> The blade name, the ~/<section>/ segment, and the userId segment are confirmed in Microsoft's own docs and answers, but not every section slug is. So only 'overview' is used below -- it's the safest, and it's also the page where BOTH follow-up actions are reachable (the Block sign-in button sits on its toolbar, and Groups is in its left nav). Deliberately not guessing at a deeper slug like ~/Groups/ that might 404. VERIFIED Entra all-users list https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/AllUsers -- used as the fallback whenever an object GUID isn't available. The Entra links need the user's Entra object GUID, which this toolkit gets for free: Get-Mailbox's ExternalDirectoryObjectId IS that GUID (confirmed -- "ExternalDirectoryObjectId is the Microsoft Entra ID object GUID for the mailbox"). That matters a lot here, because it means precise per-user deep links work with NOTHING but the Exchange connection -- no Microsoft Graph, no Microsoft.Entra module, and no tenant-admin consent, which is exactly what wasn't available when license removal and sign-in blocking had to be dropped from the automated path in the first place. URL TEMPLATES -- operator-overridable, see Set-M365PortalUrlTemplate below. WHY TEMPLATES INSTEAD OF FIXED STRINGS: the operator asked whether the Defender pages could open with a filter already applied -- Quarantine pre-filtered to a sender, Threat Explorer pre-filled with the address being investigated. Both portals DO have those filters in their UI (Quarantine filters on Message ID / sender / recipient; Explorer on sender / subject / recipient / NetworkMessageID), but Microsoft publishes NO query-string syntax for pre-filling them, and inventing parameter names would produce links that quietly land on an unfiltered page -- worse than no link, because it looks like it worked. So: these are templates. Any template may contain {sender}, {recipient}, or {messageId}. Get-M365PortalUrl substitutes (URL-encoded) and reports whether substitution actually happened, so the UI can label the link honestly either way. Templates ship WITHOUT placeholders -- i.e. the plain, documented, known-good base URLs. TO TURN ON PREPOPULATION once you've discovered the real syntax (filter in the portal, then read your address bar): Set-M365PortalUrlTemplate -Kind Quarantine ` -Template 'https://security.microsoft.com/quarantine?viewid=Email&<theparam>={sender}' Put that call in the host script after the module import and every link in the toolkit starts prepopulating. No code edit, nothing to hunt for.

## function Set-M365PortalUrlTemplate

<!-- line 724 -->

```text
.SYNOPSIS
        Overrides one portal URL template at runtime.
    .DESCRIPTION
        Lets the operator enable filter prepopulation (or fix a URL Microsoft
        has moved) without editing this module. See the region comment above
        for the discovery recipe and an example.

        Placeholders understood by Get-M365PortalUrl: {sender},
        {recipient}, {messageId}, {objectId}. Values are URL-encoded on
        substitution, so don't pre-encode them in the template.
```


## function Get-M365PortalUrl

<!-- line 746 -->

```text
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
```

- **L783-784** No GUID for a user-specific link -- fall back to the list rather than emitting a deep link with an empty id in it.
- **L803-805** {objectId} is plumbing, not a user-visible filter -- don't let it count as "prepopulated" or every Entra link would claim to be filtered.

## function Get-M365EntraObjectId

<!-- line 819 -->

```text
.SYNOPSIS
        Returns a mailbox's Entra object GUID, or $null.
    .DESCRIPTION
        Get-Mailbox's ExternalDirectoryObjectId is the Entra object GUID
        (see the region comment above). Returns $null rather than throwing if
        the mailbox can't be read or the property is empty -- callers use it
        to build an optional deep link, so a failure here should degrade the
        link to the users list, not break the operation that needed it.
```

- **L834** Deliberately swallowed -- see DESCRIPTION.

## region Suite identity + tenant guard

- **L841-855** ADDED 2026-07-24, when ANZW/SRGW web hosts were added alongside the M365 one. WHY THIS EXISTS: this module is tenant-agnostic by design -- Get-TargetSmtp asks the CONNECTED tenant for its default domain rather than reading a hardcoded constant. That's what lets one engine serve every tenant. But it also means three hosts running the same engine are indistinguishable to the engine itself: start the SRG host, connect to ANZ by mistake, and every tool will happily operate on ANZ while the window says SRG. So a host may declare which tenant it EXPECTS. Test-M365ExpectedTenant compares that against the connected tenant's real default domain, and the host surfaces a mismatch loudly. This is a guard, not a lock -- it warns and labels rather than blocking, because a legitimate mismatch is possible (a tenant with several accepted domains, for instance) and silently refusing to work would be worse than saying plainly what it sees.

## function Set-M365SuiteIdentity

<!-- line 861 -->

```text
.SYNOPSIS
        Names the running suite and, optionally, the tenant it expects.
    .DESCRIPTION
        Called once at host startup. -LogFileName keeps each suite's completed
        task history in its own CSV, so ANZ work never interleaves with SRG's
        (the same separation the WinForms suites already maintain).
        -ExpectedDomain arms the tenant guard; omit it to leave the host
        genuinely tenant-agnostic.
```

- **L879-880** A previous host in this process may have cached a different tenant's domain; drop it so the guard and Get-TargetSmtp both re-resolve.

## function Test-M365ExpectedTenant

<!-- line 887 -->

```text
.SYNOPSIS
        Is the connected tenant the one this suite expects? Never throws.
    .OUTPUTS
        PSCustomObject:
          Armed      [bool]   -- did the host declare an expected domain?
          Match      [bool]   -- true when it matches, or when not armed
          Expected   [string]
          Actual     [string] -- the connected tenant's default domain, or $null
          Detail     [string]
```

- **L910-911** Can't tell yet (probably not connected). Not a mismatch -- reporting one here would cry wolf before the operator has even signed in.

## function Write-CompletedLog

<!-- line 927 -->

```text
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
```

- **L943-946** REVISED 2026-07-24 -- filename comes from Set-M365SuiteIdentity so each suite (M365 / ANZW / SRGW) keeps its own history rather than three tenants' work interleaving in one CSV. Defaults to M365-CompletedTasks.csv when no host has set it.

## function Get-InboxRulesForMailbox

<!-- line 963 -->

```text
.SYNOPSIS
        Returns a mailbox's inbox rules, or the error that prevented it.
    .DESCRIPTION
        Returns @{ Rules = @(...); Error = $null|<string> } rather than
        throwing, so a composite caller (Get-MailboxHealthSnapshot) can
        record a per-field failure and still return everything else.
```


## region Mailbox audit (Search-UnifiedAuditLog)

- **L982-988** MOVED 2026-07-24 from MailboxAudit.ps1's "#region Constants" and "#region Data layer". Search logic and JSON parsing are unchanged; the ONLY behavioural change is how cancellation works -- see Invoke-M365MailboxAudit. The operation list was verified against Microsoft Learn's "Audit log activities" (Exchange mailbox activities table) when the WinForms tool was built, and is carried over verbatim in the same order as the source table.
- **L1022** Reverse lookup (Operation value -> FriendlyName) for display.
- **L1027-1029** Documented ReturnLargeSet ceiling (Search-UnifiedAuditLog reference) -- used to detect and warn on a possibly-truncated result set rather than presenting it silently as complete.

## function Get-M365MailboxAuditOperations

<!-- line 1034 -->

```text
.SYNOPSIS
        The FriendlyName -> Operation map, for building a picker.
```


## function Invoke-M365MailboxAudit

<!-- line 1042 -->

```text
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
```

- **L1097** short page -- no more data

## function ConvertTo-M365AuditRecord

<!-- line 1107 -->

```text
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
```

- **L1133-1136** UTC ISO-8601 rather than the original's .ToLocalTime(): a browser renders the timestamp in the viewer's own zone, so the server should hand over an unambiguous instant instead of pre-converting to the SERVER's local time (which would be wrong if they ever differ).

## function Get-M365FullAccessHolders

<!-- line 1152 -->

```text
.SYNOPSIS
        Who currently holds Full Access on a mailbox (excluding SELF/inherited).
    .DESCRIPTION
        MOVED 2026-07-24 from SharedMailboxAudit.ps1's Get-FullAccessHolders.
        Used to cross-reference audit activity against the current delegate
        list, which is the whole point of the shared-mailbox audit: "who did
        this, and are they still supposed to be able to?"

        Returns an object rather than throwing so the audit can still run and
        report the holder-list failure separately.
```


## region Inbox rule actions

- **L1181-1195** MOVED 2026-07-24 from InboxRules.ps1's "#region Actions - inbox rules", logic unchanged (unchanged there in turn from the original monolith's lines 807-982). Get-InboxRulesForMailbox, which reads the rules, is above. CARRYING FORWARD A REAL BUGFIX -- do not "tidy" this by adding -Mailbox back. All three functions accept -Smtp but deliberately do NOT pass it to the cmdlet. $RuleIdentity comes from Get-InboxRule -Mailbox <smtp>, and Exchange returns that Identity ALREADY SCOPED to the mailbox it was queried against. Passing -Mailbox alongside an already-scoped -Identity throws: "You can't use both the Identity and Mailbox parameters to specify a mailbox" -- confirmed against Microsoft Learn, where -Mailbox exists only to disambiguate a bare rule NAME. The operator hit this for real on Enable/Disable (fixed 2026-07-09); Remove had the identical latent bug. -Smtp is retained because callers use it for the completed-tasks log entry.

## function Remove-TargetInboxRule

- **L1209-1210** DESTRUCTIVE and irreversible -- a deleted inbox rule cannot be restored. NO CONFIRMATION here; the caller must confirm.

## function Get-InboxRuleSummaries

<!-- line 1218 -->

```text
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
```


## function Get-MailboxHealthSnapshot

<!-- line 1263 -->

```text
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
```


## function Format-MailboxHealthReport

<!-- line 1407 -->

```text
.SYNOPSIS
        Renders a Get-MailboxHealthSnapshot object as the plain-text report.
    .DESCRIPTION
        Pure formatting, no queries. Kept separate from the snapshot on
        purpose: a scheduled job may want the OBJECT (to filter, compare, or
        export) while an interactive UI wants this TEXT. Both UI layers use
        this identical function, so a mailbox looked up through WinForms and
        the same mailbox through the browser read exactly the same.
```


## function Get-MailboxState

<!-- line 1485 -->

```text
.SYNOPSIS
        Core properties plus Full Access / Send As grantees for one mailbox.
    .DESCRIPTION
        Per-field try/catch, same partial-data philosophy as
        Get-MailboxHealthSnapshot. ArchiveEnabled is derived from
        ArchiveGuid rather than ArchiveStatus (Microsoft KB4486848 --
        ArchiveStatus can read "None" while an archive is active), and
        ArchiveStatus is reported alongside it for reference.
```


## function Format-MailboxState

<!-- line 1568 -->

```text
.SYNOPSIS
        Renders a Get-MailboxState object as plain text under a heading.
    .DESCRIPTION
        -Heading is what makes BEFORE/AFTER pairs readable; callers pass
        'CURRENT STATE', 'BEFORE', 'AFTER'.
```

- **L1602-1607** The four grant/revoke primitives. NO CONFIRMATION -- see module header. AutoMapping is left at the cmdlet default here (unlike MailboxOneDriveAccess.ps1, which explicitly forces -AutoMapping:$true); that difference is inherited from the tool each came from and is deliberately not "harmonized" without the operator deciding which behavior should win.

## function Invoke-MailboxPermissionEdit

<!-- line 1634 -->

```text
.SYNOPSIS
        Runs one grant/revoke with a BEFORE/AFTER snapshot pair and logs it.
    .DESCRIPTION
        The orchestration both UI layers were duplicating: snapshot, act,
        snapshot, log, return the combined text. NO CONFIRMATION -- the
        caller must have already confirmed (see module header).

        -Action is a scriptblock so this stays agnostic about which of the
        four primitives is being run; -ActionLabel is what appears in the
        returned text, -Verb is what goes in the completed-tasks log.
```


## function Invoke-MessageTrace

<!-- line 1667 -->

```text
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
```


## function Invoke-MessageTraceDetail

<!-- line 1707 -->

```text
.SYNOPSIS
        Per-message delivery detail for one trace result.
```


## function Find-QuarantineMatch

<!-- line 1727 -->

```text
.SYNOPSIS
        Finds quarantine entries matching a trace result's Message-ID.
    .DESCRIPTION
        Throws (rather than returning empty) when there's no Message-ID to
        match on at all -- that's a different situation from "searched and
        found nothing," and the caller should say so differently.
```


## function Get-QuarantineMessageDetail

<!-- line 1754 -->

```text
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
```

- **L1804** Read each property only if present -- see DESCRIPTION on why.

## function Format-QuarantineMessageDetail

<!-- line 1833 -->

```text
.SYNOPSIS
        Renders Get-QuarantineMessageDetail as plain text.
```


## function Get-QuarantineMessagePreview

<!-- line 1858 -->

```text
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
```


## function Invoke-QuarantineRelease

<!-- line 1892 -->

```text
.SYNOPSIS
        Releases one quarantined message to ALL original recipients.
    .DESCRIPTION
        NO CONFIRMATION -- see module header. -ReleaseToAll is inherited
        from the original tools; this releases to every original recipient,
        not just the one whose trace row was selected.
```


## function Get-TraceTimeZone

<!-- line 1910 -->

```text
.SYNOPSIS
        Maps a friendly label to a TimeZoneInfo, falling back to UTC.
    .DESCRIPTION
        Windows time-zone IDs, so this is for the WinForms/5.1 side. (The
        browser layer converts client-side with IANA names instead --
        America/Chicago rather than 'Central Standard Time' -- because
        Windows IDs aren't available in a browser.) Never throws; an
        unknown label or a missing zone yields UTC.
```


## function ConvertTo-TraceDisplayTime

<!-- line 1938 -->

```text
.SYNOPSIS
        Converts a UTC trace timestamp into a target zone for display.
    .DESCRIPTION
        Trace timestamps come back UTC but not always Kind-tagged as such,
        so this forces DateTimeKind.Utc before converting -- otherwise the
        framework may treat an untagged value as local and double-shift it.
        Returns the input unchanged if conversion fails.
```


## function Format-TraceDetail

<!-- line 1958 -->

```text
.SYNOPSIS
        Renders Invoke-MessageTraceDetail output as plain text.
```


## region Calendar permissions

- **L1979-1983** MOVED 2026-07-24 from CalendarPermissions.ps1's "#region Actions - calendar permissions", logic unchanged. That file in turn had them unchanged from the original Mailbox-Group-Inspector.ps1 monolith (lines 1272-1455), so this is the same code that's been in production -- only its home has changed. EXO-only; no on-prem dependency.

## function Get-CalendarFolderIdentity

<!-- line 1985 -->

```text
.SYNOPSIS
        Resolves a mailbox's actual calendar folder identity.
    .DESCRIPTION
        Resolves the real calendar folder rather than hardcoding ":\Calendar",
        since the folder can be renamed or (rarely) localized. Falls back to
        the first Calendar-scope folder if none matches the default root path.
```


## function Get-CalendarPermissions

<!-- line 2002 -->

```text
.SYNOPSIS
        Current calendar folder permissions, or the error that prevented it.
    .DESCRIPTION
        Returns an object rather than throwing so a UI can show the failure
        alongside the rest of its state.
```


## function Set-CalendarPermission

<!-- line 2020 -->

```text
.SYNOPSIS
        Grants or updates one user's calendar role. NO CONFIRMATION.
    .DESCRIPTION
        Add-MailboxFolderPermission fails if the user already has an entry, so
        this checks first (best-effort) rather than parsing the resulting error
        text. If the check itself errors for any reason, it falls through to
        Add-; a genuine problem will surface clearly from that call instead.

        Returns 'Added' or 'Updated' so the caller can report which happened.
```


## function Remove-CalendarPermission

<!-- line 2052 -->

```text
.SYNOPSIS
        Removes a user's calendar permissions entirely. NO CONFIRMATION.
    .DESCRIPTION
        Remove-MailboxFolderPermission removes ALL permissions for this user on
        the folder -- there is no selective/partial removal, per Microsoft
        Learn. The caller must confirm before calling this.
```

- **L2066-2067** The documented MailboxFolderPermission roles, for populating a UI dropdown in one place instead of each front end keeping its own list.

## region Out of office

- **L2076-2078** MOVED 2026-07-24 from OOFEditor.ps1's "#region Actions - out of office", logic unchanged (and unchanged there from the original monolith's lines 1108-1269). EXO-only.

## function Get-OutOfOfficeConfig

<!-- line 2080 -->

```text
.SYNOPSIS
        Reads a mailbox's auto-reply (Out of Office) configuration.
```


## function Set-OutOfOfficeConfig

<!-- line 2089 -->

```text
.SYNOPSIS
        Writes a mailbox's auto-reply configuration. NO CONFIRMATION.
    .DESCRIPTION
        StartTime/EndTime are only sent when the state is 'Scheduled' --
        passing them otherwise is meaningless and Exchange may reject it.
        The caller must confirm before calling this.
```

- **L2122** Valid dropdown values, kept here so both front ends agree.

## function Get-OneDrivePersonalSiteUrl

<!-- line 2129 -->

```text
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
```


## function Set-M365OneDriveSiteAdmin

<!-- line 2161 -->

```text
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
```


## region Mailbox + OneDrive access (bulk)

- **L2192-2203** MOVED 2026-07-24 from MailboxOneDriveAccess.ps1 (the merged Set-MailboxAccess + onedriveaccess tool), logic unchanged. SHAPE: many-to-many. Every Mailbox in one list gets the selected access granted to every User/Manager in the other -- a full cross product, which is what both original tools did. The five Access Types are five COMPLETE, mutually exclusive end states, not independent toggles: Full, SendAs, OneDrive, FullSendAs, FullOneDrive. Send-As + OneDrive is deliberately absent per the operator's rule ("there should never be a Send-As, OneDrive need") -- it isn't validated against, it simply isn't one of the five, so it can't be requested.

## function Get-M365AccessFlags

<!-- line 2205 -->

```text
.SYNOPSIS
        Turns an Access Type name into three booleans.
    .DESCRIPTION
        Single source of truth for what each Access Type does. Both the
        pre-flight gate and the action loop read these same flags, so there's
        no second list anywhere that could drift.
```


## function Invoke-M365MailboxOneDriveAccess

<!-- line 2240 -->

```text
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
```

- **L2300** --- Pre-flight, Exchange-touching runs only.

## function Disable-M365User

<!-- line 2447 -->

```text
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
```

- **L2514-2516** Grab the Entra object GUID BEFORE the mailbox is altered, so the follow-up deep link is available even if a later step fails. Non-fatal by design -- a missing GUID just means the UI links to the users list.
- **L2519** Steps 1+2: hide + convert to shared, in one call.
- **L2531** Step 3: manager Full Access (never Send-As).
- **L2547** Step 4: manager OneDrive site-collection admin.
