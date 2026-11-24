# Common.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Shared helper library for the Admin Console launcher and every standalone
    tool script (MailboxHealth.ps1, MessageTrace.ps1, InboxRules.ps1, etc.).
.DESCRIPTION
    Dot-source this file from any tool script:  . "$PSScriptRoot\Common.ps1"

    Every function in here is MOVED, UNCHANGED, from the original monolithic
    Mailbox-Group-Inspector.ps1 -- Get-TargetSmtp, Format-ErrorDetail,
    Get-EXOCommand, Invoke-EXOCommand, Resolve-Target, Write-CompletedLog, and
    Get-InboxRulesForMailbox are copy-pasted verbatim from that script (their
    original line numbers are noted below for traceability). Nothing about
    their behavior was redesigned, simplified, or rewritten.

    Two things in this file ARE new, because the monolith never needed them
    (everything used to share one WinForms $form and its control objects):
      - Get-ADServerString: a parameter-based equivalent of the monolith's
        Get-SelectedADServer, which read $radHQ.Checked / $cboDC.SelectedItem
        directly off the shared form. Tool scripts now receive Environment /
        DomainController as plain strings (from CLI params or their own local
        prompt), so this takes those same two inputs as parameters instead of
        reading a control.
      - Show-TargetPromptDialog / Resolve-ToolTarget: the shared "manual
        entry" prompt every tool script falls back to when launched
        standalone (no CLI parameters), so each of the ~10 tool scripts
        doesn't reimplement its own copy of the input box / Environment
        radios / DC dropdown. This is orchestration/UI glue for the
        parameter-passing model described in the architecture doc, not any
        tool's business logic.

    REVISED 2026-07-23 (operator request -- auto-detect HQ vs Franchise here
    too, matching AdminConsole.ps1's same-day change): account names are
    unique across the HQ and Franchise forests, so there's no reason any tool
    script's manual-entry fallback should still make the operator pick an
    environment by hand.
      - New function Find-TargetEnvironment does the detection: Get-ADObject
        -Filter "mail -eq '<address>'" against HQ first ($script:HQDomain,
        forest-wide, smaller forest), then Franchise only if HQ comes back
        empty (using the new $script:FranchiseAutoDetectDC below). Whichever
        forest answers determines Environment + ADServer. Returns Environment
        = $null if neither forest has a match.
      - Show-TargetPromptDialog dropped its Environment radio group and DC
        dropdown entirely -- it now collects ONLY the mailbox/group address;
        Resolve-ToolTarget calls Find-TargetEnvironment once that address is
        known, instead of trusting whatever the operator picked on the
        now-removed radios.
      - Resolve-ToolTarget's existing fast path (all three of SmtpAddress /
        Environment / DomainController already valid -- the normal case when
        a tool is launched from AdminConsole.ps1, which already ran its own
        auto-detect) is UNCHANGED: it still trusts those values as-is rather
        than re-querying AD a second time for the same address. Auto-detect
        only runs when SmtpAddress is present but Environment/DomainController
        aren't (or a tool is launched fully standalone with nothing at all,
        in which case the trimmed-down Show-TargetPromptDialog collects just
        the address first).
      - $script:FranchiseAutoDetectDC is a NEW, deliberately separately-named
        constant -- NOT reusing the name $script:DefaultDC that AdminConsole.
        ps1 (and SendToDL.ps1/SendAsDL.ps1) each keep as their OWN local
        copy. This isn't the same situation as the 2026-07-21 incident where
        a shared $script:DefaultDC in Common.ps1 silently changed
        disable-intake.ps1's behavior: Find-TargetEnvironment is a brand-new
        function nothing called before this revision, so there's nothing for
        it to unexpectedly override -- every script in this codebase that
        dot-sources Common.ps1 gains a new, inert function/variable pair it
        simply doesn't call unless it already invokes Resolve-ToolTarget.
      - Scope: this only changes the manual-entry FALLBACK path shared by
        tool scripts that already call Resolve-ToolTarget (AllAttributes,
        CalendarPermissions, DLMemberships, GroupEditor, InboxRules,
        MailAttributes, MailboxAudit, MailboxHealth, MailboxPermissions,
        OOFEditor, SharedMailboxAudit -- 11 scripts, unchanged themselves,
        since Resolve-ToolTarget's own signature and contract didn't change).
        SendToDL.ps1, SendAsDL.ps1, DLManager.ps1, and UnifiedGroupMembership.
        ps1 each have their OWN separate, local Environment radio UI (they
        look up Distribution Lists by Name, not by mailbox/group SMTP
        address, so Find-TargetEnvironment's mail-attribute AD filter doesn't
        directly apply) and were NOT touched by this revision -- see the
        2026-07-23 conversation for the follow-up plan on those four.
        MessageTrace.ps1 no longer calls Resolve-ToolTarget at all (see its
        own header) and is unaffected either way. SRG/ANZ Common.ps1 copies
        (SRGCommon.ps1, ANZCommon.ps1) are entirely separate files and were
        not touched -- those suites have no HQ/Franchise concept at all.

    Why Get-InboxRulesForMailbox lives here instead of only in InboxRules.ps1:
    Get-MailboxHealthSnapshot (in MailboxHealth.ps1) calls it too ("Reuses
    Get-InboxRulesForMailbox ... rather than querying it twice" -- see that
    function's own comment). Since every tool script must be fully
    functional on its own, this one small, genuinely shared function is
    promoted to Common.ps1 rather than duplicated or left to create a
    file-to-file dependency between two otherwise-independent tool scripts.
.NOTES
    Author   : Glenn Armbruster
    Created  : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 as part
               of the launcher/tool-script architecture refactor -- see
               REFACTOR-DESIGN.md for the full migration plan).
    Revised  : 2026-07-23 (auto-detect HQ vs Franchise in
               Resolve-ToolTarget/Show-TargetPromptDialog -- see REVISED
               note above)
```


## region Constants (moved from the WinForms UI region, lines ~3672-3682)

- **L114-119** Franchise DC used ONLY by Find-TargetEnvironment's auto-detect check below. Deliberately a separately-named constant, NOT $script:DefaultDC (the name AdminConsole.ps1/SendToDL.ps1/SendAsDL.ps1 each keep as their OWN local copy) -- see the REVISED 2026-07-23 header note above for why this is safe to centralize here even though a shared $script:DefaultDC caused a real problem in disable-intake.ps1 back on 2026-07-21.

## function Format-ErrorDetail

- **L136-138** Central formatter so every catch block reports the same level of detail: cmdlet, script line, exception type, message, category, and inner exception if present. Never just "$_.Exception.Message" alone.

## function Get-EXOCommand

- **L165-172** Resolves a cmdlet name to the Exchange Online-sourced command specifically. Field-tested finding (2026-07-08): even with no explicit on-prem EMS session connected, an on-prem Exchange module/snap-in can still be loaded in the host (leftover profile/autoload) and win a plain-name lookup -- the classic symptom is a "couldn't be found on '<DC>'" error from what should be a cloud-only cmdlet, or a script-line number in the thousands (an auto-generated remoting proxy, not this script). Every EXO-flavored call in ANY tool script MUST go through Invoke-EXOCommand below instead of calling the cmdlet name directly.

## function Invoke-EXOCommand

- **L189** Splat-friendly wrapper: Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }

## function Resolve-Target

- **L201-208** Used by AdminConsole.ps1 to detect Mailbox / UnifiedGroup / OnpremGroup so it knows which tool buttons to render. Individual tool scripts do NOT need to call this themselves -- each one already assumes its own target type (e.g. MailboxHealth.ps1 assumes it was handed a mailbox); if handed the wrong type of object, the underlying Get-Mailbox / Get-ADGroup call in that tool will simply fail with a normal Format-ErrorDetail message, same as the monolith's behavior when someone tried, say, "Inbox Rules" against a group.

## region Inbox rule read (moved unchanged -- original lines 798-806)

- **L252-253** Shared by InboxRules.ps1 (its primary reason for existing) and MailboxHealth.ps1 (rule count / suspicious-rule scan) -- see header note.

## region NEW -- parameter-based AD server resolution

- **L284-288** Equivalent to the monolith's Get-SelectedADServer, but takes plain string parameters instead of reading $radHQ.Checked / $cboDC.SelectedItem off a shared form (there is no shared form anymore -- every tool script owns its own controls, or receives these values as CLI parameters with no controls at all). Same two-branch logic, nothing behavioral changed.

## region NEW -- auto-detect HQ vs Franchise

- **L300-311** Added 2026-07-23 -- see the header REVISED note. Auto-detects which AD forest a mailbox/group's on-prem object lives in: account names are unique across HQ and Franchise (operator-confirmed), so "an AD object with this mail address exists here" reliably identifies environment -- no separate identity mapping needed. HQ is checked first (forest-wide $script:HQDomain, smaller forest, no fixed-DC risk); Franchise is only checked if HQ comes back empty, using $script:FranchiseAutoDetectDC above. Returns Environment = $null if neither forest has a match -- callers treat that as "not found in HQ or Franchise", not a silent fallback to one side. Identical logic to AdminConsole.ps1's own copy of this function (kept separate there rather than refactored to call this one, to avoid touching an already-tested file for a pure deduplication with no behavior change).

## region NEW -- auto-detect HQ vs Franchise BY sAMAccountName

- **L338-356** ADDED 2026-07-24 for newuser.ps1 (Enable Remote Mailbox). WHY THIS EXISTS RATHER THAN REUSING Find-TargetEnvironment ABOVE -- this is the important bit, do not "simplify" by merging them: Find-TargetEnvironment locates an object by its MAIL ATTRIBUTE (Get-ADObject -Filter "mail -eq '<smtp>'"). That is correct for every other tool in this suite, because those tools act on mailboxes and groups that are ALREADY mail-enabled. The Enable-RemoteMailbox workflow is the exact opposite case: the accounts it targets are NOT mail-enabled yet -- that is the entire reason the tool is being run. Their mail attribute is empty, so a mail-based lookup finds nothing and would report "not found in HQ or Franchise" for every single valid new account. So this one resolves by sAMAccountName, which exists from the moment the account is created. Same HQ-first-then-Franchise order and same "Environment = $null means genuinely not found" contract as its sibling.

## function Find-UserEnvironmentBySam

<!-- line 358 -->

```text
.SYNOPSIS
        Finds which forest an account lives in, by sAMAccountName.
    .OUTPUTS
        PSCustomObject:
          Environment [string] 'HQ' | 'Franchise' | $null when not found
          ADServer    [string] the server that answered ($null when not found)
          DCParam     [string] '' for HQ (forest-wide), the DC FQDN for
                               Franchise -- ready to pass to
                               Set-ADServerSettings / -DomainController
          User        [object] the matched AD user, for display
          Errors      [string[]]
```

- **L376** HQ first: smaller forest, and forest-wide so no DC is pinned.
- **L387-388** Genuinely absent from HQ -- expected, fall through to Franchise rather than recording it as an error.
- **L404** Absent from Franchise too -- the caller reports "does not exist yet".

## region NEW -- consistent window placement

- **L417-427** ADDED 2026-07-24 (operator request): the console should always open at the top-left of the screen, and every tool it launches at the bottom-left, so the two never cover each other and both land somewhere predictable. Uses WorkingArea rather than Bounds so the taskbar is respected -- a form placed by Bounds can end up underneath it. Clamped to the working area: if a form is taller than the screen, the bottom-left calculation would otherwise push its title bar off the top, leaving a window that can't be dragged. Better to sit flush at the top than to be unreachable.

## function Set-FormScreenCorner

<!-- line 429 -->

```text
.SYNOPSIS
        Places a WinForms form in a screen corner. Call BEFORE ShowDialog().
    .PARAMETER Corner
        TopLeft | BottomLeft
```

- **L448** taller than the screen
- **L452** Placement is cosmetic -- never let it stop a tool from opening.

## region NEW -- connection state probes (for AdminConsole's Connect buttons)

- **L459-473** ADDED 2026-07-24 -- so AdminConsole.ps1 can colour its Connect Scripts buttons green when a service is connected and red when it isn't. SAFE ADDITION, same reasoning as Find-TargetEnvironment above: these are brand-new FUNCTIONS that nothing previously called, so no existing script's behaviour can change because of them. This is deliberately NOT a new shared CONSTANT -- adding a shared $script:DefaultDC here on 2026-07-21 is what broke disable-intake.ps1's own DC handling, and that lesson stands. DELIBERATE DUPLICATION with M365Logic.psm1's Test-M365ExoConnection / Test-M365SpoConnection: the four suites in this codebase (this one, SRGCommon, ANZCommon, M365Logic) are intentionally independent, sharing no runtime state, so a change in one can never surprise another. Copying ~20 lines is the price of that isolation and is the established pattern here -- do NOT "fix" this by having Common.ps1 import M365Logic.

## function Test-ExoConnectionState

<!-- line 475 -->

```text
.SYNOPSIS
        Is Exchange Online usable in THIS process? Returns an object, never
        throws.
    .DESCRIPTION
        Cheap: reads in-process state only, no service call, so it's safe to
        call on a form's startup path.

        Two stages. First, are EXO cmdlets even present (the same Get-Command
        probe Get-EXOCommand uses)? If not, nothing is connected. Second, if
        the module offers Get-ConnectionInformation (EXO v3+), read live
        session state -- that's what distinguishes "module loaded but the
        session expired" from "genuinely connected", which matters for a
        console left open all day.

        Where state can't be determined it reports CONNECTED rather than
        disconnected, on purpose: a false green is corrected the moment a
        cmdlet actually fails with its own clear error, whereas a false red
        would send the operator chasing a connection problem that isn't there.
```


## function Get-ConnectedTenantDomain

<!-- line 538 -->

```text
.SYNOPSIS
        The CONNECTED tenant's default accepted domain, or $null. Never throws.
    .DESCRIPTION
        ADDED 2026-07-24 -- so the console can show WHICH tenant it's pointed
        at, the way M365_Host_Web.ps1's top bar does. "Connected" on its own
        doesn't tell you connected to what, which matters when several tenants
        get worked in a day.

        NOTE this is the LIVE tenant, deliberately distinct from
        $script:MailDomain above. That constant is what this suite APPENDS to a
        bare name typed in the box ('expresspros.com'); this function reports
        what Exchange Online is actually connected to. They should normally
        agree -- and when they don't, that's exactly the mistake worth seeing,
        because the box would be resolving names into a tenant you aren't
        connected to.

        Cached after the first success, so repeated calls (startup, Refresh,
        each Look Up) cost nothing. Returns $null rather than throwing when
        nothing is connected, since callers use it for a label.
```

- **L565-580** REVISED 2026-07-24 -- prefer the VANITY domain, not the default one. The first version returned whichever domain was flagged Default, and in this tenant that is expresspros.onmicrosoft.com. That is completely normal (the onmicrosoft domain is very often the default even when every user's primary SMTP is the vanity domain) -- but it is not what an admin means by "which tenant am I in", and it made the label read as a mismatch against $script:MailDomain when nothing was wrong. Preference order: 1. an accepted domain equal to $script:MailDomain -- the exact domain this console appends to bare names, so the most meaningful thing to show 2. any non-*.onmicrosoft.com domain -- the vanity domain 3. the Default domain 4. whatever came first
- **L598-599** Not connected, or the lookup failed -- not fatal, the caller shows "unknown". Deliberately swallowed.

## function Clear-ConnectedTenantDomainCache

- **L605-606** Needed if the operator disconnects and reconnects to a different tenant inside one console session.

## function Test-SpoConnectionState

<!-- line 611 -->

```text
.SYNOPSIS
        Is SharePoint Online connected? Returns an object, never throws.
    .DESCRIPTION
        COSTS A REAL SERVICE CALL, and callers on a WinForms thread should know
        that: the SPO module exposes no local connection-state cmdlet (there is
        no SPO equivalent of Get-ConnectionInformation), so the only reliable
        probe is a cheap real call. Get-SPOTenant is used because it is
        tenant-scoped, needs no arguments, and returns quickly.

        Because it blocks, AdminConsole.ps1 calls this on startup and after a
        Connect click -- NOT on a timer, which would freeze the form at random
        moments during real work.
```


## region NEW -- shared "enter a target" prompt for standalone tool launches

- **L643-653** Mirrors the monolith's own top-of-form controls (original lines ~3688-3726) as a small modal dialog, so every tool script gets the exact same "manual entry" experience without each of the ~14 tool scripts re-implementing its own copy of that UI. This function contains NO business logic -- it only collects an address and returns it. REVISED 2026-07-23 -- the Environment radios and DC dropdown this dialog used to show are gone; Resolve-ToolTarget below now auto-detects that via Find-TargetEnvironment once an address is known, so there's nothing left for the operator to pick here except the address itself. See the header REVISED 2026-07-23 note for the full rationale.

## function Show-TargetPromptDialog

- **L724-737** Every tool script's param block gives it -SmtpAddress / -Environment / -DomainController. When AdminConsole launches the tool, all three arrive pre-populated (AdminConsole already ran its own Find-TargetEnvironment) and this returns immediately with no UI/no re-query, same fast path as before. REVISED 2026-07-23 -- when that fast path doesn't apply (a tool launched standalone, or with only a partial/invalid set of the three), this no longer pops the old Environment-radio/DC-dropdown prompt. Instead: if there's no usable address yet, Show-TargetPromptDialog collects ONLY the address (see its own REVISED note); either way, once an address is known, Find-TargetEnvironment auto-detects HQ vs Franchise the same way AdminConsole.ps1 does. If neither forest has a matching object, the operator sees "not found in HQ or Franchise" and this returns $null, same as a cancelled prompt.

## function Resolve-ToolTarget

- **L756** operator cancelled

## Console button geometry (2026-08-09)

`Get-ConsoleButtonWidth`, `Get-ConsolePanelWidth`, and the `$script:ConsoleBtn*`
constants live here rather than in AdminConsole.ps1 because ToolManager's
Sections dialog previews the size the console is going to draw. A preview
computed from a different pad, height or scrollbar allowance than the console
actually uses is a preview that LIES -- which is worse than having no preview,
because it will be believed. See ARCHITECTURE 14n.

### MaxLines is derived, not assumed

A line of Microsoft Sans Serif 9 is 15px at 100% scaling and 16px at some DPI
settings. The obvious implementation -- "the text area is 30px, so two lines
fit" -- would silently mean two lines on one machine and ONE line on another.
The one-line machine would not error. It would quietly return a much larger
width and produce a much wider console, and the only symptom would be a user
saying the window looks different on their laptop.

So the line height is measured at runtime with `'Hg'` (an ascender and a
descender, so it measures the full line box rather than whichever letters happen
to be in a tool name), and the line count is derived from it.

### Why the width check exists as well as the height check

`TextFormatFlags::WordBreak` breaks between WORDS. A single word longer than the
proposed box does not wrap -- it overflows, and `MeasureText` returns a Width
GREATER than the width it was given. Height alone would call that a pass and
ship a clipped caption. The width comparison is what catches an unbreakable
name.

### The inset is sized for the wider of the two button styles

A themed button is `FlatStyle::Flat` (1px border); an unthemed one is Standard
(3px plus a focus rectangle). The inset is sized for Standard. A few pixels
generous costs a few pixels of window; a few pixels short clips a caption, and a
clipped caption in a demo is not recoverable by explaining it.

### The catch returns 150

Measuring is an optimisation, not a requirement. If GDI is unavailable for any
reason the console must still open, so the fallback is the width it used before
this function existed.


## Audit log schema v3 -- Tenant (2026-08-09)

Environment was carrying a job it cannot do. It holds HQ / Franchise / EXO -- a
DIRECTORY inside one tenant, not the tenant. This console ships to six of them,
every copy writes the same CSV format, and the moment two of those files are put
side by side to answer "who changed what", nothing in a row says which tenant it
happened in. Operator and Machine do not settle it either: the same engineer on
the same jump box administers all six.

A separate column rather than something folded into Detail, for the reason the
whole v2 rewrite existed: a filter beats a grep.

### Read from config, not from the session

`Get-LogTenantLabel` reads Tenant.json rather than asking the connected session,
because it has to work on the FAILURE rows. A row logged because Connect-ExO
threw has no session to interrogate, and those are exactly the rows worth having.

### The fallback filename is keyed to the schema version

`Get-CompletedLogPath` hardcoded `CompletedTasks-v2.csv`, which worked exactly
once. Adding a column made that name stale, and any deployment already pushed
onto the v2 fallback would have had v3 rows appended under a v2 header -- every
field after Machine shifted one column left, in a file whose name still claimed
to be v2. No error, no warning, and discoverable only by someone reading the log
months later to answer a question that mattered.

The name now embeds `$script:LogSchemaVersion`. If both candidate names somehow
exist with wrong headers -- only reachable by hand-editing, since the name
encodes the version -- it date-stamps a third rather than returning `$null`.
Losing a row is the one outcome an audit log cannot have; a confusing filename
is recoverable, a missing entry is not.

### Consumers were already correct

`Convert-LogHistory.ps1` and `Dashboard.ps1` both compare against
`$script:LogHeaderLine` and select `$script:LogColumns` rather than listing
columns themselves, so they followed the schema change with no edit. That is
what that indirection was for.
