# UnifiedGroupMembership.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Bulk add/remove users to/from Microsoft 365 (Unified) Group membership --
    many-to-many, no target lookup.
.DESCRIPTION
    Two independent paste lists -- Users and Unified Groups, one entry per
    line (commas/semicolons also tolerated as separators). Every user listed
    is added to (or removed from) every group listed: a full Users x Groups
    cross-product, per the standards' batch/cross-product pattern (Section
    21). Confirmed with the operator before building that this tool would
    NOT use the single-object lookup flow every other tool in this suite
    uses (Section 1 triage) -- there is no single "target" here, there are
    two independent lists, each de-duplicated and normalized on its own.

    The two lists use DIFFERENT lookup keys, deliberately (Section 5 is
    written around a single-object tool; this is a two-list batch tool, so
    each list gets the key that matches what the operator actually has in
    hand for that list):
      - Users: PrimarySmtpAddress (Section 5 exception -- users are
        mail-enabled objects). Bare local-parts (no '@') get
        '@expresspros.com' silently appended via Common.ps1's existing
        Get-TargetSmtp; a full address is passed through unchanged.
      - Unified Groups: Name (the group's display name), passed straight
        through to -Identity with no transformation and no domain
        appended. Confirmed against the Microsoft Learn reference (Section
        2) that Add-/Remove-/Get-UnifiedGroupLinks' -Identity parameter
        resolves on Name, Alias, DN, Canonical DN, email address, or GUID --
        Name is a fully supported, documented value, not a guess. An email
        address still works too if that's easier to paste, since it's one
        of the same accepted values -- it just isn't required.
    Nothing in either list is required to already exist in the operator's
    EXO session before it's typed; unresolvable entries surface as errors
    at Add/Remove time, not as a validation step beforehand.

    Only the Members link type is touched. Owners and Subscribers are never
    read or modified by this tool -- use Group Editor (GroupEditor.ps1) for
    single-group owner delegation.

    Deliberate deviation from the standard Parse/Run/Close button triad
    (Section 13): the operator explicitly asked for exactly three buttons --
    Add, Remove, Cancel -- so there is no separate Parse button. In its
    place, the parsed-preview box (the Section 21 "show the scale before
    Run" requirement) refreshes live on every keystroke in either list box,
    so the true Users x Groups operation count is always visible before Add
    or Remove is ever clicked. Because there's no separate Parse step, there
    is also no persisted $script:Parsed state to go stale (Section 7) --
    Add/Remove each re-read and re-normalize both textboxes directly at
    click time.

    Verify mode (Section 19) is deliberately omitted, not an oversight:
    unlike on-prem AD writes elsewhere in this suite, a Get-UnifiedGroupLinks
    read run immediately after an Add-/Remove-UnifiedGroupLinks write against
    Exchange Online is not subject to the DC-replication-lag concern Verify
    mode exists to work around. The inline AFTER snapshot captured per group
    immediately after its writes (Section 17) is already authoritative.

    Already-a-member (on Add) / not-a-member (on Remove) operations are
    reported as SKIP, not errors (Section 21) -- determined from each
    group's own live BEFORE membership snapshot, never inferred from a
    cmdlet failure. BEFORE/AFTER is captured once per target GROUP, not once
    per user (Section 17/21); one CompletedTasks.csv row is written per
    group that had at least one real (non-skipped, non-failed) change,
    joining the successful members with '+' (Section 21 batch log grammar),
    never one row per user.

    Environment required: Exchange Online only. No on-prem AD/Exchange calls
    are made, so there is no DC selector and no Franchise/HQ split --
    Microsoft 365 Groups have no on-prem equivalent (Section 22 does not
    apply to this tool).

    -SmtpAddress/-Environment/-DomainController are declared, exactly as in
    PowerShellConsole.ps1, only so a future Start-ToolScript launch would
    not throw a parameter-binding error -- this tool ignores all three. It
    is not currently wired into Tools.json: ToolManager.ps1 validates every
    entry's ObjectTypes against a single resolved object (Mailbox /
    UnifiedGroup / OnpremGroup), which this many-to-many, lookup-free tool
    does not fit, the same reasoning DLManager.ps1's header gives for why it
    is launched standalone instead of from AdminConsole's per-object button
    row.

    Cmdlets used -- Add-UnifiedGroupLinks, Remove-UnifiedGroupLinks,
    Get-UnifiedGroupLinks -- are EXO-only and were verified against
    Microsoft Learn before delivery (Section 2):
      https://learn.microsoft.com/powershell/module/exchangepowershell/add-unifiedgrouplinks
      https://learn.microsoft.com/powershell/module/exchangepowershell/remove-unifiedgrouplinks
    Add-/Remove-UnifiedGroupLinks are also already in confirmed working use
    elsewhere in this suite, with the identical Identity/LinkType='Members'/
    Links/ErrorAction='Stop' shape, via GroupEditor.ps1's
    Add-UnifiedMember/Remove-UnifiedMember.

    KNOWN LIMITATION -- no PowerShell host was available to run the Section
    15 AST parse check before delivery. The script has been reviewed by eye
    for balanced braces/parens/quotes and structural consistency with the
    other tool scripts in this suite, but that is not a substitute for
    Parser::ParseFile -- run the Section 15 snippet against this file before
    first real use and treat any reported errors as blocking.
.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-07-21
    Revised         : 2026-07-21 1357
    Run from        : standard PowerShell with Exchange Online connected
                       (Connect-ExchangeOnline)
    Required modules: ExchangeOnlineManagement
```


## region Parsing

- **L118-130** Normalizes one paste box into a de-duplicated array. De-dup is case-insensitive but preserves first-seen casing and order. -AsEmail (Users box): tolerates newline, comma, and semicolon as separators (Section 21) since pasted address lists often arrive comma-delimited rather than one-per-line; runs every entry through Common.ps1's Get-TargetSmtp so a bare local-part gets '@expresspros.com' appended and a full address passes through as-is. Groups box (no -AsEmail): entries are Unified Group DISPLAY NAMES, not addresses, passed straight through with no domain-append and no other transformation. Only newline and semicolon are treated as separators here -- NOT comma -- because a group display name can legitimately contain a comma (e.g. "IT, Facilities, and Ops"); splitting on comma would silently corrupt a name like that into two bogus entries.

## function Remove-UnifiedGroupMemberLink

- **L181-194** Runs the full Users x Groups cross-product for one direction (Add or Remove). One BEFORE/AFTER snapshot per target GROUP (Section 17/21), not per operation -- each user add/remove within that group is logged as a single line between the two snapshots. Already-satisfied operations (member already present on Add; member already absent on Remove) are SKIP, not errors, determined from the live BEFORE snapshot -- never from cmdlet failure (Section 21). A group that fails to resolve at all (bad address, not actually a Unified Group, etc.) is reported and skipped in its entirety without aborting the rest of the batch -- same "one bad row doesn't hide the other nine" convention used throughout this suite (e.g. DLManager.ps1's Add-DLMemberBulk). One CompletedTasks.csv row is written per group that had at least one real change (Section 21 batch grammar), never one row per user, and never for a group where every operation was a SKIP or a failure.

## function Update-ParsedPreview

- **L440** ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
- **L442** Console, which opens top-left. Helper is in Common.ps1.
