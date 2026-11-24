# DL_SendToAs.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Lookup / Add / Remove Send-To (allowed senders) and/or Send-As permission
    on an on-prem Distribution List -- merged from the operator's two
    separate tools, SendToDL.ps1 and SendAsDL.ps1, into a single console.
.DESCRIPTION
    REVISED 2026-07-24 -- merge of two previously separate tools:
      - SendToDL.ps1 (a DL's delivery restriction --
        Set-DistributionGroup -AcceptMessagesOnlyFromSendersOrMembers -- who
        is allowed to send TO the DL)
      - SendAsDL.ps1 (Send-As extended right --
        Add-/Remove-ADPermission -ExtendedRights "Send-As" -- who can send
        mail that appears to come FROM the DL)
    Both tools already shared the identical "1 DL x N users" shape (one DL
    box, one users box, Lookup/Add/Remove/Close four-button layout instead
    of the standard Parse/Run/Close triad, both look up the DL by Name with
    no domain-append on the users box, both auto-detect HQ vs Franchise the
    same way via an identical Resolve-DLEnvironment probe), so this merge
    keeps everything about that shape unchanged and adds one thing: which of
    the two independent operations to run.

    UNLIKE the Full Access / Send-As / OneDrive merge (MailboxOneDriveAccess.
    ps1), Send-To and Send-As on a DL have no stated exclusion rule between
    them -- there's no "these two together never make sense" constraint the
    way Send-As + OneDrive access was ruled out there. A DL delegation change
    can legitimately need either, or both, in the same sitting. So Access
    Type here is two independent CHECKBOXES (Send-To, Send-As), not a
    five-way mutually exclusive radio group -- both can be checked at once,
    and Add/Remove runs whichever ones are checked, one after the other, in
    a single confirmation and a single combined result. At least one must be
    checked for Lookup/Add/Remove to do anything.

    Send-To defaults checked; Send-As defaults unchecked -- narrowest single
    default, same reasoning as MailboxOneDriveAccess.ps1 defaulting to Full
    Access alone rather than a combined option.

    What stayed exactly as each source tool built it, moved here unchanged:
      - Get-SendToState / Format-SendToState / Invoke-SendToChange (from
        SendToDL.ps1) -- full-REPLACE semantics on
        AcceptMessagesOnlyFromSendersOrMembers, current-list and typed users
        both resolved to DistinguishedName via Get-Recipient before diffing
        (so differently-typed identities for the same recipient still
        match), one Set-DistributionGroup call per run, SKIP (not error) for
        already-satisfied requests, nothing written at all if nothing
        actually changes.
      - The "On Add: also enable require sender authentication" checkbox --
        AcceptMessagesOnlyFromSendersOrMembers has no effect unless
        RequireSenderAuthenticationEnabled is also true, a well-known
        Exchange gotcha SendToDL.ps1 surfaced explicitly rather than silently
        fixing. It only applies to a Send-To Add, so it's disabled (and
        unchecked, so stale state can't linger) whenever the Send-To
        checkbox itself is unchecked.
      - Get-SamAccountNameFromUserString / Resolve-SamAccountName /
        Get-SendAsState / Format-SendAsState / Add-SendAsGrant /
        Remove-SendAsGrant / Invoke-SendAsChange (from SendAsDL.ps1) --
        naturally incremental Add-/Remove-ADPermission per resolved user
        (not a full-list replace like Send-To), inherited ACEs and
        NT AUTHORITY\SELF excluded from every BEFORE/AFTER snapshot, the
        hyphenated "Send-As" ExtendedRights string (not "Send As" with a
        space -- see SendAsDL.ps1's original KNOWN SYNTAX GOTCHA note, still
        the standards doc's field-tested correction over current public
        Microsoft Learn examples), SKIP (not error) for already-satisfied
        requests.
      - Resolve-DLEnvironment -- byte-for-byte identical in both source
        files already (auto-detect HQ vs Franchise by trying
        Get-DistributionGroup against HQ first, no DC pinned, then Franchise
        pinned to $script:DefaultDC only if HQ doesn't have it), so there was
        nothing to reconcile; kept as the superset shape SendAsDL.ps1's copy
        already returned (Environment, DCParam for the Exchange-side calls,
        AND a separate plain ADServer string for Get-ADUser's -Server
        parameter -- see each source file's own BUG FIX note for why those
        two shapes differ for HQ). Called once per Lookup/Add/Remove click,
        same as both originals; if neither forest has the DL,
        Lookup/Add/Remove report "not found in HQ or Franchise" instead of
        guessing.
      - ConvertTo-NormalizedUserList -- byte-for-byte identical in both
        source files already, kept as one copy.
      - The -SmtpAddress / -Environment / -DomainController params at the
        top: both source tools already declared these ONLY for
        calling-convention consistency with AdminConsole.ps1 (which passes
        the same three parameters to every tool it launches) and never
        actually wired any of them to anything -- the DL box is typed fresh
        every time, matching "standalone, no target lookup" in both
        originals' own headers. That is preserved exactly; this merge does
        not add a use for them that didn't exist before.

    Combining the two into one Add/Remove: when both Send-To and Send-As are
    checked, one click runs Send-To's full BEFORE/action/AFTER block, then
    Send-As's, into one combined result box, behind ONE confirmation dialog
    that names both operations up front. Each operation still gets its own
    Write-CompletedLog line, independently, so a run where only one of the
    two actually changed anything (e.g. Send-To was already satisfied but
    Send-As wasn't) logs exactly the one that did.

    Warning label: kept SendToDL.ps1's stronger "conflict risk" wording
    (Get-/Set-DistributionGroup and Get-Recipient exist in both on-prem
    Exchange and Exchange Online, so an also-connected EXO session in this
    same host could silently resolve against the wrong DL), not SendAsDL.
    ps1's plain "connect before running" wording -- because Resolve-
    DLEnvironment's own HQ/Franchise probe calls Get-DistributionGroup on
    every single Lookup/Add/Remove regardless of which Access Type
    checkbox(es) are selected, so that collision risk applies to the merged
    tool unconditionally, not only when Send-To is checked.

    Deliberate deviation from the standard Parse/Run/Close button triad
    (Section 13), preserved from both originals: Lookup / Add Access /
    Remove Access / Close. No separate Parse step; Lookup IS the preview,
    and Add/Remove each re-run their own BEFORE snapshot(s) internally at
    the start of every run regardless of whether Lookup was clicked first.

    CMDLET COLLISION WARNING (Section 3), same as both originals: every
    cmdlet here is called by its bare name (Get-/Set-DistributionGroup,
    Get-Recipient, Get-/Add-/Remove-ADPermission, Get-ADUser), never through
    Invoke-EXOCommand, because this tool wants whatever on-prem Exchange
    Management Shell session is already loaded and has no Exchange Online
    dependency at all -- same as DLManager.ps1, SendToDL.ps1, and
    SendAsDL.ps1 before it.
.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-07-21 (SendToDL.ps1, SendAsDL.ps1, as two separate
                       tools)
    Revised         : 2026-07-24 (merged into DLAccessManager.ps1 -- see
                       REVISED note above)
    Run from        : on-prem Exchange Management Shell (or Exchange snap-in
                       loaded), with the ActiveDirectory (RSAT) module
                       available for Get-ADUser (needed only if Send-As is
                       used)
    Required modules: ActiveDirectory (Send-As only; Send-To alone needs only
                       the on-prem Exchange Management Shell)
```


## region Parsing

- **L144-147** Tolerates newline, comma, and semicolon as separators (Section 21 convention). No domain-append -- these are AD logons (firstname.lastname), not email addresses. Byte-for-byte identical to both source tools' own copy of this function.

## region Send-To actions -- from SendToDL.ps1, unchanged

- **L162-165** $DCParam is a splatted hashtable -- @{ DomainController = <fqdn> } for Franchise, @{} for HQ. Every Exchange cmdlet call below uses @DCParam instead of a hardcoded -DomainController so HQ never passes a bare domain name where a real server is required.

## function Get-SendToState

- **L168** on-prem
- **L172** on-prem

## function Format-SendToState

- **L205-207** Orchestrates one Add or Remove run against a single DL for Send-To. Resolves every current entry AND every typed user to a DistinguishedName so the diff is never fooled by two different-looking identities for the same recipient.

## function Invoke-SendToChange

- **L224** on-prem
- **L269** on-prem
- **L272** on-prem

## region Send-As actions -- from SendAsDL.ps1, unchanged

- **L286-287** Pulls the SamAccountName back out of whatever form Get-ADPermission reports an existing grantee in (commonly "DOMAIN\sam", sometimes just "sam").

## function Get-SendAsState

- **L303-305** Explicit (non-inherited) Send-As grants only -- inherited ACEs and NT AUTHORITY\SELF are noise for a delegation-tracking tool, not permissions anyone granted.

## function Add-SendAsGrant

- **L338** ExtendedRights "Send-As" (hyphenated) -- see header KNOWN SYNTAX GOTCHA.

## function Remove-SendAsGrant

- **L347-351** Orchestrates one Add or Remove run against a single DL for Send-As, one Add-/Remove-ADPermission call per resolved user (naturally incremental, unlike Send-To's full-list replace). $ADServer feeds Get-ADUser -Server (genuine ActiveDirectory module, bare domain name is fine); $DCParam is the splat used for the three Exchange *-ADPermission calls (empty for HQ).

## region Combined orchestration

- **L422-426** ADDED 2026-07-24 -- the only genuinely new logic in this merge: runs whichever of Send-To / Send-As are selected, in that order, into one combined output, keeping each side's own Changed/ActedOn separate so the caller can log each operation type independently (see header note on why a run that only changed one of the two should only log that one).

## region WinForms UI

- **L476-478** Kept SendToDL.ps1's stronger "conflict risk" wording -- see header note on why this applies to the merged tool unconditionally, not only when Send-To is checked.
- **L487-490** Environment radio group and DC selector -- never existed in this merged file; both source tools had already removed them (commented out, not deleted, in each source file) before this merge, per the same 2026-07-21 / 2026-07-23 revisions described in the header. Nothing to carry over here.
- **L492-494** Fixed Franchise DC. LOCAL to this file (not Common.ps1) -- see AdminConsole.ps1's header for why a shared copy there caused a real breakage in disable-intake.ps1. Used only by Resolve-DLEnvironment below.
- **L497-502** Byte-for-byte identical in both source files already -- see header note. Tries Get-DistributionGroup against HQ first (no DC pinned, forest-wide), then Franchise (pinned to $script:DefaultDC) only if HQ doesn't have it. Returns BOTH shapes this file needs -- $DCParam (empty-for-HQ splat) for every Exchange cmdlet call, and $ADServer (a plain string) for Get-ADUser's -Server parameter, needed only when Send-As is used.

## function Resolve-DLEnvironment

- **L507** on-prem, HQ (no DC pinned)
- **L513** on-prem, Franchise
- **L545-548** Access Type -- two INDEPENDENT checkboxes, not a mutually exclusive radio group. See header DESIGN NOTE for why this merge is different from MailboxOneDriveAccess.ps1's radio-button approach: there is no rule here saying Send-To and Send-As can never both be needed.
- **L574-576** Disabled (and unchecked, so no stale state lingers) whenever Send-To itself is unchecked -- this checkbox only ever matters inside the Send-To Add path.

## function Get-AccessTypeSummary

- **L764** ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
- **L766** Console, which opens top-left. Helper is in Common.ps1.
