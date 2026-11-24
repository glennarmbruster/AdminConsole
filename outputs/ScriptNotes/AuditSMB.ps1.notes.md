# AuditSMB.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Shared MB Audit -- traces Exchange Online activity on a shared mailbox,
    cross-referenced against its current Full Access holder list.
.DESCRIPTION
    Layer 2 tool script, same conventions as the rest of this suite: dot-sources
    Common.ps1, uses Resolve-ToolTarget (prompts for a target itself when
    launched with no -SmtpAddress), Tier2 entry in Tools.json.
 
    ObjectTypes = ['Mailbox'], same as every other mailbox tool -- and per an
    explicit design decision, this button is meant to ALWAYS show for any
    Mailbox-type lookup, not just ones Exchange happens to flag as
    RecipientTypeDetails=SharedMailbox. Do not "fix" this later to gate on
    shared-mailbox-only detection -- that was considered and deliberately
    rejected: showing it always is harmless (a regular mailbox just returns a
    short or empty Full Access holder list), and adding a extra type check
    was judged not worth the complexity.
 
    -Environment / -DomainController are accepted only for calling-convention
    consistency with AdminConsole.ps1 -- this tool is pure Exchange Online /
    Microsoft Purview, no on-prem AD dependency, same as MailboxAudit.ps1.
 
    THIS IS A SEPARATE DESIGN FROM MailboxAudit.ps1, NOT A COPY -- the single
    most important difference, verified against Microsoft Learn ("Search the
    audit log for mailbox activities in specific mailboxes," July 2026):
    Microsoft's own guidance is explicit that shared-mailbox searches should
    put the mailbox's SMTP address in the FreeText/Keywords field, NOT UserIds
    ("When searching for shared mailbox activities, enter the SMTP address
    under the Keywords field, not the Users field"). MailboxAudit.ps1 uses
    -UserIds because that's Microsoft's recommended filter for a REGULAR
    mailbox; this script uses -FreeText because a shared mailbox behaves
    differently in the underlying audit data (nobody "owns" it the way a
    regular mailbox has one owner UPN that every record can be filtered
    against). Everything else -- date range, Operations list, paging, subject
    filter, retention, permissions prerequisite, ingestion delay -- carries
    over unchanged from MailboxAudit.ps1; see that script's header for the
    full sourcing on those.
 
    ATTRIBUTION (the second half of this module's design goal -- "one trace
    of the shared mailbox, holders used to attribute results"): on open (and
    on demand via Refresh Holders), this script pulls the shared mailbox's
    current Full Access holder list using the EXACT SAME verified pattern
    already proven in MailboxPermissions.ps1's Get-MailboxState -- Invoke-
    EXOCommand -Name 'Get-MailboxPermission', filtered to non-inherited,
    non-deny, AccessRights -contains 'FullAccess', excluding 'NT AUTHORITY\
    SELF'. Get-MailboxPermission is deliberately called through Invoke-
    EXOCommand (not bare), same reasoning as everywhere else in this suite
    that does so: it exists both on-prem and in EXO, so the wrapper forces
    resolution to the cloud-sourced copy regardless of what else might be
    loaded in the session.
    Each result row is then flagged against that holder list: if the row's
    Actor matches a current Full Access holder (case-insensitive), it's shown
    as a known holder; if not, it's flagged "NOT a current Full Access
    holder" -- which can mean former access since revoked, an admin/owner
    action, or access via a permission type other than FullAccess (SendAs,
    SendOnBehalf, eDiscovery, impersonation) that isn't reflected in this
    list. This is intentionally a FACTUAL comparison against current state,
    not a probabilistic guess -- it does NOT attempt the "85% confidence it
    was Bob" scoring discussed separately; that remains a deliberately later,
    out-of-scope phase (see MailboxAudit.ps1 header). The holder list itself
    is a live snapshot taken when the window opened/last refreshed -- if
    permissions changed since then, click Refresh Holders before trusting the
    flag on results from after that change.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-18
```

- **L81** operator cancelled the prompt

## region Constants

- **L84-87** Identical to MailboxAudit.ps1's list -- see that script's header for full sourcing/exclusion notes (verified against Microsoft Learn "Audit log activities," July 2026; 3 rows excluded: no documented Operation value, duplicate Operation value, and a mismatched SharePoint operation).

## region Data layer

- **L130-132** Pulls current Full Access holders for the shared mailbox -- same verified filter as MailboxPermissions.ps1's Get-MailboxState (see header). Returns a simple string array of holder identities (User.ToString()).

## function Get-FullAccessHolders

- **L145-148** Pages through Search-UnifiedAuditLog for one shared mailbox, scoped via -FreeText (NOT -UserIds -- see header) until it returns no further records or the documented 50,000-record ReturnLargeSet ceiling is hit. Returns a PSCustomObject: Records (raw rows, AuditData still a JSON string) + HitCeiling.

## function Invoke-SharedMailboxAuditSearch

- **L150-151** Cancellation checks $script:AuditCancelRequested directly, matching the convention used in MailboxAudit.ps1 and DLManager.ps1's count-scan loop.
- **L189-193** Parses one raw Search-UnifiedAuditLog row's AuditData JSON -- identical field set to MailboxAudit.ps1's ConvertTo-MailboxAuditRecord (see that script's header for the verified field-name sourcing). HolderStatus is computed separately, after this, by the caller -- it depends on the holder snapshot, which is per-window state, not something this parsing function should own.

## function ConvertTo-SharedMailboxAuditRecord

- **L214** filled in by the caller against the current holder snapshot

## function Update-HolderList

- **L457-459** Attribute each row against the holder snapshot taken when this window opened/last refreshed -- see header for why this is a factual current-state comparison, not a probability score.
