# M365_View_SharedMailboxAudit.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Shared Mailbox Audit view for M365_Host_Web.ps1. Read-only.
.DESCRIPTION
    ADDED 2026-07-24. Browser equivalent of the WinForms SharedMailboxAudit.ps1.

    Contains no logic -- Invoke-M365MailboxAudit, ConvertTo-M365AuditRecord, and
    Get-M365FullAccessHolders all live in M365Logic.psm1.

    WHAT MAKES THIS DIFFERENT FROM THE PLAIN MAILBOX AUDIT, and why it's worth
    being its own view: it cross-references every actor in the audit results
    against the mailbox's CURRENT Full Access holder list. On a shared mailbox
    the interesting question isn't just "who did this" but "who did this, and
    are they still supposed to be able to?" -- so any actor who acted on the
    mailbox but is NOT in the current holder list is flagged. Common innocent
    explanations exist (their access was removed after the fact, or they acted
    via a different mechanism such as Send-As or an admin action), which is why
    it's labelled "not a current Full Access holder" rather than anything
    accusatory.

    ObjectTypes note carried forward from the WinForms tool: it deliberately
    does NOT gate on RecipientTypeDetails=SharedMailbox. That was considered and
    rejected -- running it against a regular mailbox is harmless (the holder
    list just comes back short or empty) and the extra type check bought
    nothing.

    SAME TWO CAVEATS AS MAILBOX AUDIT, surfaced in the UI: the 50,000-record
    ceiling can truncate results, and a wide search blocks every other view
    while it runs.
```


## function applyFilter

- **L243-244** Holder list first -- it's cheap, and it's useful even if the audit search then returns nothing.
- **L249-251** Case-insensitive holder set for the cross-reference. Holder names can come back in several forms, so compare on the local part too rather than requiring an exact string match.
- **L268-269** The mailbox acting on itself is not a delegate action, so don't flag it as an outsider.


## Embedded HTML / JavaScript comments

- Holder panel -- shown even when the audit itself found nothing, since
- "who can get in here" is useful on its own.
