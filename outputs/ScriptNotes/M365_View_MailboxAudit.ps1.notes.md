# M365_View_MailboxAudit.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Mailbox Audit view for M365_Host_Web.ps1. Read-only.
.DESCRIPTION
    ADDED 2026-07-24. Browser equivalent of the WinForms MailboxAudit.ps1.

    Contains no logic -- Get-M365MailboxAuditOperations, Invoke-M365MailboxAudit,
    and ConvertTo-M365AuditRecord live in M365Logic.psm1. The operation list is
    served from the module so the picker can't drift from what the search
    actually accepts.

    TWO THINGS THE OPERATOR SHOULD KNOW, surfaced in the UI rather than buried:

    1. TRUNCATION. Search-UnifiedAuditLog has a documented 50,000-record
       ReturnLargeSet ceiling. When the module reports HitCeiling, this view
       says the results may be incomplete instead of presenting a capped set as
       though it were the whole story. Narrow the date range or the operation
       list in that case.

    2. IT BLOCKS. The host is single-threaded by design (EXO's session isn't
       thread-safe), and a wide audit search can take a long while and page
       repeatedly. While it runs, every other view waits. The UI says so up
       front rather than looking frozen. No Cancel button is offered -- and
       that's deliberate, not an omission: the WinForms tool's Cancel worked by
       flipping a flag its modal dialog could set mid-search, which a
       one-request-at-a-time server genuinely cannot do. Bounding the search
       before starting it (dates, operations) is the honest equivalent.

    Timestamps come from the server as UTC ISO-8601 and are rendered in the
    viewer's own timezone by the browser, rather than being pre-converted to
    the server's local time.
```

- **L38-39** Operation checkboxes are built server-side from the module's map so there's exactly one list of valid operations in the whole toolkit.

## function applyFilter

- **L237-238** Treat the supplied dates as UTC day boundaries, matching the label on the form. End is inclusive of the whole day.
