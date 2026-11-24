# M365_View_MailboxOneDriveAccess.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Mailbox & OneDrive Access view for M365_Host_Web.ps1. Bulk, many-to-many.
.DESCRIPTION
    ADDED 2026-07-24. Browser equivalent of MailboxOneDriveAccess.ps1 (itself
    the merge of Set-MailboxAccess.ps1 and onedriveaccess.ps1).

    Contains no logic -- Get-M365AccessFlags, Get-M365AccessTypeText, and
    Invoke-M365MailboxOneDriveAccess live in M365Logic.psm1.

    FIVE ACCESS TYPES as mutually exclusive radios, matching the WinForms
    tool: Full Access / Send-As / OneDrive / Full & Send-As / Full & OneDrive.
    Send-As + OneDrive is deliberately NOT one of them, per the operator's
    rule that such a combination is never needed -- so it's structurally
    impossible to request rather than merely validated against.

    CONFIRMATION IS THIS VIEW'S JOB. The module function does exactly what
    it's told with no prompt, and this is a CROSS PRODUCT -- three mailboxes
    and four users is twelve permission changes -- so the confirm shows the
    operation count explicitly, not just the lists.

    ONEDRIVE LINK, per the original tool's requirement: when a run grants
    OneDrive access and at least one grant on a given mailbox succeeded, that
    mailbox's OneDrive URL is surfaced at the top of the result as a clickable
    link. The module reports which mailboxes actually succeeded
    (.OneDriveGrantedFor), so a link never appears for access that doesn't
    exist.
```

- **L193-195** Normalize + domain-append both lists, and drop any user that is also a mailbox -- a mailbox granting itself access is meaningless (same dedup the original tool applied).


## Embedded HTML / JavaScript comments

- OneDrive links, only for mailboxes where a grant actually succeeded.
