# M365_View_InboxRules.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Inbox Rules view for M365_Host_Web.ps1.
.DESCRIPTION
    ADDED 2026-07-24. Browser equivalent of the WinForms InboxRules.ps1.

    Contains no logic -- Get-InboxRuleSummaries, Enable-TargetInboxRule,
    Disable-TargetInboxRule, and Remove-TargetInboxRule live in M365Logic.psm1.

    SUSPICIOUS RULES are flagged with the same criteria Mailbox Health uses (an
    ENABLED rule that forwards, redirects, or deletes), because that shared
    definition lives in the module. It's a prompt for review, NOT an accusation
    -- plenty of legitimate rules forward or delete, and the UI says so.

    CONFIRMATION IS THIS VIEW'S JOB. Delete is the one that matters: a removed
    inbox rule cannot be restored, so its confirm says that explicitly rather
    than asking a generic "are you sure".
```
