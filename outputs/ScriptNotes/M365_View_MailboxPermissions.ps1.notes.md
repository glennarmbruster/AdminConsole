# M365_View_MailboxPermissions.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Mailbox Permissions view (Full Access / Send As delegation) for
    M365_Host_Web.ps1.
.DESCRIPTION
    ADDED 2026-07-24 as part of the single-process host rewrite. Replaces
    standalone M365_MailboxPermissions_Web.ps1 -- see M365WebCommon.ps1's
    header for why the per-tool-process model was abandoned.

    Contains no logic. Everything is M365Logic.psm1:
      Get-MailboxState / Format-MailboxState for display, and
      Invoke-MailboxPermissionEdit for each grant/revoke (which does the
      BEFORE snapshot, the change, the AFTER snapshot, and the
      completed-tasks log entry in one call).

    CONFIRMATION IS THIS VIEW'S JOB. The module's grant/revoke functions
    deliberately never prompt -- so a scheduled job can call them -- which
    means the confirm() in the JS below is the only thing standing between a
    misclick and a real permission change. Do not remove it without
    replacing it with something equivalent.
```
