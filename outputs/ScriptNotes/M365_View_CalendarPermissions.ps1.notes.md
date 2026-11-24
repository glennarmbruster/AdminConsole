# M365_View_CalendarPermissions.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Calendar Permissions view for M365_Host_Web.ps1.
.DESCRIPTION
    ADDED 2026-07-24. Browser equivalent of the WinForms CalendarPermissions.ps1.

    Contains no logic -- Get-CalendarFolderIdentity, Get-CalendarPermissions,
    Set-CalendarPermission, Remove-CalendarPermission, and
    Get-CalendarPermissionRoles all live in M365Logic.psm1, moved there
    unchanged from the WinForms tool (which in turn had them unchanged from the
    original Mailbox-Group-Inspector monolith).

    CONFIRMATION IS THIS VIEW'S JOB. The module's Set-/Remove- functions never
    prompt. Note especially that Remove-MailboxFolderPermission removes ALL of
    a user's permissions on the calendar -- there is no partial removal -- so
    the confirm text says so explicitly rather than implying something
    narrower.

    The role list comes from the module rather than being hardcoded in this
    HTML, so the WinForms tool and this page can't drift apart on which roles
    are offered.
```

- **L28-29** Interpolating here-string: the role dropdown is built server-side from Get-CalendarPermissionRoles so there is one source of truth.

## function setActEnabled

- **L177-178** $_ may already be a formatted string (rethrown from Get-CalendarPermissions' own .Error), so don't double-format.


## Embedded HTML / JavaScript comments

- Clicking a row loads that user into the grantee box -- saves
- retyping an address just to change a role.
