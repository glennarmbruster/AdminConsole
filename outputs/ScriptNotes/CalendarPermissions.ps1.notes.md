# CalendarPermissions.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Standalone Calendar Permissions tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\CalendarPermissions.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Calendar Permissions window.

    Fully functional from AdminConsole.ps1 too:
        .\CalendarPermissions.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Calendar Permissions
    window for that mailbox.

    Get-CalendarFolderIdentity, Get-CalendarPermissions, Set-CalendarPermission,
    Remove-CalendarPermission, and Show-CalendarPermissionsForm below are
    MOVED, UNCHANGED, from the original monolithic Mailbox-Group-Inspector.ps1
    (original lines 1272-1455, region "Actions - calendar permissions"). The
    only edit anywhere in this file that touches original tool behavior is
    mechanical: Show-CalendarPermissionsForm's closing call changed from
        [void]$calForm.ShowDialog($form)
    to
        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
        # Console, which opens top-left. Helper is in Common.ps1.
        Set-FormScreenCorner -Form $calForm -Corner 'BottomLeft'
        [void]$calForm.ShowDialog()
    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. Every other line, including all comments, is
    byte-for-byte the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
```

- **L48** operator cancelled the prompt

## function Get-CalendarFolderIdentity

- **L52-54** Resolves the actual calendar folder rather than hardcoding ":\Calendar", since the folder can be renamed or (rarely) localized. Falls back to the first Calendar-scope folder if none matches the default root path.

## function Set-CalendarPermission

- **L73-76** Add-MailboxFolderPermission fails if the user already has an entry, so check first (best-effort) rather than parsing the resulting error text. If the check itself errors for any reason, fall through to Add-; a genuine problem will surface clearly from that call instead.

## function Remove-CalendarPermission

- **L97-98** Remove-MailboxFolderPermission removes ALL permissions for this user on the folder -- there's no selective/partial removal, per Microsoft Learn.

## function Update-CalGridFromCurrent

- **L233-236** CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the monolith's shared parent window. This script has no parent window, so it shows itself as the top-level dialog instead. Nothing else in this function changed.
