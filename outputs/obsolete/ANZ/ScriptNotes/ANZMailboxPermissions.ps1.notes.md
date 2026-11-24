# ANZMailboxPermissions.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    ANZ (cloud-only, expresspros.com.au) Mailbox Permissions (Full Access /
    Send As delegation) tool.
.DESCRIPTION
    Fully functional on its own:
        .\ANZMailboxPermissions.ps1
    prompts for a mailbox via ANZCommon.ps1's Show-TargetPromptDialog, then
    opens the Mailbox Permissions window.

    Fully functional from ANZAdminConsole.ps1 too:
        .\ANZMailboxPermissions.ps1 -SmtpAddress user@expresspros.com.au
    skips the prompt entirely and opens straight to the Mailbox Permissions
    window for that mailbox.

    Built the same way as SRGMailboxPermissions.ps1 -- both are ported from
    the original suite's MailboxPermissions.ps1. Get-MailboxState,
    Format-MailboxState, Grant-/Revoke-FullAccessPermission, and
    Grant-/Revoke-SendAsPermission are unchanged (Add-/Remove-MailboxPermission
    and Add-/Remove-RecipientPermission are all EXO cmdlets -- there was
    never any on-prem dependency here to remove). Show-MailboxPermissionsForm
    and Invoke-MailboxEdit are also unchanged apart from the same mechanical
    edits every cloud-only tool script in this family needs:
      - Dot-sources ANZCommon.ps1, not Common.ps1 or SRGCommon.ps1.
      - No -Environment / -DomainController parameters. The Write-CompletedLog
        line inside Invoke-MailboxEdit (which logged "($Environment; grantee:
        ...)" in the original) now just logs "(grantee: ...)" -- there's only
        one environment in this suite to distinguish.
      - Resolve-ToolTarget is called with only -SmtpAddress.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21 (ported from MailboxPermissions.ps1, same way as
              SRGMailboxPermissions.ps1)
```

- **L46** operator cancelled the prompt

## function Get-MailboxState

- **L76-77** Per Microsoft KB4486848, ArchiveStatus can read "None" even when an archive is active. ArchiveGuid is the reliable indicator.

## function Update-StateDisplay

- **L250-253** CHANGED FROM ORIGINAL: the original's Write-CompletedLog call logged "($Environment; grantee: $grantee)". ANZ has no Environment concept, so that label is simply omitted -- the confirm-dialog / BEFORE-AFTER mechanics themselves are otherwise unchanged.
