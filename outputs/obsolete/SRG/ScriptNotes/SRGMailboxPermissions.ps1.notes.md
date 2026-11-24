# SRGMailboxPermissions.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    SRG (cloud-only, srgpros.com) Mailbox Permissions (Full Access / Send As
    delegation) tool.
.DESCRIPTION
    Fully functional on its own:
        .\SRGMailboxPermissions.ps1
    prompts for a mailbox via SRGCommon.ps1's Show-TargetPromptDialog, then
    opens the Mailbox Permissions window.

    Fully functional from SRGAdminConsole.ps1 too:
        .\SRGMailboxPermissions.ps1 -SmtpAddress user@srgpros.com
    skips the prompt entirely and opens straight to the Mailbox Permissions
    window for that mailbox.

    Ported from the original suite's MailboxPermissions.ps1 -- Get-MailboxState,
    Format-MailboxState, Grant-/Revoke-FullAccessPermission, and
    Grant-/Revoke-SendAsPermission are unchanged (Add-/Remove-MailboxPermission
    and Add-/Remove-RecipientPermission are all EXO cmdlets -- there was never
    any on-prem dependency here to remove). Show-MailboxPermissionsForm and
    Invoke-MailboxEdit are also unchanged apart from the same mechanical edits
    every SRG tool script needs:
      - Dot-sources SRGCommon.ps1, not Common.ps1.
      - No -Environment / -DomainController parameters. The Write-CompletedLog
        line inside Invoke-MailboxEdit (which logged "($Environment; grantee:
        ...)" in the original) now just logs "(grantee: ...)" -- there's only
        one environment in this suite to distinguish.
      - Resolve-ToolTarget is called with only -SmtpAddress.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21 (ported from MailboxPermissions.ps1)
```

- **L44** operator cancelled the prompt

## function Get-MailboxState

- **L74-75** Per Microsoft KB4486848, ArchiveStatus can read "None" even when an archive is active. ArchiveGuid is the reliable indicator.

## function Update-StateDisplay

- **L248-251** CHANGED FROM ORIGINAL: the original's Write-CompletedLog call logged "($Environment; grantee: $grantee)". SRG has no Environment concept, so that label is simply omitted -- the confirm-dialog / BEFORE-AFTER mechanics themselves are otherwise unchanged.
