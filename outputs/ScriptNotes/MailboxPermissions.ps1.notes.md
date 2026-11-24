# MailboxPermissions.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Standalone Mailbox Permissions (Full Access / Send As delegation) tool --
    Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\MailboxPermissions.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Mailbox Permissions window.

    Fully functional from AdminConsole.ps1 too:
        .\MailboxPermissions.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Mailbox Permissions
    window for that mailbox.

    Get-MailboxState, Format-MailboxState, Grant-FullAccessPermission,
    Revoke-FullAccessPermission, Grant-SendAsPermission, and
    Revoke-SendAsPermission below are MOVED, UNCHANGED, from the original
    monolithic Mailbox-Group-Inspector.ps1 (original lines 671-795, region
    "Actions - mailbox").

    UNLIKE most of the other extracted tools, this one's editable controls
    (the grantee textbox and the four Grant/Revoke buttons) never had their
    own Show-XForm in the monolith -- they lived directly on the monolith's
    shared main form as the "$grpMbx" group box, and Invoke-MailboxEdit
    (original lines 4083-4114) referenced $grpMbx / $txtMbxGrantee / $txtOut /
    $btnRun / $script:LastSmtp as outer-scope variables belonging to that
    shared form. So this file adds ONE new function, Show-MailboxPermissionsForm,
    that did not exist in the monolith -- a small new wrapper window built to
    hold exactly the same grantee-textbox-plus-four-buttons layout the
    monolith drew directly on its main form, plus a state display box on top
    (populated via Get-MailboxState / Format-MailboxState, the same
    before/after pattern Invoke-MailboxEdit already used). This is the same
    treatment InboxRules.ps1 needed for its own main-form-embedded grid, and
    GroupEditor.ps1 needs for its own main-form-embedded owner/member
    controls -- see REFACTOR-DESIGN.md Section 7.

    Invoke-MailboxEdit itself is reproduced with the following NECESSARY,
    mechanical adaptations (the monolith's version cannot run unmodified
    outside its main form, since the controls/variables it referenced no
    longer exist):
      - $grpMbx / $txtMbxGrantee / $txtOut are now the controls built by
        Show-MailboxPermissionsForm below, instead of the monolith's main
        form controls of the same names.
      - $btnRun.Enabled = $true/$false (disabling the main form's Lookup
        button during the edit) is removed -- there is no main-form Lookup
        button in this process.
      - $script:LastSmtp is replaced with the -Smtp parameter this script
        already resolved via Resolve-ToolTarget (Invoke-MailboxEdit is now a
        nested function inside Show-MailboxPermissionsForm and reads $Smtp
        from that enclosing scope instead).
      - The "Run a Lookup first" guard is removed -- Resolve-ToolTarget
        already guarantees a resolved mailbox before this window opens at
        all, so that state is no longer reachable here.
      - The Write-CompletedLog line's "(Get-SelectedEnvLabel)" call is
        replaced with the -Environment parameter this script already
        resolved (Get-SelectedEnvLabel read $radHQ.Checked off the
        monolith's main form, which no longer exists).
    Every other line of every function, including all comments and the
    confirm-dialog / before-after-log mechanics, is unchanged in behavior.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
```

- **L79** operator cancelled the prompt

## function Get-MailboxState

- **L109-110** Per Microsoft KB4486848, ArchiveStatus can read "None" even when an archive is active. ArchiveGuid is the reliable indicator.

## region NEW - wrapper window for mailbox delegation (see .DESCRIPTION above)

- **L210-214** This function did not exist in the monolith. It exists only to give Invoke-MailboxEdit the $grpMbx/$txtMbxGrantee/$txtOut controls it already expected to find in outer scope, using the exact same grantee-textbox + four-button layout the monolith drew directly on its shared main form (original "$grpMbx" group box), plus a BEFORE/AFTER state display on top.

## function Update-StateDisplay

- **L288-295** Reproduced from Invoke-MailboxEdit (original lines 4083-4114) -- see the .DESCRIPTION header above for the exact list of mechanical adaptations (control names now resolve to this form's own controls, $btnRun disabling removed, $script:LastSmtp replaced by the enclosing $Smtp, the "Run a Lookup first" guard removed since Resolve-ToolTarget already guarantees a resolved mailbox, Get-SelectedEnvLabel replaced by the enclosing $Environment). The confirm-dialog / BEFORE-AFTER-log mechanics themselves are unchanged.

## function Invoke-MailboxEdit

- **L349-350** ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin Console, which opens top-left. Helper is in Common.ps1.
