# InboxRules.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Standalone Inbox Rules tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\InboxRules.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Inbox Rules window.

    Fully functional from AdminConsole.ps1 too:
        .\InboxRules.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Inbox Rules window
    for that mailbox.

    Enable-TargetInboxRule, Disable-TargetInboxRule, Remove-TargetInboxRule,
    Update-RulesGrid, and Show-RuleDetailForm below are MOVED, UNCHANGED,
    from the original monolithic Mailbox-Group-Inspector.ps1 (original lines
    807-982, region "Actions - inbox rules"; Get-InboxRulesForMailbox from
    that same region now lives in Common.ps1 -- see its header for why).

    UNLIKE the other tool scripts extracted so far, the grid this tool
    displays (grdRules) never had its own Show-XForm in the monolith -- it
    lived directly on the monolith's shared main form, and Update-RulesGrid
    referenced $grdRules/$txtOut as outer-scope variables belonging to that
    shared form (see Update-RulesGrid's own original comment: "Referencing
    $grdRules/$txtOut here without $script: works the same way
    Invoke-MailboxEdit/Invoke-GroupEdit already do -- they're script-scope
    (top-level) variables by the time any button handler calls into this
    function"). So this file adds ONE new function, Show-InboxRulesForm,
    that did not exist in the monolith -- it is a small new wrapper window
    built to hold exactly the same grid layout (same columns, same
    double-click-for-detail interaction) the monolith drew directly on its
    main form, so that Update-RulesGrid and Show-RuleDetailForm can be moved
    here completely unchanged. This is the same treatment
    MailboxPermissions.ps1 and GroupEditor.ps1 need for their own
    main-form-embedded controls -- see REFACTOR-DESIGN.md Section 7.

    Mechanical-only changes to the moved code:
      1. Show-RuleDetailForm's closing call changed from
             [void]$ruleForm.ShowDialog($form)
         to
             [void]$ruleForm.ShowDialog()
         because there is no longer a parent $form to own it.
      2. Update-RulesGrid's $grdRules/$txtOut references now resolve to the
         controls built in the new Show-InboxRulesForm wrapper instead of
         the monolith's main form -- same "script-scope variables by the
         time any handler calls into this function" mechanism the original
         comment already describes, just with a different (also top-level,
         relative to this process) form supplying them.
    Every other line, including all comments, is byte-for-byte the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
```

- **L69** operator cancelled the prompt

## function Enable-TargetInboxRule

- **L73-82** BUGFIX (2026-07-09): do NOT also pass -Mailbox here. $RuleIdentity comes from Get-InboxRule -Mailbox $Smtp (see Update-RulesGrid), and Exchange returns that Identity already scoped to the mailbox it was queried against. Passing -Mailbox in addition to that already-scoped -Identity throws "You can't use both the Identity and Mailbox parameters to specify a mailbox" -- confirmed against the Enable-InboxRule syntax on Microsoft Learn, where -Mailbox exists only to disambiguate a bare rule NAME (per the Disable-InboxRule/Remove-InboxRule examples on the same page), which is not what's happening here. $Smtp is still accepted as a parameter and used by the caller for logging (Write-CompletedLog).

## function Disable-TargetInboxRule

- **L89-90** BUGFIX (2026-07-09): see the comment above Enable-TargetInboxRule -- same fix, same reason (this is the exact cmdlet/error the operator hit).

## function Remove-TargetInboxRule

- **L97-98** BUGFIX (2026-07-09): see the comment above Enable-TargetInboxRule -- same latent bug (not yet hit in the field, but identical cause).

## function Update-RulesGrid

- **L105-109** Repopulates the grdRules grid (defined in the WinForms UI region) for the given mailbox. Referencing $grdRules/$txtOut here without $script: works the same way Invoke-MailboxEdit/Invoke-GroupEdit already do -- they're script-scope (top-level) variables by the time any button handler calls into this function.

## function Show-RuleDetailForm

- **L246-249** CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the monolith's shared parent window. This script has no parent window, so it shows itself as the top-level dialog instead. Nothing else in this function changed.

## region NEW - wrapper window for the grid (see .DESCRIPTION above)

- **L255-258** This function did not exist in the monolith. It exists only to give Update-RulesGrid the $grdRules/$txtOut controls it already expected to find in outer scope, using the exact same grid layout/columns the monolith drew directly on its shared main form.

## function Show-InboxRulesForm

- **L332-333** ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin Console, which opens top-left. Helper is in Common.ps1.
