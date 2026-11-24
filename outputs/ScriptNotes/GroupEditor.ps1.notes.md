# GroupEditor.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Standalone Group Editor (owner / member delegation for M365 Groups and
    on-prem synced groups) tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\GroupEditor.ps1
    prompts for a group / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Group Editor window.

    Fully functional from AdminConsole.ps1 too:
        .\GroupEditor.ps1 -SmtpAddress group@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Group Editor window
    for that group.

    Get-GroupState, Format-GroupState, Add-UnifiedOwner, Remove-UnifiedOwner,
    Add-UnifiedMember, Remove-UnifiedMember, Set-OnpremOwner, Add-OnpremMember,
    and Remove-OnpremMember below are MOVED, UNCHANGED, from the original
    monolithic Mailbox-Group-Inspector.ps1 (original lines 985-1105, region
    "Actions - group").

    THIS SCRIPT IS DIFFERENT FROM EVERY OTHER TOOL SCRIPT in one respect:
    AdminConsole.ps1 only ever passes -SmtpAddress/-Environment/-DomainController
    to a tool -- it does NOT pass the object Type (UnifiedGroup vs
    OnpremGroup) or the on-prem ADGroupDN it already resolved while deciding
    which buttons to show (see Resolve-Target's own comment in Common.ps1:
    "Individual tool scripts do NOT need to call this themselves -- each one
    already assumes its own target type"). Every other tool assumes Mailbox
    and never needs to ask. Group editing genuinely cannot proceed without
    knowing Type (it changes which cmdlets run) and, for an on-prem group,
    ADGroupDN (Get-ADGroup/Get-ADGroupMember/Set-ADGroup/Add-ADGroupMember/
    Remove-ADGroupMember all key off it, not off the mail-address Identity).
    So this script calls Resolve-Target itself, immediately after
    Resolve-ToolTarget resolves Smtp/Environment/ADServer, to get exactly
    the Type/ADGroupDN it needs -- the one exception to the "tool scripts
    don't call Resolve-Target" rule, made necessary by what this specific
    tool does, not a change to that rule for anyone else.

    LIKE MailboxPermissions.ps1 and InboxRules.ps1, this tool's editable
    controls (the grantee textbox and the four Owner/Member buttons) never
    had their own Show-XForm in the monolith -- they lived directly on the
    monolith's shared main form as the "$grpGrp" group box, and
    Invoke-GroupEdit (original lines 4136-4168) referenced $grpGrp /
    $txtGrpGrantee / $txtOut / $btnRun / $script:LastSmtp / $script:Target /
    Get-SelectedADServer / Get-SelectedEnvLabel as outer-scope
    variables/functions belonging to that shared form. So this file adds ONE
    new function, Show-GroupEditorForm, that did not exist in the monolith --
    a small new wrapper window built to hold exactly the same
    grantee-textbox-plus-four-button layout the monolith drew directly on
    its main form, plus a state display box on top (populated via
    Get-GroupState / Format-GroupState, the same before/after pattern
    Invoke-GroupEdit already used).

    Invoke-GroupEdit itself is reproduced with the following NECESSARY,
    mechanical adaptations (the monolith's version cannot run unmodified
    outside its main form, since the controls/variables/functions it
    referenced no longer exist):
      - $grpGrp / $txtGrpGrantee / $txtOut are now the controls built by
        Show-GroupEditorForm below, instead of the monolith's main form
        controls of the same names.
      - $btnRun.Enabled = $true/$false (disabling the main form's Lookup
        button during the edit) is removed -- there is no main-form Lookup
        button in this process.
      - $script:LastSmtp / $script:Target are replaced with the -Smtp /
        -Type / -ADGroupDN / -ADServer parameters this script already
        resolved (Invoke-GroupEdit is now a nested function inside
        Show-GroupEditorForm and reads $Smtp/$Type/$ADGroupDN/$ADServer from
        that enclosing scope instead).
      - The "Run a Lookup first" guard is removed -- Resolve-ToolTarget plus
        this script's own Resolve-Target call already guarantee a resolved
        group before this window opens at all, so that state is no longer
        reachable here.
      - Get-SelectedADServer() calls are replaced with the enclosing
        $ADServer parameter (Get-SelectedADServer read $cboDC.SelectedItem /
        $radHQ.Checked off the monolith's main form, which no longer exists;
        Common.ps1's Get-ADServerString already produced the equivalent
        value once, during this script's own Resolve-Target call above).
      - The Write-CompletedLog line's "(Get-SelectedEnvLabel)" call is
        replaced with the -Environment parameter this script already
        resolved.
    Every other line of every function, including all comments and the
    confirm-dialog / before-after-log mechanics, is unchanged in behavior.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
```

- **L101** operator cancelled the prompt
- **L103-106** See .DESCRIPTION above: this is the one tool script that must call Resolve-Target itself, since AdminConsole.ps1 doesn't hand off the object Type/ADGroupDN it already resolved while deciding whether to show this tool's button at all.

## function Format-GroupState

- **L199** Unified Group (Microsoft 365 Group) edits -- EXO, resolved via Invoke-EXOCommand

## function Remove-UnifiedMember

- **L224** On-prem group edits -- ActiveDirectory module, unambiguous cmdlets

## function Set-OnpremOwner

- **L227** Replaces the single AD ManagedBy value -- see header "KNOWN LIMITATION".

## region NEW - wrapper window for group delegation (see .DESCRIPTION above)

- **L241-245** This function did not exist in the monolith. It exists only to give Invoke-GroupEdit the $grpGrp/$txtGrpGrantee/$txtOut controls it already expected to find in outer scope, using the exact same grantee-textbox + four-button layout the monolith drew directly on its shared main form (original "$grpGrp" group box), plus a BEFORE/AFTER state display on top.

## function Update-StateDisplay

- **L319-323** Reproduced from Invoke-GroupEdit (original lines 4136-4168) -- see the .DESCRIPTION header above for the exact list of mechanical adaptations (control names now resolve to this form's own controls, $btnRun disabling removed, $script:LastSmtp/$script:Target replaced by the enclosing $Smtp/$Type/$ADGroupDN/$ADServer, the "Run a Lookup first"
- **L325-328** Resolve-Target call already guarantee a resolved group, Get-SelectedADServer replaced by the enclosing $ADServer, Get-SelectedEnvLabel replaced by the enclosing $Environment). The confirm-dialog / BEFORE-AFTER-log mechanics themselves are unchanged.

## function Invoke-GroupEdit

- **L390-391** ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin Console, which opens top-left. Helper is in Common.ps1.


## Trailing notes

- guard removed since Resolve-ToolTarget + this script's own
