# OOFEditor.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Standalone Out of Office tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\OOFEditor.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Out of Office window.

    Fully functional from AdminConsole.ps1 too:
        .\OOFEditor.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Out of Office window
    for that mailbox.

    Get-OutOfOfficeConfig, Set-OutOfOfficeConfig, and Show-OutOfOfficeForm
    below are MOVED, UNCHANGED, from the original monolithic
    Mailbox-Group-Inspector.ps1 (original lines 1108-1269, region "Actions -
    out of office"). The only edit anywhere in this file that touches
    original tool behavior is mechanical: Show-OutOfOfficeForm's closing call
    changed from
        [void]$oofForm.ShowDialog($form)
    to
        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
        # Console, which opens top-left. Helper is in Common.ps1.
        Set-FormScreenCorner -Form $oofForm -Corner 'BottomLeft'
        [void]$oofForm.ShowDialog()
    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. Every other line, including all comments, is
    byte-for-byte the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
```

- **L48** operator cancelled the prompt

## function Update-OofFormFromCurrent

- **L211-214** CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the monolith's shared parent window. This script has no parent window, so it shows itself as the top-level dialog instead. Nothing else in this function changed.
