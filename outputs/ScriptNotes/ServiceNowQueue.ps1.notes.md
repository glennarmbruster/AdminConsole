# ServiceNowQueue.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Opens the team's ServiceNow ticket queue in the default browser.
.DESCRIPTION
    Deliberately the simplest possible tool script in this suite -- one
    Start-Process call, nothing else. No lookup, no Exchange/AD cmdlets, no
    WinForms UI of its own. Accepts the same -SmtpAddress / -Environment /
    -DomainController parameters every other tool script accepts (so
    AdminConsole.ps1's Start-ToolScript can launch it exactly like any other
    tool, with no special-casing there), but ignores all three -- the
    ServiceNow queue isn't target-specific.

    The AdminConsole "always available regardless of last lookup" behavior
    for this button comes from its Tools.json ObjectTypes entry listing all
    three known types (Mailbox, UnifiedGroup, OnpremGroup), not from
    anything in this script -- see Common.ps1/AdminConsole.ps1's
    Update-ToolButtonState for why that's what keeps a button from being
    greyed out after a lookup.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21
```
