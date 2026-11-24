# M365_View_OutOfOffice.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Out of Office view for M365_Host_Web.ps1.
.DESCRIPTION
    ADDED 2026-07-24. Browser equivalent of the WinForms OOFEditor.ps1.

    Contains no logic -- Get-OutOfOfficeConfig / Set-OutOfOfficeConfig live in
    M365Logic.psm1, moved there unchanged from the WinForms tool (unchanged
    there in turn from the original monolith's lines 1108-1269). The State and
    Audience dropdown values come from the module too
    (Get-OutOfOfficeStates / Get-OutOfOfficeAudiences) so the two front ends
    can't drift on what's offered.

    CONFIRMATION IS THIS VIEW'S JOB -- Set-OutOfOfficeConfig never prompts.

    ONE GENUINE IMPROVEMENT over the WinForms version, not just a port: the
    original used free-text boxes for Start/End and called [DateTime]::Parse
    on whatever was typed, which throws on a typo and is ambiguous across
    locales (is 03/04 March 4th or April 3rd?). This uses
    <input type="datetime-local">, so the browser supplies a real picker and
    submits an unambiguous ISO-8601 value. The server still parses defensively
    and reports a clear error if it somehow gets something unparseable.

    Start/End are only sent when State is 'Scheduled' -- the module enforces
    that too, but the page also disables the fields so the UI can't imply
    they matter when they don't.
```


## function syncSchedule

- **L189** <input type="datetime-local"> wants yyyy-MM-ddTHH:mm exactly.
- **L208-210** Parse defensively even though the browser sends ISO-8601 -- a hand-crafted request could send anything, and a clear error beats a confusing Exchange one.


## Embedded HTML / JavaScript comments

- Schedule fields only matter for 'Scheduled' -- grey them out otherwise
- rather than letting the form imply they're used.
