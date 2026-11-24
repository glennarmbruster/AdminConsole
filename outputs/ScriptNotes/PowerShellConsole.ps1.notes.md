# PowerShellConsole.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    PowerShell Console -- ad hoc command execution, in-session, logged. Meant
    for ONE console only, never distributed.
.DESCRIPTION
    NOT part of the standard tiered-delegation model this suite otherwise
    follows. Every other tool script has guardrails, confirm dialogs, or at
    minimum Write-CompletedLog on its own writes -- this one has none of that
    by design. It runs exactly what's typed, exactly as typed, with the full
    privileges of whatever is connected in this PowerShell session. That's
    the entire point of it, and also exactly why the operator who requested
    this was explicit that it installs on their own console only and is never
    copied into a Tools.json handed to another admin. Nothing in AdminConsole.
    ps1, ToolManager.ps1, or Tools.json needed to change to support that --
    Start-ToolScript already handles a Tools.json entry pointing at a script
    that isn't present in a given folder gracefully (a "Tool Script Not
    Found" message, not a crash -- see AdminConsole.ps1's own header). Simply
    don't include this file, or this file's Tools.json entry, in any copy of
    this suite handed to anyone else.

    Runs IN-SESSION via the same call-operator mechanism Start-ToolScript
    already uses for every tool in this suite -- inherits whatever Exchange
    Online connection / ActiveDirectory module are already loaded here,
    exactly like every other tool. Nothing special was needed to make that
    work: module state and a live EXO connection are process-level, not
    scoped to a particular script or function, so they're visible here
    automatically.

    -SmtpAddress / -Environment / -DomainController are declared only because
    Start-ToolScript always calls every tool with those three named
    parameters -- omitting them here would make every launch from
    AdminConsole throw a parameter-binding error. This tool ignores all
    three: it isn't object-scoped, never calls Resolve-ToolTarget, and never
    prompts for a target the way every other tool does.

    MANDATORY LOGGING, NOT OPTIONAL: every non-blank command is written to
    Write-CompletedLog BEFORE it's executed, not after -- so even a command
    that crashes this window, hangs it, or takes AdminConsole down with it
    (a real risk here more than anywhere else in this suite -- see below)
    still leaves a record that it was attempted. This was the one hard
    requirement from the design discussion that led to building this script.
    Do not add a way to skip or disable it.

    Write-CompletedLog (Common.ps1) wraps its Task string in literal double
    quotes with no CSV-standard escaping for embedded quotes -- perfectly
    fine for every other tool's short, internally-built log strings, which
    never contain a stray " or a line break, but a real risk for arbitrary
    typed PowerShell, which very commonly contains its own quotes and can
    span multiple lines. ConvertTo-LogSafeCommandText below flattens
    multi-line input to one line and swaps embedded double quotes for single
    quotes before logging -- a workaround local to this script, NOT a fix to
    Write-CompletedLog itself. That remains the broader, still-open
    "enterprise logging" gap flagged separately (operator identity, CSV
    escaping, centralized location) -- out of scope here.

    Execution uses Invoke-Expression, deliberately. Most PowerShell security
    guidance flags Invoke-Expression as something to avoid -- but that
    guidance is specifically about not running strings sourced from untrusted
    or remote input through it. Here the "untrusted string" is quite
    literally typed by the operator, by hand, into a box on their own
    console, for the express purpose of running it immediately. That's the
    textbook legitimate use case for Invoke-Expression, not the anti-pattern
    the usual warnings are about.

    Output -- both success output and errors, captured together via 2>&1 --
    is APPENDED to a running transcript in this window rather than cleared
    on every run, so a whole session's worth of ad hoc commands and their
    results stays visible and scrollable. Clear Output resets it.

    Trade-off worth naming plainly, carried over from AdminConsole's own
    documented reasoning for running tools in-session: an unhandled or
    interactive-prompting command typed here can hang or crash this window,
    and because this runs in AdminConsole's own process (not a separate one),
    that can take AdminConsole down with it too. That risk exists for every
    tool in this suite already -- it's simply far more exposed here, since
    arbitrary typed commands are exactly the kind of thing most likely to
    trigger it. No guardrails, no confirmation dialogs, no undo. A destructive
    command run here runs immediately, exactly as typed.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-18
```


## region Data layer

- **L96-99** Write-CompletedLog has no CSV-standard escaping (no doubled "" for embedded quotes) -- see header. Flattens to one line and swaps " for ' before it ever reaches Write-CompletedLog, so a typed command containing quotes or line breaks can't corrupt CompletedTasks.csv.

## function Invoke-AdHocCommand

- **L182-183** Logged BEFORE execution, unconditionally -- see header. Do not move this after the try/catch below, and do not make it conditional.
- **L202-203** Ctrl+Enter runs without adding a newline first -- plain Enter still just adds a line, matching how a multi-line command box should behave.
- **L213-214** ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin Console, which opens top-left. Helper is in Common.ps1.
