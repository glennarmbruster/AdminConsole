# M365_View_MailboxHealth.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Mailbox Health view for M365_Host_Web.ps1 -- HTML fragment plus its own
    API route. Dot-sourced INTO the host process.
.DESCRIPTION
    ADDED 2026-07-24 as part of the single-process host rewrite. Replaces
    M365_MailboxHealth_Web.ps1, which was a standalone script with its own
    HttpListener on its own port -- see M365WebCommon.ps1's header for why
    that model was abandoned (per-process Exchange connections, and a real
    macOS hang).

    This file is dot-sourced by the host, so it runs IN the host process and
    shares the host's single Exchange Online connection. It contains no
    listener, no port, no connection handling, and no logic -- only:
      - the HTML fragment for this view
      - one route, GET /api/health

    All actual work is M365Logic.psm1's Get-MailboxHealthSnapshot and
    Format-MailboxHealthReport, exactly the same two functions the WinForms
    M365_MailboxHealth.ps1 calls. One implementation, two front ends.

    The fragment is a literal here-string (@' '@) rather than an
    interpolating one, deliberately: it's full of JavaScript using $ and {}
    and template syntax, and a literal here-string means PowerShell won't
    try to interpret any of it. Nothing in this fragment needs PowerShell
    interpolation -- the CSS comes from the host once, and the mailbox
    address is typed by the operator rather than injected at render time.
```
