# M365_Host_Web.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    M365 Admin Host -- ONE long-running process that connects to Exchange
    Online (and optionally SharePoint Online and a directory module) once,
    and serves every tool from a single port.
.DESCRIPTION
    ADDED 2026-07-24. This is the whole browser-based toolkit now.

    RUN IT IN WINDOWS POWERSHELL 5.1 for the full experience:

        powershell
        cd <this folder>
        .\M365_Host_Web.ps1

    REVISED 2026-07-24, later -- retargeted from #Requires -Version 7.0 to
    5.1. This was not cosmetic. The operator's actual daily routine is three
    connects, and only ONE of them works under PowerShell 7:
        Connect-ExchangeOnline   PS7-native                        OK
        Connect-SPOService       needs -UseWindowsPowerShell in PS7
        Connect-AzureAD          Windows PowerShell 5.1 ONLY
    (Both limitations confirmed against Microsoft's own module-compatibility
    documentation.) A PS7-only host therefore could never light up two of the
    three connection pills -- it would have been shipping indicators wired to
    nothing. System.Net.HttpListener is .NET Framework, so hosting works fine
    in 5.1, and 5.1 is where the operator's workflow already lives.

    It is still version-AWARE rather than 5.1-only: run it under PowerShell 7
    (including on macOS) and everything Exchange-based works, while the SPO
    and directory pills honestly report NotConnected because those modules
    can't load there. Nothing was given up by retargeting.

    WINDOWS NOTE: HttpListener sits on http.sys, which may want a one-time URL
    reservation. If startup fails with access-denied, the error prints the
    exact netsh command to run once.

    It will pause and ask you to run Connect-ExchangeOnline in that same
    window if you haven't already, then open http://localhost:8787/ with the
    console home page. Click a tool, use it, click "Back to Console" -- all
    in one browser tab, all against that one Exchange connection.

    WHY THIS REPLACED THE LAUNCHER (M365_AdminConsole_Web.ps1 plus the four
    standalone M365_*_Web.ps1 tools -- all now superseded, see OBSOLETE
    below): the launcher spawned each tool as its own pwsh process on its own
    port. Connect-ExchangeOnline is per-process, so every tool demanded its
    own interactive sign-in, which the operator correctly rejected as
    unusable. Worse, on macOS Start-Process does not allocate a new terminal
    -- the spawned child inherited the console's own TTY, so the child's
    connection prompt hung the console's terminal and the tool never started
    at all. The full diagnosis is in M365WebCommon.ps1's header.

    This design is the operator's own specification:
        "The preferred design is a long-running PowerShell 7 web host that
         owns authentication and stays alive all day. The web host process
         would connect to Exchange Online once, load all common modules and
         tool functions into memory, and then expose browser endpoints that
         call functions directly rather than launching separate scripts."

    No child processes exist anywhere in this design, which makes the macOS
    TTY failure structurally impossible rather than worked around. It also
    restores what the WinForms suite always had and the launcher gave up:
    connect once, every tool inherits it.

    STRUCTURE -- four layers, each with one job:
        M365Logic.psm1        logic. No UI. Imported once, here.
        M365_View_*.ps1       one file per tool: HTML fragment + its own API
                              routes. Dot-sourced INTO this process, so they
                              share the connection. Adding a tool means
                              adding one view file and one line below --
                              nothing else changes.
        M365WebCommon.ps1     host infrastructure: view registry, shared CSS,
                              connection gate, the listener loop.
        M365_Host_Web.ps1     this file: compose and serve.

    Views live in separate files specifically because the operator's stated
    reason for wanting a launcher was "so it doesn't get unwieldy -- that was
    starting to happen before already." That concern was about file size, not
    process count. Splitting per-tool UI into its own file addresses it
    without giving up the shared connection.

    ONE REQUEST AT A TIME, deliberately -- see M365WebCommon.ps1's
    concurrency note. EXO's session is not thread-safe, so the listener is
    synchronous. A long Message Trace will block other views until it
    finishes; every view shows an explicit "working..." state rather than
    looking frozen.

    OBSOLETE AS OF THIS FILE -- safe to delete, kept only so nothing is
    removed without the operator's say-so:
        M365_AdminConsole_Web.ps1        (launcher console)
        M365_MailboxHealth_Web.ps1       -> M365_View_MailboxHealth.ps1
        M365_MessageTrace_Web.ps1        -> M365_View_MessageTrace.ps1
        M365_MailboxPermissions_Web.ps1  -> M365_View_MailboxPermissions.ps1
        M365_Disable_Web.ps1             -> M365_View_Disable.ps1
    Still current and NOT obsolete: M365Logic.psm1/.psd1 (shared logic),
    M365Common.ps1 and M365_MailboxHealth.ps1 (the WinForms side, which is
    independent of all of this).

    SECURITY: binds loopback only (http://localhost:<port>/), never a
    wildcard. But this process holds a live privileged Exchange session all
    day behind endpoints with no authentication, so anything running as the
    same OS user can drive it while it's up. Fine on a personal
    workstation; think about it before running on a shared admin box.

    NOT RUN AGAINST A LIVE TENANT from here -- syntax-verified only. The
    pieces most worth watching on first real use: the connection gate's
    nested prompt (proven to work when running a tool directly, which is the
    same situation this is now in), and whether any view's JSON payload hits
    a ConvertTo-Json depth limit on unusually large results.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24
```

- **L117-124** OFF BY DEFAULT, deliberately. Loads M365_View_PowerShellConsole.ps1, which executes arbitrary PowerShell in this process with its live privileged connections. Every other view exposes a fixed set of operations; that one exposes everything the connected account can do. The host's per-run token + Origin check (see M365WebCommon.ps1) is what makes it defensible, but enabling it should be a conscious choice rather than a default, so it is a switch. See that view file's header for the full reasoning.
- **L127-136** REVISED 2026-07-24 -- the directory (AzureAD / Graph / Entra) pill is now OFF by default. An audit of every script in the toolkit found that not one of them calls a directory cmdlet for any functional purpose; the only such calls anywhere were the read-only probes that fed this pill. It was reporting on a connection nothing used, and the operator confirmed they only need Exchange Online and SharePoint Online now. Pass this to bring it back (the probe code was kept, not deleted, so that's all it takes) if directory work resumes -- most likely as Graph rather than AzureAD, and the pill relabels itself to whichever answers.
- **L139-144** --- Suite branding + tenant guard, added 2026-07-24 for ANZW/SRGW ------ This host is the ONE engine behind every web suite. ANZW_Host_Web.ps1 and SRGW_Host_Web.ps1 are thin launchers that call this file with different values here. That is deliberate and departs from the duplicate-per-suite pattern used by the WinForms suites -- see the WHY ONE ENGINE note in the .DESCRIPTION above.
- **L147-150** Arms the tenant guard. When set, the page shows a loud banner if the connected tenant's default domain is not this one -- the hazard being that three identical hosts against a domain-detecting engine would otherwise let you act on the wrong tenant without noticing.
- **L153** Keeps each suite's completed-task history in its own CSV.
- **L157-158** --- Infrastructure. M365WebCommon.ps1 imports M365Logic.psm1 itself, so --- this one dot-source brings in both the host plumbing and all the logic.
- **L161-163** Declare who we are BEFORE anything queries the tenant, so the log file and tenant guard are in place from the first request. Also clears any cached default domain, which matters if a different suite ran in this process first.

## region View loading

- **L172-175** Dot-source order = display order on the home page. To add a tool: write M365_View_<Name>.ps1 (copy an existing one for the shape), add it here. That's the whole registration process -- no ports, no registry entry, no changes to the host or the infrastructure.
- **L177** Read-only / investigative first, then the ones that change things.
- **L190-192** The console is appended ONLY when explicitly enabled -- see the -EnableConsole parameter comment. Without the switch this file is never even dot-sourced, so its route does not exist and cannot be called at all.

## region Route composition

- **L223-225** Merge every view's routes into one table. Duplicate keys are a programming error (two views claiming the same endpoint), so fail loudly at startup rather than letting one silently win at request time.

## function Get-HostPageHtml

- **L243-245** This run's API token, embedded so the page's fetch wrapper can attach it to every /api/ call. Generated by Start-M365WebHost before the listener begins, so it's always set by the time this renders.
- **L248-251** Directory pill is only emitted when -ShowDirectoryPill was passed. Not rendering it at all is better than rendering it greyed: a permanently grey light invites "is that broken?" for something that is simply not being watched.
- **L256** Home page: one card per view.
- **L268** One <section> per view, each holding that view's fragment.


## Embedded HTML / JavaScript comments

- WHICH tenant, not just whether something is connected. Given several tenants get worked in a day, this is the bit that stops you doing the right thing to the wrong place.
- Tenant guard banner. Hidden unless the connected tenant differs from the one this host declared -- see Test-M365ExpectedTenant. Sits above every view, not just home, because acting on the wrong tenant is equally bad from any of them.
- --- API token, and a global fetch wrapper that attaches it.
- Every /api/ call must carry X-M365-Token or the host rejects it with 403
- (see M365WebCommon.ps1's request-authorization region for the threat this
- addresses). Rather than editing the fetch() call in every view, fetch is
- wrapped ONCE here: same-origin /api/ requests get the header injected
- automatically. Views stay completely unaware of it, which also means a
- future view can't forget to include it.
- --- View switching. Exactly one section carries .active at a time.
- Delegated so buttons inside view fragments work too.
- --- Three connection pills: ExO / SPO / AzureAD.
- The operator's real daily routine is three separate connects, so there
- are three separate lights rather than one aggregate "M365 connected"
- that can't tell you WHICH service dropped.
- CHEAP vs EXPENSIVE probes: the Exchange check reads in-process state and
- is effectively free, so it polls on a timer. SharePoint has no local
- state cmdlet at all -- the only way to know is to make a real service
- call -- and the AzureAD check also touches the service. Those run only
- on page load and on explicit Refresh (?full=1), so an all-day host isn't
- hammering the tenant every minute just to keep a light green. A pill
- whose probe was skipped this pass shows a neutral "?" rather than
- claiming disconnected, because "not checked" and "not connected" are
- different facts.
- Tenant identity. Prefers the default accepted domain (what an admin
- actually calls "the tenant"), falls back to the *.onmicrosoft.com name
- from the connection, and only says "unknown" when neither is available
- -- which effectively means not connected yet.
- Set directly rather than via paintPill: that helper appends
- "Connected"/"NotConnected" to a label, which is the wrong grammar here
- -- this chip states WHICH tenant, it isn't a second health light.
- Green because it's a positive fact. A red tenant chip would read as
- "wrong tenant" rather than "here is the tenant".
- The directory pill only exists when the host was started with
- -ShowDirectoryPill, so everything here is guarded on aad being
- present. Label follows whichever module actually answered, so a
- Graph-connected session doesn't sit under a deprecated module name.
- Tenant guard. Only ever shown on a real mismatch -- an unarmed host,
- or one that simply can't tell yet because nothing is connected, stays
- silent rather than nagging.
- full check once on load
- cheap check thereafter
- --- Clean shutdown from the page.
