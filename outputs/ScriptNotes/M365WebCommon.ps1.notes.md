# M365WebCommon.ps1 -- design notes

> Private. Not distributed with the scripts.


## (top of file)

<!-- line 2 -->

```text
.SYNOPSIS
    Host infrastructure for M365_Host_Web.ps1 -- ONE long-running process,
    one port, one Exchange Online connection, all tools served as views.
.DESCRIPTION
    REWRITTEN 2026-07-24 -- this file previously supported a "console as
    launcher" model where each tool ran as its own spawned pwsh process on
    its own port. That model is GONE, and this rewrite is why:

    WHAT WENT WRONG WITH THE LAUNCHER (confirmed on a live macOS run, not
    theorized):
      1. Connect-ExchangeOnline is per-process. A spawned tool process
         started with no connection and no way to inherit one, so every tool
         needed its own interactive sign-in -- which the operator correctly
         called out as an unusable workflow.
      2. Worse, the Wait-M365ExoConnection gate added to paper over (1)
         made it actively break on macOS. Start-Process 'pwsh' there does
         NOT allocate a new terminal window the way it does on Windows -- the
         child inherits the parent's TTY. So the child's
         $host.EnterNestedPrompt() appeared in the CONSOLE's own terminal
         (a stray '>>' nested prompt), the child never reached its listener,
         and the console's port poll failed with "did not respond on port
         8787 within 5 seconds." The launcher hung the operator's terminal
         instead of merely being inconvenient.

    THE FIX IS ARCHITECTURAL, per the operator's own specification:
        "The preferred design is a long-running PowerShell 7 web host that
         owns authentication and stays alive all day. The web host process
         would connect to Exchange Online once, load all common modules and
         tool functions into memory, and then expose browser endpoints that
         call functions directly rather than launching separate scripts."

    So: ONE process. It connects once, imports M365Logic.psm1 once, and
    serves every tool from a single port. There are NO child processes at
    all, which makes the macOS TTY problem structurally impossible rather
    than patched. This also restores the property the WinForms suite always
    had -- connect once, every tool inherits it -- which is exactly what the
    launcher had traded away.

    THE THREE LAYERS, unchanged in spirit from the module refactor:
        M365Logic.psm1          -- logic. No UI. Imported here, once.
        M365_View_*.ps1         -- one file per tool: its HTML fragment and
                                   its own API routes. Dot-sourced INTO this
                                   host process (not spawned), so they share
                                   the connection.
        M365_Host_Web.ps1       -- composes the views, owns the listener.

    WHY VIEWS ARE SEPARATE FILES rather than all inlined into the host: the
    operator's stated reason for preferring a launcher in the first place was
    "so it doesn't get unwieldy -- that was starting to happen before
    already." That concern was about FILE SIZE, not process count, and it's
    a fair one. Splitting per-tool UI into its own view file keeps each file
    modest while still running everything in one process. Best of both.

    CONCURRENCY -- IMPORTANT, and deliberate: the listener loop below is
    synchronous. It handles exactly one request at a time. That is not
    laziness; EXO's implicit-remoting session is not thread-safe, and firing
    parallel cmdlets down one connection can corrupt it. Serializing through
    one loop is the safe design. The visible consequence is that a slow
    operation (a Message Trace pulling thousands of rows) blocks every other
    request -- including plain page navigation -- until it finishes. Every
    view's JS therefore shows an explicit "working..." state rather than
    appearing frozen. If that ever becomes intolerable, the fix is a proper
    job queue with polling, NOT parallel handlers.

    SESSION EXPIRY: a host that stays up all day will eventually lose its
    EXO session to token expiry or idle teardown. The
    'GET /api/connection' route below (served for every host) reports live
    state via M365Logic's Test-M365ExoConnection, and the host page polls it
    so the operator sees an honest banner instead of confusing
    "no command named Get-Mailbox" errors from individual tools.

    SECURITY, unchanged from before but now more relevant: the listener
    binds http://localhost:<port>/ only -- loopback, never a wildcard. But
    this process holds a live privileged Exchange session all day with no
    authentication in front of its endpoints, so anything running as the
    same OS user on this machine can drive it while it's up. Fine on a
    personal workstation; think twice on a shared admin box.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24 (as launcher infrastructure)
    Revised : 2026-07-24, later -- rewritten for the single-process host.
              Removed: $script:M365ToolRegistry / $script:M365ConsolePort
              (no per-tool ports any more), Start-M365WebTool (replaced by
              Start-M365WebHost), Get-M365CloseButtonHtml (no per-tool
              shutdown -- "Back to Console" is client-side navigation now).
```


## region Logic module import

- **L91-106** RESTORED 2026-07-24 -- this block existed in the pre-rewrite version of this file and was accidentally dropped when the file was rewritten for the single-process host. Nothing then imported M365Logic.psm1, so the very first logic call the host makes (Test-M365ExoConnection, from Wait-M365ExoConnection below) failed with "not recognized as a name of a cmdlet" -- and because an unrecognized command is non-terminating by default, $conn came back $null, $null.Connected read as falsey, and the host fell through to the connect prompt as though nothing were connected. Confusing symptom, trivial cause. This import is the ONLY place the logic module is loaded for the web side: the host dot-sources this file, and the view files are dot-sourced into that same process, so all of them see the module's functions from here. -Force so .psm1 edits are picked up without restarting the process (a second Import-Module without it is a no-op). The .psd1 manifest is preferred over the raw .psm1 so its explicit FunctionsToExport applies.
- **L118-124** Post-import sanity check. Added alongside the restore above: the bug that prompted it was silent precisely because a MISSING function is a non-terminating error in PowerShell, so the host limped onward and reported a wrong diagnosis ("not connected") instead of the real one. Verifying a few representative exports here means any future breakage of the module path, manifest, or FunctionsToExport list fails immediately, at startup, with a message that names the actual problem.

## region View registry

- **L141-150** Each M365_View_*.ps1 file appends one entry here when dot-sourced by the host. Order of dot-sourcing determines order on the home page. Entry shape: Id [string] -- short slug, used as the JS view id. Unique. Name [string] -- card title / nav label. Description [string] -- one line on the home card. Destructive [bool] -- tints the card so risky tools read as risky. GetHtml [scriptblock] -- returns this view's HTML fragment. Routes [hashtable] -- "METHOD /path" -> scriptblock($Request).

## function Register-M365View

<!-- line 154 -->

```text
.SYNOPSIS
        Adds one view to the registry. Called by each M365_View_*.ps1 file.
```


## function Get-M365Views

<!-- line 180 -->

```text
.SYNOPSIS
        Returns every registered view, in registration order.
```


## function Get-M365SharedCss

<!-- line 190 -->

```text
.SYNOPSIS
        One stylesheet for the whole host page.
    .DESCRIPTION
        Emitted ONCE by the host, so view fragments never carry their own
        <style> block and can't disagree about what a class means. Note the
        history: .card and .status previously lived in individual tool files,
        which meant the console page used .card without defining it and
        rendered its whole tool list unstyled. Everything shared lives here
        now for exactly that reason.
```


## region Request authorization (CSRF defence)

- **L254-281** ADDED 2026-07-24, before the PowerShell console view was built -- but this protects EVERY view, and is worth having regardless. THE THREAT, concretely: this host listens on localhost and holds a live privileged Exchange session all day. Any web page the operator happens to have open in the same browser can issue requests to http://localhost:<port>/ -- the browser will happily send them. Without a check, a malicious or merely compromised page could quietly POST to /api/disable/run or /api/perms/revoke-full and act with full admin authority. Loopback binding does NOT prevent this; it only stops other machines, not other tabs. TWO INDEPENDENT DEFENCES, both cheap: 1. Origin/Referer check -- browsers set these on cross-site requests and forbid pages from forging them. Anything claiming an origin other than this host's own is rejected. This alone blocks the classic "malicious site posts to localhost" attack. 2. Per-run secret token -- generated fresh each time the host starts, embedded in the page it serves, and required on every API call as the X-M365-Token header. A page that never loaded from this host cannot know it. This also covers requests that carry no Origin at all. Deliberately NOT a login: there is one operator, on their own machine, and a password prompt would add friction without adding protection against the threat that actually exists here. What this stops is OTHER SOFTWARE on the machine driving the admin session through the browser. GET / is always allowed (it's how the page bootstraps and receives the token). Everything under /api/ is gated.

## function New-M365RequestToken

<!-- line 285 -->

```text
.SYNOPSIS
        Generates and stores this run's API token. Called once by the host.
    .DESCRIPTION
        Cryptographically random via RNGCryptoServiceProvider rather than
        Get-Random -- Get-Random is not a cryptographic RNG and its output
        would be guessable, which for a security token defeats the purpose.
        RNGCryptoServiceProvider exists in .NET Framework and .NET Core, so
        this works in both Windows PowerShell 5.1 and PowerShell 7.
```


## function Test-M365RequestAuthorized

<!-- line 311 -->

```text
.SYNOPSIS
        Returns $true if a request may proceed; $false to reject with 403.
    .DESCRIPTION
        See the region comment for the reasoning. Applies to /api/ paths
        only; the host lets GET / through before calling this.
```

- **L323** --- Defence 1: Origin / Referer must be this host, when present.
- **L335-337** Neither header present: not conclusive on its own (some legitimate same-origin GETs omit both), so fall through to the token check, which is the stronger of the two defences anyway.
- **L340** --- Defence 2: the per-run token.

## function Wait-M365ExoConnection

<!-- line 351 -->

```text
.SYNOPSIS
        Pauses the HOST at an interactive sub-prompt so the operator can
        connect, before the listener starts.
    .DESCRIPTION
        Only M365_Host_Web.ps1 calls this, exactly once, at startup -- and
        that is now safe in a way it was not before. Under the old launcher
        model this ran inside SPAWNED CHILD processes which, on macOS, share
        the parent's terminal; the nested prompt appeared in the console's
        own window and hung it. With a single process there are no children:
        this runs in the terminal the operator themselves just typed into,
        which is the same situation as running a tool directly -- already
        proven to work.

        Uses M365Logic's Test-M365ExoConnection, so it distinguishes "never
        connected" from "session expired," not just "are cmdlets present."

        Never runs Connect-ExchangeOnline for the operator -- modern EXO auth
        is an interactive, MFA-capable browser sign-in, and entering
        credentials on someone's behalf isn't something to script. This only
        pauses and gets out of the way.
```

- **L393-395** The operator may have connected to a different tenant than a previous run -- drop the cached default domain so bare names resolve against whatever is connected now.

## function Start-M365WebHost

<!-- line 409 -->

```text
.SYNOPSIS
        Starts the single listener that serves every registered view.
    .DESCRIPTION
        $Routes is the MERGED route table the host builds from every view,
        plus whatever the host itself adds. Keys are "METHOD /path"; values
        are scriptblocks taking one [System.Net.HttpListenerRequest] and
        returning @{ ContentType = '...'; Body = '...' }.

        Handled here automatically, so no view needs to implement them:
          GET  /                 -- the composed single page ($GetHtml)
          GET  /api/connection   -- live EXO connection state, for the pill
          POST /api/shutdown     -- clean stop from the page
        Ctrl+C also stops cleanly -- see the CancelKeyPress handler below.
```

- **L430-432** Passed through to Get-M365ConnectionSummary. Off by default -- see that function's -IncludeDirectory help for why the directory pill is no longer monitored unless asked for.
- **L436-437** Fresh token per run -- see the request-authorization region. Generated BEFORE the page is ever served, since the page embeds it.
- **L449-452** On Windows, HttpListener sits on http.sys, which requires a URL reservation for a prefix unless the process is elevated. Localhost prefixes often work unelevated, but not always -- so name the exact fix rather than leaving an access-denied to interpret.
- **L466-476** REVISED 2026-07-24 -- version-aware, because $IsWindows / $IsMacOS / $IsLinux are PowerShell 7 AUTOMATIC VARIABLES that simply do not exist in Windows PowerShell 5.1. In 5.1 all three evaluate to $null, so the old if/elseif chain fell through silently and no browser ever opened -- a bug that only appeared once this host started targeting 5.1 (which it must, so the AzureAD and SPO modules can load; see Get-M365ConnectionSummary in M365Logic for the full reasoning). 5.1 only ever runs on Windows, so "PSVersion major < 6" is a sound Windows test. Under PS7 the real variables are used.
- **L493-498** Ctrl+C handling: GetContext() is a blocking synchronous call and PowerShell's default Ctrl+C does not reliably interrupt one. Forcibly Stop()ing the listener from this handler makes the blocked call throw, which the loop treats as a clean shutdown. $e.Cancel = $true keeps PowerShell from hard-killing the process so the finally block still runs.
- **L524-525** Always allowed -- this is how the page bootstraps and receives its token. See the request-authorization region.
- **L533-534** Rejected before ANY handler runs, so an unauthorized request never reaches a cmdlet.
- **L543-545** ?full=1 opts into the probes that cost a service call (SharePoint especially) -- the page polls without it and passes it on load / explicit Refresh.
- **L550-552** Lets the page hide the directory pill entirely rather than render a meaningless state for something nobody is monitoring.
- **L554-568** WHICH TENANT is connected -- added 2026-07-24 at the operator's request, and it's the single most useful thing on this bar. "Connected" alone doesn't tell you connected to WHAT, which matters a lot when you work several tenants in a day. Two sources, most recognisable first: tenantDomain -- the default ACCEPTED domain (expresspros.com). This is what an admin actually thinks of as "the tenant". Cached after the first lookup, so the cheap 60s poll costs nothing. organization -- what Get-ConnectionInformation reports, usually the *.onmicrosoft.com name. Used as a fallback when the accepted-domain lookup hasn't succeeded yet.
- **L570-573** Tenant guard -- see Test-M365ExpectedTenant. Only meaningful once something is connected, which is why it rides along on this route rather than being checked once at startup.
- **L630-632** A handler blew up. Log it to the host console (the operator can see it) and return 500 rather than killing the host -- one bad request must not take down an all-day session.


## Embedded HTML / JavaScript comments

- Top bar: host identity + live connection state + Back to Console
- Client-side view switching -- exactly one section visible at a time.
