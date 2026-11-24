#Requires -Version 5.1
<#
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
#>

#region Logic module import -------------------------------------------------------
# RESTORED 2026-07-24 -- this block existed in the pre-rewrite version of
# this file and was accidentally dropped when the file was rewritten for the
# single-process host. Nothing then imported M365Logic.psm1, so the very
# first logic call the host makes (Test-M365ExoConnection, from
# Wait-M365ExoConnection below) failed with "not recognized as a name of a
# cmdlet" -- and because an unrecognized command is non-terminating by
# default, $conn came back $null, $null.Connected read as falsey, and the
# host fell through to the connect prompt as though nothing were connected.
# Confusing symptom, trivial cause.
#
# This import is the ONLY place the logic module is loaded for the web side:
# the host dot-sources this file, and the view files are dot-sourced into
# that same process, so all of them see the module's functions from here.
# -Force so .psm1 edits are picked up without restarting the process
# (a second Import-Module without it is a no-op). The .psd1 manifest is
# preferred over the raw .psm1 so its explicit FunctionsToExport applies.
$script:M365LogicManifest = Join-Path $PSScriptRoot 'M365Logic.psd1'
if (Test-Path -LiteralPath $script:M365LogicManifest) {
    Import-Module $script:M365LogicManifest -Force -ErrorAction Stop
} else {
    $script:M365LogicModule = Join-Path $PSScriptRoot 'M365Logic.psm1'
    if (-not (Test-Path -LiteralPath $script:M365LogicModule)) {
        throw "M365Logic.psd1/.psm1 not found in '$PSScriptRoot'. The web host needs the logic module alongside this file -- copy M365Logic.psm1 and M365Logic.psd1 into the same folder."
    }
    Import-Module $script:M365LogicModule -Force -ErrorAction Stop
}

# Post-import sanity check. Added alongside the restore above: the bug that
# prompted it was silent precisely because a MISSING function is a
# non-terminating error in PowerShell, so the host limped onward and reported
# a wrong diagnosis ("not connected") instead of the real one. Verifying a
# few representative exports here means any future breakage of the module
# path, manifest, or FunctionsToExport list fails immediately, at startup,
# with a message that names the actual problem.
$requiredLogic = @(
    'Test-M365ExoConnection'
    'Get-TargetSmtp'
    'Format-ErrorDetail'
    'Get-MailboxHealthSnapshot'
    'Invoke-MessageTrace'
    'Get-MailboxState'
    'Disable-M365User'
)
$missingLogic = @($requiredLogic | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
if ($missingLogic.Count -gt 0) {
    throw ("M365Logic imported but these expected functions are not available: {0}. Check M365Logic.psd1's FunctionsToExport matches the functions defined in M365Logic.psm1." -f ($missingLogic -join ', '))
}
#endregion

#region View registry ------------------------------------------------------------
# Each M365_View_*.ps1 file appends one entry here when dot-sourced by the
# host. Order of dot-sourcing determines order on the home page.
#
# Entry shape:
#   Id          [string]      -- short slug, used as the JS view id. Unique.
#   Name        [string]      -- card title / nav label.
#   Description [string]      -- one line on the home card.
#   Destructive [bool]        -- tints the card so risky tools read as risky.
#   GetHtml     [scriptblock] -- returns this view's HTML fragment.
#   Routes      [hashtable]   -- "METHOD /path" -> scriptblock($Request).
$script:M365Views = @()

function Register-M365View {
    <#
    .SYNOPSIS
        Adds one view to the registry. Called by each M365_View_*.ps1 file.
    #>
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$Name,
        [string]$Description = '',
        [switch]$Destructive,
        [Parameter(Mandatory)][scriptblock]$GetHtml,
        [hashtable]$Routes = @{}
    )
    if ($script:M365Views | Where-Object { $_.Id -eq $Id }) {
        throw "A view with Id '$Id' is already registered -- view Ids must be unique (they're used as HTML element ids)."
    }
    $script:M365Views += [PSCustomObject]@{
        Id          = $Id
        Name        = $Name
        Description = $Description
        Destructive = [bool]$Destructive
        GetHtml     = $GetHtml
        Routes      = $Routes
    }
}

function Get-M365Views {
    <#
    .SYNOPSIS
        Returns every registered view, in registration order.
    #>
    return $script:M365Views
}
#endregion

#region Shared CSS ----------------------------------------------------------------
function Get-M365SharedCss {
    <#
    .SYNOPSIS
        One stylesheet for the whole host page.
    .DESCRIPTION
        Emitted ONCE by the host, so view fragments never carry their own
        <style> block and can't disagree about what a class means. Note the
        history: .card and .status previously lived in individual tool files,
        which meant the console page used .card without defining it and
        rendered its whole tool list unstyled. Everything shared lives here
        now for exactly that reason.
    #>
    return @'
  body { font-family: "Segoe UI", -apple-system, sans-serif; margin:0; background:#f5f5f5; color:#222; }
  .wrap { max-width:1200px; margin:0 auto; padding:20px 24px 40px 24px; }
  h1 { font-size:19px; margin:0 0 4px 0; }
  h2 { font-size:17px; margin:0 0 10px 0; }
  h3 { font-size:15px; margin:0 0 4px 0; }

  /* Top bar: host identity + live connection state + Back to Console */
  .topbar { background:#1f2937; color:#f9fafb; padding:10px 24px; display:flex; align-items:center; gap:16px; flex-wrap:wrap; }
  .topbar .brand { font-weight:700; font-size:15px; }
  .topbar .spacer { flex:1; }
  .connpill { font-size:12px; font-weight:700; padding:4px 10px; border-radius:999px; background:#374151; color:#e5e7eb; }
  .connpill.ok { background:#065f46; color:#d1fae5; }
  .connpill.bad { background:#7f1d1d; color:#fee2e2; }

  .warn { background:#fdecea; color:#b3261e; border:1px solid #f5c2c0; padding:10px 14px; border-radius:6px; font-weight:600; margin-bottom:16px; }
  .note { background:#eff6ff; color:#1e40af; border:1px solid #bfdbfe; padding:10px 14px; border-radius:6px; margin-bottom:16px; font-size:13px; }

  .row { display:flex; gap:8px; margin-bottom:12px; align-items:flex-end; flex-wrap:wrap; }
  label { font-size:13px; color:#444; }
  input[type=text], input[type=date], select, textarea { padding:8px; font-family: Consolas, monospace; font-size:13px; border:1px solid #ccc; border-radius:4px; background:#fff; }
  textarea { width:100%; box-sizing:border-box; }
  button { padding:8px 16px; border:0; border-radius:4px; background:#2563eb; color:#fff; font-weight:600; cursor:pointer; font-size:13px; }
  button:hover:not(:disabled) { background:#1d4ed8; }
  button:disabled { background:#9ca3af; cursor:default; }
  button.secondary { background:#6b7280; }
  button.secondary:hover:not(:disabled) { background:#4b5563; }
  button.danger { background:#b3261e; }
  button.danger:hover:not(:disabled) { background:#991b1b; }

  pre, .out { background:#fff; border:1px solid #ddd; border-radius:6px; padding:14px; font-family: Consolas, monospace; font-size:13px; white-space:pre-wrap; overflow-x:auto; }
  .card { background:#fff; border:1px solid #ddd; border-radius:6px; padding:14px 18px; margin-bottom:10px; }
  .card.destructive { border-left:4px solid #b3261e; }
  .card .desc { color:#555; font-size:13px; margin-bottom:10px; }
  .status { font-size:13px; color:#444; margin:8px 0; }
  .status.err { color:#b3261e; }
  fieldset { border:1px solid #ccc; border-radius:6px; margin:12px 0; padding:12px 16px; }
  legend { font-size:13px; font-weight:700; color:#374151; }

  table { border-collapse:collapse; width:100%; background:#fff; }
  th, td { border:1px solid #ddd; padding:6px 8px; font-size:13px; text-align:left; }
  th { background:#eef1f5; cursor:pointer; user-select:none; }
  tbody tr:hover { background:#f3f4f6; }
  tbody tr.selected { background:#dbeafe; }

  /* Client-side view switching -- exactly one section visible at a time. */
  .view { display:none; }
  .view.active { display:block; }
'@
}
#endregion

#region Request authorization (CSRF defence) ---------------------------------------
# ADDED 2026-07-24, before the PowerShell console view was built -- but this
# protects EVERY view, and is worth having regardless.
#
# THE THREAT, concretely: this host listens on localhost and holds a live
# privileged Exchange session all day. Any web page the operator happens to
# have open in the same browser can issue requests to http://localhost:<port>/
# -- the browser will happily send them. Without a check, a malicious or
# merely compromised page could quietly POST to /api/disable/run or
# /api/perms/revoke-full and act with full admin authority. Loopback binding
# does NOT prevent this; it only stops other machines, not other tabs.
#
# TWO INDEPENDENT DEFENCES, both cheap:
#   1. Origin/Referer check -- browsers set these on cross-site requests and
#      forbid pages from forging them. Anything claiming an origin other than
#      this host's own is rejected. This alone blocks the classic
#      "malicious site posts to localhost" attack.
#   2. Per-run secret token -- generated fresh each time the host starts,
#      embedded in the page it serves, and required on every API call as the
#      X-M365-Token header. A page that never loaded from this host cannot
#      know it. This also covers requests that carry no Origin at all.
#
# Deliberately NOT a login: there is one operator, on their own machine, and
# a password prompt would add friction without adding protection against the
# threat that actually exists here. What this stops is OTHER SOFTWARE on the
# machine driving the admin session through the browser.
#
# GET / is always allowed (it's how the page bootstraps and receives the
# token). Everything under /api/ is gated.
$script:M365RequestToken = $null

function New-M365RequestToken {
    <#
    .SYNOPSIS
        Generates and stores this run's API token. Called once by the host.
    .DESCRIPTION
        Cryptographically random via RNGCryptoServiceProvider rather than
        Get-Random -- Get-Random is not a cryptographic RNG and its output
        would be guessable, which for a security token defeats the purpose.
        RNGCryptoServiceProvider exists in .NET Framework and .NET Core, so
        this works in both Windows PowerShell 5.1 and PowerShell 7.
    #>
    $bytes = New-Object 'byte[]' 32
    $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::Create()
    try {
        $rng.GetBytes($bytes)
    } finally {
        if ($rng -is [System.IDisposable]) { $rng.Dispose() }
    }
    $script:M365RequestToken = ([System.BitConverter]::ToString($bytes) -replace '-', '').ToLowerInvariant()
    return $script:M365RequestToken
}

function Get-M365RequestToken {
    return $script:M365RequestToken
}

function Test-M365RequestAuthorized {
    <#
    .SYNOPSIS
        Returns $true if a request may proceed; $false to reject with 403.
    .DESCRIPTION
        See the region comment for the reasoning. Applies to /api/ paths
        only; the host lets GET / through before calling this.
    #>
    param(
        [Parameter(Mandatory)]$Request,
        [Parameter(Mandatory)][int]$Port
    )

    # --- Defence 1: Origin / Referer must be this host, when present.
    $expected = @("http://localhost:$Port", "http://127.0.0.1:$Port")
    $origin = $Request.Headers['Origin']
    if ($origin) {
        if ($expected -notcontains $origin.TrimEnd('/')) { return $false }
    } else {
        $referer = $Request.Headers['Referer']
        if ($referer) {
            $ok = $false
            foreach ($e in $expected) { if ($referer.StartsWith($e)) { $ok = $true; break } }
            if (-not $ok) { return $false }
        }
        # Neither header present: not conclusive on its own (some legitimate
        # same-origin GETs omit both), so fall through to the token check,
        # which is the stronger of the two defences anyway.
    }

    # --- Defence 2: the per-run token.
    $token = $Request.Headers['X-M365-Token']
    if (-not $token) { return $false }
    if ($token -ne $script:M365RequestToken) { return $false }

    return $true
}
#endregion

#region Connection gate -----------------------------------------------------------
function Wait-M365ExoConnection {
    <#
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
    #>
    param([string]$AlsoNote = '')

    $conn = Test-M365ExoConnection
    if ($conn.Connected) {
        Write-Host "Exchange Online: $($conn.Detail)" -ForegroundColor Green
        return
    }

    Write-Host ""
    Write-Host "Exchange Online is not usable in this session yet." -ForegroundColor Yellow
    Write-Host "  $($conn.Detail)" -ForegroundColor Yellow
    Write-Host "Run Connect-ExchangeOnline below (sign in when prompted)$AlsoNote, then type 'exit' to continue starting the host." -ForegroundColor Yellow
    Write-Host ""
    try {
        $host.EnterNestedPrompt()
    } catch {
        Write-Host "This host doesn't support pausing here ($($_.Exception.Message)) -- connect first, then re-run this script." -ForegroundColor Yellow
        return
    }

    # The operator may have connected to a different tenant than a previous
    # run -- drop the cached default domain so bare names resolve against
    # whatever is connected now.
    Clear-M365DefaultDomainCache

    $conn2 = Test-M365ExoConnection
    if ($conn2.Connected) {
        Write-Host "Connected. Starting host..." -ForegroundColor Green
    } else {
        Write-Host "Still not connected -- starting anyway. The page will show a red connection banner, and lookups will fail with a clear error until you connect." -ForegroundColor Yellow
    }
}
#endregion

#region The web host --------------------------------------------------------------
function Start-M365WebHost {
    <#
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
    #>
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][int]$Port,
        [Parameter(Mandatory)][scriptblock]$GetHtml,
        [hashtable]$Routes = @{},
        [switch]$NoBrowser,
        # Passed through to Get-M365ConnectionSummary. Off by default -- see
        # that function's -IncludeDirectory help for why the directory pill is
        # no longer monitored unless asked for.
        [switch]$ShowDirectoryPill
    )

    # Fresh token per run -- see the request-authorization region. Generated
    # BEFORE the page is ever served, since the page embeds it.
    [void](New-M365RequestToken)

    $prefix = "http://localhost:$Port/"
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add($prefix)
    try {
        $listener.Start()
    } catch {
        Write-Host "Could not start the web host on $prefix -- $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Two common causes:" -ForegroundColor Yellow
        Write-Host "  1. A previous run is still holding the port. Stop it, or pass a different -Port." -ForegroundColor Yellow
        # On Windows, HttpListener sits on http.sys, which requires a URL
        # reservation for a prefix unless the process is elevated. Localhost
        # prefixes often work unelevated, but not always -- so name the exact
        # fix rather than leaving an access-denied to interpret.
        Write-Host "  2. Access denied: Windows needs a one-time URL reservation for this prefix." -ForegroundColor Yellow
        Write-Host "     Run ONCE in an elevated prompt (then re-run this normally):" -ForegroundColor Yellow
        Write-Host "       netsh http add urlacl url=$prefix user=$env:USERDOMAIN\$env:USERNAME" -ForegroundColor Cyan
        return
    }

    Write-Host ""
    Write-Host "$HostName running at $prefix" -ForegroundColor Green
    Write-Host "All tools are served from this one process, sharing this one Exchange connection." -ForegroundColor Gray
    Write-Host "Press Ctrl+C here, or use Shut Down on the page, to stop." -ForegroundColor Yellow

    if (-not $NoBrowser) {
        try {
            # REVISED 2026-07-24 -- version-aware, because $IsWindows /
            # $IsMacOS / $IsLinux are PowerShell 7 AUTOMATIC VARIABLES that
            # simply do not exist in Windows PowerShell 5.1. In 5.1 all three
            # evaluate to $null, so the old if/elseif chain fell through
            # silently and no browser ever opened -- a bug that only appeared
            # once this host started targeting 5.1 (which it must, so the
            # AzureAD and SPO modules can load; see Get-M365ConnectionSummary
            # in M365Logic for the full reasoning).
            #
            # 5.1 only ever runs on Windows, so "PSVersion major < 6" is a
            # sound Windows test. Under PS7 the real variables are used.
            $isPs5   = ($PSVersionTable.PSVersion.Major -lt 6)
            $onWin   = $isPs5 -or ($IsWindows -eq $true)
            $onMac   = (-not $isPs5) -and ($IsMacOS -eq $true)
            $onLinux = (-not $isPs5) -and ($IsLinux -eq $true)

            if     ($onWin)   { Start-Process $prefix }
            elseif ($onMac)   { Start-Process 'open' -ArgumentList $prefix }
            elseif ($onLinux) { Start-Process 'xdg-open' -ArgumentList $prefix }
            else {
                Write-Host "Could not determine the platform to open a browser -- open $prefix manually." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Could not auto-open a browser -- open $prefix manually. ($($_.Exception.Message))" -ForegroundColor Yellow
        }
    }

    # Ctrl+C handling: GetContext() is a blocking synchronous call and
    # PowerShell's default Ctrl+C does not reliably interrupt one. Forcibly
    # Stop()ing the listener from this handler makes the blocked call throw,
    # which the loop treats as a clean shutdown. $e.Cancel = $true keeps
    # PowerShell from hard-killing the process so the finally block still
    # runs.
    $script:M365HostStopRequested = $false
    $cancelHandler = {
        param($sender, $e)
        $e.Cancel = $true
        Write-Host "`nCtrl+C received -- stopping host..." -ForegroundColor Yellow
        $script:M365HostStopRequested = $true
        try { $listener.Stop() } catch { }
    }
    [Console]::add_CancelKeyPress($cancelHandler)

    try {
        while ($listener.IsListening -and -not $script:M365HostStopRequested) {
            try {
                $context = $listener.GetContext()
            } catch [System.Net.HttpListenerException] {
                break
            } catch [System.ObjectDisposedException] {
                break
            }
            $request  = $context.Request
            $response = $context.Response
            try {
                $routeKey = "$($request.HttpMethod) $($request.Url.AbsolutePath)"

                if ($request.HttpMethod -eq 'GET' -and $request.Url.AbsolutePath -eq '/') {
                    # Always allowed -- this is how the page bootstraps and
                    # receives its token. See the request-authorization region.
                    $html  = & $GetHtml
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
                    $response.ContentType = 'text/html; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                elseif ($request.Url.AbsolutePath -like '/api/*' -and -not (Test-M365RequestAuthorized -Request $request -Port $Port)) {
                    # Rejected before ANY handler runs, so an unauthorized
                    # request never reaches a cmdlet.
                    Write-Host "Rejected unauthorized $routeKey (bad/missing token or foreign origin)." -ForegroundColor Red
                    $response.StatusCode = 403
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes('{"ok":false,"error":"Request rejected: missing or invalid API token, or the request did not originate from this host''s own page. Reload the page in your browser."}')
                    $response.ContentType = 'application/json; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                elseif ($routeKey -eq 'GET /api/connection') {
                    # ?full=1 opts into the probes that cost a service call
                    # (SharePoint especially) -- the page polls without it and
                    # passes it on load / explicit Refresh.
                    $full = ($request.QueryString['full'] -eq '1')
                    $sum  = Get-M365ConnectionSummary -IncludeExpensive:$full -IncludeDirectory:$ShowDirectoryPill
                    $payload = [ordered]@{
                        checkedExpensive = $full
                        # Lets the page hide the directory pill entirely rather
                        # than render a meaningless state for something nobody
                        # is monitoring.
                        directoryMonitored = [bool]$ShowDirectoryPill
                        # WHICH TENANT is connected -- added 2026-07-24 at the
                        # operator's request, and it's the single most useful
                        # thing on this bar. "Connected" alone doesn't tell you
                        # connected to WHAT, which matters a lot when you work
                        # several tenants in a day.
                        #
                        # Two sources, most recognisable first:
                        #   tenantDomain -- the default ACCEPTED domain
                        #     (expresspros.com). This is what an admin actually
                        #     thinks of as "the tenant". Cached after the first
                        #     lookup, so the cheap 60s poll costs nothing.
                        #   organization -- what Get-ConnectionInformation
                        #     reports, usually the *.onmicrosoft.com name. Used
                        #     as a fallback when the accepted-domain lookup
                        #     hasn't succeeded yet.
                        tenantDomain = "$(Get-M365DefaultDomain)"
                        # Tenant guard -- see Test-M365ExpectedTenant. Only
                        # meaningful once something is connected, which is why
                        # it rides along on this route rather than being
                        # checked once at startup.
                        tenant = $(
                            $tg = Test-M365ExpectedTenant
                            [ordered]@{
                                armed    = [bool]$tg.Armed
                                match    = [bool]$tg.Match
                                expected = "$($tg.Expected)"
                                actual   = "$($tg.Actual)"
                                detail   = "$($tg.Detail)"
                            }
                        )
                        exo = [ordered]@{
                            connected    = [bool]$sum.Exo.Connected
                            detail       = "$($sum.Exo.Detail)"
                            organization = "$($sum.Exo.Organization)"
                            skipped      = $false
                        }
                        spo = [ordered]@{
                            connected = [bool]$sum.Spo.Connected
                            detail    = "$($sum.Spo.Detail)"
                            skipped   = [bool]($sum.Spo.PSObject.Properties['Skipped'] -and $sum.Spo.Skipped)
                        }
                        aad = [ordered]@{
                            connected = [bool]$sum.Aad.Connected
                            detail    = "$($sum.Aad.Detail)"
                            provider  = "$($sum.Aad.Provider)"
                            account   = "$($sum.Aad.Account)"
                            skipped   = [bool]($sum.Aad.PSObject.Properties['Skipped'] -and $sum.Aad.Skipped)
                        }
                    }
                    $json  = $payload | ConvertTo-Json -Depth 5 -Compress
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $response.ContentType = 'application/json; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                elseif ($routeKey -eq 'POST /api/shutdown' -or $routeKey -eq 'GET /api/shutdown') {
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes('{"ok":true}')
                    $response.ContentType = 'application/json; charset=utf-8'
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                    $script:M365HostStopRequested = $true
                }
                elseif ($Routes.ContainsKey($routeKey)) {
                    $result = & $Routes[$routeKey] $request
                    $bytes  = [System.Text.Encoding]::UTF8.GetBytes($result.Body)
                    $response.ContentType = $result.ContentType
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
                else {
                    $response.StatusCode = 404
                    $bytes = [System.Text.Encoding]::UTF8.GetBytes("No route for $routeKey")
                    $response.ContentLength64 = $bytes.Length
                    $response.OutputStream.Write($bytes, 0, $bytes.Length)
                }
            } catch {
                # A handler blew up. Log it to the host console (the operator
                # can see it) and return 500 rather than killing the host --
                # one bad request must not take down an all-day session.
                Write-Host "Handler error on $($request.HttpMethod) $($request.Url.AbsolutePath): $($_.Exception.Message)" -ForegroundColor Red
                try {
                    $response.StatusCode = 500
                    $errBytes = [System.Text.Encoding]::UTF8.GetBytes("Server error: $($_.Exception.Message)")
                    $response.OutputStream.Write($errBytes, 0, $errBytes.Length)
                } catch { }
            } finally {
                $response.OutputStream.Close()
            }
        }
    } finally {
        [Console]::remove_CancelKeyPress($cancelHandler)
        try { $listener.Stop() } catch { }
        try { $listener.Close() } catch { }
        Write-Host "$HostName stopped." -ForegroundColor Yellow
    }
}
#endregion
