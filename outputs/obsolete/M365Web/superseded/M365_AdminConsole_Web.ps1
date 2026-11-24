#Requires -Version 7.0
<#
.SYNOPSIS
    M365 Admin Console (web) -- a launcher page, not a unified app. Lists
    every tool in M365WebCommon.ps1's registry; clicking one spawns that
    tool as its own independent PowerShell process on its own port and
    opens it in a new browser tab.
.DESCRIPTION
    ADDED 2026-07-24 -- built after the operator picked "console as
    launcher" over a single unified app or a standalone-but-mountable
    hybrid, explicitly to avoid repeating how large the original WinForms
    AdminConsole.ps1 had already gotten ("i would still think we'd want it
    as a launcher - so it doesn't get unwieldy - that was starting to
    happen before already"). Each tool -- M365_MailboxHealth_Web.ps1,
    M365_MessageTrace_Web.ps1, M365_MailboxPermissions_Web.ps1,
    M365_Disable_Web.ps1 -- stays a fully independent script with its own
    listener on its own fixed port (see M365WebCommon.ps1's
    $script:M365ToolRegistry), runnable directly with no console involved
    at all, same "fully functional on its own" property every WinForms tool
    in the original suite already had.

    IMPORTANT OPERATIONAL NOTE -- READ BEFORE RELYING ON THE LAUNCH BUTTON:
    clicking Launch spawns a BRAND NEW, NOT-YET-CONNECTED PowerShell
    process for that tool (Start-Process 'pwsh' -File <tool script> -Port
    <port> -NoBrowser, then this console opens http://localhost:<port>/ in
    a new tab once that process's listener responds). Connect-ExchangeOnline
    is per-process state -- being connected in THIS console's own process,
    or in any tool you've already launched, does NOT carry over to a
    freshly spawned one.

    REVISED 2026-07-24, confirmed against a live run (a launched tool
    failed with "no command named Get-Mailbox exists in this session" even
    though the operator HAD connected -- just in a different process):
    every M365_*_Web.ps1 tool that touches EXO now calls
    M365WebCommon.ps1's Wait-M365ExoConnection right before starting its
    listener. If that spawned process isn't already connected, it pauses
    with a real interactive sub-prompt in that exact window
    ($host.EnterNestedPrompt()) so you can type Connect-ExchangeOnline (and
    Connect-SPOService, for the Disable tool's OneDrive step) right there,
    then 'exit' to let the tool continue starting. This only works if
    Start-Process actually gives that spawned process a visible, interactive
    window to pause in -- confirmed true on Windows (a normal console
    window opens). NOT verified on macOS/Linux, where Start-Process for a
    bare CLI executable often does not allocate an interactive terminal at
    all -- if a launched tool never shows a connect prompt and fails
    "not connected" anyway on one of those platforms, the reliable fallback
    is: open your own terminal, cd to this folder, run
    Connect-ExchangeOnline (and Connect-SPOService if needed) yourself,
    then run that tool script directly (e.g. ./M365_MailboxHealth_Web.ps1)
    instead of using this console's Launch button for that session.
    This console does not, and cannot safely, automate
    Connect-ExchangeOnline itself -- modern EXO auth is normally an
    interactive, MFA-capable browser sign-in, not something to script
    invisibly, and entering credentials on the operator's behalf is outside
    what this assistant will do regardless.

    "Closing a script and going back to the primary admin page" (the
    operator's own description of the desired behavior): since every tool
    is its own browser tab, "closing" is closing that tab -- each tool page
    includes M365WebCommon.ps1's "Close & Return to Console" button, which
    stops that tool's own listener/process (POST /api/shutdown) and then
    attempts window.close() (works because the console opened it via
    window.open(), which is what allows a script to close a tab in most
    browsers). If the browser blocks the close, the page says so plainly
    rather than leaving the operator guessing -- the Console's own tab, if
    still open, is what you'd switch back to either way.

    This console has no lookup/target flow of its own (no equivalent of the
    original ANZAdminConsole.ps1's "enter a mailbox, detect its type, grey
    out tools that don't apply" step) -- with tools running as independent
    processes rather than one shared session, there's no single "last
    looked up" state to grey buttons against the way the WinForms consoles
    had. Every tool is just a link; each tool's own page still has its own
    mailbox/DL input field, same as running it standalone.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24
#>
param(
    [int]$Port = 8786,
    [switch]$NoBrowser
)

. (Join-Path $PSScriptRoot 'M365WebCommon.ps1')

#region Child process management
# Quick, short-timeout check for "is something already listening on this
# port" -- used both to avoid spawning a duplicate process for a tool
# that's already running, and to poll for readiness right after spawning
# one.
function Test-M365PortOpen {
    param([int]$Port)
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect('localhost', $Port, $null, $null)
        $connected = $iar.AsyncWaitHandle.WaitOne(300)
        if ($connected -and $client.Connected) {
            $client.EndConnect($iar)
            $client.Close()
            return $true
        }
        $client.Close()
        return $false
    } catch {
        return $false
    }
}

# Spawns the tool's script as an independent pwsh process if its port isn't
# already answering, then polls (up to ~5 seconds) for it to come up before
# handing back its URL. See header CAVEAT -- this does NOT connect Exchange
# Online/SharePoint Online for you.
function Start-M365ChildTool {
    param($Tool)
    $scriptPath = Join-Path $PSScriptRoot $Tool.ScriptFile
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        return [ordered]@{ ok = $false; url = $null; error = "'$($Tool.Name)' points to '$($Tool.ScriptFile)', which does not exist in this folder." }
    }
    if (-not (Test-M365PortOpen -Port $Tool.Port)) {
        try {
            Start-Process -FilePath 'pwsh' -ArgumentList @('-NoLogo', '-File', $scriptPath, '-Port', $Tool.Port, '-NoBrowser') -ErrorAction Stop | Out-Null
        } catch {
            return [ordered]@{ ok = $false; url = $null; error = "Could not launch '$($Tool.Name)': $($_.Exception.Message) -- is 'pwsh' on PATH?" }
        }
        $tries = 0
        while (-not (Test-M365PortOpen -Port $Tool.Port) -and $tries -lt 20) {
            Start-Sleep -Milliseconds 250
            $tries++
        }
        if (-not (Test-M365PortOpen -Port $Tool.Port)) {
            return [ordered]@{ ok = $false; url = $null; error = "'$($Tool.Name)' was launched but did not respond on port $($Tool.Port) within 5 seconds. Check its window for an error." }
        }
    }
    return [ordered]@{ ok = $true; url = "http://localhost:$($Tool.Port)/"; error = $null }
}
#endregion

#region HTML page
function Get-IndexHtml {
    $css = Get-M365SharedCss
    $cards = ($script:M365ToolRegistry | ForEach-Object {
        @"
    <div class="card">
      <h3 style="margin:0 0 4px 0;">$($_.Name)</h3>
      <div style="color:#555;font-size:13px;margin-bottom:8px;">$($_.Description)</div>
      <button class="launchBtn" data-script="$($_.ScriptFile)" data-name="$($_.Name)">Launch</button>
      <span class="launchStatus" style="margin-left:8px;font-size:13px;color:#666;"></span>
    </div>
"@
    }) -join "`n"
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>M365 Admin Console</title>
<style>
$css
</style>
</head>
<body>
  <h1>M365 Admin Console</h1>
  <div class="warn">Each Launch below opens its tool as its own PowerShell process on its own port. Exchange Online (and SharePoint Online, for Disable) must be connected in THAT tool's own process -- being connected here doesn't carry over. See this file's own header comment for platform-specific detail.</div>
  $cards
<script>
  document.querySelectorAll('.launchBtn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const statusEl = btn.nextElementSibling;
      btn.disabled = true;
      statusEl.textContent = 'Launching...';
      try {
        const res = await fetch('/api/launch', {
          method: 'POST', headers: {'Content-Type':'application/json'},
          body: JSON.stringify({ scriptFile: btn.dataset.script })
        });
        const data = await res.json();
        if (data.ok) {
          statusEl.textContent = 'Opened.';
          window.open(data.url, '_blank');
        } else {
          statusEl.textContent = 'ERROR: ' + data.error;
        }
      } catch (err) {
        statusEl.textContent = 'Request failed: ' + err;
      } finally {
        btn.disabled = false;
      }
    });
  });
</script>
</body>
</html>
"@
}
#endregion

#region Routes
$routes = @{
    'POST /api/launch' = {
        param($Request)
        $result = [ordered]@{ ok = $false; url = $null; error = '' }
        try {
            $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
            $body = $reader.ReadToEnd() | ConvertFrom-Json
            $tool = $script:M365ToolRegistry | Where-Object { $_.ScriptFile -eq $body.scriptFile } | Select-Object -First 1
            if (-not $tool) {
                $result.error = "'$($body.scriptFile)' is not in the tool registry."
            } else {
                $launch = Start-M365ChildTool -Tool $tool
                $result.ok    = $launch.ok
                $result.url   = $launch.url
                $result.error = $launch.error
            }
        } catch {
            $result.error = "Launch failed: $($_.Exception.Message)"
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
    }
}

Start-M365WebTool -ToolName 'M365 Admin Console' -Port $Port -GetHtml ${function:Get-IndexHtml} -Routes $routes -NoBrowser:$NoBrowser
#endregion
