#Requires -Version 7.0
<#
.SYNOPSIS
    Headless, browser-based counterpart to M365_MailboxHealth.ps1 -- runs a
    small local web server instead of a WinForms window, so it works under
    PowerShell 7 on Windows, macOS, or Linux, with no WinForms dependency
    anywhere in this file.
.DESCRIPTION
    ADDED 2026-07-24 -- this is the proof-of-concept for the "could the
    cloud-only tools run on a Mac" question: WinForms itself is Windows-only
    regardless of PowerShell version, so instead of a Form this script opens
    a tiny local HTTP listener (System.Net.HttpListener -- a managed,
    genuinely cross-platform implementation under .NET Core/PS7, not the
    old Windows-only http.sys-backed one) and serves a single HTML page.
    That page's JavaScript calls back into this same PowerShell process over
    a local API for the actual mailbox lookup. The browser is just a
    renderer; every EXO call still happens in PowerShell, in this process,
    same as every other tool in this suite.

    #Requires -Version 7.0, not 5.1 like every other script in this family --
    the auto-detect browser launch below uses $IsWindows / $IsMacOS /
    $IsLinux, PowerShell 7 automatic variables that don't exist in Windows
    PowerShell 5.1. This script has never been run against a live tenant or
    tested on an actual Mac -- only syntax-checked. The HttpListener and
    browser-launch mechanics specifically are unproven; if either misbehaves
    on your actual machine, that's the part to debug first.

    WHAT'S IDENTICAL TO M365_MailboxHealth.ps1, copied over unchanged:
      - Get-MailboxHealthSnapshot -- byte-for-byte the same function, same
        EXO calls (Get-MailboxStatistics, Get-Mailbox, Get-MailboxPermission,
        Get-RecipientPermission, Get-InboxRule via Get-InboxRulesForMailbox,
        Get-CASMailbox), same interpretation/verification notes in its
        header comment.
      - The report text itself. Update-HealthFromCurrent's formatting logic
        (the WinForms click handler that built up $txtHealth.Text) is
        refactored here into a standalone Format-MailboxHealthReport
        function that takes the same snapshot object and returns the same
        string, so a mailbox looked up through this web page and the same
        mailbox looked up through the WinForms tool read identically. Only
        the delivery mechanism changed (an HTML <pre> block instead of a
        TextBox), not the content or the logic that builds it.
      - The shared logic (Invoke-EXOCommand / Get-EXOCommand collision
        guard, Format-ErrorDetail, Get-InboxRulesForMailbox,
        Get-M365DefaultDomain / Get-TargetSmtp) is used exactly as it
        already existed. REVISED 2026-07-24: it now comes from
        M365Logic.psm1, imported by M365WebCommon.ps1, rather than from
        dot-sourcing M365Common.ps1 -- see below.

    WHAT'S DELIBERATELY DIFFERENT:
      - This file no longer dot-sources M365Common.ps1 at all (REVISED
        2026-07-24). That file has become the WinForms UI layer -- its only
        remaining contents are Show-TargetPromptDialog / Resolve-ToolTarget,
        which build a System.Windows.Forms dialog. A web page has no use for
        a modal desktop dialog (the mailbox address comes from the ?smtp=
        query string), and a PS7 browser host on a non-Windows machine
        cannot load WinForms at all, so importing it here would be actively
        wrong rather than merely unnecessary. Get-TargetSmtp -- the pure
        string-parsing half, no WinForms -- moved into M365Logic.psm1 and is
        still used directly, so "bob.smith" with no @ still resolves against
        the connected tenant's default domain exactly as it does in the
        WinForms tool.
      - No -SmtpAddress launch parameter, since there's no console launching
        this the way AdminConsole.ps1 launches its tools -- -Port and
        -NoBrowser instead (see below).
      - Runs until Ctrl+C, like any local dev server, instead of opening and
        closing a modal dialog. -Port <int> (default 8787) picks the local
        port; -NoBrowser skips the automatic browser launch if you'd rather
        open the URL yourself.
      - Binds to "http://localhost:<port>/" specifically, not a wildcard or
        0.0.0.0 -- loopback-only on purpose, so nothing else on the network
        can reach an endpoint that runs EXO queries under whatever
        credentials this PowerShell session is connected with. There is no
        authentication on top of that; anything else running as the same OS
        user on the same machine could still hit it while it's running.
      - Connect-ExchangeOnline still has to happen yourself, in the SAME
        PowerShell 7 process before running this script -- same "connect
        before running" convention as every other tool in this suite, just
        without a WinForms warning label to put it on. The console this
        script prints to is where that warning lives instead.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24 (headless/browser port of M365_MailboxHealth.ps1)
    Revised : 2026-07-24, later -- retrofitted onto M365WebCommon.ps1's
              shared Start-M365WebTool once four more tools were about to
              repeat this file's own listener boilerplate. Port (8787) is
              unchanged and must keep matching its entry in
              M365WebCommon.ps1's $script:M365ToolRegistry. Added the
              "Close & Return to Console" button (Get-M365CloseButtonHtml)
              and a shared <style> block (Get-M365SharedCss) -- no change to
              Get-MailboxHealthSnapshot or Format-MailboxHealthReport.
#>
param(
    [int]$Port = 8787,
    [switch]$NoBrowser
)

. (Join-Path $PSScriptRoot 'M365WebCommon.ps1')

# REVISED 2026-07-24 -- the logic that used to live here MOVED, unchanged,
# into M365Logic.psm1, which M365WebCommon.ps1 imports (so it is already
# available below with nothing further to import here). This file is now a
# pure browser UI layer: build the page, call module functions, render what
# comes back. The WinForms tools call the exact same functions -- one
# implementation, two front ends, no duplicated logic to drift apart.
# Moved out of this file: Get-MailboxHealthSnapshot, Format-MailboxHealthReport.

#region The single HTML page -- inline CSS/JS, no external assets, no CDN calls (runs fully offline/local)
function Get-IndexHtml {
    $css = Get-M365SharedCss
    $closeBtn = Get-M365CloseButtonHtml
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>M365 Mailbox Health</title>
<style>
$css
  pre { min-height:420px; }
</style>
</head>
<body>
  <h1>M365 Mailbox Health</h1>
  <div class="warn">Connect Exchange Online (Connect-ExchangeOnline) in the PowerShell window running this server before looking anything up.</div>
  <div class="row">
    <input id="smtp" type="text" placeholder="mailbox email address" autofocus>
    <button id="lookup">Look Up</button>
    <button id="refresh" class="secondary" disabled>Refresh</button>
    <button id="copy" class="secondary" disabled>Copy</button>
    $closeBtn
  </div>
  <pre id="out">Enter a mailbox address above and click Look Up.</pre>
<script>
  const smtpEl = document.getElementById('smtp');
  const outEl = document.getElementById('out');
  const lookupBtn = document.getElementById('lookup');
  const refreshBtn = document.getElementById('refresh');
  const copyBtn = document.getElementById('copy');
  let currentSmtp = null;

  async function runLookup(smtp) {
    lookupBtn.disabled = true;
    outEl.textContent = 'Loading mailbox health...';
    try {
      const res = await fetch('/api/health?smtp=' + encodeURIComponent(smtp));
      const data = await res.json();
      if (data.ok) {
        outEl.textContent = data.text;
        currentSmtp = smtp;
        refreshBtn.disabled = false;
        copyBtn.disabled = false;
      } else {
        outEl.textContent = 'ERROR: ' + data.error;
      }
    } catch (err) {
      outEl.textContent = 'Request failed: ' + err;
    } finally {
      lookupBtn.disabled = false;
    }
  }

  lookupBtn.addEventListener('click', () => {
    const smtp = smtpEl.value.trim();
    if (!smtp) { outEl.textContent = 'Enter a mailbox address first.'; return; }
    runLookup(smtp);
  });
  smtpEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') lookupBtn.click(); });
  refreshBtn.addEventListener('click', () => { if (currentSmtp) runLookup(currentSmtp); });
  copyBtn.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(outEl.textContent);
      copyBtn.textContent = 'Copied!';
      setTimeout(() => { copyBtn.textContent = 'Copy'; }, 1200);
    } catch (err) {
      alert('Copy failed: ' + err);
    }
  });
</script>
</body>
</html>
"@
}
#endregion

#region Routes + server startup -- delegates all listener plumbing to M365WebCommon.ps1
$routes = @{
    'GET /api/health' = {
        param($Request)
        $smtpParam = $Request.QueryString['smtp']
        $result = [ordered]@{ ok = $false; text = ''; error = '' }
        if (-not $smtpParam) {
            $result.error = 'Missing ?smtp= parameter.'
        } else {
            $smtp = Get-TargetSmtp -Text $smtpParam
            if (-not $smtp) {
                $result.error = "Could not determine a domain for '$smtpParam' -- type the full email address, or connect Exchange Online first."
            } else {
                try {
                    $h = Get-MailboxHealthSnapshot -Smtp $smtp
                    $result.ok   = $true
                    $result.text = Format-MailboxHealthReport -Smtp $smtp -H $h
                } catch {
                    $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Health (web)'
                }
            }
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
    }
}

Wait-M365ExoConnection
Start-M365WebTool -ToolName 'M365 Mailbox Health' -Port $Port -GetHtml ${function:Get-IndexHtml} -Routes $routes -NoBrowser:$NoBrowser
#endregion
