#Requires -Version 5.1
<#
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
#>

Register-M365View -Id 'health' -Name 'Mailbox Health' `
    -Description 'Read-only composite snapshot: quotas, holds, forwarding, delegates, inbox rules, protocol access.' `
    -GetHtml {
        return @'
  <h2>Mailbox Health</h2>
  <div class="note">Read-only. Nothing is modified by this view.</div>

  <div class="row">
    <label>Mailbox:<br><input id="hz-smtp" type="text" placeholder="user@domain, or just the alias" style="width:300px"></label>
    <button id="hz-lookup">Look Up</button>
    <button id="hz-refresh" class="secondary" disabled>Refresh</button>
    <button id="hz-copy" class="secondary" disabled>Copy</button>
  </div>

  <div id="hz-out" class="out" style="min-height:380px;">Enter a mailbox above and click Look Up.</div>

<script>
(function () {
  const smtpEl    = document.getElementById('hz-smtp');
  const outEl     = document.getElementById('hz-out');
  const lookupBtn = document.getElementById('hz-lookup');
  const refreshBtn= document.getElementById('hz-refresh');
  const copyBtn   = document.getElementById('hz-copy');
  let current = null;

  async function run(smtp) {
    lookupBtn.disabled = true;
    refreshBtn.disabled = true;
    outEl.textContent = 'Loading mailbox health...';
    try {
      const res  = await fetch('/api/health?smtp=' + encodeURIComponent(smtp));
      const data = await res.json();
      if (data.ok) {
        outEl.textContent = data.text;
        current = smtp;
        copyBtn.disabled = false;
      } else {
        outEl.textContent = 'ERROR: ' + data.error;
      }
    } catch (err) {
      outEl.textContent = 'Request failed: ' + err;
    } finally {
      lookupBtn.disabled = false;
      refreshBtn.disabled = (current === null);
    }
  }

  lookupBtn.addEventListener('click', () => {
    const smtp = smtpEl.value.trim();
    if (!smtp) { outEl.textContent = 'Enter a mailbox address first.'; return; }
    run(smtp);
  });
  smtpEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') lookupBtn.click(); });
  refreshBtn.addEventListener('click', () => { if (current) run(current); });
  copyBtn.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(outEl.textContent);
      copyBtn.textContent = 'Copied!';
      setTimeout(() => { copyBtn.textContent = 'Copy'; }, 1200);
    } catch (err) { alert('Copy failed: ' + err); }
  });
})();
</script>
'@
    } `
    -Routes @{
        'GET /api/health' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            $smtpParam = $Request.QueryString['smtp']
            if (-not $smtpParam) {
                $result.error = 'Missing ?smtp= parameter.'
            } else {
                $smtp = Get-TargetSmtp -Text $smtpParam
                if (-not $smtp) {
                    $result.error = "Could not determine a domain for '$smtpParam' -- type the full email address, or make sure Exchange Online is connected."
                } else {
                    try {
                        $h = Get-MailboxHealthSnapshot -Smtp $smtp
                        $result.ok   = $true
                        $result.text = Format-MailboxHealthReport -Smtp $smtp -H $h
                    } catch {
                        $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Health'
                    }
                }
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
    }
