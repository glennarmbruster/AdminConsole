#Requires -Version 5.1
<#
.SYNOPSIS
    PowerShell Console view for M365_Host_Web.ps1. OFF BY DEFAULT -- the host
    only loads this file when started with -EnableConsole.
.DESCRIPTION
    ADDED 2026-07-24. Browser equivalent of the WinForms PowerShellConsole.ps1.

    WHY THIS ONE IS OPT-IN AND THE OTHER SEVEN AREN'T -- read before enabling:

    Every other view exposes a FIXED set of operations. This one executes
    whatever PowerShell you type, inside the host process, which is holding a
    live privileged Exchange Online session (and possibly SPO and a directory
    module too). That is a categorically larger exposure than the rest of the
    toolkit, for two reasons:

      1. Scope. The other views can grant a permission or convert a mailbox.
         This can run anything the connected account can do -- across every
         connected service, plus the local filesystem, plus the network.
      2. Reachability. The host listens on localhost, and any web page open in
         the same browser can attempt requests to localhost. Loopback binding
         stops other MACHINES, not other TABS.

    Reason 2 is why M365WebCommon.ps1's request-authorization layer exists
    (per-run random token + Origin check, required on every /api/ call). That
    layer is what makes this view defensible rather than reckless -- a page that
    never loaded from this host cannot know the token. It is a genuine
    mitigation, not a fig leaf. But it is also the ONLY thing standing between
    a hostile page and arbitrary code execution as you, which is why enabling
    this is a deliberate switch rather than the default.

    NOT ENABLED unless the operator passes -EnableConsole to M365_Host_Web.ps1.
    If you don't need it, don't enable it -- the WinForms PowerShellConsole.ps1
    has none of this exposure because it isn't reachable over HTTP at all.

    HOW IT EXECUTES: the submitted text is run in the host's own runspace via
    the call operator on a scriptblock created from it, so it sees the host's
    live connections and imported modules -- which is the entire point (a
    separate runspace would have no Exchange session, making the view useless).
    Output is captured with Out-String so objects render the way they would in
    a console.

    NO SANDBOX, deliberately and explicitly: there is no allow-list, no
    language-mode restriction, and no attempt to block dangerous cmdlets. A
    partial restriction would give a false sense of safety while being trivial
    to bypass, so the honest design is an unrestricted console that is off by
    default and clearly labelled, rather than a "safe" one that isn't.

    Errors are captured and returned as text rather than thrown, so a typo
    doesn't take down an all-day host.
#>

Register-M365View -Id 'psconsole' -Name 'PowerShell Console' -Destructive `
    -Description 'Run arbitrary PowerShell in the host process, with its live connections. Unrestricted -- enabled only because you passed -EnableConsole.' `
    -GetHtml {
        return @'
  <h2>PowerShell Console</h2>
  <div class="warn">
    UNRESTRICTED. This runs in the host process with its live Exchange (and SPO / directory) connections.
    There is no sandbox, no allow-list, and no cmdlet filtering -- anything the connected account can do,
    this can do. Requests are protected by the host's per-run token and origin check, which is what makes
    this defensible, but treat this window as a fully privileged shell.
  </div>

  <div class="row">
    <label style="flex:1;">Command(s) -- Ctrl+Enter to run:<br>
      <textarea id="ps-cmd" rows="8" spellcheck="false"
        placeholder="Get-Mailbox -ResultSize 5 | Select DisplayName,PrimarySmtpAddress"></textarea>
    </label>
  </div>

  <div class="row">
    <button id="ps-run" class="danger">Run</button>
    <button id="ps-clear" class="secondary">Clear Output</button>
    <label style="margin-left:8px;"><input type="checkbox" id="ps-append" checked> Append output (keep history)</label>
    <span id="ps-status" class="status" style="margin:0;"></span>
  </div>

  <div id="ps-out" class="out" style="min-height:300px;font-size:12px;">(nothing run yet)</div>

<script>
(function () {
  const cmdEl    = document.getElementById('ps-cmd');
  const outEl    = document.getElementById('ps-out');
  const statusEl = document.getElementById('ps-status');
  const runBtn   = document.getElementById('ps-run');
  const appendEl = document.getElementById('ps-append');

  document.getElementById('ps-clear').addEventListener('click', () => {
    outEl.textContent = '(nothing run yet)';
    statusEl.textContent = '';
  });

  // Ctrl+Enter runs; plain Enter stays a newline so multi-line scripts work.
  cmdEl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); runBtn.click(); }
  });

  runBtn.addEventListener('click', async () => {
    const cmd = cmdEl.value;
    if (!cmd.trim()) { statusEl.className = 'status err'; statusEl.textContent = 'Nothing to run.'; return; }

    runBtn.disabled = true;
    statusEl.className = 'status';
    statusEl.textContent = 'Running... the host is single-threaded, so other views wait until this returns.';
    try {
      const res = await fetch('/api/psconsole/run', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ command: cmd })
      });
      const data = await res.json();
      const stamp = new Date().toLocaleTimeString();
      const block = 'PS [' + stamp + '] > ' + cmd + '\n' +
                    (data.text && data.text.length ? data.text : '(no output)') + '\n';
      if (appendEl.checked && outEl.textContent !== '(nothing run yet)') {
        outEl.textContent = outEl.textContent + '\n' + block;
      } else {
        outEl.textContent = block;
      }
      outEl.scrollTop = outEl.scrollHeight;
      statusEl.className = data.hadError ? 'status err' : 'status';
      statusEl.textContent = data.hadError ? 'Completed with errors.' : 'Completed.';
    } catch (err) {
      statusEl.className = 'status err';
      statusEl.textContent = 'Request failed: ' + err;
    } finally {
      runBtn.disabled = false;
    }
  });
})();
</script>
'@
    } `
    -Routes @{
        'POST /api/psconsole/run' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; hadError = $false; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                $command = "$($body.command)"
                if ([string]::IsNullOrWhiteSpace($command)) { throw "No command supplied." }

                # Log every command BEFORE running it, so the completed-tasks
                # log records what was attempted even if it then fails or the
                # host dies partway through. This is the only audit trail this
                # view has, and an unrestricted console deserves one.
                $firstLine = ($command -split "`n")[0].Trim()
                if ($firstLine.Length -gt 200) { $firstLine = $firstLine.Substring(0, 200) + '...' }
                Write-CompletedLog -Task "PowerShellConsole (web) ran: $firstLine"

                # Executed in the host's own runspace on purpose -- it must see
                # the host's live connections and imported modules. See the
                # header on why there is no sandbox.
                $sb = [scriptblock]::Create($command)
                $errs = @()
                $out = & $sb 2>&1 | Out-String -Width 4096

                # 2>&1 folds error records into the stream above, so they're
                # already in $out. $Error is checked only to set the flag that
                # colours the status line.
                if ($Error.Count -gt 0) {
                    $errs = @($Error | Select-Object -First 5 | ForEach-Object { "$_" })
                    $Error.Clear()
                }

                $result.ok       = $true
                $result.text     = $out
                $result.hadError = ($errs.Count -gt 0)
            } catch {
                # A parse failure (bad syntax) lands here rather than above.
                $result.ok       = $true
                $result.hadError = $true
                $result.text     = (Format-ErrorDetail -ErrorRecord $_ -Context 'PowerShell Console')
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
    }
