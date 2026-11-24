#Requires -Version 5.1
<#
    M365_View_InboxRules.ps1
    Inbox Rules view for M365_Host_Web.ps1.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

Register-M365View -Id 'rules' -Name 'Inbox Rules' `
    -Description 'List a mailbox inbox rules, flag suspicious ones, and enable, disable, or delete them.' `
    -GetHtml {
        return @'
  <h2>Inbox Rules</h2>
  <div class="note">
    Rules that are <strong>enabled</strong> and forward, redirect, or delete mail are highlighted for review.
    That is a prompt to look, not a verdict -- many legitimate rules do exactly that.
  </div>

  <div class="row">
    <label>Mailbox:<br><input id="ir-smtp" type="text" placeholder="user@domain, or just the alias" style="width:300px"></label>
    <button id="ir-load">Load Rules</button>
    <span id="ir-status" class="status" style="margin:0;">Nothing loaded yet.</span>
  </div>

  <table id="ir-table">
    <thead><tr>
      <th>Name</th><th>Enabled</th><th>Flagged Actions</th><th>Description</th>
    </tr></thead>
    <tbody id="ir-body"></tbody>
  </table>

  <h3 style="margin-top:16px;">Selected rule</h3>
  <div id="ir-detail" class="out" style="min-height:120px;">Click a rule above to select it.</div>

  <div class="row">
    <button id="ir-enable" disabled>Enable</button>
    <button id="ir-disable" class="secondary" disabled>Disable</button>
    <button id="ir-delete" class="danger" disabled>Delete</button>
    <span id="ir-action-status" class="status" style="margin:0;"></span>
  </div>

<script>
(function () {
  const smtpEl   = document.getElementById('ir-smtp');
  const statusEl = document.getElementById('ir-status');
  const bodyEl   = document.getElementById('ir-body');
  const detailEl = document.getElementById('ir-detail');
  const actEl    = document.getElementById('ir-action-status');
  const loadBtn  = document.getElementById('ir-load');
  const enBtn    = document.getElementById('ir-enable');
  const disBtn   = document.getElementById('ir-disable');
  const delBtn   = document.getElementById('ir-delete');
  let selected = null;

  function setActEnabled(on) {
    enBtn.disabled = !on; disBtn.disabled = !on; delBtn.disabled = !on;
  }

  async function load() {
    const smtp = smtpEl.value.trim();
    if (!smtp) { statusEl.className = 'status err'; statusEl.textContent = 'Enter a mailbox first.'; return; }
    loadBtn.disabled = true;
    setActEnabled(false);
    selected = null;
    detailEl.textContent = 'Click a rule above to select it.';
    bodyEl.innerHTML = '';
    statusEl.className = 'status';
    statusEl.textContent = 'Loading...';
    try {
      const res  = await fetch('/api/rules/list?smtp=' + encodeURIComponent(smtp));
      const data = await res.json();
      if (!data.ok) { statusEl.className = 'status err'; statusEl.textContent = data.error; return; }
      const flagged = data.rules.filter(r => r.Suspicious).length;
      statusEl.className = 'status';
      statusEl.textContent = data.rules.length + ' rule(s)' + (flagged ? ('  --  ' + flagged + ' flagged for review') : '');
      for (const r of data.rules) {
        const tr = document.createElement('tr');
        if (r.Suspicious) { tr.style.background = '#fdecea'; }
        [r.Name, (r.Enabled ? 'Yes' : 'No'), (r.Actions || ''), r.Description].forEach(v => {
          const td = document.createElement('td'); td.textContent = v; tr.appendChild(td);
        });
        tr.addEventListener('click', () => {
          selected = r;
          Array.from(bodyEl.children).forEach(x => x.classList.remove('selected'));
          tr.classList.add('selected');
          detailEl.textContent =
            'Rule        : ' + r.Name + '\n' +
            'Mailbox     : ' + smtpEl.value.trim() + '\n' +
            'Enabled     : ' + r.Enabled + '\n' +
            'Priority    : ' + (r.Priority === null ? '' : r.Priority) + '\n' +
            'Flagged     : ' + (r.Suspicious ? ('yes -- ' + r.Actions) : 'no') + '\n' +
            'Identity    : ' + r.Identity + '\n\n' +
            'Full description:\n' + (r.FullDescription || '(none)');
          setActEnabled(true);
          actEl.textContent = '';
        });
        bodyEl.appendChild(tr);
      }
    } catch (err) {
      statusEl.className = 'status err';
      statusEl.textContent = 'Request failed: ' + err;
    } finally {
      loadBtn.disabled = false;
    }
  }
  loadBtn.addEventListener('click', load);
  smtpEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') load(); });

  async function act(kind, label) {
    if (!selected) return;
    const smtp = smtpEl.value.trim();
    let msg = label + " rule '" + selected.Name + "' on " + smtp + '?';
    if (kind === 'delete') {
      msg = "DELETE rule '" + selected.Name + "' on " + smtp +
            '?\n\nThis cannot be undone -- a deleted inbox rule is not recoverable.';
    }
    // GUARD: confirmation is required here. Do not remove.
    if (!confirm(msg)) return;

    setActEnabled(false);
    loadBtn.disabled = true;
    actEl.className = 'status';
    actEl.textContent = 'Working...';
    try {
      const res = await fetch('/api/rules/' + kind, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ smtp: smtp, identity: selected.Identity, name: selected.Name })
      });
      const data = await res.json();
      if (data.ok) {
        actEl.className = 'status';
        actEl.textContent = data.text;
        await load();
      } else {
        actEl.className = 'status err';
        actEl.textContent = data.error;
      }
    } catch (err) {
      actEl.className = 'status err';
      actEl.textContent = 'Request failed: ' + err;
    } finally {
      loadBtn.disabled = false;
    }
  }
  enBtn.addEventListener('click',  () => act('enable',  'Enable'));
  disBtn.addEventListener('click', () => act('disable', 'Disable'));
  delBtn.addEventListener('click', () => act('delete',  'Delete'));
})();
</script>
'@
    } `
    -Routes @{
        'GET /api/rules/list' = {
            param($Request)
            $result = [ordered]@{ ok = $false; rules = @(); error = '' }
            try {
                $smtp = Get-TargetSmtp -Text $Request.QueryString['smtp']
                if (-not $smtp) { throw "Could not determine a domain -- type the full email address, or make sure Exchange Online is connected." }
                $res = Get-InboxRuleSummaries -Smtp $smtp
                if ($res.Error) { $result.error = $res.Error }
                else {
                    $result.ok = $true
                    $result.rules = @($res.Rules | ForEach-Object {
                        [ordered]@{
                            Name            = $_.Name
                            Identity        = $_.Identity
                            Enabled         = $_.Enabled
                            Priority        = $_.Priority
                            Description     = $_.Description
                            FullDescription = $_.FullDescription
                            Actions         = $_.Actions
                            Suspicious      = $_.Suspicious
                        }
                    })
                }
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Inbox Rules'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 6 -Compress) }
        }
        'POST /api/rules/enable' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                Enable-TargetInboxRule -Smtp $body.smtp -RuleIdentity $body.identity
                Write-CompletedLog -Task "Enabled InboxRule '$($body.name)' on $($body.smtp)"
                $result.ok = $true; $result.text = "Enabled '$($body.name)'."
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Enable-InboxRule'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'POST /api/rules/disable' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                Disable-TargetInboxRule -Smtp $body.smtp -RuleIdentity $body.identity
                Write-CompletedLog -Task "Disabled InboxRule '$($body.name)' on $($body.smtp)"
                $result.ok = $true; $result.text = "Disabled '$($body.name)'."
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Disable-InboxRule'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'POST /api/rules/delete' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                Remove-TargetInboxRule -Smtp $body.smtp -RuleIdentity $body.identity
                Write-CompletedLog -Task "Deleted InboxRule '$($body.name)' on $($body.smtp)"
                $result.ok = $true; $result.text = "Deleted '$($body.name)'. This is not recoverable."
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-InboxRule'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
    }
