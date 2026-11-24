#Requires -Version 5.1
<#
    M365_View_MailboxOneDriveAccess.ps1
    Mailbox & OneDrive Access view for M365_Host_Web.ps1. Bulk,
    many-to-many.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

Register-M365View -Id 'mbxod' -Name 'Mailbox & OneDrive Access' -Destructive `
    -Description 'Bulk grant/remove Full Access, Send-As, and OneDrive site admin across a set of mailboxes and users.' `
    -GetHtml {
        return @'
  <h2>Mailbox &amp; OneDrive Access</h2>
  <div class="warn">
    MANY-TO-MANY: every user listed gets the selected access on every mailbox listed. Three mailboxes and
    four users is twelve changes. The confirmation shows the exact count before anything runs.
  </div>

  <div class="row">
    <label style="flex:1;">Mailbox(es) -- one per line:<br>
      <textarea id="od-mailboxes" rows="5"></textarea>
    </label>
    <label style="flex:1;">User(s) / Manager(s) receiving access -- one per line:<br>
      <textarea id="od-users" rows="5"></textarea>
    </label>
  </div>

  <fieldset>
    <legend>Access Type</legend>
    <label><input type="radio" name="od-type" value="Full" checked> Full Access</label>
    <label style="margin-left:14px;"><input type="radio" name="od-type" value="SendAs"> Send-As</label>
    <label style="margin-left:14px;"><input type="radio" name="od-type" value="OneDrive"> OneDrive</label>
    <label style="margin-left:14px;"><input type="radio" name="od-type" value="FullSendAs"> Full &amp; Send-As</label>
    <label style="margin-left:14px;"><input type="radio" name="od-type" value="FullOneDrive"> Full &amp; OneDrive</label>
    <div class="status" style="margin-top:6px;">Send-As + OneDrive is not offered -- that combination is never needed.</div>
  </fieldset>

  <div class="row">
    <fieldset style="margin:0;">
      <legend>Mode</legend>
      <label><input type="radio" name="od-mode" value="Grant" checked> Grant</label>
      <label style="margin-left:14px;"><input type="radio" name="od-mode" value="Remove"> Remove</label>
    </fieldset>
    <label>SharePoint tenant prefix (OneDrive types only -- the "contoso" in https://contoso-my.sharepoint.com):<br>
      <input id="od-spo" type="text" style="width:260px">
    </label>
  </div>

  <div class="row">
    <button id="od-run" class="danger">Run</button>
    <button id="od-clear" class="secondary">Clear</button>
    <span id="od-status" class="status" style="margin:0;"></span>
  </div>

  <div id="od-links"></div>
  <div id="od-out" class="out" style="min-height:220px;">(nothing run yet)</div>

<script>
(function () {
  const mbxEl    = document.getElementById('od-mailboxes');
  const usersEl  = document.getElementById('od-users');
  const spoEl    = document.getElementById('od-spo');
  const outEl    = document.getElementById('od-out');
  const linksEl  = document.getElementById('od-links');
  const statusEl = document.getElementById('od-status');
  const runBtn   = document.getElementById('od-run');

  const splitList = (s) => s.split(/[\r\n,;]+/).map(x => x.trim()).filter(x => x !== '');
  const accessType = () => document.querySelector('input[name=od-type]:checked').value;
  const mode       = () => document.querySelector('input[name=od-mode]:checked').value;
  const needsOneDrive = () => ['OneDrive', 'FullOneDrive'].indexOf(accessType()) !== -1;

  const TYPE_TEXT = {
    Full: 'Full Access', SendAs: 'Send-As', OneDrive: 'OneDrive',
    FullSendAs: 'Full Access & Send-As', FullOneDrive: 'Full Access & OneDrive'
  };

  document.getElementById('od-clear').addEventListener('click', () => {
    mbxEl.value = ''; usersEl.value = ''; spoEl.value = '';
    document.querySelector('input[name=od-type][value=Full]').checked = true;
    document.querySelector('input[name=od-mode][value=Grant]').checked = true;
    statusEl.textContent = '';
    linksEl.innerHTML = '';
    outEl.textContent = '(nothing run yet)';
  });

  runBtn.addEventListener('click', async () => {
    const mailboxes = splitList(mbxEl.value);
    const users     = splitList(usersEl.value);
    statusEl.className = 'status';

    if (!mailboxes.length) { statusEl.className = 'status err'; statusEl.textContent = 'Enter at least one mailbox.'; return; }
    if (!users.length)     { statusEl.className = 'status err'; statusEl.textContent = 'Enter at least one user/manager.'; return; }
    if (needsOneDrive() && !spoEl.value.trim()) {
      statusEl.className = 'status err';
      statusEl.textContent = 'This Access Type includes OneDrive, so a SharePoint tenant prefix is required.';
      return;
    }

    const ops = mailboxes.length * users.length;
    const msg = mode() + ' ' + TYPE_TEXT[accessType()] + '?\n\n' +
      'Mailboxes (' + mailboxes.length + '):\n  ' + mailboxes.join('\n  ') + '\n\n' +
      'Users (' + users.length + '):\n  ' + users.join('\n  ') + '\n\n' +
      'That is ' + ops + ' mailbox/user combination' + (ops === 1 ? '' : 's') + '.';
    // GUARD: confirmation is required here. Do not remove.
    if (!confirm(msg)) { outEl.textContent = 'Cancelled.'; return; }

    runBtn.disabled = true;
    linksEl.innerHTML = '';
    statusEl.textContent = 'Running ' + ops + ' operation' + (ops === 1 ? '' : 's') + '...';
    outEl.textContent = '';
    try {
      const res = await fetch('/api/mbxod/run', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          mailboxes: mbxEl.value, users: usersEl.value,
          accessType: accessType(), mode: mode(), spoPrefix: spoEl.value.trim()
        })
      });
      const data = await res.json();
      if (!data.ok) {
        statusEl.className = 'status err';
        statusEl.textContent = 'Failed.';
        outEl.textContent = data.error;
        return;
      }
      statusEl.className = 'status';
      statusEl.textContent = data.summary;
      outEl.textContent = data.text;

      if (data.oneDriveLinks && data.oneDriveLinks.length) {
        const box = document.createElement('div');
        box.className = 'card';
        const h = document.createElement('h3');
        h.textContent = 'OneDrive can be accessed via the following link' + (data.oneDriveLinks.length === 1 ? '' : 's');
        box.appendChild(h);
        data.oneDriveLinks.forEach(l => {
          const row = document.createElement('div');
          row.style.cssText = 'padding:4px 0;font-size:13px;';
          const who = document.createElement('span');
          who.style.cssText = 'font-family:Consolas,monospace;margin-right:10px;';
          who.textContent = l.mailbox;
          const a = document.createElement('a');
          a.href = l.url; a.target = '_blank'; a.rel = 'noopener';
          a.textContent = l.url;
          a.style.cssText = 'color:#1d4ed8;font-weight:600;';
          row.appendChild(who); row.appendChild(a);
          box.appendChild(row);
        });
        linksEl.appendChild(box);
      }
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
        'POST /api/mbxod/run' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; summary = ''; oneDriveLinks = @(); error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json

                $mailboxes = @(ConvertTo-NormalizedUserList -Text $body.mailboxes | ForEach-Object { Get-TargetSmtp -Text $_ })
                $users     = @(ConvertTo-NormalizedUserList -Text $body.users     | ForEach-Object { Get-TargetSmtp -Text $_ })
                if ($mailboxes -contains $null -or $users -contains $null) {
                    throw "One or more entries could not be resolved to an email address -- type full addresses, or make sure Exchange Online is connected so the default domain can be detected."
                }
                $users = @($users | Where-Object { $mailboxes -notcontains $_ })
                if ($mailboxes.Count -eq 0) { throw "Enter at least one mailbox." }
                if ($users.Count -eq 0)     { throw "Enter at least one user/manager that is not also listed as a mailbox." }

                $r = Invoke-M365MailboxOneDriveAccess -Mailboxes $mailboxes -Users $users `
                        -AccessType $body.accessType -Mode $body.mode -SpoTenantPrefix $body.spoPrefix

                $result.ok      = $true
                $result.text    = $r.Text
                $result.summary = "$($r.Succeeded) succeeded, $($r.Failed) failed, $($r.Skipped) with nothing to do."
                $result.oneDriveLinks = @($r.OneDriveGrantedFor | ForEach-Object {
                    [ordered]@{
                        mailbox = $_
                        url     = (Get-OneDrivePersonalSiteUrl -Smtp $_ -SpoTenantPrefix $body.spoPrefix)
                    }
                })
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox & OneDrive Access'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 6 -Compress) }
        }
    }
