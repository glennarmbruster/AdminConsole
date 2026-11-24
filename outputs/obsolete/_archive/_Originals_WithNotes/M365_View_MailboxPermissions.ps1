#Requires -Version 5.1
<#
.SYNOPSIS
    Mailbox Permissions view (Full Access / Send As delegation) for
    M365_Host_Web.ps1.
.DESCRIPTION
    ADDED 2026-07-24 as part of the single-process host rewrite. Replaces
    standalone M365_MailboxPermissions_Web.ps1 -- see M365WebCommon.ps1's
    header for why the per-tool-process model was abandoned.

    Contains no logic. Everything is M365Logic.psm1:
      Get-MailboxState / Format-MailboxState for display, and
      Invoke-MailboxPermissionEdit for each grant/revoke (which does the
      BEFORE snapshot, the change, the AFTER snapshot, and the
      completed-tasks log entry in one call).

    CONFIRMATION IS THIS VIEW'S JOB. The module's grant/revoke functions
    deliberately never prompt -- so a scheduled job can call them -- which
    means the confirm() in the JS below is the only thing standing between a
    misclick and a real permission change. Do not remove it without
    replacing it with something equivalent.
#>

Register-M365View -Id 'perms' -Name 'Mailbox Permissions' `
    -Description 'View, grant, and revoke Full Access and Send As on a single mailbox.' `
    -GetHtml {
        return @'
  <h2>Mailbox Permissions</h2>
  <div class="note">Look up a mailbox first. Grant/revoke buttons enable once a lookup succeeds, and each asks for confirmation before changing anything.</div>

  <div class="row">
    <label>Mailbox:<br><input id="pm-smtp" type="text" placeholder="user@domain, or just the alias" style="width:300px"></label>
    <button id="pm-lookup">Look Up</button>
  </div>

  <div id="pm-out" class="out" style="min-height:240px;">Enter a mailbox above and click Look Up.</div>

  <fieldset>
    <legend>Grant / Revoke</legend>
    <div class="row">
      <label>Grantee (email or logon):<br><input id="pm-grantee" type="text" style="width:280px"></label>
    </div>
    <div class="row">
      <button class="pm-act" data-action="grant-full"    data-label="Grant Full Access"  disabled>Grant Full Access</button>
      <button class="pm-act secondary" data-action="revoke-full"   data-label="Revoke Full Access" disabled>Revoke Full Access</button>
      <button class="pm-act" data-action="grant-sendas"  data-label="Grant Send As"      disabled>Grant Send As</button>
      <button class="pm-act secondary" data-action="revoke-sendas" data-label="Revoke Send As"     disabled>Revoke Send As</button>
    </div>
  </fieldset>

<script>
(function () {
  const smtpEl    = document.getElementById('pm-smtp');
  const granteeEl = document.getElementById('pm-grantee');
  const outEl     = document.getElementById('pm-out');
  const lookupBtn = document.getElementById('pm-lookup');
  const actBtns   = Array.from(document.querySelectorAll('.pm-act'));

  function setActEnabled(on) { actBtns.forEach(b => b.disabled = !on); }

  async function lookup() {
    const smtp = smtpEl.value.trim();
    if (!smtp) { outEl.textContent = 'Enter a mailbox first.'; return; }
    lookupBtn.disabled = true;
    setActEnabled(false);
    outEl.textContent = 'Loading...';
    try {
      const res  = await fetch('/api/perms/state?smtp=' + encodeURIComponent(smtp));
      const data = await res.json();
      if (!data.ok) { outEl.textContent = 'ERROR: ' + data.error; return; }
      outEl.textContent = data.text;
      setActEnabled(true);
    } catch (err) {
      outEl.textContent = 'Request failed: ' + err;
    } finally {
      lookupBtn.disabled = false;
    }
  }
  lookupBtn.addEventListener('click', lookup);
  smtpEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') lookup(); });

  actBtns.forEach(btn => {
    btn.addEventListener('click', async () => {
      const smtp    = smtpEl.value.trim();
      const grantee = granteeEl.value.trim();
      const label   = btn.dataset.label;
      if (!smtp)    { outEl.textContent = 'Enter a mailbox first.'; return; }
      if (!grantee) { outEl.textContent = 'Enter a grantee (email or logon) first.'; return; }
      if (!confirm(label + " for '" + grantee + "' on mailbox '" + smtp + "'?")) return;
      setActEnabled(false);
      lookupBtn.disabled = true;
      outEl.textContent = 'Working...';
      try {
        const res = await fetch('/api/perms/' + btn.dataset.action, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ smtp: smtp, grantee: grantee })
        });
        const data = await res.json();
        outEl.textContent = data.ok ? data.text : ('ERROR: ' + data.error);
      } catch (err) {
        outEl.textContent = 'Request failed: ' + err;
      } finally {
        setActEnabled(true);
        lookupBtn.disabled = false;
      }
    });
  });
})();
</script>
'@
    } `
    -Routes @{
        'GET /api/perms/state' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $smtp = Get-TargetSmtp -Text $Request.QueryString['smtp']
                if (-not $smtp) { throw "Could not determine a domain -- type the full email address, or make sure Exchange Online is connected." }
                $state = Get-MailboxState -Smtp $smtp
                $result.ok   = $true
                $result.text = Format-MailboxState -Heading 'CURRENT STATE' -State $state
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Permissions'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'POST /api/perms/grant-full' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                $smtp = Get-TargetSmtp -Text $body.smtp
                $result.text = Invoke-MailboxPermissionEdit -Smtp $smtp -Grantee $body.grantee `
                    -ActionLabel 'Grant Full Access' -Verb 'Granted FullAccess' -Action {
                        Grant-FullAccessPermission -Smtp $smtp -Grantee $body.grantee
                    }
                $result.ok = $true
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Grant Full Access'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'POST /api/perms/revoke-full' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                $smtp = Get-TargetSmtp -Text $body.smtp
                $result.text = Invoke-MailboxPermissionEdit -Smtp $smtp -Grantee $body.grantee `
                    -ActionLabel 'Revoke Full Access' -Verb 'Revoked FullAccess' -Action {
                        Revoke-FullAccessPermission -Smtp $smtp -Grantee $body.grantee
                    }
                $result.ok = $true
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Revoke Full Access'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'POST /api/perms/grant-sendas' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                $smtp = Get-TargetSmtp -Text $body.smtp
                $result.text = Invoke-MailboxPermissionEdit -Smtp $smtp -Grantee $body.grantee `
                    -ActionLabel 'Grant Send As' -Verb 'Granted SendAs' -Action {
                        Grant-SendAsPermission -Smtp $smtp -Grantee $body.grantee
                    }
                $result.ok = $true
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Grant Send As'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'POST /api/perms/revoke-sendas' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                $smtp = Get-TargetSmtp -Text $body.smtp
                $result.text = Invoke-MailboxPermissionEdit -Smtp $smtp -Grantee $body.grantee `
                    -ActionLabel 'Revoke Send As' -Verb 'Revoked SendAs' -Action {
                        Revoke-SendAsPermission -Smtp $smtp -Grantee $body.grantee
                    }
                $result.ok = $true
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Revoke Send As'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
    }
