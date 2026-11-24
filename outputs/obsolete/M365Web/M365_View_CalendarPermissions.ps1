#Requires -Version 5.1
<#
    M365_View_CalendarPermissions.ps1
    Calendar Permissions view for M365_Host_Web.ps1.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

Register-M365View -Id 'calperms' -Name 'Calendar Permissions' `
    -Description 'View, grant, update, and remove calendar folder permissions on a mailbox.' `
    -GetHtml {
        $roleOptions = ((Get-CalendarPermissionRoles) | ForEach-Object {
            $sel = if ($_ -eq 'Reviewer') { ' selected' } else { '' }
            "        <option$sel>$_</option>"
        }) -join "`n"
        return @"
  <h2>Calendar Permissions</h2>
  <div class="note">
    The calendar folder is resolved by name rather than assumed to be ":\Calendar", so renamed or
    localized calendars still work. Removing a user removes <strong>all</strong> of their calendar
    permissions -- Exchange has no partial removal.
  </div>

  <div class="row">
    <label>Mailbox:<br><input id="cal-smtp" type="text" placeholder="user@domain, or just the alias" style="width:300px"></label>
    <button id="cal-lookup">Look Up</button>
  </div>

  <div id="cal-folder" class="status">No mailbox looked up yet.</div>

  <table id="cal-table">
    <thead><tr><th>User</th><th>Access Rights</th></tr></thead>
    <tbody id="cal-body"></tbody>
  </table>

  <fieldset>
    <legend>Grant / Update / Remove</legend>
    <div class="row">
      <label>User (email or logon):<br><input id="cal-grantee" type="text" style="width:260px"></label>
      <label>Role:<br>
        <select id="cal-role">
$roleOptions
        </select>
      </label>
      <button id="cal-set" disabled>Grant / Update</button>
      <button id="cal-remove" class="danger" disabled>Remove All</button>
    </div>
  </fieldset>

  <div id="cal-out" class="out" style="min-height:110px;">(no action run yet)</div>

<script>
(function () {
  const smtpEl    = document.getElementById('cal-smtp');
  const granteeEl = document.getElementById('cal-grantee');
  const roleEl    = document.getElementById('cal-role');
  const folderEl  = document.getElementById('cal-folder');
  const bodyEl    = document.getElementById('cal-body');
  const outEl     = document.getElementById('cal-out');
  const lookupBtn = document.getElementById('cal-lookup');
  const setBtn    = document.getElementById('cal-set');
  const remBtn    = document.getElementById('cal-remove');
  let calendarIdentity = null;

  function setActEnabled(on) { setBtn.disabled = !on; remBtn.disabled = !on; }

  async function lookup() {
    const smtp = smtpEl.value.trim();
    if (!smtp) { folderEl.className = 'status err'; folderEl.textContent = 'Enter a mailbox first.'; return; }
    lookupBtn.disabled = true;
    setActEnabled(false);
    folderEl.className = 'status';
    folderEl.textContent = 'Loading...';
    bodyEl.innerHTML = '';
    try {
      const res  = await fetch('/api/calperms/state?smtp=' + encodeURIComponent(smtp));
      const data = await res.json();
      if (!data.ok) { folderEl.className = 'status err'; folderEl.textContent = data.error; return; }
      calendarIdentity = data.calendarIdentity;
      folderEl.className = 'status';
      folderEl.textContent = 'Calendar folder: ' + calendarIdentity + '  (' + data.permissions.length + ' entr' + (data.permissions.length === 1 ? 'y' : 'ies') + ')';
      for (const p of data.permissions) {
        const tr = document.createElement('tr');
        [p.User, p.AccessRights].forEach(v => {
          const td = document.createElement('td'); td.textContent = v; tr.appendChild(td);
        });
        tr.addEventListener('click', () => { granteeEl.value = p.User; });
        bodyEl.appendChild(tr);
      }
      setActEnabled(true);
    } catch (err) {
      folderEl.className = 'status err';
      folderEl.textContent = 'Request failed: ' + err;
    } finally {
      lookupBtn.disabled = false;
    }
  }
  lookupBtn.addEventListener('click', lookup);
  smtpEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') lookup(); });

  async function act(kind) {
    const grantee = granteeEl.value.trim();
    if (!calendarIdentity) { outEl.textContent = 'Look up a mailbox first.'; return; }
    if (!grantee) { outEl.textContent = 'Enter a user (email or logon) first.'; return; }

    let msg;
    if (kind === 'set') {
      msg = "Set '" + grantee + "' to role " + roleEl.value + " on " + calendarIdentity + '?';
    } else {
      msg = 'Remove ALL calendar permissions for ' + grantee + ' on ' + calendarIdentity +
            '?\n\nExchange has no partial removal -- every right this user has on the calendar goes away.';
    }
    // GUARD: confirmation is required here. Do not remove.
    if (!confirm(msg)) return;

    setActEnabled(false);
    lookupBtn.disabled = true;
    outEl.textContent = 'Working...';
    try {
      const res = await fetch('/api/calperms/' + kind, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ calendarIdentity: calendarIdentity, grantee: grantee, role: roleEl.value })
      });
      const data = await res.json();
      outEl.textContent = data.ok ? data.text : ('ERROR: ' + data.error);
      if (data.ok) await lookup();
    } catch (err) {
      outEl.textContent = 'Request failed: ' + err;
    } finally {
      setActEnabled(true);
      lookupBtn.disabled = false;
    }
  }
  setBtn.addEventListener('click', () => act('set'));
  remBtn.addEventListener('click', () => act('remove'));
})();
</script>
"@
    } `
    -Routes @{
        'GET /api/calperms/state' = {
            param($Request)
            $result = [ordered]@{ ok = $false; calendarIdentity = $null; permissions = @(); error = '' }
            try {
                $smtp = Get-TargetSmtp -Text $Request.QueryString['smtp']
                if (-not $smtp) { throw "Could not determine a domain -- type the full email address, or make sure Exchange Online is connected." }
                $state = Get-CalendarPermissions -Smtp $smtp
                if ($state.Error) { throw $state.Error }
                $result.ok = $true
                $result.calendarIdentity = $state.CalendarIdentity
                $result.permissions = @($state.Permissions | ForEach-Object {
                    [ordered]@{
                        User         = "$($_.User)"
                        AccessRights = (@($_.AccessRights) -join ', ')
                    }
                })
            } catch {
                $result.error = if ($_.Exception.Message -like 'ERROR *') { $_.Exception.Message } else { Format-ErrorDetail -ErrorRecord $_ -Context 'Calendar Permissions' }
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'POST /api/calperms/set' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                $what = Set-CalendarPermission -CalendarIdentity $body.calendarIdentity -Grantee $body.grantee -Role $body.role
                Write-CompletedLog -Task "$what CalendarPermission $($body.calendarIdentity) (user: $($body.grantee); role: $($body.role))"
                $result.ok   = $true
                $result.text = "$what : $($body.grantee) now has role '$($body.role)' on $($body.calendarIdentity)."
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Set calendar permission'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'POST /api/calperms/remove' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                Remove-CalendarPermission -CalendarIdentity $body.calendarIdentity -Grantee $body.grantee
                Write-CompletedLog -Task "Removed CalendarPermission $($body.calendarIdentity) (user: $($body.grantee))"
                $result.ok   = $true
                $result.text = "Removed all calendar permissions for $($body.grantee) on $($body.calendarIdentity)."
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove calendar permission'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
    }
