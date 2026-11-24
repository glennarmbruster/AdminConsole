#Requires -Version 5.1
<#
.SYNOPSIS
    Disable / termination view for M365_Host_Web.ps1. DESTRUCTIVE.
.DESCRIPTION
    ADDED 2026-07-24 as part of the single-process host rewrite. Replaces
    standalone M365_Disable_Web.ps1.

    Contains no logic -- ConvertTo-NormalizedUserList,
    Get-OneDrivePersonalSiteUrl, and Disable-M365User all live in
    M365Logic.psm1. Per user, Disable-M365User: hides from address lists and
    converts to a shared mailbox, optionally grants a manager Full Access
    (never Send-As), and optionally grants the manager site-collection admin
    on the user's OneDrive.

    CONFIRMATION IS THIS VIEW'S JOB, and it matters more here than anywhere
    else in the toolkit. Disable-M365User does exactly what it's told
    immediately, with no prompt, so it can be driven by a scheduled job --
    which means the confirm() below listing every affected user is the only
    guard against a misclick converting real mailboxes to shared. Do not
    remove it.

    SHAREPOINT TENANT PREFIX is an operator-supplied field, not a constant.
    The original ANZ tool hardcoded 'expressprosau', and per that tool's own
    header it is NOT derivable from the mail domain by any reliable rule --
    naively stripping the TLD from 'expresspros.com.au' gives 'expresspros',
    which was confirmed wrong for that exact tenant. Guessing would point a
    site-collection-admin grant at the wrong SharePoint site, so it's asked
    for instead. Only needed when the OneDrive step is used; the module
    fails that step with a clear message (and still completes the others) if
    it's missing.

    LICENSE REMOVAL AND SIGN-IN BLOCK ARE NOT AUTOMATED, matching the
    original's final state. CORRECTED 2026-07-24 -- earlier notes in this
    codebase (including this file) said the AzureAD module was "permanently
    dead," which overstates it. What's true: the module is DEPRECATED
    (announced 2024-03-30), and cmdlets that call the retired Azure AD Graph
    endpoint fail -- which is precisely what broke Get-AzureADGroup on a real
    run. But Connect-AzureAD itself still authenticates, confirmed by the
    operator's daily use. The reason these two steps stay manual is therefore
    NOT "the module is gone"; it's that the specific group/user-write cmdlets
    needed are the ones hitting the retired endpoint, and the Microsoft.Entra
    alternative needed tenant-admin consent that wasn't available. The run
    reports both as NOT ATTEMPTED, and this view now renders a per-user Entra
    deep link for each so finishing them by hand is two clicks.
#>

Register-M365View -Id 'disable' -Name 'Disable' -Destructive `
    -Description 'Hide and convert one or more mailboxes to shared, and optionally grant a manager Full Access and/or OneDrive access. Destructive -- confirms first.' `
    -GetHtml {
        return @'
  <h2>Disable</h2>
  <div class="warn">DESTRUCTIVE. Each listed user is hidden from address lists and converted to a shared mailbox. License removal and sign-in block are NOT performed here -- a reminder lists which users still need those done manually.</div>

  <div class="row">
    <label style="flex:1;">User(s) to disable -- one per line (commas and semicolons also work):<br>
      <textarea id="ds-users" rows="6"></textarea>
    </label>
  </div>

  <div class="row">
    <label>Manager (optional):<br><input id="ds-manager" type="text" style="width:240px"></label>
    <label>SharePoint tenant prefix (only for OneDrive access -- the "contoso" in https://contoso-my.sharepoint.com):<br>
      <input id="ds-spo" type="text" style="width:260px">
    </label>
  </div>

  <fieldset id="ds-access" disabled>
    <legend>Manager Access (applies only if a manager is entered)</legend>
    <label><input type="radio" name="ds-mgr" value="both" checked> Mailbox &amp; OneDrive</label>
    <label style="margin-left:14px;"><input type="radio" name="ds-mgr" value="mailbox"> Mailbox Only</label>
    <label style="margin-left:14px;"><input type="radio" name="ds-mgr" value="onedrive"> OneDrive Only</label>
  </fieldset>

  <div class="row">
    <button id="ds-run" class="danger">Disable</button>
    <button id="ds-clear" class="secondary">Clear</button>
  </div>

  <div id="ds-label" class="status">Result:</div>

  <!-- Follow-up actions, populated after a run. The two steps this tool
       deliberately does not automate get a per-user deep link each, so
       finishing them is two clicks instead of a hunt through the portal. -->
  <div id="ds-followup"></div>

  <div id="ds-out" class="out" style="min-height:200px;">(nothing run yet)</div>

<script>
(function () {
  const usersEl   = document.getElementById('ds-users');
  const managerEl = document.getElementById('ds-manager');
  const spoEl     = document.getElementById('ds-spo');
  const accessEl  = document.getElementById('ds-access');
  const outEl     = document.getElementById('ds-out');
  const labelEl   = document.getElementById('ds-label');
  const runBtn    = document.getElementById('ds-run');

  const followEl = document.getElementById('ds-followup');

  function syncAccess() { accessEl.disabled = (managerEl.value.trim() === ''); }
  managerEl.addEventListener('input', syncAccess);
  syncAccess();

  // Renders the two manual follow-up steps as per-user deep links. Built
  // from the Entra object GUID the server captured during the run (from
  // Get-Mailbox's ExternalDirectoryObjectId), so each link lands on that
  // specific user's page rather than dumping the operator in the users list.
  // When a GUID couldn't be read, the server already substituted the
  // all-users URL, so the link still works -- it's just less direct, and the
  // row says so.
  function renderFollowUp(users) {
    followEl.innerHTML = '';
    if (!users || !users.length) return;

    const box = document.createElement('div');
    box.className = 'card destructive';

    const h = document.createElement('h3');
    h.textContent = 'Still to do manually (' + users.length + ' user' + (users.length === 1 ? '' : 's') + ')';
    box.appendChild(h);

    const d = document.createElement('div');
    d.className = 'desc';
    d.textContent = 'This tool does not remove licenses or block sign-in. Both are done on the user’s Entra page: ' +
                    'Block sign-in is a button on the overview toolbar, and license groups are under Groups in the left nav. ' +
                    'Links open in a new tab.';
    box.appendChild(d);

    users.forEach(u => {
      const row = document.createElement('div');
      row.style.cssText = 'display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:6px 0;border-top:1px solid #eee;';

      const who = document.createElement('span');
      who.style.cssText = 'font-family:Consolas,monospace;font-size:13px;min-width:260px;';
      who.textContent = u.smtp;
      row.appendChild(who);

      const mkLink = (text, url) => {
        const a = document.createElement('a');
        a.href = url; a.target = '_blank'; a.rel = 'noopener';
        a.textContent = text;
        a.style.cssText = 'font-size:13px;font-weight:600;color:#1d4ed8;text-decoration:none;border:1px solid #bfdbfe;background:#eff6ff;padding:5px 10px;border-radius:4px;';
        return a;
      };
      row.appendChild(mkLink('Block sign-in ↗', u.entraUrl));
      row.appendChild(mkLink('License groups ↗', u.entraUrl));

      if (!u.hasObjectId) {
        const warn = document.createElement('span');
        warn.className = 'status err';
        warn.style.margin = '0';
        warn.textContent = '(object ID unavailable -- links open the all-users list; search for this address)';
        row.appendChild(warn);
      }
      box.appendChild(row);
    });

    followEl.appendChild(box);
  }

  document.getElementById('ds-clear').addEventListener('click', () => {
    usersEl.value = ''; managerEl.value = ''; spoEl.value = '';
    document.querySelector('input[name=ds-mgr][value=both]').checked = true;
    syncAccess();
    labelEl.className = 'status';
    labelEl.textContent = 'Result:';
    outEl.textContent = '(nothing run yet)';
    followEl.innerHTML = '';
  });

  runBtn.addEventListener('click', async () => {
    const raw = usersEl.value;
    if (!raw.trim()) { outEl.textContent = 'Enter at least one user to disable first.'; return; }

    const manager  = managerEl.value.trim();
    const mode     = document.querySelector('input[name=ds-mgr]:checked').value;
    const doMbx    = manager !== '' && (mode === 'both' || mode === 'mailbox');
    const doOneDr  = manager !== '' && (mode === 'both' || mode === 'onedrive');
    const spo      = spoEl.value.trim();

    if (doOneDr && !spo) {
      outEl.textContent = 'Manager Access includes OneDrive but no SharePoint tenant prefix was entered. Fill that in, or switch to Mailbox Only.';
      return;
    }

    const list = raw.split(/[\r\n,;]+/).map(s => s.trim()).filter(s => s !== '');
    let accessNote;
    if (manager) {
      const bits = [];
      if (doMbx)   bits.push('mailbox Full Access (never Send-As)');
      if (doOneDr) bits.push('OneDrive site collection admin');
      accessNote = "Manager '" + manager + "' will be granted: " + bits.join(' and ') + '.';
    } else {
      accessNote = 'No manager entered -- no Full Access or OneDrive access will be granted to anyone.';
    }

    const msg = 'Disable ' + list.length + ' user(s)?\n\n' + list.join('\n') +
      '\n\nFor each: hide from address lists, convert to shared mailbox.' +
      '\n\nLicense removal and sign-in block are NOT performed -- handle those manually.' +
      '\n\n' + accessNote;
    if (!confirm(msg)) { outEl.textContent = 'Cancelled.'; return; }

    runBtn.disabled = true;
    labelEl.className = 'status';
    labelEl.textContent = 'Running...';
    outEl.textContent = '';
    followEl.innerHTML = '';
    try {
      const res = await fetch('/api/disable/run', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ users: raw, manager: manager, grantMailbox: doMbx, grantOneDrive: doOneDr, spoPrefix: spo })
      });
      const data = await res.json();
      if (!data.ok) {
        labelEl.className = 'status err';
        labelEl.textContent = 'Failed:';
        outEl.textContent = data.error;
        return;
      }
      labelEl.className = 'status';
      labelEl.textContent = 'Done -- processed ' + data.users.length + ' user(s).';
      outEl.textContent = data.text;
      renderFollowUp(data.users);
    } catch (err) {
      labelEl.className = 'status err';
      labelEl.textContent = 'Failed:';
      outEl.textContent = 'Request failed: ' + err;
    } finally {
      runBtn.disabled = false;
    }
  });
})();
</script>
'@
    } `
    -Routes @{
        'POST /api/disable/run' = {
            param($Request)
            # REVISED 2026-07-24 -- returns per-user objects (with the Entra
            # deep link already built server-side) instead of a flat list of
            # addresses, so the page can render the two manual follow-up
            # steps as real links. Building the URL here rather than in JS
            # keeps every portal URL in one place: M365Logic's
            # Get-M365PortalUrl.
            $result = [ordered]@{ ok = $false; text = ''; users = @(); error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json

                $users = ConvertTo-NormalizedUserList -Text $body.users
                if ($users.Count -eq 0) { throw "Enter at least one user to disable first." }
                $userSmtps = @($users | ForEach-Object { Get-TargetSmtp -Text $_ })
                if ($userSmtps -contains $null) {
                    throw "One or more entries could not be resolved to an email address -- type full addresses, or make sure Exchange Online is connected so the default domain can be detected."
                }

                $managerSmtp = $null
                if ($body.manager) { $managerSmtp = Get-TargetSmtp -Text $body.manager }

                $out = New-Object System.Text.StringBuilder
                $userRows = @()
                foreach ($smtp in $userSmtps) {
                    $r = Disable-M365User -Smtp $smtp -ManagerSmtp $managerSmtp `
                        -GrantMailboxAccess ([bool]$body.grantMailbox) `
                        -GrantOneDriveAccess ([bool]$body.grantOneDrive) `
                        -SpoTenantPrefix $body.spoPrefix
                    [void]$out.Append($r.Text)
                    [void]$out.AppendLine("")
                    $userRows += [ordered]@{
                        smtp        = $r.Smtp
                        hasObjectId = [bool]$r.ObjectId
                        entraUrl    = (Get-M365PortalUrl -Kind 'EntraUser' -ObjectId $r.ObjectId)
                    }
                }
                $result.ok    = $true
                $result.text  = $out.ToString()
                $result.users = $userRows
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'M365 Disable'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 6 -Compress) }
        }
    }
