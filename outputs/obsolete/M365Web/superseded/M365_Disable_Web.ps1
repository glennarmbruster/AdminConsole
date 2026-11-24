#Requires -Version 7.0
<#
.SYNOPSIS
    Tenant-agnostic, browser-based termination/disable tool -- hides and
    converts one or more mailboxes to shared, and optionally grants a
    manager Full Access (mailbox) and/or Site Collection Admin (OneDrive).
    Ported from ANZDisable.ps1. DESTRUCTIVE -- confirms before running.
.DESCRIPTION
    ADDED 2026-07-24 -- part of the same batch as M365_MessageTrace_Web.ps1,
    M365_MailboxPermissions_Web.ps1, and M365_AdminConsole_Web.ps1. Uses
    M365WebCommon.ps1's Start-M365WebTool, which also imports the logic
    module. Disable-M365User (renamed from Disable-AnzUser) is otherwise the
    same four-steps-per-user function as the original, unchanged in
    behavior -- as of 2026-07-24 it lives in M365Logic.psm1 rather than in
    this file, so the same function is available to WinForms tools and
    scheduled jobs too. NOTE that the module's copy performs NO
    confirmation: this page's client-side confirm() is what guards it, and
    any new caller must supply its own.

    THE ONE PIECE THAT COULD NOT JUST BE COPIED OVER: SPO TENANT PREFIX.
    ANZDisable.ps1 hardcoded $script:SpoTenantPrefix = 'expressprosau' --
    genuinely tenant-specific, and per that file's own header, NOT
    derivable from the mail domain by any reliable rule ("expresspros.com.au"
    naively stripped would give "expresspros", which was already confirmed
    wrong for this exact tenant). Rather than guess at an SPO cmdlet that
    might return it automatically -- Get-SPOTenant's exact output shape
    hasn't been verified here, and getting this wrong would point a
    site-collection-admin grant at the WRONG SharePoint site, which is a
    real safety issue, not a cosmetic one -- this version makes it a
    required text field on the page instead: "SharePoint tenant prefix
    (e.g. contoso if your OneDrive URLs look like
    https://contoso-my.sharepoint.com/...)". Same algorithmic slug
    construction as before (dots and @ replaced with underscores) once the
    prefix is known -- that part of the original IS a genuinely standard,
    documented Microsoft pattern, not tenant-specific, so it stayed as
    logic rather than becoming another input field.

    CMDLET VERIFICATION CARRIED OVER FROM THE ORIGINAL, still true here:
    Set-Mailbox -Type Shared / -HiddenFromAddressListsEnabled,
    Add-MailboxPermission -AccessRights FullAccess, and Set-SPOUser
    -IsSiteCollectionAdmin were all checked against Microsoft Learn before
    ANZDisable.ps1 was delivered. What's NEW and UNVERIFIED here specifically:
    whether Microsoft.Online.SharePoint.PowerShell (Connect-SPOService,
    Set-SPOUser -- what the OneDrive step needs) actually loads and works
    under PowerShell 7 cross-platform. This module has historically been a
    Windows PowerShell 5.1-oriented one; if Set-SPOUser fails to load in a
    pwsh 7 session, the OneDrive-access step specifically is the part that
    would need PnP.PowerShell instead (different cmdlets, different auth
    flow) -- not attempted here since that's a real rewrite, not a
    substitution, and hasn't been verified against a live tenant. The
    mailbox-only path (hide+shared, manager Full Access) has no such
    dependency and should be unaffected either way.

    License removal and sign-in block are STILL NOT AUTOMATED, matching the
    original's final state (AzureAD module permanently dead; Microsoft.Entra
    needed one-time tenant-admin consent that wasn't available when this was
    built at Express Personnel). A completion message on the page reminds
    the operator which users need both handled manually after every run,
    same wording as the original's popup. If this is ever run against a
    tenant where full admin consent is already available (e.g. a personally
    administered tenant), that reminder-only behavior can be revisited and
    the Microsoft.Entra automation re-added -- not done here since that's a
    scope decision beyond "port this tool to the web," not something to
    silently re-enable.

    Same shape as the original: NOT routed through Resolve-ToolTarget/a
    single-target flow -- one Users textarea (one or more, newline/comma/
    semicolon separated), one optional Manager field, one Manager Access
    selection (Mailbox & OneDrive / Mailbox Only / OneDrive Only) that
    applies to the WHOLE batch, same as before. Confirms once, client-side,
    before running the whole batch -- same as the original's single
    MessageBox confirmation covering every user in the batch.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24 (ported from ANZDisable.ps1)
    Requires: ExchangeOnlineManagement, and Microsoft.Online.SharePoint.
              PowerShell if the OneDrive step is used (see UNVERIFIED note
              above) -- connected before running, in the same PowerShell
              session/process this script runs in.
#>
param(
    [string]$SmtpAddress,
    [int]$Port = 8790,
    [switch]$NoBrowser
)

. (Join-Path $PSScriptRoot 'M365WebCommon.ps1')

# REVISED 2026-07-24 -- the logic that used to live here MOVED, unchanged,
# into M365Logic.psm1, which M365WebCommon.ps1 imports (so it is already
# available below with nothing further to import here). This file is now a
# pure browser UI layer: build the page, call module functions, render what
# comes back. The WinForms tools call the exact same functions -- one
# implementation, two front ends, no duplicated logic to drift apart.
# Moved out of this file: ConvertTo-NormalizedUserList, Get-OneDrivePersonalSiteUrl, Disable-M365User.

#region HTML page
function Get-IndexHtml {
    $css = Get-M365SharedCss
    $closeBtn = Get-M365CloseButtonHtml
    $prefill = if ($SmtpAddress) { $SmtpAddress } else { '' }
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>M365 Disable</title>
<style>
$css
  .out { min-height:200px; }
</style>
</head>
<body>
  <h1>M365 Disable</h1>
  <div class="warn">Connect Exchange Online AND SharePoint Online (Connect-SPOService) before running. License removal and sign-in block are NOT performed by this tool -- handle those manually (a reminder appears after every run listing exactly which users).</div>

  <div class="row">
    <label style="flex:1;">User(s) to disable -- one per line (commas/semicolons also tolerated):<br>
      <textarea id="users" rows="6" style="width:100%;">$prefill</textarea>
    </label>
  </div>

  <div class="row">
    <label>Manager (optional):<br><input id="manager" type="text" style="width:240px"></label>
    <label>SharePoint tenant prefix (needed only for OneDrive access -- e.g. "contoso" if your OneDrive URLs look like https://contoso-my.sharepoint.com/...):<br>
      <input id="spoPrefix" type="text" style="width:260px">
    </label>
  </div>

  <fieldset id="managerAccessFs" disabled>
    <legend>Manager Access (only applies if a manager is entered above)</legend>
    <label><input type="radio" name="mgrAccess" value="both" checked> Mailbox &amp; OneDrive</label>
    <label style="margin-left:12px;"><input type="radio" name="mgrAccess" value="mailbox"> Mailbox Only</label>
    <label style="margin-left:12px;"><input type="radio" name="mgrAccess" value="onedrive"> OneDrive Only</label>
  </fieldset>

  <div class="row" style="margin-top:12px;">
    <button id="disableBtn" class="danger">Disable</button>
    <button id="clearBtn" class="secondary">Clear</button>
    $closeBtn
  </div>

  <div id="resultLabel" class="status">Result:</div>
  <div id="out" class="out">(nothing run yet)</div>

<script>
  const usersEl = document.getElementById('users');
  const managerEl = document.getElementById('manager');
  const spoEl = document.getElementById('spoPrefix');
  const fsEl = document.getElementById('managerAccessFs');
  const outEl = document.getElementById('out');
  const labelEl = document.getElementById('resultLabel');

  function updateManagerAccessEnabled() {
    fsEl.disabled = (managerEl.value.trim() === '');
  }
  managerEl.addEventListener('input', updateManagerAccessEnabled);
  updateManagerAccessEnabled();

  document.getElementById('clearBtn').addEventListener('click', () => {
    usersEl.value = '';
    managerEl.value = '';
    spoEl.value = '';
    document.querySelector('input[name=mgrAccess][value=both]').checked = true;
    updateManagerAccessEnabled();
    labelEl.textContent = 'Result:';
    outEl.textContent = '(nothing run yet)';
  });

  document.getElementById('disableBtn').addEventListener('click', async () => {
    const raw = usersEl.value;
    if (!raw.trim()) { outEl.textContent = 'Enter at least one user to disable first.'; return; }
    const manager = managerEl.value.trim();
    const mgrAccess = document.querySelector('input[name=mgrAccess]:checked').value;
    const grantMailbox = manager !== '' && (mgrAccess === 'both' || mgrAccess === 'mailbox');
    const grantOneDrive = manager !== '' && (mgrAccess === 'both' || mgrAccess === 'onedrive');
    const spoPrefix = spoEl.value.trim();
    if (grantOneDrive && !spoPrefix) {
      outEl.textContent = 'Manager Access includes OneDrive, but no SharePoint tenant prefix was entered. Fill that in, or switch Manager Access to Mailbox Only.';
      return;
    }

    const userLines = raw.split(/[\\r\\n,;]+/).map(s => s.trim()).filter(s => s !== '');
    let accessNote;
    if (manager) {
      const bits = [];
      if (grantMailbox) bits.push('mailbox Full Access (no Send-As)');
      if (grantOneDrive) bits.push('OneDrive Site Collection Admin');
      accessNote = "Manager '" + manager + "' will be granted: " + bits.join(' and ') + '.';
    } else {
      accessNote = 'No manager entered -- no Full Access or OneDrive access will be granted to anyone.';
    }
    const confirmMsg = 'Disable ' + userLines.length + ' user(s)?\\n\\n' + userLines.join('\\n') +
      '\\n\\nFor each: hide from address lists, convert to shared mailbox.' +
      '\\n\\nNOTE: license removal and sign-in block are NOT performed by this tool -- handle those manually.' +
      '\\n\\n' + accessNote;
    if (!confirm(confirmMsg)) { outEl.textContent = 'Cancelled.'; return; }

    document.getElementById('disableBtn').disabled = true;
    labelEl.textContent = 'Running...';
    outEl.textContent = '';
    try {
      const res = await fetch('/api/disable', {
        method: 'POST', headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ users: raw, manager, grantMailbox, grantOneDrive, spoPrefix })
      });
      const data = await res.json();
      if (!data.ok) { labelEl.textContent = 'Failed:'; outEl.textContent = data.error; return; }
      labelEl.textContent = 'Done -- processed ' + data.processed.length + ' user(s).';
      outEl.textContent = data.text;
      alert('Complete:\\n\\nManually remove from license group and block sign-in for these users.\\n\\n' + data.processed.join('\\n'));
    } catch (err) {
      labelEl.textContent = 'Failed:';
      outEl.textContent = 'Request failed: ' + err;
    } finally {
      document.getElementById('disableBtn').disabled = false;
    }
  });
</script>
</body>
</html>
"@
}
#endregion

#region Routes
$routes = @{
    'POST /api/disable' = {
        param($Request)
        $result = [ordered]@{ ok = $false; text = ''; processed = @(); error = '' }
        try {
            $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
            $body = $reader.ReadToEnd() | ConvertFrom-Json

            $users = ConvertTo-NormalizedUserList -Text $body.users
            if ($users.Count -eq 0) { throw "Enter at least one user to disable first." }
            $userSmtps = @($users | ForEach-Object { Get-TargetSmtp -Text $_ })

            $managerSmtp = $null
            if ($body.manager) { $managerSmtp = Get-TargetSmtp -Text $body.manager }

            $out = New-Object System.Text.StringBuilder
            foreach ($smtp in $userSmtps) {
                [void]$out.Append((Disable-M365User -Smtp $smtp -ManagerSmtp $managerSmtp `
                    -GrantMailboxAccess ([bool]$body.grantMailbox) -GrantOneDriveAccess ([bool]$body.grantOneDrive) `
                    -SpoTenantPrefix $body.spoPrefix))
                [void]$out.AppendLine("")
            }
            $result.ok        = $true
            $result.text      = $out.ToString()
            $result.processed = $userSmtps
        } catch {
            $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'M365 Disable'
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 6 -Compress) }
    }
}

Wait-M365ExoConnection -AlsoNote ' (and Connect-SPOService, if you will use the OneDrive step)'
Start-M365WebTool -ToolName 'M365 Disable' -Port $Port -GetHtml ${function:Get-IndexHtml} -Routes $routes -NoBrowser:$NoBrowser
#endregion
