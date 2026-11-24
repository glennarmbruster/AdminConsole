#Requires -Version 7.0
<#
.SYNOPSIS
    Tenant-agnostic, browser-based Mailbox Permissions (Full Access / Send As
    delegation) tool -- ported from ANZMailboxPermissions.ps1.
.DESCRIPTION
    ADDED 2026-07-24 -- part of the same batch as M365_MessageTrace_Web.ps1,
    M365_Disable_Web.ps1, and M365_AdminConsole_Web.ps1. Uses
    M365WebCommon.ps1's Start-M365WebTool, which also imports the logic
    module -- same as every other M365_*_Web tool.

    REVISED 2026-07-24 -- Get-MailboxState, Format-MailboxState, the four
    Grant-/Revoke- primitives, and the BEFORE/act/AFTER orchestration all
    MOVED into M365Logic.psm1 (the last one as
    Invoke-MailboxPermissionEdit). They were pure EXO logic with nothing
    WinForms- or web-specific about them, so both front ends now call the
    same copy instead of each carrying their own. Format-MailboxState's text
    output is still reused as-is for this page's result box, same "same
    report, different delivery mechanism" approach as
    M365_MailboxHealth_Web.ps1.

    Grant/Revoke Full Access and Send As are each their own POST route
    (/api/grant, /api/revoke with an "action" field) that reruns
    Get-MailboxState before and after, same BEFORE/AFTER pairing the
    WinForms version's Invoke-MailboxEdit built, and confirms client-side
    (a JS confirm()) before firing, same as the WinForms MessageBox
    confirmation it's replacing.

    No -Environment / -DomainController parameters -- this toolkit has no
    on-prem concept at all. -SmtpAddress is an optional convenience prefill
    into the mailbox field (there's no console-driven auto-launch-with-target
    the way the WinForms suite's AdminConsole did; this toolkit's console is
    a launcher, not a per-target dispatcher -- see M365_AdminConsole_Web.
    ps1's header), typed manually otherwise.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24 (ported from ANZMailboxPermissions.ps1)
#>
param(
    [string]$SmtpAddress,
    [int]$Port = 8789,
    [switch]$NoBrowser
)

. (Join-Path $PSScriptRoot 'M365WebCommon.ps1')

# REVISED 2026-07-24 -- the logic that used to live here MOVED, unchanged,
# into M365Logic.psm1, which M365WebCommon.ps1 imports (so it is already
# available below with nothing further to import here). This file is now a
# pure browser UI layer: build the page, call module functions, render what
# comes back. The WinForms tools call the exact same functions -- one
# implementation, two front ends, no duplicated logic to drift apart.
# Moved out of this file: Get-MailboxState, Format-MailboxState, Grant-/Revoke-FullAccessPermission,
# Grant-/Revoke-SendAsPermission. Invoke-PermissionEdit (the BEFORE/act/AFTER
# orchestration) also moved, as the module's Invoke-MailboxPermissionEdit.

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
<title>M365 Mailbox Permissions</title>
<style>
$css
  .out { min-height:260px; }
</style>
</head>
<body>
  <h1>M365 Mailbox Permissions</h1>
  <div class="warn">Connect Exchange Online (Connect-ExchangeOnline) in the PowerShell window running this server before running.</div>

  <div class="row">
    <label>Mailbox:<br><input id="smtp" type="text" value="$prefill" style="width:280px"></label>
    <button id="lookupBtn" style="align-self:flex-end;">Look Up</button>
    $closeBtn
  </div>

  <div id="out" class="out">Enter a mailbox above and click Look Up.</div>

  <fieldset style="margin-top:12px;">
    <legend>Mailbox Permissions</legend>
    <div class="row">
      <label>Grantee (email or logon):<br><input id="grantee" type="text" style="width:260px"></label>
    </div>
    <div class="row">
      <button id="faGrant" disabled>Grant Full Access</button>
      <button id="faRevoke" class="secondary" disabled>Revoke Full Access</button>
      <button id="saGrant" disabled>Grant Send As</button>
      <button id="saRevoke" class="secondary" disabled>Revoke Send As</button>
    </div>
  </fieldset>

<script>
  const smtpEl = document.getElementById('smtp');
  const outEl = document.getElementById('out');
  const granteeEl = document.getElementById('grantee');
  const actionBtns = ['faGrant','faRevoke','saGrant','saRevoke'].map(id => document.getElementById(id));

  async function lookup() {
    const smtp = smtpEl.value.trim();
    if (!smtp) { outEl.textContent = 'Enter a mailbox first.'; return; }
    outEl.textContent = 'Loading...';
    actionBtns.forEach(b => b.disabled = true);
    try {
      const res = await fetch('/api/state?smtp=' + encodeURIComponent(smtp));
      const data = await res.json();
      if (!data.ok) { outEl.textContent = 'ERROR: ' + data.error; return; }
      outEl.textContent = data.text;
      actionBtns.forEach(b => b.disabled = false);
    } catch (err) {
      outEl.textContent = 'Request failed: ' + err;
    }
  }
  document.getElementById('lookupBtn').addEventListener('click', lookup);
  smtpEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') lookup(); });

  async function doEdit(action, label) {
    const smtp = smtpEl.value.trim();
    const grantee = granteeEl.value.trim();
    if (!smtp) { outEl.textContent = 'Enter a mailbox first.'; return; }
    if (!grantee) { outEl.textContent = 'Enter a grantee (email or logon) first.'; return; }
    if (!confirm(label + " for '" + grantee + "' on mailbox '" + smtp + "'?")) return;
    actionBtns.forEach(b => b.disabled = true);
    outEl.textContent = 'Working...';
    try {
      const res = await fetch('/api/' + action, {
        method: 'POST', headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ smtp, grantee })
      });
      const data = await res.json();
      outEl.textContent = data.ok ? data.text : ('ERROR: ' + data.error);
    } catch (err) {
      outEl.textContent = 'Request failed: ' + err;
    } finally {
      actionBtns.forEach(b => b.disabled = false);
    }
  }
  document.getElementById('faGrant').addEventListener('click', () => doEdit('grant-full', 'Grant Full Access'));
  document.getElementById('faRevoke').addEventListener('click', () => doEdit('revoke-full', 'Revoke Full Access'));
  document.getElementById('saGrant').addEventListener('click', () => doEdit('grant-sendas', 'Grant Send As'));
  document.getElementById('saRevoke').addEventListener('click', () => doEdit('revoke-sendas', 'Revoke Send As'));
</script>
</body>
</html>
"@
}
#endregion

#region Routes
# The local Invoke-PermissionEdit that used to sit here was moved into
# M365Logic.psm1 as Invoke-MailboxPermissionEdit (identical
# BEFORE/act/AFTER/log orchestration) so the WinForms side can use it too.
# The four routes below call the module's copy.

$routes = @{
    'GET /api/state' = {
        param($Request)
        $result = [ordered]@{ ok = $false; text = ''; error = '' }
        $smtp = $Request.QueryString['smtp']
        try {
            $state = Get-MailboxState -Smtp $smtp
            $result.ok   = $true
            $result.text = Format-MailboxState -Heading "CURRENT STATE" -State $state
        } catch {
            $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Permissions'
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
    }
    'POST /api/grant-full' = {
        param($Request)
        $result = [ordered]@{ ok = $false; text = ''; error = '' }
        try {
            $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
            $body = $reader.ReadToEnd() | ConvertFrom-Json
            $result.text = Invoke-MailboxPermissionEdit -Smtp $body.smtp -Grantee $body.grantee -ActionLabel 'Grant Full Access' -Verb 'Granted FullAccess' -Action {
                Grant-FullAccessPermission -Smtp $body.smtp -Grantee $body.grantee
            }
            $result.ok = $true
        } catch {
            $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Grant Full Access'
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
    }
    'POST /api/revoke-full' = {
        param($Request)
        $result = [ordered]@{ ok = $false; text = ''; error = '' }
        try {
            $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
            $body = $reader.ReadToEnd() | ConvertFrom-Json
            $result.text = Invoke-MailboxPermissionEdit -Smtp $body.smtp -Grantee $body.grantee -ActionLabel 'Revoke Full Access' -Verb 'Revoked FullAccess' -Action {
                Revoke-FullAccessPermission -Smtp $body.smtp -Grantee $body.grantee
            }
            $result.ok = $true
        } catch {
            $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Revoke Full Access'
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
    }
    'POST /api/grant-sendas' = {
        param($Request)
        $result = [ordered]@{ ok = $false; text = ''; error = '' }
        try {
            $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
            $body = $reader.ReadToEnd() | ConvertFrom-Json
            $result.text = Invoke-MailboxPermissionEdit -Smtp $body.smtp -Grantee $body.grantee -ActionLabel 'Grant Send As' -Verb 'Granted SendAs' -Action {
                Grant-SendAsPermission -Smtp $body.smtp -Grantee $body.grantee
            }
            $result.ok = $true
        } catch {
            $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Grant Send As'
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
    }
    'POST /api/revoke-sendas' = {
        param($Request)
        $result = [ordered]@{ ok = $false; text = ''; error = '' }
        try {
            $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
            $body = $reader.ReadToEnd() | ConvertFrom-Json
            $result.text = Invoke-MailboxPermissionEdit -Smtp $body.smtp -Grantee $body.grantee -ActionLabel 'Revoke Send As' -Verb 'Revoked SendAs' -Action {
                Revoke-SendAsPermission -Smtp $body.smtp -Grantee $body.grantee
            }
            $result.ok = $true
        } catch {
            $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Revoke Send As'
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
    }
}

Wait-M365ExoConnection
Start-M365WebTool -ToolName 'M365 Mailbox Permissions' -Port $Port -GetHtml ${function:Get-IndexHtml} -Routes $routes -NoBrowser:$NoBrowser
#endregion
