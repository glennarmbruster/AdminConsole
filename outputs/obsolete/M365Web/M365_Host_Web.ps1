#Requires -Version 5.1
<#
    M365_Host_Web.ps1
    M365 Admin Host -- ONE long-running process that connects to Exchange
    Online (and optionally SharePoint Online and a directory module) once,
    and serves every tool from a single port.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

param(
    [int]$Port = 8787,
    [switch]$NoBrowser,

    [switch]$EnableConsole,

    [switch]$ShowDirectoryPill,

    [string]$SuiteName = 'M365',

    [string]$ExpectedDomain,

    [string]$LogFileName = 'M365-CompletedTasks.csv'
)

. (Join-Path $PSScriptRoot 'M365WebCommon.ps1')

Set-M365SuiteIdentity -SuiteName $SuiteName -LogFileName $LogFileName -ExpectedDomain $ExpectedDomain
if ($ExpectedDomain) {
    Write-Host "Suite: $SuiteName   (expecting tenant $ExpectedDomain)" -ForegroundColor Cyan
} else {
    Write-Host "Suite: $SuiteName   (tenant-agnostic -- works against whatever is connected)" -ForegroundColor Cyan
}

#region View loading ---------------------------------------------------------------
$viewFiles = @(
    'M365_View_MailboxHealth.ps1'
    'M365_View_MessageTrace.ps1'
    'M365_View_MailboxAudit.ps1'
    'M365_View_SharedMailboxAudit.ps1'
    'M365_View_InboxRules.ps1'
    'M365_View_MailboxPermissions.ps1'
    'M365_View_CalendarPermissions.ps1'
    'M365_View_OutOfOffice.ps1'
    'M365_View_MailboxOneDriveAccess.ps1'
    'M365_View_Disable.ps1'
)

if ($EnableConsole) {
    $viewFiles += 'M365_View_PowerShellConsole.ps1'
}

foreach ($vf in $viewFiles) {
    $path = Join-Path $PSScriptRoot $vf
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Host "View file missing, skipping: $vf" -ForegroundColor Yellow
        continue
    }
    . $path
}

if ($EnableConsole) {
    Write-Host ""
    Write-Host "PowerShell Console view is ENABLED. This exposes arbitrary code execution in this process," -ForegroundColor Red
    Write-Host "with its live Exchange/SPO/directory connections, to anything that can present this run's" -ForegroundColor Red
    Write-Host "API token. Re-run without -EnableConsole to turn it off." -ForegroundColor Red
    Write-Host ""
}

$views = Get-M365Views
if (@($views).Count -eq 0) {
    Write-Host "No views loaded -- there is nothing to serve. Check that the M365_View_*.ps1 files are in '$PSScriptRoot'." -ForegroundColor Red
    return
}
Write-Host "Loaded $(@($views).Count) view(s): $(($views | ForEach-Object { $_.Name }) -join ', ')" -ForegroundColor Gray
#endregion

#region Route composition ----------------------------------------------------------
$allRoutes = @{}
foreach ($v in $views) {
    foreach ($key in $v.Routes.Keys) {
        if ($allRoutes.ContainsKey($key)) {
            Write-Host "Duplicate route '$key' -- view '$($v.Name)' collides with an earlier view. Give each view its own /api/<tool>/... prefix." -ForegroundColor Red
            return
        }
        $allRoutes[$key] = $v.Routes[$key]
    }
}
Write-Host "Composed $($allRoutes.Count) API route(s)." -ForegroundColor Gray
#endregion

#region Page composition -----------------------------------------------------------
function Get-HostPageHtml {
    $css = Get-M365SharedCss
    $views = Get-M365Views
    $token = Get-M365RequestToken

    $aadPill = if ($ShowDirectoryPill) {
        '    <span class="connpill" id="pill-aad" title="Directory (AzureAD / Graph / Entra)">AzureAD ...</span>'
    } else { '' }

    $cards = ($views | ForEach-Object {
        $cls = if ($_.Destructive) { 'card destructive' } else { 'card' }
        @"
      <div class="$cls">
        <h3>$($_.Name)</h3>
        <div class="desc">$($_.Description)</div>
        <button class="navbtn" data-target="$($_.Id)">Open</button>
      </div>
"@
    }) -join "`n"

    $sections = ($views | ForEach-Object {
        $frag = & $_.GetHtml
        @"
    <section class="view" id="view-$($_.Id)">
      <div class="row" style="margin-bottom:14px;">
        <button class="navbtn secondary" data-target="home">&larr; Back to Console</button>
      </div>
$frag
    </section>
"@
    }) -join "`n"

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>M365 Admin Host</title>
<style>
$css
</style>
</head>
<body>
  <div class="topbar">
    <span class="brand">$SuiteName Admin Host</span>
    <span class="connpill" id="pill-exo" title="Exchange Online">ExO ...</span>
    <span class="connpill" id="pill-spo" title="SharePoint Online">SPO ...</span>
    
    <span class="connpill" id="pill-tenant" title="Connected tenant">tenant ...</span>
$aadPill
    <button id="connRefresh" class="secondary" style="padding:4px 10px;font-size:12px;">Refresh</button>
    <span class="spacer"></span>
    <button class="navbtn secondary" data-target="home">Console Home</button>
    <button id="shutdownBtn" class="danger">Shut Down</button>
  </div>

  <div id="tenantWarn" style="display:none;"></div>

  <div class="wrap">
    <section class="view active" id="view-home">
      <h1>Console</h1>
      <div class="note">
        Every tool below runs in this one process and shares its single Exchange Online connection --
        connect once and they all work. The host handles one request at a time, so a long-running
        trace will make other views wait rather than running in parallel.
      </div>
$cards
    </section>

$sections
  </div>

<script>
  const M365_TOKEN = '$token';
  (function () {
    const rawFetch = window.fetch.bind(window);
    window.fetch = function (input, init) {
      init = init || {};
      let url = (typeof input === 'string') ? input : (input && input.url) || '';
      const isApi = url.startsWith('/api/') || url.indexOf(window.location.origin + '/api/') === 0;
      if (isApi) {
        const h = new Headers(init.headers || (typeof input !== 'string' && input.headers) || {});
        h.set('X-M365-Token', M365_TOKEN);
        init.headers = h;
      }
      return rawFetch(input, init);
    };
  })();

  function showView(id) {
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    const el = document.getElementById('view-' + id);
    if (el) el.classList.add('active');
    window.scrollTo(0, 0);
  }
  document.addEventListener('click', (e) => {
    const btn = e.target.closest('.navbtn');
    if (!btn) return;
    showView(btn.dataset.target);
  });

  function paintPill(el, label, state, detail) {
    el.classList.remove('ok', 'bad');
    if (state === 'ok')       { el.classList.add('ok');  el.textContent = label + ' Connected'; }
    else if (state === 'bad') { el.classList.add('bad'); el.textContent = label + ' NotConnected'; }
    else                      {                          el.textContent = label + ' ?'; }
    el.title = detail || '';
  }

  async function refreshConn(full) {
    const exo = document.getElementById('pill-exo');
    const spo = document.getElementById('pill-spo');
    const aad = document.getElementById('pill-aad');
    const ten = document.getElementById('pill-tenant');
    try {
      const res  = await fetch('/api/connection' + (full ? '?full=1' : ''));
      const data = await res.json();

      paintPill(exo, 'ExO', data.exo.connected ? 'ok' : 'bad', data.exo.detail);

      const tenantName = data.tenantDomain || data.exo.organization || '';
      ten.classList.remove('ok', 'bad');
      if (tenantName) {
        ten.classList.add('ok');
        ten.textContent = tenantName;
        const alsoOrg = (data.exo.organization && data.tenantDomain &&
                         data.exo.organization !== data.tenantDomain)
                      ? '  (organization: ' + data.exo.organization + ')' : '';
        ten.title = 'Connected tenant: ' + tenantName + alsoOrg;
      } else {
        ten.textContent = 'tenant unknown';
        ten.title = 'Tenant unknown -- connect Exchange Online, then click Refresh.';
      }

      paintPill(spo, 'SPO',
        data.spo.skipped ? 'unknown' : (data.spo.connected ? 'ok' : 'bad'),
        data.spo.skipped ? 'Not checked on this pass -- click Refresh (checking SharePoint costs a service call).' : data.spo.detail);

      if (aad && data.directoryMonitored) {
        const aadLabel = (data.aad.provider && !data.aad.skipped) ? data.aad.provider : 'AzureAD';
        paintPill(aad, aadLabel,
          data.aad.skipped ? 'unknown' : (data.aad.connected ? 'ok' : 'bad'),
          data.aad.skipped ? 'Not checked on this pass -- click Refresh.' : data.aad.detail);
      }

      const tw = document.getElementById('tenantWarn');
      if (data.tenant && data.tenant.armed && !data.tenant.match) {
        tw.style.display = 'block';
        tw.innerHTML = '';
        const w = document.createElement('div');
        w.className = 'warn';
        w.style.margin = '10px 24px 0 24px';
        w.textContent = data.tenant.detail;
        tw.appendChild(w);
      } else {
        tw.style.display = 'none';
      }

    } catch (err) {
      [exo, spo, aad, ten].forEach(p => {
        if (!p) return;
        p.classList.remove('ok');
        p.classList.add('bad');
      });
      exo.textContent = 'host unreachable';
      exo.title = String(err);
      if (ten) { ten.textContent = 'tenant unknown'; ten.title = String(err); }
    }
  }

  document.getElementById('connRefresh').addEventListener('click', () => refreshConn(true));
  refreshConn(true);
  setInterval(() => refreshConn(false), 60000);

  document.getElementById('shutdownBtn').addEventListener('click', async () => {
    // GUARD: confirmation is required here. Do not remove.
    if (!confirm('Shut down the host? Every tool stops and you will need to re-run the script (and reconnect) to use them again.')) return;
    try { await fetch('/api/shutdown', { method: 'POST' }); } catch (e) {}
    document.body.innerHTML =
      '<div class="wrap"><h1>Host stopped.</h1><div class="note">You can close this tab. ' +
      'Re-run ./M365_Host_Web.ps1 to start it again.</div></div>';
  });
</script>
</body>
</html>
"@
}
#endregion

#region Start ----------------------------------------------------------------------
Wait-M365ExoConnection -AlsoNote ' (and Connect-SPOService too, if you plan to use the Disable view''s OneDrive step)'

Start-M365WebHost -HostName 'M365 Admin Host' -Port $Port `
    -GetHtml ${function:Get-HostPageHtml} -Routes $allRoutes -NoBrowser:$NoBrowser `
    -ShowDirectoryPill:$ShowDirectoryPill
#endregion
