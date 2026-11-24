#Requires -Version 7.0
<#
.SYNOPSIS
    Tenant-agnostic, browser-based Message Trace / Quarantine Release tool --
    ported from ANZMessageTrace.ps1 (via M365_MailboxHealth_Web.ps1's shared
    web-hosting pattern), with every tenant-specific reference removed.
.DESCRIPTION
    ADDED 2026-07-24 -- part of the same batch as M365_MailboxPermissions_Web.
    ps1, M365_Disable_Web.ps1, and M365_AdminConsole_Web.ps1. Uses
    M365WebCommon.ps1's Start-M365WebTool for the listener/routing/browser-
    launch boilerplate (see that file's header), which also imports
    M365Logic.psm1 for Invoke-EXOCommand / Format-ErrorDetail /
    Write-CompletedLog / Get-TargetSmtp.

    WHAT'S IDENTICAL TO ANZMessageTrace.ps1's LOGIC (as of 2026-07-24 those
    functions live in M365Logic.psm1, shared with every other front end,
    rather than being copied into this file):
      - Invoke-MessageTrace / Invoke-MessageTraceDetail / Find-QuarantineMatch
        / Invoke-QuarantineRelease -- the same four EXO-only functions,
        same parameters, same Invoke-EXOCommand wrapping.
      - Same field set per result (Received, SenderAddress, RecipientAddress,
        Subject, Status, MessageTraceId, MessageId) and the same "From/To
        both independently optional, leave both blank to trace everything in
        the date range" behavior the original settled on after dropping its
        old single-mailbox "Direction" radio.
      - Same quarantine flow: look up matches for a trace result via its
        Message-ID, confirm, release one at a time, log each release.

    WHAT'S DIFFERENT, beyond the WinForms -> browser swap itself:
      - Results grid becomes an HTML table; the "Filter loaded results" box
        and the timezone dropdown are now pure client-side JavaScript over
        the already-fetched JSON array -- genuinely no new server round trip
        either way, same as the original's own filter box already did (it
        also never re-queried EXO), just implemented where it naturally
        belongs now that there's a real client-side layer to put it in.
        Column headers are click-to-sort (ascending/descending toggle),
        which the original's DataGridView got for free from
        DataGridViewColumnSortMode.Automatic -- reimplemented in JS here.
      - Timezone conversion moved client-side too (Intl.DateTimeFormat with
        IANA zone names -- America/New_York, America/Chicago, America/Denver,
        America/Los_Angeles, UTC -- mapped from the same five labels the
        original's Windows time-zone-ID dropdown used) so switching the
        dropdown re-renders instantly against data already in the browser,
        matching the original's instant re-render on
        $cboTimeZone.Add_SelectedIndexChanged with no new EXO call.
      - Export to CSV is now a client-side Blob download of whatever's
        currently in the table (filtered or not) -- there's no
        SaveFileDialog equivalent in a browser, so this is the closest
        honest analog; it preserves the original's "export what's on
        screen, not the full unfiltered set" behavior.
      - Quarantine release renders each match as its own small card with its
        own "Release" button and its own confirm() -- same one-at-a-time,
        confirm-per-match flow as the original's foreach loop over $matches,
        just laid out as cards instead of sequential MessageBox popups.
      - Includes M365WebCommon.ps1's "Close & Return to Console" button.
    -SmtpAddress remains a purely optional convenience prefill into the From
    field, same as the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24 (ported from ANZMessageTrace.ps1)
#>
param(
    [string]$SmtpAddress,
    [int]$Port = 8788,
    [switch]$NoBrowser
)

. (Join-Path $PSScriptRoot 'M365WebCommon.ps1')

# REVISED 2026-07-24 -- the logic that used to live here MOVED, unchanged,
# into M365Logic.psm1, which M365WebCommon.ps1 imports (so it is already
# available below with nothing further to import here). This file is now a
# pure browser UI layer: build the page, call module functions, render what
# comes back. The WinForms tools call the exact same functions -- one
# implementation, two front ends, no duplicated logic to drift apart.
# Moved out of this file: Invoke-MessageTrace, Invoke-MessageTraceDetail, Find-QuarantineMatch,
# Invoke-QuarantineRelease.

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
<title>M365 Message Trace</title>
<style>
$css
  #traceTable th, #traceTable td { white-space:nowrap; }
  #traceTable td.subject { white-space:normal; max-width:320px; }
  /* .card and .status/.status.err moved into Get-M365SharedCss 2026-07-24 --
     other pages needed them and didn't have them. Local copies removed so
     there's one definition of each. */
</style>
</head>
<body>
  <h1>M365 Message Trace</h1>
  <div class="warn">Connect Exchange Online (Connect-ExchangeOnline) in the PowerShell window running this server before tracing. Read-only -- no objects are modified except an explicit quarantine release below.</div>

  <div class="row">
    <label>From:<br><input id="from" type="text" value="$prefill" style="width:220px"></label>
    <label>To:<br><input id="to" type="text" style="width:220px"></label>
    <label>Subject contains:<br><input id="subject" type="text" style="width:220px"></label>
    <label>Status:<br>
      <select id="status">
        <option>All Messages</option><option>Delivered</option><option>Failed</option>
        <option>Quarantined</option><option>FilteredAsSpam</option><option>Pending</option>
        <option>Expanded</option><option>GettingStatus</option>
      </select>
    </label>
  </div>
  <div class="row">
    <label>Search From:<br><input id="startDate" type="date"></label>
    <label>Through:<br><input id="endDate" type="date"></label>
    <label>Show time in:<br>
      <select id="tz">
        <option>Eastern</option><option selected>Central</option><option>Mountain</option>
        <option>Pacific</option><option>UTC</option>
      </select>
    </label>
    <button id="traceBtn" style="align-self:flex-end;">Trace</button>
    $closeBtn
  </div>
  <div id="traceStatus" class="status">Leave From/To blank to trace all mail in the date range. Max 30 days back, 9-day window.</div>

  <div class="row">
    <label>Filter loaded results:<br><input id="filter" type="text" style="width:280px"></label>
    <button id="filterBtn" style="align-self:flex-end;">Search</button>
    <button id="clearFilterBtn" class="secondary" style="align-self:flex-end;">Clear Filter</button>
    <button id="exportBtn" class="secondary" style="align-self:flex-end;" disabled>Export to CSV</button>
  </div>

  <table id="traceTable">
    <thead><tr>
      <th data-col="Received">Received</th>
      <th data-col="SenderAddress">Sender</th>
      <th data-col="RecipientAddress">Recipient</th>
      <th data-col="Subject">Subject</th>
      <th data-col="Status">Status</th>
    </tr></thead>
    <tbody id="traceBody"></tbody>
  </table>

  <h3>Detail (click a row to select, double-click to load detail)</h3>
  <div id="detailOut" class="out" style="min-height:120px;">No row selected.</div>

  <div class="row">
    <button id="releaseBtn" class="danger" disabled>Look Up Quarantine for Selected Row</button>
  </div>
  <div id="quarantineOut"></div>

<script>
  const ZONE_MAP = { Eastern:'America/New_York', Central:'America/Chicago', Mountain:'America/Denver', Pacific:'America/Los_Angeles', UTC:'UTC' };
  let allResults = [];      // full fetched set
  let shownResults = [];    // currently filtered/sorted subset (what the table renders)
  let selected = null;      // currently selected row's data object
  let sortCol = null, sortAsc = true;
  let searchStart = null, searchEnd = null;

  const startEl = document.getElementById('startDate');
  const endEl = document.getElementById('endDate');
  const today = new Date();
  const isoDate = (d) => d.toISOString().slice(0,10);
  startEl.max = isoDate(today);
  startEl.min = isoDate(new Date(today.getTime() - 30*86400000));
  startEl.value = isoDate(new Date(today.getTime() - 7*86400000));
  function clampEnd() {
    const s = new Date(startEl.value);
    let maxEnd = new Date(s.getTime() + 9*86400000);
    if (maxEnd > today) maxEnd = today;
    endEl.min = startEl.value;
    endEl.max = isoDate(maxEnd);
    if (!endEl.value || new Date(endEl.value) > maxEnd || new Date(endEl.value) < s) endEl.value = isoDate(maxEnd);
  }
  startEl.addEventListener('change', clampEnd);
  clampEnd();

  function fmtTime(iso) {
    if (!iso) return '';
    const zone = ZONE_MAP[document.getElementById('tz').value] || 'UTC';
    try {
      return new Intl.DateTimeFormat('en-US', { timeZone: zone, dateStyle: 'short', timeStyle: 'medium' }).format(new Date(iso));
    } catch (e) { return iso; }
  }

  function renderTable() {
    const tbody = document.getElementById('traceBody');
    tbody.innerHTML = '';
    for (const r of shownResults) {
      const tr = document.createElement('tr');
      tr.dataset.key = r.MessageTraceId + '|' + r.RecipientAddress;
      tr.innerHTML = '<td>' + fmtTime(r.Received) + '</td><td>' + (r.SenderAddress||'') + '</td><td>' +
        (r.RecipientAddress||'') + '</td><td class="subject">' + (r.Subject||'') + '</td><td>' + (r.Status||'') + '</td>';
      tr.addEventListener('click', () => selectRow(r, tr));
      tr.addEventListener('dblclick', () => loadDetail(r));
      tbody.appendChild(tr);
    }
    document.getElementById('exportBtn').disabled = (shownResults.length === 0);
  }

  function selectRow(r, tr) {
    selected = r;
    document.querySelectorAll('#traceBody tr').forEach(x => x.style.background = '');
    tr.style.background = '#dbeafe';
    document.getElementById('releaseBtn').disabled = false;
  }

  document.querySelectorAll('#traceTable th').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      sortAsc = (sortCol === col) ? !sortAsc : true;
      sortCol = col;
      shownResults.sort((a, b) => {
        const av = (a[col] || ''), bv = (b[col] || '');
        return (av > bv ? 1 : av < bv ? -1 : 0) * (sortAsc ? 1 : -1);
      });
      renderTable();
    });
  });

  function applyFilter() {
    const needle = document.getElementById('filter').value.trim().toLowerCase();
    const statusEl = document.getElementById('traceStatus');
    if (!needle) {
      shownResults = allResults.slice();
      statusEl.className = 'status';
      statusEl.textContent = allResults.length + ' result(s).';
    } else {
      shownResults = allResults.filter(r =>
        (r.SenderAddress||'').toLowerCase().includes(needle) ||
        (r.RecipientAddress||'').toLowerCase().includes(needle) ||
        (r.Subject||'').toLowerCase().includes(needle) ||
        (r.Status||'').toLowerCase().includes(needle));
      statusEl.className = 'status';
      statusEl.textContent = shownResults.length + ' of ' + allResults.length + " result(s) match '" + needle + "'.";
    }
    renderTable();
  }
  document.getElementById('filterBtn').addEventListener('click', applyFilter);
  document.getElementById('filter').addEventListener('keydown', (e) => { if (e.key === 'Enter') applyFilter(); });
  document.getElementById('clearFilterBtn').addEventListener('click', () => {
    document.getElementById('filter').value = '';
    applyFilter();
  });
  document.getElementById('tz').addEventListener('change', renderTable);

  document.getElementById('traceBtn').addEventListener('click', async () => {
    const btn = document.getElementById('traceBtn');
    const statusEl = document.getElementById('traceStatus');
    btn.disabled = true;
    document.getElementById('releaseBtn').disabled = true;
    document.getElementById('detailOut').textContent = 'No row selected.';
    document.getElementById('quarantineOut').innerHTML = '';
    document.getElementById('filter').value = '';
    statusEl.className = 'status';
    statusEl.textContent = 'Running trace...';
    try {
      const params = new URLSearchParams({
        from: document.getElementById('from').value.trim(),
        to: document.getElementById('to').value.trim(),
        subject: document.getElementById('subject').value.trim(),
        status: document.getElementById('status').value,
        start: startEl.value,
        end: endEl.value
      });
      const res = await fetch('/api/trace?' + params.toString());
      const data = await res.json();
      if (!data.ok) { statusEl.className = 'status err'; statusEl.textContent = data.error; return; }
      allResults = data.results;
      shownResults = allResults.slice();
      searchStart = data.start; searchEnd = data.end;
      renderTable();
      statusEl.className = 'status';
      const statusNote = (document.getElementById('status').value !== 'All Messages') ? ('  Status: ' + document.getElementById('status').value + '.') : '';
      statusEl.textContent = allResults.length + ' result(s). ' + startEl.value + ' through ' + endEl.value + '.' + statusNote;
    } catch (err) {
      statusEl.className = 'status err';
      statusEl.textContent = 'Request failed: ' + err;
    } finally {
      btn.disabled = false;
    }
  });

  async function loadDetail(r) {
    const el = document.getElementById('detailOut');
    el.textContent = 'Loading...';
    try {
      const params = new URLSearchParams({ traceId: r.MessageTraceId, recipient: r.RecipientAddress, start: searchStart, end: searchEnd });
      const res = await fetch('/api/tracedetail?' + params.toString());
      const data = await res.json();
      if (!data.ok) { el.textContent = 'ERROR: ' + data.error; return; }
      if (data.rows.length === 0) { el.textContent = '(no detail records returned)'; return; }
      el.textContent = data.rows.map(d => fmtTime(d.Date) + '  [' + d.Event + ']  ' + d.Action + (d.Detail ? ('\\n    ' + d.Detail) : '')).join('\\n');
    } catch (err) {
      el.textContent = 'Request failed: ' + err;
    }
  }

  document.getElementById('releaseBtn').addEventListener('click', async () => {
    if (!selected) return;
    const out = document.getElementById('quarantineOut');
    out.innerHTML = 'Looking up quarantine entries...';
    try {
      const params = new URLSearchParams({ messageId: selected.MessageId || '', recipient: selected.RecipientAddress, start: searchStart, end: searchEnd });
      const res = await fetch('/api/quarantine?' + params.toString());
      const data = await res.json();
      if (!data.ok) { out.innerHTML = '<div class="status err">' + data.error + '</div>'; return; }
      if (data.matches.length === 0) {
        out.innerHTML = '<div class="status">No quarantine entry found (may not have been quarantined, may already be released, or may have expired).</div>';
        return;
      }
      out.innerHTML = '';
      data.matches.forEach((m, idx) => {
        const card = document.createElement('div');
        card.className = 'card';
        card.innerHTML = '<b>Type:</b> ' + m.Type + '  <b>Policy:</b> ' + m.PolicyName + '<br><b>Received:</b> ' + m.ReceivedTime +
          '  <b>Release status:</b> ' + m.ReleaseStatus + '<br><button data-idx="' + idx + '">Release this message</button> <span class="rstatus"></span>';
        const relBtn = card.querySelector('button');
        relBtn.addEventListener('click', async () => {
          if (!confirm('Release this message to ALL original recipients?')) return;
          relBtn.disabled = true;
          const rstatus = card.querySelector('.rstatus');
          rstatus.textContent = 'Releasing...';
          try {
            const relRes = await fetch('/api/quarantine/release', {
              method: 'POST', headers: {'Content-Type':'application/json'},
              body: JSON.stringify({ identity: m.Identity, sender: selected.SenderAddress, recipient: selected.RecipientAddress, subject: selected.Subject })
            });
            const relData = await relRes.json();
            rstatus.textContent = relData.ok ? 'Released.' : ('FAILED: ' + relData.error);
          } catch (err) {
            rstatus.textContent = 'Request failed: ' + err;
          }
        });
        out.appendChild(card);
      });
    } catch (err) {
      out.innerHTML = '<div class="status err">Request failed: ' + err + '</div>';
    }
  });

  document.getElementById('exportBtn').addEventListener('click', () => {
    if (shownResults.length === 0) return;
    const header = ['Received (' + document.getElementById('tz').value + ')','SenderAddress','RecipientAddress','Subject','Status','MessageTraceId'];
    const esc = (v) => '"' + String(v == null ? '' : v).replace(/"/g,'""') + '"';
    const lines = [header.map(esc).join(',')];
    for (const r of shownResults) {
      lines.push([fmtTime(r.Received), r.SenderAddress, r.RecipientAddress, r.Subject, r.Status, r.MessageTraceId].map(esc).join(','));
    }
    const blob = new Blob([lines.join('\\r\\n')], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'M365_MessageTrace_' + new Date().toISOString().replace(/[:T]/g,'').slice(0,13) + '.csv';
    a.click();
  });
</script>
</body>
</html>
"@
}
#endregion

#region Routes
$routes = @{
    'GET /api/trace' = {
        param($Request)
        $result = [ordered]@{ ok = $false; results = @(); start = $null; end = $null; error = '' }
        try {
            $q = $Request.QueryString
            $sd = [datetime]::Parse($q['start']).Date
            $ed = [datetime]::Parse($q['end']).Date.AddDays(1).AddSeconds(-1)
            $traceParams = @{ StartDate = $sd; EndDate = $ed }
            if ($q['from'])    { $traceParams.SenderAddress    = $q['from'] }
            if ($q['to'])      { $traceParams.RecipientAddress = $q['to'] }
            if ($q['subject']) { $traceParams.Subject          = $q['subject'] }
            if ($q['status'] -and $q['status'] -ne 'All Messages') { $traceParams.Status = $q['status'] }
            $rows = @(Invoke-MessageTrace @traceParams)
            $result.ok      = $true
            $result.start   = $sd.ToString('o')
            $result.end     = $ed.ToString('o')
            $result.results = @($rows | ForEach-Object {
                [ordered]@{
                    Received         = if ($_.Received) { ([datetime]$_.Received).ToUniversalTime().ToString('o') } else { $null }
                    SenderAddress    = $_.SenderAddress
                    RecipientAddress = $_.RecipientAddress
                    Subject          = $_.Subject
                    Status           = $_.Status
                    MessageTraceId   = $_.MessageTraceId.ToString()
                    MessageId        = $_.MessageId
                }
            })
        } catch {
            $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceV2'
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 6 -Compress) }
    }
    'GET /api/tracedetail' = {
        param($Request)
        $result = [ordered]@{ ok = $false; rows = @(); error = '' }
        try {
            $q = $Request.QueryString
            $details = Invoke-MessageTraceDetail -MessageTraceId ([guid]$q['traceId']) -RecipientAddress $q['recipient'] `
                -StartDate ([datetime]::Parse($q['start'])) -EndDate ([datetime]::Parse($q['end']))
            $result.ok   = $true
            $result.rows = @($details | ForEach-Object {
                [ordered]@{
                    Date   = if ($_.Date) { ([datetime]$_.Date).ToUniversalTime().ToString('o') } else { $null }
                    Event  = $_.Event
                    Action = $_.Action
                    Detail = $_.Detail
                }
            })
        } catch {
            $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceDetailV2'
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 6 -Compress) }
    }
    'GET /api/quarantine' = {
        param($Request)
        $result = [ordered]@{ ok = $false; matches = @(); error = '' }
        try {
            $q = $Request.QueryString
            $matches = @(Find-QuarantineMatch -MessageId $q['messageId'] -RecipientAddress $q['recipient'] `
                -StartReceivedDate ([datetime]::Parse($q['start'])) -EndReceivedDate ([datetime]::Parse($q['end'])))
            $result.ok      = $true
            $result.matches = @($matches | ForEach-Object {
                [ordered]@{
                    Identity      = $_.Identity.ToString()
                    Type          = $_.Type
                    PolicyName    = $_.PolicyName
                    ReceivedTime  = $_.ReceivedTime
                    ReleaseStatus = $_.ReleaseStatus
                }
            })
        } catch {
            $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-QuarantineMessage'
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 6 -Compress) }
    }
    'POST /api/quarantine/release' = {
        param($Request)
        $result = [ordered]@{ ok = $false; error = '' }
        try {
            $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
            $body = $reader.ReadToEnd() | ConvertFrom-Json
            Invoke-QuarantineRelease -QuarantineIdentity $body.identity
            Write-CompletedLog -Task "Released QuarantineMessage $($body.recipient) (sender: $($body.sender); subject: $($body.subject))"
            $result.ok = $true
        } catch {
            $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Quarantine release'
        }
        return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
    }
}

Wait-M365ExoConnection
Start-M365WebTool -ToolName 'M365 Message Trace' -Port $Port -GetHtml ${function:Get-IndexHtml} -Routes $routes -NoBrowser:$NoBrowser
#endregion
