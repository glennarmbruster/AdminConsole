#Requires -Version 5.1
<#
    M365_View_MessageTrace.ps1
    Message Trace / Quarantine Release view for M365_Host_Web.ps1.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

Register-M365View -Id 'trace' -Name 'Message Trace' `
    -Description 'Trace mail by sender, recipient, subject, or status over a date range. View per-message detail and release from quarantine.' `
    -GetHtml {
        return @'
  <h2>Message Trace</h2>
  <div class="note">
    Leave From and To both blank to trace all mail in the range. Max 30 days back; the window is capped at 9 days.
    Read-only except an explicit quarantine release.
    <br><br>
    <strong>Need more than a trace?</strong> Message trace answers "what happened to this message." For threat
    hunting -- URL/attachment verdicts, campaign views, saved queries, bulk remediation -- use Threat Explorer
    instead. It needs Defender for Office 365 Plan 2.
    <a id="tr-explorer-link" href="#" target="_blank" rel="noopener"
       style="font-weight:600;color:#1d4ed8;">Open Threat Explorer ↗</a>
  </div>

  <div class="row">
    <label>From:<br><input id="tr-from" type="text" style="width:200px"></label>
    <label>To:<br><input id="tr-to" type="text" style="width:200px"></label>
    <label>Subject contains:<br><input id="tr-subject" type="text" style="width:200px"></label>
    <label>Status:<br>
      <select id="tr-status">
        <option>All Messages</option><option>Delivered</option><option>Failed</option>
        <option>Quarantined</option><option>FilteredAsSpam</option><option>Pending</option>
        <option>Expanded</option><option>GettingStatus</option>
      </select>
    </label>
  </div>
  <div class="row">
    <label>Search From:<br><input id="tr-start" type="date"></label>
    <label>Through:<br><input id="tr-end" type="date"></label>
    <label>Show time in:<br>
      <select id="tr-tz">
        <option>Eastern</option><option selected>Central</option><option>Mountain</option>
        <option>Pacific</option><option>UTC</option>
      </select>
    </label>
    <button id="tr-run">Trace</button>
  </div>
  <div id="tr-status-line" class="status">No trace run yet.</div>

  <div class="row">
    <label>Filter loaded results:<br><input id="tr-filter" type="text" style="width:260px"></label>
    <button id="tr-filter-btn">Search</button>
    <button id="tr-filter-clear" class="secondary">Clear Filter</button>
    <button id="tr-export" class="secondary" disabled>Export to CSV</button>
  </div>

  <table id="tr-table">
    <thead><tr>
      <th data-col="Received">Received</th>
      <th data-col="SenderAddress">Sender</th>
      <th data-col="RecipientAddress">Recipient</th>
      <th data-col="Subject">Subject</th>
      <th data-col="Status">Status</th>
    </tr></thead>
    <tbody id="tr-body"></tbody>
  </table>

  <h3 style="margin-top:16px;">Detail</h3>
  <div class="status">Click a row to select it; double-click to load its delivery detail.</div>
  <div id="tr-detail" class="out" style="min-height:110px;">No row selected.</div>

  <div class="row">
    <button id="tr-quarantine" class="danger" disabled>Look Up Quarantine for Selected Row</button>
    <button id="tr-copy-msgid" class="secondary" disabled>Copy Message-ID</button>
    <button id="tr-copy-sender" class="secondary" disabled>Copy Sender</button>
    <button id="tr-open-explorer" class="secondary" disabled>Open Threat Explorer ↗</button>
    <span id="tr-copy-note" class="status" style="margin:0;"></span>
  </div>
  <div id="tr-quarantine-out"></div>

<script>
(function () {
  const ZONE = { Eastern:'America/New_York', Central:'America/Chicago', Mountain:'America/Denver', Pacific:'America/Los_Angeles', UTC:'UTC' };
  let all = [], shown = [], selected = null, sortCol = null, sortAsc = true;
  let searchStart = null, searchEnd = null;
  let portal = {};

  (async function loadPortalUrls() {
    try {
      const res = await fetch('/api/trace/portal-urls');
      portal = await res.json();
      const ex = document.getElementById('tr-explorer-link');
      if (portal.threatExplorer) ex.href = portal.threatExplorer;
    } catch (e) {  }
  })();

  const startEl  = document.getElementById('tr-start');
  const endEl    = document.getElementById('tr-end');
  const tzEl     = document.getElementById('tr-tz');
  const lineEl   = document.getElementById('tr-status-line');
  const bodyEl   = document.getElementById('tr-body');
  const exportBtn= document.getElementById('tr-export');
  const qBtn     = document.getElementById('tr-quarantine');
  const qOut     = document.getElementById('tr-quarantine-out');
  const detailEl = document.getElementById('tr-detail');
  const copyMsgIdBtn  = document.getElementById('tr-copy-msgid');
  const copySenderBtn = document.getElementById('tr-copy-sender');
  const explorerBtn   = document.getElementById('tr-open-explorer');
  const noteEl        = document.getElementById('tr-copy-note');

  async function copyToClip(text, label) {
    try {
      await navigator.clipboard.writeText(text);
      noteEl.className = 'status';
      noteEl.textContent = label + ' copied -- paste it into the portal’s filter.';
    } catch (err) {
      noteEl.className = 'status err';
      noteEl.textContent = 'Copy failed: ' + err;
    }
  }
  copyMsgIdBtn.addEventListener('click', () => {
    if (selected && selected.MessageId) copyToClip(selected.MessageId, 'Message-ID');
  });
  copySenderBtn.addEventListener('click', () => {
    if (selected && selected.SenderAddress) copyToClip(selected.SenderAddress, 'Sender');
  });

  explorerBtn.addEventListener('click', async () => {
    if (!selected) return;
    try {
      const params = new URLSearchParams({
        sender:    selected.SenderAddress || '',
        recipient: selected.RecipientAddress || '',
        messageId: selected.MessageId || ''
      });
      const res  = await fetch('/api/trace/portal-urls?' + params.toString());
      const data = await res.json();
      window.open(data.threatExplorer, '_blank', 'noopener');
      noteEl.className = 'status';
      if (data.threatExplorerPrepopulated) {
        noteEl.textContent = 'Threat Explorer opened with your filter applied.';
      } else {
        noteEl.textContent = 'Threat Explorer opened unfiltered (no prepopulating URL is configured) -- use Copy Sender / Copy Message-ID and paste into its filter.';
      }
    } catch (err) {
      noteEl.className = 'status err';
      noteEl.textContent = 'Could not open Threat Explorer: ' + err;
    }
  });

  const iso = (d) => d.toISOString().slice(0, 10);
  const today = new Date();
  startEl.max = iso(today);
  startEl.min = iso(new Date(today.getTime() - 30 * 86400000));
  startEl.value = iso(new Date(today.getTime() - 7 * 86400000));
  function clampEnd() {
    const s = new Date(startEl.value);
    let maxEnd = new Date(s.getTime() + 9 * 86400000);
    if (maxEnd > today) maxEnd = today;
    endEl.min = startEl.value;
    endEl.max = iso(maxEnd);
    if (!endEl.value || new Date(endEl.value) > maxEnd || new Date(endEl.value) < s) endEl.value = iso(maxEnd);
  }
  startEl.addEventListener('change', clampEnd);
  clampEnd();

  function fmtTime(isoStr) {
    if (!isoStr) return '';
    try {
      return new Intl.DateTimeFormat('en-US', { timeZone: ZONE[tzEl.value] || 'UTC', dateStyle:'short', timeStyle:'medium' }).format(new Date(isoStr));
    } catch (e) { return isoStr; }
  }

  function render() {
    bodyEl.innerHTML = '';
    for (const r of shown) {
      const tr = document.createElement('tr');
      const cells = [fmtTime(r.Received), r.SenderAddress || '', r.RecipientAddress || '', r.Subject || '', r.Status || ''];
      cells.forEach(v => { const td = document.createElement('td'); td.textContent = v; tr.appendChild(td); });
      tr.addEventListener('click', () => {
        selected = r;
        Array.from(bodyEl.children).forEach(x => x.classList.remove('selected'));
        tr.classList.add('selected');
        qBtn.disabled = false;
        copyMsgIdBtn.disabled = !r.MessageId;
        copySenderBtn.disabled = !r.SenderAddress;
        explorerBtn.disabled = false;
        noteEl.textContent = '';
      });
      tr.addEventListener('dblclick', () => loadDetail(r));
      bodyEl.appendChild(tr);
    }
    exportBtn.disabled = (shown.length === 0);
  }

  document.querySelectorAll('#tr-table th').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      sortAsc = (sortCol === col) ? !sortAsc : true;
      sortCol = col;
      shown.sort((a, b) => {
        const av = a[col] || '', bv = b[col] || '';
        return (av > bv ? 1 : av < bv ? -1 : 0) * (sortAsc ? 1 : -1);
      });
      render();
    });
  });

  function applyFilter() {
    const needle = document.getElementById('tr-filter').value.trim().toLowerCase();
    if (!needle) {
      shown = all.slice();
      lineEl.className = 'status';
      lineEl.textContent = all.length + ' result(s).';
    } else {
      shown = all.filter(r =>
        (r.SenderAddress || '').toLowerCase().includes(needle) ||
        (r.RecipientAddress || '').toLowerCase().includes(needle) ||
        (r.Subject || '').toLowerCase().includes(needle) ||
        (r.Status || '').toLowerCase().includes(needle));
      lineEl.className = 'status';
      lineEl.textContent = shown.length + ' of ' + all.length + " result(s) match '" + needle + "'.";
    }
    render();
  }
  document.getElementById('tr-filter-btn').addEventListener('click', applyFilter);
  document.getElementById('tr-filter').addEventListener('keydown', (e) => { if (e.key === 'Enter') applyFilter(); });
  document.getElementById('tr-filter-clear').addEventListener('click', () => {
    document.getElementById('tr-filter').value = '';
    applyFilter();
  });
  tzEl.addEventListener('change', render);

  document.getElementById('tr-run').addEventListener('click', async () => {
    const btn = document.getElementById('tr-run');
    btn.disabled = true;
    qBtn.disabled = true;
    copyMsgIdBtn.disabled = true;
    copySenderBtn.disabled = true;
    explorerBtn.disabled = true;
    noteEl.textContent = '';
    selected = null;
    detailEl.textContent = 'No row selected.';
    qOut.innerHTML = '';
    document.getElementById('tr-filter').value = '';
    lineEl.className = 'status';
    lineEl.textContent = 'Running trace... (the host handles one request at a time, so this blocks other views until it finishes)';
    try {
      const params = new URLSearchParams({
        from:    document.getElementById('tr-from').value.trim(),
        to:      document.getElementById('tr-to').value.trim(),
        subject: document.getElementById('tr-subject').value.trim(),
        status:  document.getElementById('tr-status').value,
        start:   startEl.value,
        end:     endEl.value
      });
      const res  = await fetch('/api/trace/search?' + params.toString());
      const data = await res.json();
      if (!data.ok) { lineEl.className = 'status err'; lineEl.textContent = data.error; return; }
      all = data.results; shown = all.slice();
      searchStart = data.start; searchEnd = data.end;
      render();
      const sNote = (document.getElementById('tr-status').value !== 'All Messages')
        ? '  Status: ' + document.getElementById('tr-status').value + '.' : '';
      lineEl.className = 'status';
      lineEl.textContent = all.length + ' result(s). ' + startEl.value + ' through ' + endEl.value + '.' + sNote;
    } catch (err) {
      lineEl.className = 'status err';
      lineEl.textContent = 'Request failed: ' + err;
    } finally {
      btn.disabled = false;
    }
  });

  async function loadDetail(r) {
    detailEl.textContent = 'Loading...';
    try {
      const params = new URLSearchParams({ traceId: r.MessageTraceId, recipient: r.RecipientAddress, start: searchStart, end: searchEnd });
      const res  = await fetch('/api/trace/detail?' + params.toString());
      const data = await res.json();
      if (!data.ok) { detailEl.textContent = 'ERROR: ' + data.error; return; }
      if (!data.rows.length) { detailEl.textContent = '(no detail records returned)'; return; }
      detailEl.textContent = data.rows.map(d =>
        fmtTime(d.Date) + '  [' + d.Event + ']  ' + d.Action + (d.Detail ? ('\n    ' + d.Detail) : '')
      ).join('\n');
    } catch (err) {
      detailEl.textContent = 'Request failed: ' + err;
    }
  }

  qBtn.addEventListener('click', async () => {
    if (!selected) return;
    qOut.innerHTML = '<div class="status">Looking up quarantine entries...</div>';
    try {
      const params = new URLSearchParams({ messageId: selected.MessageId || '', recipient: selected.RecipientAddress, start: searchStart, end: searchEnd });
      const res  = await fetch('/api/trace/quarantine?' + params.toString());
      const data = await res.json();
      if (!data.ok) { qOut.innerHTML = '<div class="status err"></div>'; qOut.firstChild.textContent = data.error; return; }
      if (!data.matches.length) {
        qOut.innerHTML = '<div class="status">No quarantine entry found (may not have been quarantined, may already be released, or may have expired).</div>';
        return;
      }
      qOut.innerHTML = '';
      data.matches.forEach(m => {
        const card = document.createElement('div');
        card.className = 'card destructive';
        const info = document.createElement('div');
        info.style.fontSize = '13px';
        info.textContent = 'Type: ' + m.Type + '   Policy: ' + m.PolicyName + '   Received: ' + m.ReceivedTime + '   Release status: ' + m.ReleaseStatus;
        const btnRow = document.createElement('div');
        btnRow.style.cssText = 'display:flex;align-items:center;gap:8px;margin-top:8px;flex-wrap:wrap;';

        const relBtn = document.createElement('button');
        relBtn.className = 'danger';
        relBtn.textContent = 'Release message';

        const detailsBtn = document.createElement('button');
        detailsBtn.className = 'secondary';
        detailsBtn.textContent = 'Show details';

        const previewBtn = document.createElement('button');
        previewBtn.className = 'secondary';
        previewBtn.textContent = 'Preview message';

        const portalBtn = document.createElement('button');
        portalBtn.className = 'secondary';
        portalBtn.textContent = 'Open Quarantine ↗';
        portalBtn.title = 'Opens the Defender quarantine list. Defender has no documented per-message link, so you may still need to filter there.';
        portalBtn.addEventListener('click', () => {
          if (portal.quarantine) window.open(portal.quarantine, '_blank', 'noopener');
        });

        const detailOut = document.createElement('div');
        detailOut.className = 'out';
        detailOut.style.cssText = 'display:none;margin-top:8px;font-size:12px;';

        detailsBtn.addEventListener('click', async () => {
          if (detailOut.style.display === 'block' && detailOut.dataset.kind === 'detail') {
            detailOut.style.display = 'none'; return;
          }
          detailOut.dataset.kind = 'detail';
          detailOut.style.display = 'block';
          detailOut.textContent = 'Loading details...';
          try {
            const res  = await fetch('/api/trace/quarantine-detail?identity=' + encodeURIComponent(m.Identity));
            const data = await res.json();
            detailOut.textContent = data.ok ? data.text : ('ERROR: ' + data.error);
          } catch (err) { detailOut.textContent = 'Request failed: ' + err; }
        });

        previewBtn.addEventListener('click', async () => {
          if (detailOut.style.display === 'block' && detailOut.dataset.kind === 'preview') {
            detailOut.style.display = 'none'; return;
          }
          detailOut.dataset.kind = 'preview';
          detailOut.style.display = 'block';
          detailOut.textContent = 'Loading preview... (read-only -- previewing does not release the message)';
          try {
            const params = new URLSearchParams({ identity: m.Identity, recipient: selected.RecipientAddress || '' });
            const res  = await fetch('/api/trace/quarantine-preview?' + params.toString());
            const data = await res.json();
            detailOut.textContent = data.ok ? data.text : ('ERROR: ' + data.error);
          } catch (err) { detailOut.textContent = 'Request failed: ' + err; }
        });

        const st = document.createElement('span');
        st.className = 'status';
        st.style.margin = '0';
        relBtn.addEventListener('click', async () => {
          // GUARD: confirmation is required here. Do not remove.
          if (!confirm('Release this message to ALL original recipients?')) return;
          relBtn.disabled = true;
          st.textContent = 'Releasing...';
          try {
            const rr = await fetch('/api/trace/release', {
              method: 'POST', headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ identity: m.Identity, sender: selected.SenderAddress, recipient: selected.RecipientAddress, subject: selected.Subject })
            });
            const rd = await rr.json();
            st.textContent = rd.ok ? 'Released.' : ('FAILED: ' + rd.error);
          } catch (err) { st.textContent = 'Request failed: ' + err; }
        });
        btnRow.appendChild(relBtn);
        btnRow.appendChild(detailsBtn);
        btnRow.appendChild(previewBtn);
        btnRow.appendChild(portalBtn);
        btnRow.appendChild(st);
        card.appendChild(info); card.appendChild(btnRow); card.appendChild(detailOut);
        qOut.appendChild(card);
      });
    } catch (err) {
      qOut.innerHTML = '<div class="status err"></div>';
      qOut.firstChild.textContent = 'Request failed: ' + err;
    }
  });

  exportBtn.addEventListener('click', () => {
    if (!shown.length) return;
    const header = ['Received (' + tzEl.value + ')', 'SenderAddress', 'RecipientAddress', 'Subject', 'Status', 'MessageTraceId'];
    const esc = (v) => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
    const lines = [header.map(esc).join(',')];
    for (const r of shown) {
      lines.push([fmtTime(r.Received), r.SenderAddress, r.RecipientAddress, r.Subject, r.Status, r.MessageTraceId].map(esc).join(','));
    }
    const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'M365_MessageTrace_' + new Date().toISOString().replace(/[:T]/g, '').slice(0, 13) + '.csv';
    a.click();
  });
})();
</script>
'@
    } `
    -Routes @{
        'GET /api/trace/portal-urls' = {
            param($Request)
            $q = $Request.QueryString
            $ex = Get-M365PortalUrl -Kind 'ThreatExplorer' -Detailed `
                    -Sender $q['sender'] -Recipient $q['recipient'] -MessageId $q['messageId']
            $qt = Get-M365PortalUrl -Kind 'Quarantine' -Detailed `
                    -Sender $q['sender'] -Recipient $q['recipient'] -MessageId $q['messageId']
            $payload = [ordered]@{
                threatExplorer             = $ex.Url
                threatExplorerPrepopulated = $ex.Prepopulated
                quarantine                 = $qt.Url
                quarantinePrepopulated     = $qt.Prepopulated
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($payload | ConvertTo-Json -Compress) }
        }
        'GET /api/trace/quarantine-detail' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $d = Get-QuarantineMessageDetail -Identity $Request.QueryString['identity']
                if ($d.Error) { $result.error = $d.Error }
                else {
                    $result.ok   = $true
                    $result.text = Format-QuarantineMessageDetail -Detail $d
                }
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Quarantine detail'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'GET /api/trace/quarantine-preview' = {
            param($Request)
            $result = [ordered]@{ ok = $false; text = ''; error = '' }
            try {
                $q = $Request.QueryString
                $p = Get-QuarantineMessagePreview -Identity $q['identity'] -RecipientAddress $q['recipient']
                if ($p.Error) {
                    $result.error = $p.Error
                } else {
                    $result.ok = $true
                    $result.text = ($p.Preview | Format-List | Out-String).Trim()
                    if (-not $result.text) { $result.text = '(the preview returned no readable content for this message)' }
                }
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Quarantine preview'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'GET /api/trace/search' = {
            param($Request)
            $result = [ordered]@{ ok = $false; results = @(); start = $null; end = $null; error = '' }
            try {
                $q  = $Request.QueryString
                $sd = [datetime]::Parse($q['start']).Date
                $ed = [datetime]::Parse($q['end']).Date.AddDays(1).AddSeconds(-1)
                $p  = @{ StartDate = $sd; EndDate = $ed }
                if ($q['from'])    { $p.SenderAddress    = $q['from'] }
                if ($q['to'])      { $p.RecipientAddress = $q['to'] }
                if ($q['subject']) { $p.Subject          = $q['subject'] }
                if ($q['status'] -and $q['status'] -ne 'All Messages') { $p.Status = $q['status'] }
                $rows = @(Invoke-MessageTrace @p)
                $result.ok    = $true
                $result.start = $sd.ToString('o')
                $result.end   = $ed.ToString('o')
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
        'GET /api/trace/detail' = {
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
        'GET /api/trace/quarantine' = {
            param($Request)
            $result = [ordered]@{ ok = $false; matches = @(); error = '' }
            try {
                $q = $Request.QueryString
                $found = @(Find-QuarantineMatch -MessageId $q['messageId'] -RecipientAddress $q['recipient'] `
                    -StartReceivedDate ([datetime]::Parse($q['start'])) -EndReceivedDate ([datetime]::Parse($q['end'])))
                $result.ok = $true
                $result.matches = @($found | ForEach-Object {
                    [ordered]@{
                        Identity      = $_.Identity.ToString()
                        Type          = "$($_.Type)"
                        PolicyName    = "$($_.PolicyName)"
                        ReceivedTime  = "$($_.ReceivedTime)"
                        ReleaseStatus = "$($_.ReleaseStatus)"
                    }
                })
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-QuarantineMessage'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 6 -Compress) }
        }
        'POST /api/trace/release' = {
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
