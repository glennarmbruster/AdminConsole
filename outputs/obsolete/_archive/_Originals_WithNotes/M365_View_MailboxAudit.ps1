#Requires -Version 5.1
<#
.SYNOPSIS
    Mailbox Audit view for M365_Host_Web.ps1. Read-only.
.DESCRIPTION
    ADDED 2026-07-24. Browser equivalent of the WinForms MailboxAudit.ps1.

    Contains no logic -- Get-M365MailboxAuditOperations, Invoke-M365MailboxAudit,
    and ConvertTo-M365AuditRecord live in M365Logic.psm1. The operation list is
    served from the module so the picker can't drift from what the search
    actually accepts.

    TWO THINGS THE OPERATOR SHOULD KNOW, surfaced in the UI rather than buried:

    1. TRUNCATION. Search-UnifiedAuditLog has a documented 50,000-record
       ReturnLargeSet ceiling. When the module reports HitCeiling, this view
       says the results may be incomplete instead of presenting a capped set as
       though it were the whole story. Narrow the date range or the operation
       list in that case.

    2. IT BLOCKS. The host is single-threaded by design (EXO's session isn't
       thread-safe), and a wide audit search can take a long while and page
       repeatedly. While it runs, every other view waits. The UI says so up
       front rather than looking frozen. No Cancel button is offered -- and
       that's deliberate, not an omission: the WinForms tool's Cancel worked by
       flipping a flag its modal dialog could set mid-search, which a
       one-request-at-a-time server genuinely cannot do. Bounding the search
       before starting it (dates, operations) is the honest equivalent.

    Timestamps come from the server as UTC ISO-8601 and are rendered in the
    viewer's own timezone by the browser, rather than being pre-converted to
    the server's local time.
#>

Register-M365View -Id 'audit' -Name 'Mailbox Audit' `
    -Description 'Search the unified audit log for one mailbox: who did what, when, from where. Read-only.' `
    -GetHtml {
        # Operation checkboxes are built server-side from the module's map so
        # there's exactly one list of valid operations in the whole toolkit.
        $opBoxes = ((Get-M365MailboxAuditOperations).GetEnumerator() | ForEach-Object {
            $friendly = $_.Key
            $value    = $_.Value
            "      <label style=""display:inline-block;width:340px;font-size:12px;""><input type=""checkbox"" class=""au-op"" value=""$value""> $friendly</label>"
        }) -join "`n"
        return @"
  <h2>Mailbox Audit</h2>
  <div class="note">
    Read-only. Leave every operation unchecked to search <strong>all</strong> of them.
    <br><br>
    <strong>This one can be slow.</strong> The host handles one request at a time, so a wide search makes every
    other view wait until it finishes. Narrow the dates and operations if you can. There is no cancel button --
    a single-threaded server can't honour one mid-search, so bound the query up front instead.
  </div>

  <div class="row">
    <label>Mailbox:<br><input id="au-smtp" type="text" placeholder="user@domain, or just the alias" style="width:280px"></label>
    <label>From (UTC date):<br><input id="au-start" type="date"></label>
    <label>Through (UTC date):<br><input id="au-end" type="date"></label>
    <button id="au-run">Search</button>
  </div>

  <fieldset>
    <legend>Operations (none checked = all)</legend>
    <div class="row" style="margin-bottom:6px;">
      <button id="au-op-all" class="secondary" style="padding:4px 10px;font-size:12px;">Check all</button>
      <button id="au-op-none" class="secondary" style="padding:4px 10px;font-size:12px;">Uncheck all</button>
    </div>
    <div style="max-height:180px;overflow-y:auto;">
$opBoxes
    </div>
  </fieldset>

  <div id="au-status" class="status">No search run yet.</div>
  <div id="au-warn"></div>

  <div class="row">
    <label>Filter loaded results:<br><input id="au-filter" type="text" style="width:260px"></label>
    <button id="au-filter-btn">Filter</button>
    <button id="au-filter-clear" class="secondary">Clear</button>
    <button id="au-export" class="secondary" disabled>Export to CSV</button>
  </div>

  <table id="au-table">
    <thead><tr>
      <th data-col="TimeUtc">Time</th>
      <th data-col="OperationFriendly">Operation</th>
      <th data-col="Actor">Actor</th>
      <th data-col="ClientIP">Client IP</th>
      <th data-col="Subject">Subject</th>
      <th data-col="ResultStatus">Result</th>
    </tr></thead>
    <tbody id="au-body"></tbody>
  </table>

  <h3 style="margin-top:16px;">Raw record</h3>
  <div class="status">Click a row to see the full audit record JSON.</div>
  <div id="au-detail" class="out" style="min-height:140px;font-size:12px;">No row selected.</div>

<script>
(function () {
  let all = [], shown = [], sortCol = null, sortAsc = true;
  const smtpEl   = document.getElementById('au-smtp');
  const startEl  = document.getElementById('au-start');
  const endEl    = document.getElementById('au-end');
  const statusEl = document.getElementById('au-status');
  const warnEl   = document.getElementById('au-warn');
  const bodyEl   = document.getElementById('au-body');
  const detailEl = document.getElementById('au-detail');
  const exportBtn= document.getElementById('au-export');
  const runBtn   = document.getElementById('au-run');

  const iso = (d) => d.toISOString().slice(0, 10);
  const today = new Date();
  endEl.value   = iso(today);
  startEl.value = iso(new Date(today.getTime() - 7 * 86400000));

  document.getElementById('au-op-all').addEventListener('click',
    () => document.querySelectorAll('.au-op').forEach(c => c.checked = true));
  document.getElementById('au-op-none').addEventListener('click',
    () => document.querySelectorAll('.au-op').forEach(c => c.checked = false));

  function fmtTime(isoStr) {
    if (!isoStr) return '';
    try { return new Date(isoStr).toLocaleString(); } catch (e) { return isoStr; }
  }

  function render() {
    bodyEl.innerHTML = '';
    for (const r of shown) {
      const tr = document.createElement('tr');
      [fmtTime(r.TimeUtc), r.OperationFriendly, r.Actor, r.ClientIP, r.Subject, r.ResultStatus]
        .forEach(v => { const td = document.createElement('td'); td.textContent = v || ''; tr.appendChild(td); });
      tr.addEventListener('click', () => {
        Array.from(bodyEl.children).forEach(x => x.classList.remove('selected'));
        tr.classList.add('selected');
        detailEl.textContent = r.RawJson || '(no raw record)';
      });
      bodyEl.appendChild(tr);
    }
    exportBtn.disabled = (shown.length === 0);
  }

  document.querySelectorAll('#au-table th').forEach(th => {
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
    const n = document.getElementById('au-filter').value.trim().toLowerCase();
    if (!n) { shown = all.slice(); statusEl.textContent = all.length + ' record(s).'; }
    else {
      shown = all.filter(r => ['OperationFriendly','Actor','ClientIP','ClientInfo','Subject','Folder','ResultStatus','MailboxOwnerUPN']
        .some(k => (r[k] || '').toLowerCase().includes(n)));
      statusEl.textContent = shown.length + ' of ' + all.length + " record(s) match '" + n + "'.";
    }
    statusEl.className = 'status';
    render();
  }
  document.getElementById('au-filter-btn').addEventListener('click', applyFilter);
  document.getElementById('au-filter').addEventListener('keydown', (e) => { if (e.key === 'Enter') applyFilter(); });
  document.getElementById('au-filter-clear').addEventListener('click', () => {
    document.getElementById('au-filter').value = ''; applyFilter();
  });

  runBtn.addEventListener('click', async () => {
    const smtp = smtpEl.value.trim();
    if (!smtp) { statusEl.className = 'status err'; statusEl.textContent = 'Enter a mailbox first.'; return; }
    const ops = Array.from(document.querySelectorAll('.au-op:checked')).map(c => c.value);

    runBtn.disabled = true;
    warnEl.innerHTML = '';
    bodyEl.innerHTML = '';
    detailEl.textContent = 'No row selected.';
    document.getElementById('au-filter').value = '';
    statusEl.className = 'status';
    statusEl.textContent = 'Searching the audit log... this can take a while, and other views will wait until it finishes.';
    try {
      const params = new URLSearchParams({
        smtp: smtp, start: startEl.value, end: endEl.value, ops: ops.join(',')
      });
      const res  = await fetch('/api/audit/search?' + params.toString());
      const data = await res.json();
      if (!data.ok) { statusEl.className = 'status err'; statusEl.textContent = data.error; return; }
      all = data.records; shown = all.slice();
      render();
      statusEl.className = 'status';
      statusEl.textContent = all.length + ' record(s) over ' + data.pages + ' page(s). ' +
        startEl.value + ' through ' + endEl.value + ' (UTC).';
      if (data.hitCeiling) {
        const w = document.createElement('div');
        w.className = 'warn';
        w.textContent = 'RESULTS MAY BE INCOMPLETE. Search-UnifiedAuditLog stops at 50,000 records and that ceiling was reached, ' +
                        'so older activity in this range is missing. Narrow the date range or pick specific operations and search again.';
        warnEl.appendChild(w);
      }
    } catch (err) {
      statusEl.className = 'status err';
      statusEl.textContent = 'Request failed: ' + err;
    } finally {
      runBtn.disabled = false;
    }
  });

  exportBtn.addEventListener('click', () => {
    if (!shown.length) return;
    const cols = ['TimeUtc','Operation','OperationFriendly','Actor','MailboxOwnerUPN','ClientIP','ClientInfo','Subject','Folder','ResultStatus'];
    const esc = (v) => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
    const lines = [cols.map(esc).join(',')];
    for (const r of shown) lines.push(cols.map(c => esc(r[c])).join(','));
    const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'M365_MailboxAudit_' + new Date().toISOString().replace(/[:T]/g, '').slice(0, 13) + '.csv';
    a.click();
  });
})();
</script>
"@
    } `
    -Routes @{
        'GET /api/audit/search' = {
            param($Request)
            $result = [ordered]@{ ok = $false; records = @(); hitCeiling = $false; pages = 0; error = '' }
            try {
                $q = $Request.QueryString
                $smtp = Get-TargetSmtp -Text $q['smtp']
                if (-not $smtp) { throw "Could not determine a domain -- type the full email address, or make sure Exchange Online is connected." }

                # Treat the supplied dates as UTC day boundaries, matching the
                # label on the form. End is inclusive of the whole day.
                $sd = [datetime]::SpecifyKind([datetime]::Parse($q['start'], [System.Globalization.CultureInfo]::InvariantCulture).Date, [DateTimeKind]::Utc)
                $ed = [datetime]::SpecifyKind([datetime]::Parse($q['end'],   [System.Globalization.CultureInfo]::InvariantCulture).Date.AddDays(1).AddSeconds(-1), [DateTimeKind]::Utc)
                if ($ed -le $sd) { throw "The 'Through' date must be on or after the 'From' date." }

                $ops = @()
                if ($q['ops']) { $ops = @($q['ops'] -split ',' | Where-Object { $_ -ne '' }) }

                $search = Invoke-M365MailboxAudit -Mailbox $smtp -StartUtc $sd -EndUtc $ed -Operations $ops
                $result.ok         = $true
                $result.hitCeiling = [bool]$search.HitCeiling
                $result.pages      = $search.Pages
                $result.records    = @($search.Records | ForEach-Object {
                    $rec = ConvertTo-M365AuditRecord -RawRecord $_
                    [ordered]@{
                        TimeUtc           = $rec.TimeUtc
                        Operation         = $rec.Operation
                        OperationFriendly = $rec.OperationFriendly
                        Actor             = $rec.Actor
                        MailboxOwnerUPN   = $rec.MailboxOwnerUPN
                        ClientIP          = $rec.ClientIP
                        ClientInfo        = $rec.ClientInfo
                        Subject           = $rec.Subject
                        Folder            = $rec.Folder
                        ResultStatus      = $rec.ResultStatus
                        RawJson           = $rec.RawJson
                    }
                })
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Search-UnifiedAuditLog (Mailbox Audit)'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 6 -Compress) }
        }
    }
