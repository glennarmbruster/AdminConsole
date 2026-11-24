#Requires -Version 5.1
<#
.SYNOPSIS
    Shared Mailbox Audit view for M365_Host_Web.ps1. Read-only.
.DESCRIPTION
    ADDED 2026-07-24. Browser equivalent of the WinForms SharedMailboxAudit.ps1.

    Contains no logic -- Invoke-M365MailboxAudit, ConvertTo-M365AuditRecord, and
    Get-M365FullAccessHolders all live in M365Logic.psm1.

    WHAT MAKES THIS DIFFERENT FROM THE PLAIN MAILBOX AUDIT, and why it's worth
    being its own view: it cross-references every actor in the audit results
    against the mailbox's CURRENT Full Access holder list. On a shared mailbox
    the interesting question isn't just "who did this" but "who did this, and
    are they still supposed to be able to?" -- so any actor who acted on the
    mailbox but is NOT in the current holder list is flagged. Common innocent
    explanations exist (their access was removed after the fact, or they acted
    via a different mechanism such as Send-As or an admin action), which is why
    it's labelled "not a current Full Access holder" rather than anything
    accusatory.

    ObjectTypes note carried forward from the WinForms tool: it deliberately
    does NOT gate on RecipientTypeDetails=SharedMailbox. That was considered and
    rejected -- running it against a regular mailbox is harmless (the holder
    list just comes back short or empty) and the extra type check bought
    nothing.

    SAME TWO CAVEATS AS MAILBOX AUDIT, surfaced in the UI: the 50,000-record
    ceiling can truncate results, and a wide search blocks every other view
    while it runs.
#>

Register-M365View -Id 'sharedaudit' -Name 'Shared MB Audit' `
    -Description 'Audit activity on a shared mailbox and cross-reference every actor against the current Full Access holders.' `
    -GetHtml {
        return @'
  <h2>Shared Mailbox Audit</h2>
  <div class="note">
    Read-only. Audits delegate-relevant activity and compares each actor against who currently holds
    Full Access. Actors <strong>not</strong> in the current holder list are flagged -- often innocent
    (access removed later, or they acted via Send-As or an admin path), but worth a look.
    <br><br>
    <strong>Can be slow</strong>, and while it runs every other view waits. Narrow the dates if you can.
  </div>

  <div class="row">
    <label>Shared mailbox:<br><input id="sa-smtp" type="text" placeholder="shared@domain, or just the alias" style="width:280px"></label>
    <label>From (UTC date):<br><input id="sa-start" type="date"></label>
    <label>Through (UTC date):<br><input id="sa-end" type="date"></label>
    <button id="sa-run">Search</button>
  </div>

  <div id="sa-status" class="status">No search run yet.</div>
  <div id="sa-warn"></div>

  <div id="sa-holders"></div>

  <div class="row">
    <label>Filter loaded results:<br><input id="sa-filter" type="text" style="width:240px"></label>
    <button id="sa-filter-btn">Filter</button>
    <button id="sa-filter-clear" class="secondary">Clear</button>
    <label style="margin-left:8px;"><input type="checkbox" id="sa-only-flagged"> Only show flagged actors</label>
    <button id="sa-export" class="secondary" disabled>Export to CSV</button>
  </div>

  <table id="sa-table">
    <thead><tr>
      <th data-col="TimeUtc">Time</th>
      <th data-col="OperationFriendly">Operation</th>
      <th data-col="Actor">Actor</th>
      <th data-col="IsHolder">Current holder?</th>
      <th data-col="ClientIP">Client IP</th>
      <th data-col="Subject">Subject</th>
    </tr></thead>
    <tbody id="sa-body"></tbody>
  </table>

  <h3 style="margin-top:16px;">Raw record</h3>
  <div id="sa-detail" class="out" style="min-height:130px;font-size:12px;">No row selected.</div>

<script>
(function () {
  let all = [], shown = [], holders = [], sortCol = null, sortAsc = true;
  const smtpEl   = document.getElementById('sa-smtp');
  const startEl  = document.getElementById('sa-start');
  const endEl    = document.getElementById('sa-end');
  const statusEl = document.getElementById('sa-status');
  const warnEl   = document.getElementById('sa-warn');
  const holdEl   = document.getElementById('sa-holders');
  const bodyEl   = document.getElementById('sa-body');
  const detailEl = document.getElementById('sa-detail');
  const exportBtn= document.getElementById('sa-export');
  const runBtn   = document.getElementById('sa-run');
  const onlyFlag = document.getElementById('sa-only-flagged');

  const iso = (d) => d.toISOString().slice(0, 10);
  const today = new Date();
  endEl.value   = iso(today);
  startEl.value = iso(new Date(today.getTime() - 14 * 86400000));

  function fmtTime(s) { if (!s) return ''; try { return new Date(s).toLocaleString(); } catch (e) { return s; } }

  function render() {
    bodyEl.innerHTML = '';
    for (const r of shown) {
      const tr = document.createElement('tr');
      if (!r.IsHolder) { tr.style.background = '#fdecea'; }
      [fmtTime(r.TimeUtc), r.OperationFriendly, r.Actor,
       (r.IsHolder ? 'yes' : 'NOT a current holder'), r.ClientIP, r.Subject]
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

  document.querySelectorAll('#sa-table th').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      sortAsc = (sortCol === col) ? !sortAsc : true;
      sortCol = col;
      shown.sort((a, b) => {
        let av = a[col], bv = b[col];
        if (typeof av === 'boolean') { av = av ? 1 : 0; bv = bv ? 1 : 0; }
        av = av == null ? '' : av; bv = bv == null ? '' : bv;
        return (av > bv ? 1 : av < bv ? -1 : 0) * (sortAsc ? 1 : -1);
      });
      render();
    });
  });

  function applyFilter() {
    const n = document.getElementById('sa-filter').value.trim().toLowerCase();
    shown = all.filter(r => {
      if (onlyFlag.checked && r.IsHolder) return false;
      if (!n) return true;
      return ['OperationFriendly','Actor','ClientIP','ClientInfo','Subject','Folder','ResultStatus']
        .some(k => (r[k] || '').toLowerCase().includes(n));
    });
    const flagged = all.filter(r => !r.IsHolder).length;
    statusEl.className = 'status';
    statusEl.textContent = shown.length + ' of ' + all.length + ' record(s) shown  --  ' +
                           flagged + ' by actors who are not current Full Access holders.';
    render();
  }
  document.getElementById('sa-filter-btn').addEventListener('click', applyFilter);
  document.getElementById('sa-filter').addEventListener('keydown', (e) => { if (e.key === 'Enter') applyFilter(); });
  document.getElementById('sa-filter-clear').addEventListener('click', () => {
    document.getElementById('sa-filter').value = ''; onlyFlag.checked = false; applyFilter();
  });
  onlyFlag.addEventListener('change', applyFilter);

  runBtn.addEventListener('click', async () => {
    const smtp = smtpEl.value.trim();
    if (!smtp) { statusEl.className = 'status err'; statusEl.textContent = 'Enter a mailbox first.'; return; }
    runBtn.disabled = true;
    warnEl.innerHTML = ''; holdEl.innerHTML = ''; bodyEl.innerHTML = '';
    detailEl.textContent = 'No row selected.';
    document.getElementById('sa-filter').value = '';
    onlyFlag.checked = false;
    statusEl.className = 'status';
    statusEl.textContent = 'Searching... other views will wait until this finishes.';
    try {
      const params = new URLSearchParams({ smtp: smtp, start: startEl.value, end: endEl.value });
      const res  = await fetch('/api/sharedaudit/search?' + params.toString());
      const data = await res.json();
      if (!data.ok) { statusEl.className = 'status err'; statusEl.textContent = data.error; return; }

      all = data.records; holders = data.holders || [];

      // Holder panel -- shown even when the audit itself found nothing, since
      // "who can get in here" is useful on its own.
      const box = document.createElement('div');
      box.className = 'card';
      const h = document.createElement('h3');
      h.textContent = 'Current Full Access holders (' + holders.length + ')';
      box.appendChild(h);
      const d = document.createElement('div');
      d.className = 'desc';
      d.textContent = holders.length ? holders.join(', ') : '(none -- nobody currently holds explicit Full Access)';
      box.appendChild(d);
      if (data.holdersError) {
        const e = document.createElement('div');
        e.className = 'status err';
        e.textContent = 'Holder list could not be read, so the flagging below is unreliable: ' + data.holdersError;
        box.appendChild(e);
      }
      holdEl.appendChild(box);

      applyFilter();

      if (data.hitCeiling) {
        const w = document.createElement('div');
        w.className = 'warn';
        w.textContent = 'RESULTS MAY BE INCOMPLETE -- the 50,000-record audit ceiling was reached, so older activity ' +
                        'in this range is missing. Narrow the date range and search again.';
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
    const cols = ['TimeUtc','Operation','OperationFriendly','Actor','IsHolder','MailboxOwnerUPN','ClientIP','ClientInfo','Subject','Folder','ResultStatus'];
    const esc = (v) => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
    const lines = [cols.map(esc).join(',')];
    for (const r of shown) lines.push(cols.map(c => esc(r[c])).join(','));
    const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'M365_SharedMailboxAudit_' + new Date().toISOString().replace(/[:T]/g, '').slice(0, 13) + '.csv';
    a.click();
  });
})();
</script>
'@
    } `
    -Routes @{
        'GET /api/sharedaudit/search' = {
            param($Request)
            $result = [ordered]@{
                ok = $false; records = @(); holders = @(); holdersError = ''
                hitCeiling = $false; pages = 0; error = ''
            }
            try {
                $q = $Request.QueryString
                $smtp = Get-TargetSmtp -Text $q['smtp']
                if (-not $smtp) { throw "Could not determine a domain -- type the full email address, or make sure Exchange Online is connected." }

                $sd = [datetime]::SpecifyKind([datetime]::Parse($q['start'], [System.Globalization.CultureInfo]::InvariantCulture).Date, [DateTimeKind]::Utc)
                $ed = [datetime]::SpecifyKind([datetime]::Parse($q['end'],   [System.Globalization.CultureInfo]::InvariantCulture).Date.AddDays(1).AddSeconds(-1), [DateTimeKind]::Utc)
                if ($ed -le $sd) { throw "The 'Through' date must be on or after the 'From' date." }

                # Holder list first -- it's cheap, and it's useful even if the
                # audit search then returns nothing.
                $hold = Get-M365FullAccessHolders -Smtp $smtp
                $result.holders = @($hold.Holders)
                if ($hold.Error) { $result.holdersError = $hold.Error }

                # Case-insensitive holder set for the cross-reference. Holder
                # names can come back in several forms, so compare on the local
                # part too rather than requiring an exact string match.
                $holderKeys = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($h in $hold.Holders) {
                    [void]$holderKeys.Add("$h")
                    if ("$h" -match '\\([^\\]+)$') { [void]$holderKeys.Add($Matches[1]) }
                    if ("$h" -like '*@*')          { [void]$holderKeys.Add((("$h" -split '@')[0])) }
                }

                $search = Invoke-M365MailboxAudit -Mailbox $smtp -StartUtc $sd -EndUtc $ed
                $result.ok         = $true
                $result.hitCeiling = [bool]$search.HitCeiling
                $result.pages      = $search.Pages
                $result.records    = @($search.Records | ForEach-Object {
                    $rec = ConvertTo-M365AuditRecord -RawRecord $_
                    $actor = "$($rec.Actor)"
                    $isHolder = $holderKeys.Contains($actor)
                    if (-not $isHolder -and $actor -like '*@*') { $isHolder = $holderKeys.Contains((($actor -split '@')[0])) }
                    # The mailbox acting on itself is not a delegate action, so
                    # don't flag it as an outsider.
                    if (-not $isHolder -and $actor -and ($actor -eq $smtp)) { $isHolder = $true }
                    [ordered]@{
                        TimeUtc           = $rec.TimeUtc
                        Operation         = $rec.Operation
                        OperationFriendly = $rec.OperationFriendly
                        Actor             = $actor
                        IsHolder          = [bool]$isHolder
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
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Shared Mailbox Audit'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 6 -Compress) }
        }
    }
