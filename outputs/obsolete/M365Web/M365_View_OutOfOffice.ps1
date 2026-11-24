#Requires -Version 5.1
<#
    M365_View_OutOfOffice.ps1
    Out of Office view for M365_Host_Web.ps1.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

Register-M365View -Id 'oof' -Name 'Out of Office' `
    -Description 'Read and set a mailbox auto-reply: state, external audience, schedule, and both message bodies.' `
    -GetHtml {
        $stateOptions = ((Get-OutOfOfficeStates) | ForEach-Object { "        <option>$_</option>" }) -join "`n"
        $audOptions   = ((Get-OutOfOfficeAudiences) | ForEach-Object { "        <option>$_</option>" }) -join "`n"
        return @"
  <h2>Out of Office</h2>
  <div class="note">Load a mailbox first, adjust, then Save. Start/End apply only when State is <em>Scheduled</em>.</div>

  <div class="row">
    <label>Mailbox:<br><input id="oof-smtp" type="text" placeholder="user@domain, or just the alias" style="width:300px"></label>
    <button id="oof-load">Load</button>
  </div>

  <div class="row">
    <label>State:<br>
      <select id="oof-state">
$stateOptions
      </select>
    </label>
    <label>External audience:<br>
      <select id="oof-aud">
$audOptions
      </select>
    </label>
    <label>Start (Scheduled only):<br><input id="oof-start" type="datetime-local"></label>
    <label>End (Scheduled only):<br><input id="oof-end" type="datetime-local"></label>
  </div>

  <div class="row">
    <label style="flex:1;">Internal message:<br><textarea id="oof-int" rows="7"></textarea></label>
  </div>
  <div class="row">
    <label style="flex:1;">External message:<br><textarea id="oof-ext" rows="7"></textarea></label>
  </div>

  <div class="row">
    <button id="oof-save" disabled>Save</button>
    <span id="oof-status" class="status" style="margin:0;">Nothing loaded yet.</span>
  </div>

<script>
(function () {
  const smtpEl  = document.getElementById('oof-smtp');
  const stateEl = document.getElementById('oof-state');
  const audEl   = document.getElementById('oof-aud');
  const startEl = document.getElementById('oof-start');
  const endEl   = document.getElementById('oof-end');
  const intEl   = document.getElementById('oof-int');
  const extEl   = document.getElementById('oof-ext');
  const statusEl= document.getElementById('oof-status');
  const loadBtn = document.getElementById('oof-load');
  const saveBtn = document.getElementById('oof-save');

  function syncSchedule() {
    const on = (stateEl.value === 'Scheduled');
    startEl.disabled = !on;
    endEl.disabled   = !on;
  }
  stateEl.addEventListener('change', syncSchedule);
  syncSchedule();

  async function load() {
    const smtp = smtpEl.value.trim();
    if (!smtp) { statusEl.className = 'status err'; statusEl.textContent = 'Enter a mailbox first.'; return; }
    loadBtn.disabled = true;
    saveBtn.disabled = true;
    statusEl.className = 'status';
    statusEl.textContent = 'Loading...';
    try {
      const res  = await fetch('/api/oof/state?smtp=' + encodeURIComponent(smtp));
      const data = await res.json();
      if (!data.ok) { statusEl.className = 'status err'; statusEl.textContent = data.error; return; }
      stateEl.value = data.autoReplyState || 'Disabled';
      audEl.value   = data.externalAudience || 'None';
      startEl.value = data.startTime || '';
      endEl.value   = data.endTime || '';
      intEl.value   = data.internalMessage || '';
      extEl.value   = data.externalMessage || '';
      syncSchedule();
      saveBtn.disabled = false;
      statusEl.className = 'status';
      statusEl.textContent = 'Loaded current settings for ' + smtp + '.';
    } catch (err) {
      statusEl.className = 'status err';
      statusEl.textContent = 'Request failed: ' + err;
    } finally {
      loadBtn.disabled = false;
    }
  }
  loadBtn.addEventListener('click', load);
  smtpEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') load(); });

  saveBtn.addEventListener('click', async () => {
    const smtp = smtpEl.value.trim();
    if (!smtp) return;
    if (stateEl.value === 'Scheduled' && (!startEl.value || !endEl.value)) {
      statusEl.className = 'status err';
      statusEl.textContent = 'State is Scheduled, so both Start and End are required.';
      return;
    }
    // GUARD: confirmation is required here. Do not remove.
    if (!confirm("Save Out of Office settings for '" + smtp + "'?\n\nState: " + stateEl.value +
                 '\nExternal audience: ' + audEl.value)) return;

    saveBtn.disabled = true;
    loadBtn.disabled = true;
    statusEl.className = 'status';
    statusEl.textContent = 'Saving...';
    try {
      const res = await fetch('/api/oof/save', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          smtp: smtp,
          state: stateEl.value,
          audience: audEl.value,
          start: startEl.value,
          end: endEl.value,
          internal: intEl.value,
          external: extEl.value
        })
      });
      const data = await res.json();
      if (data.ok) {
        statusEl.className = 'status';
        statusEl.textContent = 'Saved. Reloading to confirm what Exchange actually stored...';
        await load();
      } else {
        statusEl.className = 'status err';
        statusEl.textContent = data.error;
      }
    } catch (err) {
      statusEl.className = 'status err';
      statusEl.textContent = 'Request failed: ' + err;
    } finally {
      saveBtn.disabled = false;
      loadBtn.disabled = false;
    }
  });
})();
</script>
"@
    } `
    -Routes @{
        'GET /api/oof/state' = {
            param($Request)
            $result = [ordered]@{
                ok = $false; autoReplyState = $null; externalAudience = $null
                startTime = $null; endTime = $null
                internalMessage = $null; externalMessage = $null; error = ''
            }
            try {
                $smtp = Get-TargetSmtp -Text $Request.QueryString['smtp']
                if (-not $smtp) { throw "Could not determine a domain -- type the full email address, or make sure Exchange Online is connected." }
                $cfg = Get-OutOfOfficeConfig -Smtp $smtp
                $result.ok               = $true
                $result.autoReplyState   = "$($cfg.AutoReplyState)"
                $result.externalAudience = "$($cfg.ExternalAudience)"
                if ($cfg.StartTime) { $result.startTime = ([datetime]$cfg.StartTime).ToString('yyyy-MM-ddTHH:mm') }
                if ($cfg.EndTime)   { $result.endTime   = ([datetime]$cfg.EndTime).ToString('yyyy-MM-ddTHH:mm') }
                $result.internalMessage  = "$($cfg.InternalMessage)"
                $result.externalMessage  = "$($cfg.ExternalMessage)"
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxAutoReplyConfiguration'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
        'POST /api/oof/save' = {
            param($Request)
            $result = [ordered]@{ ok = $false; error = '' }
            try {
                $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
                $body = $reader.ReadToEnd() | ConvertFrom-Json
                $smtp = Get-TargetSmtp -Text $body.smtp
                if (-not $smtp) { throw "Could not determine a domain for '$($body.smtp)'." }

                $st = $null; $et = $null
                if ($body.state -eq 'Scheduled') {
                    if ($body.start) {
                        try { $st = [datetime]::Parse($body.start, [System.Globalization.CultureInfo]::InvariantCulture) }
                        catch { throw "Could not read the Start value '$($body.start)' as a date/time." }
                    }
                    if ($body.end) {
                        try { $et = [datetime]::Parse($body.end, [System.Globalization.CultureInfo]::InvariantCulture) }
                        catch { throw "Could not read the End value '$($body.end)' as a date/time." }
                    }
                    if ($st -and $et -and $et -le $st) { throw "End must be after Start." }
                }

                Set-OutOfOfficeConfig -Smtp $smtp -AutoReplyState $body.state -ExternalAudience $body.audience `
                    -InternalMessage $body.internal -ExternalMessage $body.external -StartTime $st -EndTime $et
                Write-CompletedLog -Task "Set OutOfOffice $smtp (EXO; state: $($body.state))"
                $result.ok = $true
            } catch {
                $result.error = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-MailboxAutoReplyConfiguration'
            }
            return @{ ContentType = 'application/json; charset=utf-8'; Body = ($result | ConvertTo-Json -Depth 5 -Compress) }
        }
    }
