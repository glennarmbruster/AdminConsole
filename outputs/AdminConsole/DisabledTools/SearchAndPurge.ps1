#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
#
# NOT OwnProcess, though an earlier version declared it.
#
# The collision is real: on-premises Exchange exports its own
# Get-ComplianceSearch, so in a console that has connected EMS for directory
# work the bare name can resolve to the wrong service -- and it fails as a
# FALSE NEGATIVE, searching on-premises, finding nothing, and reporting
# "nothing found" while the message sits in the cloud.
#
# Process isolation was the wrong answer to it. It cost the console theme, and
# the console's own launcher then ran Connect-ExSE.ps1 into the fresh process,
# importing the very on-premises session the isolation was avoiding.
# Invoke-SccCommand pins every call to the compliance module instead, which
# fixes it wherever the tool runs.
# ---------------------------------------------------------------------------
<#
    SearchAndPurge.ps1
    Find a message across mailboxes and remove it.

    Built for one job: a phishing or misdirected message has landed and needs
    to come back out. It is not an eDiscovery tool and deliberately does not
    try to be one -- no cases, no holds, no export.

    Author   : Glenn Armbruster
    Revised  : 2026-08-10
#>
# TOOLKIT-HELP-BEGIN
# Finds a message across mailboxes and deletes it.
#
# Built for incident response -- a phish lands, you already know what it looks
# like, and it needs to come out of every mailbox that received it. The fields
# match Message Trace on purpose: trace it first, then paste the sender, subject
# and date straight across.
#
# The order is fixed and cannot be skipped: connect, search, look at what was
# found, then purge. Purge stays disabled until a search has completed and
# returned something.
#
# CONNECT FIRST, using the Connect S&C button on the console. This is a
# separate connection from Exchange Online and is not made automatically,
# because most days nothing wants it.
#
# It cannot be done from inside this window. An interactive sign-in needs a
# moment with no modal form in front of it -- the same reason connecting to
# Exchange Online closes and reopens the console.
#
#   SoftDelete  moves items to Recoverable Items. Users can still get them back
#               until the deleted-item retention window expires. This is the
#               default and is what you want almost every time.
#   HardDelete  marks items for permanent removal. Asks a second time before it
#               runs, because nothing brings them back.
#
# WHAT IT CANNOT DO, and these are service limits rather than tool limits:
#
#   * 10 items per mailbox per purge. If a mailbox shows more than 10 hits the
#     tool says so -- run the purge again to take the next 10.
#   * Unindexed items are never purged. A message the service could not index
#     will survive this and has to be dealt with by hand.
#   * Recently delivered mail may not be searchable yet. If a message that
#     definitely arrived returns nothing, wait a few minutes and search again.
#   * There is no preview of the actual messages. That was retired from
#     PowerShell in May 2025. You get per-mailbox hit counts, not subjects, so
#     write a query narrow enough that the count alone tells you it is right.
#
# The search is kept afterwards, named with the date and who ran it, so there
# is a record in the tenant of what was removed and by whom.
#
# Re-check re-runs the same search after a purge and shows what is left. Zero
# is the answer you want.
# TOOLKIT-HELP-END

param(
    [string]$Sender,
    [string]$Subject,
    [string]$DateFrom,
    [string]$DateTo
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# The module floor is real, not defensive padding: -EnableSearchOnlySession
# does not exist before 3.9.0, and without that switch the purge is rejected
# after the search has already run.
$script:MinExoModule = [version]'3.9.0'
$script:SearchName   = ''
$script:LastCounts   = @()
$script:Cancel       = $false

#region Prerequisites ----------------------------------------------------------

function Get-ConnectionSummary {
    <#
        One place that decides whether the tool is usable, so the status line
        and the Search button can never disagree.
    #>
    if (-not (Get-Command -Name 'Test-SccConnectionState' -ErrorAction SilentlyContinue)) {
        return [PSCustomObject]@{
            Ok   = $false
            Text = 'This tool needs a newer Common.ps1 (missing Test-SccConnectionState).'
        }
    }
    $state = Test-SccConnectionState -Refresh
    if ($state.Connected) {
        return [PSCustomObject]@{ Ok = $true; Text = 'Connected to Security and Compliance.' }
    }
    return [PSCustomObject]@{
        Ok   = $false
        Text = 'Not connected. Use the Connect S&C button on the console, then press Re-check.'
    }
}

#endregion

#region Query ------------------------------------------------------------------

function New-ContentQuery {
    <#
        Assemble the KQL from the fields.

        Every supplied field is ANDed, and at least one is required. An empty
        query would match every message in the tenant, and this tool ends in a
        delete -- so "no criteria" is refused rather than treated as "all".
    #>
    param(
        [string]$From,
        [string]$Subject,
        [string]$DateFrom,
        [string]$DateTo
    )

    $parts = @()
    if ("$From".Trim())    { $parts += ('from:"{0}"' -f "$From".Trim()) }
    if ("$Subject".Trim()) { $parts += ('subject:"{0}"' -f ("$Subject".Trim() -replace '"', '')) }

    $df = "$DateFrom".Trim()
    $dt = "$DateTo".Trim()
    if ($df -and $dt) { $parts += ("received:{0}..{1}" -f $df, $dt) }
    elseif ($df)      { $parts += ("received>={0}" -f $df) }
    elseif ($dt)      { $parts += ("received<={0}" -f $dt) }

    return ($parts -join ' AND ')
}

function Test-QueryReady {
    param([string]$Query, [string]$Scope, [string[]]$Mailboxes)

    if (-not "$Query".Trim()) {
        return [PSCustomObject]@{ Ok = $false; Msg =
            "Enter at least one of sender, subject or a date range. A search with no criteria matches everything." }
    }
    if ($Scope -eq 'List' -and $Mailboxes.Count -eq 0) {
        return [PSCustomObject]@{ Ok = $false; Msg = "Specific mailboxes is selected but the list is empty." }
    }
    return [PSCustomObject]@{ Ok = $true; Msg = '' }
}

function New-PurgeSearchName {
    <#
        Dated and attributed, because the search object outlives the incident
        and is the only record in the tenant of who removed what.
    #>
    $who = 'unknown'
    if (Get-Command -Name 'Get-LogOperator' -ErrorAction SilentlyContinue) {
        $who = "$(Get-LogOperator)"
    } elseif ($env:USERNAME) {
        $who = $env:USERNAME
    }
    $who = ($who -split '@')[0] -replace '[^A-Za-z0-9._-]', ''
    if (-not $who) { $who = 'unknown' }
    return ("Purge-{0}-{1}" -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $who)
}

#endregion

#region Search -----------------------------------------------------------------

function ConvertFrom-SuccessResults {
    <#
        Turn the SuccessResults blob into rows.

        The service returns one line per location as free text -- there is no
        structured per-mailbox property to read instead, so this is parsed
        rather than projected. Locations with zero hits are dropped: on an
        all-mailbox search they are the overwhelming majority and they are not
        what anyone is looking at.
    #>
    param([string]$Raw)

    $rows = @()
    if (-not "$Raw".Trim()) { return $rows }

    $rx = [regex]'Location:\s*(?<loc>[^,]+),\s*Item count:\s*(?<cnt>\d+),\s*Total size:\s*(?<size>\d+)'
    foreach ($m in $rx.Matches($Raw)) {
        $cnt = [int]$m.Groups['cnt'].Value
        if ($cnt -le 0) { continue }
        $rows += [PSCustomObject]@{
            Mailbox = $m.Groups['loc'].Value.Trim()
            Items   = $cnt
            Size    = [int64]$m.Groups['size'].Value
        }
    }
    return @($rows | Sort-Object -Property Items -Descending)
}

function Wait-ComplianceJob {
    <#
        Poll until the job stops running.

        Takes a job kind and a name rather than a scriptblock to run. A
        scriptblock defined in one function and invoked inside another resolves
        its variables by rules subtle enough that this script cannot be tested
        against a real PowerShell here -- so the construct is avoided instead of
        trusted. Parameters resolve one way only.

        DoEvents keeps the form repainting and the Stop button clickable. The
        timeout is a backstop, not an expectation -- a search that has genuinely
        stalled should give the operator control back rather than a frozen
        window.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Search', 'Purge')][string]$Kind,

        # ValidateNotNullOrEmpty, not just Mandatory. Learn is explicit that
        # Get-ComplianceSearch with an empty Identity returns EVERY search --
        # so a blank name here would make $obj an array and $obj.Status
        # meaningless. Mandatory alone would prompt for the value, and a
        # WinForms process has no console to answer the prompt in: it would
        # hang. This throws instead.
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Name,
        [Parameter(Mandatory)][System.Windows.Forms.Label]$Status,
        [string]$Verb = 'Searching',
        [int]$TimeoutMinutes = 30
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $spin = 0
    while ($true) {
        if ($Kind -eq 'Search') {
            $obj = Invoke-SccCommand -Name 'Get-ComplianceSearch' -Parameters @{
                Identity = $Name; ErrorAction = 'Stop' }
        } else {
            # The service names the action, not us, so it is found by its search
            # rather than by assuming the '<name>_Purge' convention holds.
            $obj = Invoke-SccCommand -Name 'Get-ComplianceSearchAction' -Parameters @{
                       ErrorAction = 'Stop' } |
                   Where-Object { "$($_.SearchName)" -eq $Name -and "$($_.Action)" -match '(?i)purge' } |
                   Sort-Object CreatedTime -Descending | Select-Object -First 1
        }
        $state = "$($obj.Status)"
        if ($state -match '(?i)^(Completed|Failed)$') { return $obj }

        if ($script:Cancel) {
            $Status.Text = "$Verb -- stopped watching. The job keeps running in the service."
            return $null
        }
        if ((Get-Date) -gt $deadline) {
            $Status.Text = "$Verb -- still running after $TimeoutMinutes minutes. Stopped watching."
            return $null
        }

        $spin++
        $Status.Text = ("{0}... {1} ({2}s)" -f $Verb, $state, ($spin * 5))
        for ($i = 0; $i -lt 10; $i++) {
            Start-Sleep -Milliseconds 500
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
}

function Invoke-ContentSearch {
    <#
        Create, run and wait. Returns the completed search object or $null.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Query,
        [Parameter(Mandatory)][string]$Scope,
        [string[]]$Mailboxes,
        [Parameter(Mandatory)][System.Windows.Forms.Label]$Status
    )

    $params = @{ Name = $Name; ContentMatchQuery = $Query; ErrorAction = 'Stop' }
    if ($Scope -eq 'All') {
        $params['ExchangeLocation'] = 'All'
    } else {
        $params['ExchangeLocation'] = $Mailboxes
        # A list pasted from a ticket will contain at least one address that no
        # longer resolves. Without this the whole search is refused for it.
        $params['AllowNotFoundExchangeLocationsEnabled'] = $true
    }

    $Status.Text = 'Creating the search...'
    Invoke-SccCommand -Name 'New-ComplianceSearch' -Parameters $params | Out-Null

    $Status.Text = 'Starting the search...'
    Invoke-SccCommand -Name 'Start-ComplianceSearch' -Parameters @{
        Identity = $Name; ErrorAction = 'Stop' } | Out-Null

    return (Wait-ComplianceJob -Kind 'Search' -Name $Name -Status $Status -Verb 'Searching')
}

#endregion

#region Purge ------------------------------------------------------------------

function Invoke-PurgeAction {
    <#
        Apply the purge and wait for it.

        -Confirm:$false is required, not tidiness: this cmdlet has a built-in
        pause, and a WinForms process has no console to answer it in -- the tool
        would hang with no visible prompt.

        -WhatIf is deliberately absent. It is accepted by the cmdlet and does
        nothing at all in Security & Compliance PowerShell, which makes it worse
        than useless here: it would read as a safety net that is not there.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][ValidateSet('SoftDelete', 'HardDelete')][string]$PurgeType,
        [Parameter(Mandatory)][System.Windows.Forms.Label]$Status
    )

    $Status.Text = "Submitting the $PurgeType purge..."
    Invoke-SccCommand -Name 'New-ComplianceSearchAction' -Parameters @{
        SearchName = $Name; Purge = $true; PurgeType = $PurgeType
        Confirm = $false; ErrorAction = 'Stop' } | Out-Null

    return (Wait-ComplianceJob -Kind 'Purge' -Name $Name -Status $Status -Verb 'Purging')
}

#endregion

#region UI ---------------------------------------------------------------------

function Show-SearchPurgeForm {
    param([string]$SeedSender, [string]$SeedSubject,
          [string]$SeedFrom, [string]$SeedTo)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Search and Purge -- find a message and remove it'
    $form.Size = New-Object System.Drawing.Size(940, 760)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12

    # --- Connection ---------------------------------------------------------
    $grpConn = New-Object System.Windows.Forms.GroupBox
    $grpConn.Text = '1.  Connection'
    $grpConn.Location = New-Object System.Drawing.Point(12, $y)
    $grpConn.Size = New-Object System.Drawing.Size(898, 62)
    $form.Controls.Add($grpConn)

    $btnConn = New-Object System.Windows.Forms.Button
    # Re-check, not Connect. An interactive sign-in cannot open while this form
    # holds the thread in a modal loop -- the sign-in window draws and then
    # refuses input, which reads as a button that does nothing. Connecting is
    # the console's job, via Connect-SCC.ps1, which is why that script opens no
    # window of its own.
    $btnConn.Text = 'Re-check'
    $btnConn.Location = New-Object System.Drawing.Point(12, 24)
    $btnConn.Size = New-Object System.Drawing.Size(110, 26)
    $grpConn.Controls.Add($btnConn)

    $lblConn = New-Object System.Windows.Forms.Label
    $lblConn.Text = 'Checking...'
    $lblConn.Location = New-Object System.Drawing.Point(134, 29)
    $lblConn.Size = New-Object System.Drawing.Size(750, 18)
    $lblConn.ForeColor = [System.Drawing.Color]::DimGray
    $grpConn.Controls.Add($lblConn)

    $y += 74

    # --- Criteria -----------------------------------------------------------
    $grpFind = New-Object System.Windows.Forms.GroupBox
    $grpFind.Text = '2.  What to look for'
    $grpFind.Location = New-Object System.Drawing.Point(12, $y)
    $grpFind.Size = New-Object System.Drawing.Size(898, 176)
    $form.Controls.Add($grpFind)

    $lblFrom = New-Object System.Windows.Forms.Label
    $lblFrom.Text = 'Sender:'
    $lblFrom.Location = New-Object System.Drawing.Point(12, 28)
    $lblFrom.Size = New-Object System.Drawing.Size(70, 18)
    $grpFind.Controls.Add($lblFrom)

    $txtFrom = New-Object System.Windows.Forms.TextBox
    $txtFrom.Location = New-Object System.Drawing.Point(86, 25)
    $txtFrom.Size = New-Object System.Drawing.Size(330, 24)
    $txtFrom.Text = "$SeedSender"
    $grpFind.Controls.Add($txtFrom)

    $lblSubj = New-Object System.Windows.Forms.Label
    $lblSubj.Text = 'Subject:'
    $lblSubj.Location = New-Object System.Drawing.Point(438, 28)
    $lblSubj.Size = New-Object System.Drawing.Size(70, 18)
    $grpFind.Controls.Add($lblSubj)

    $txtSubj = New-Object System.Windows.Forms.TextBox
    $txtSubj.Location = New-Object System.Drawing.Point(512, 25)
    $txtSubj.Size = New-Object System.Drawing.Size(370, 24)
    $txtSubj.Text = "$SeedSubject"
    $grpFind.Controls.Add($txtSubj)

    $lblDf = New-Object System.Windows.Forms.Label
    $lblDf.Text = 'Received from:'
    $lblDf.Location = New-Object System.Drawing.Point(12, 60)
    $lblDf.Size = New-Object System.Drawing.Size(100, 18)
    $grpFind.Controls.Add($lblDf)

    $txtDf = New-Object System.Windows.Forms.TextBox
    $txtDf.Location = New-Object System.Drawing.Point(116, 57)
    $txtDf.Size = New-Object System.Drawing.Size(110, 24)
    # A seeded range comes from Message Trace and is already widened to bracket
    # the message. Only fall back to a default when nothing was handed over.
    $txtDf.Text = $(if ("$SeedFrom".Trim()) { "$SeedFrom".Trim() }
                    else { (Get-Date).AddDays(-2).ToString('yyyy-MM-dd') })
    $grpFind.Controls.Add($txtDf)

    $lblDt = New-Object System.Windows.Forms.Label
    $lblDt.Text = 'to:'
    $lblDt.Location = New-Object System.Drawing.Point(236, 60)
    $lblDt.Size = New-Object System.Drawing.Size(24, 18)
    $grpFind.Controls.Add($lblDt)

    $txtDt = New-Object System.Windows.Forms.TextBox
    $txtDt.Location = New-Object System.Drawing.Point(262, 57)
    $txtDt.Size = New-Object System.Drawing.Size(110, 24)
    $txtDt.Text = $(if ("$SeedTo".Trim()) { "$SeedTo".Trim() }
                    else { (Get-Date).ToString('yyyy-MM-dd') })
    $grpFind.Controls.Add($txtDt)

    $lblDhint = New-Object System.Windows.Forms.Label
    $lblDhint.Text = 'yyyy-mm-dd. Leave both empty to search all dates.'
    $lblDhint.Location = New-Object System.Drawing.Point(384, 60)
    $lblDhint.Size = New-Object System.Drawing.Size(500, 18)
    $lblDhint.ForeColor = [System.Drawing.Color]::DimGray
    $grpFind.Controls.Add($lblDhint)

    $rbAll = New-Object System.Windows.Forms.RadioButton
    $rbAll.Text = 'All mailboxes'
    $rbAll.Location = New-Object System.Drawing.Point(12, 88)
    $rbAll.Size = New-Object System.Drawing.Size(120, 22)
    $rbAll.Checked = $true
    $grpFind.Controls.Add($rbAll)

    $rbList = New-Object System.Windows.Forms.RadioButton
    $rbList.Text = 'These mailboxes only:'
    $rbList.Location = New-Object System.Drawing.Point(140, 88)
    $rbList.Size = New-Object System.Drawing.Size(160, 22)
    $grpFind.Controls.Add($rbList)

    $txtMbx = New-Object System.Windows.Forms.TextBox
    $txtMbx.Location = New-Object System.Drawing.Point(306, 86)
    $txtMbx.Size = New-Object System.Drawing.Size(576, 52)
    $txtMbx.Multiline = $true
    $txtMbx.ScrollBars = 'Vertical'
    $txtMbx.Enabled = $false
    $grpFind.Controls.Add($txtMbx)

    $lblMhint = New-Object System.Windows.Forms.Label
    $lblMhint.Text = 'One address per line, or comma separated.'
    $lblMhint.Location = New-Object System.Drawing.Point(140, 114)
    $lblMhint.Size = New-Object System.Drawing.Size(160, 32)
    $lblMhint.ForeColor = [System.Drawing.Color]::DimGray
    $grpFind.Controls.Add($lblMhint)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = 'Search'
    $btnSearch.Location = New-Object System.Drawing.Point(12, 142)
    $btnSearch.Size = New-Object System.Drawing.Size(110, 26)
    $btnSearch.Enabled = $false
    $grpFind.Controls.Add($btnSearch)

    $btnStop = New-Object System.Windows.Forms.Button
    $btnStop.Text = 'Stop watching'
    $btnStop.Location = New-Object System.Drawing.Point(128, 142)
    $btnStop.Size = New-Object System.Drawing.Size(110, 26)
    $btnStop.Enabled = $false
    $grpFind.Controls.Add($btnStop)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Connect first.'
    $lblStatus.Location = New-Object System.Drawing.Point(248, 147)
    $lblStatus.Size = New-Object System.Drawing.Size(634, 18)
    $grpFind.Controls.Add($lblStatus)

    $y += 188

    # --- Results ------------------------------------------------------------
    $grpRes = New-Object System.Windows.Forms.GroupBox
    $grpRes.Text = '3.  What was found'
    $grpRes.Location = New-Object System.Drawing.Point(12, $y)
    $grpRes.Size = New-Object System.Drawing.Size(898, 236)
    $form.Controls.Add($grpRes)

    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location = New-Object System.Drawing.Point(12, 24)
    $dgv.Size = New-Object System.Drawing.Size(870, 168)
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.ReadOnly = $true
    $dgv.SelectionMode = 'FullRowSelect'
    $dgv.AutoSizeColumnsMode = 'Fill'
    [void]$dgv.Columns.Add('Mailbox', 'Mailbox')
    [void]$dgv.Columns.Add('Items', 'Hits')
    [void]$dgv.Columns.Add('Note', 'Note')
    $dgv.Columns['Items'].FillWeight = 22
    $dgv.Columns['Note'].FillWeight  = 46
    $grpRes.Controls.Add($dgv)

    $lblTotals = New-Object System.Windows.Forms.Label
    $lblTotals.Text = ''
    $lblTotals.Location = New-Object System.Drawing.Point(12, 200)
    $lblTotals.Size = New-Object System.Drawing.Size(870, 28)
    $grpRes.Controls.Add($lblTotals)

    $y += 248

    # --- Purge --------------------------------------------------------------
    $grpPurge = New-Object System.Windows.Forms.GroupBox
    $grpPurge.Text = '4.  Remove it'
    $grpPurge.Location = New-Object System.Drawing.Point(12, $y)
    $grpPurge.Size = New-Object System.Drawing.Size(898, 100)
    $form.Controls.Add($grpPurge)

    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Text = 'Delete type:'
    $lblType.Location = New-Object System.Drawing.Point(12, 30)
    $lblType.Size = New-Object System.Drawing.Size(80, 18)
    $grpPurge.Controls.Add($lblType)

    $cboType = New-Object System.Windows.Forms.ComboBox
    $cboType.Location = New-Object System.Drawing.Point(96, 27)
    $cboType.Size = New-Object System.Drawing.Size(160, 24)
    $cboType.DropDownStyle = 'DropDownList'
    [void]$cboType.Items.Add('SoftDelete')
    [void]$cboType.Items.Add('HardDelete')
    $cboType.SelectedIndex = 0
    $grpPurge.Controls.Add($cboType)

    $lblTypeNote = New-Object System.Windows.Forms.Label
    $lblTypeNote.Text = 'Recoverable by the user until the retention window expires.'
    $lblTypeNote.Location = New-Object System.Drawing.Point(266, 30)
    $lblTypeNote.Size = New-Object System.Drawing.Size(616, 18)
    $lblTypeNote.ForeColor = [System.Drawing.Color]::DimGray
    $grpPurge.Controls.Add($lblTypeNote)

    $lblConfirm = New-Object System.Windows.Forms.Label
    $lblConfirm.Text = 'Type PURGE to enable:'
    $lblConfirm.Location = New-Object System.Drawing.Point(12, 62)
    $lblConfirm.Size = New-Object System.Drawing.Size(148, 18)
    $grpPurge.Controls.Add($lblConfirm)

    $txtConfirm = New-Object System.Windows.Forms.TextBox
    $txtConfirm.Location = New-Object System.Drawing.Point(164, 59)
    $txtConfirm.Size = New-Object System.Drawing.Size(92, 24)
    $txtConfirm.Enabled = $false
    $grpPurge.Controls.Add($txtConfirm)

    $btnPurge = New-Object System.Windows.Forms.Button
    $btnPurge.Text = 'Purge'
    $btnPurge.Location = New-Object System.Drawing.Point(266, 58)
    $btnPurge.Size = New-Object System.Drawing.Size(110, 26)
    $btnPurge.Enabled = $false
    $grpPurge.Controls.Add($btnPurge)

    $btnRecheck = New-Object System.Windows.Forms.Button
    $btnRecheck.Text = 'Re-check'
    $btnRecheck.Location = New-Object System.Drawing.Point(382, 58)
    $btnRecheck.Size = New-Object System.Drawing.Size(110, 26)
    $btnRecheck.Enabled = $false
    $grpPurge.Controls.Add($btnRecheck)

    $lblPurgeNote = New-Object System.Windows.Forms.Label
    $lblPurgeNote.Text = 'Purge unlocks after a search finds something.'
    $lblPurgeNote.Location = New-Object System.Drawing.Point(502, 63)
    $lblPurgeNote.Size = New-Object System.Drawing.Size(380, 18)
    $lblPurgeNote.ForeColor = [System.Drawing.Color]::DimGray
    $grpPurge.Controls.Add($lblPurgeNote)

    $y += 112

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, $y)
    $txtOut.Size = New-Object System.Drawing.Size(898, 60)
    $txtOut.Multiline = $true
    $txtOut.ReadOnly = $true
    $txtOut.ScrollBars = 'Vertical'
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)

    # --- Behaviour ----------------------------------------------------------

    $syncPurgeGate = {
        # Three conditions, all required: a completed search that found
        # something, and the word typed. Anything less leaves the button off.
        $hasHits = ($script:LastCounts.Count -gt 0)
        $txtConfirm.Enabled = $hasHits
        $btnRecheck.Enabled = ($script:SearchName -ne '')
        $btnPurge.Enabled   = ($hasHits -and ($txtConfirm.Text.Trim().ToUpper() -eq 'PURGE'))
    }

    $renderCounts = {
        $dgv.Rows.Clear()
        $totalItems = 0
        $capped = 0
        foreach ($r in $script:LastCounts) {
            $note = ''
            if ($r.Items -gt 10) {
                $note = 'more than 10 -- one purge takes 10'
                $capped++
            }
            $i = $dgv.Rows.Add()
            $dgv.Rows[$i].Cells['Mailbox'].Value = $r.Mailbox
            $dgv.Rows[$i].Cells['Items'].Value   = $r.Items
            $dgv.Rows[$i].Cells['Note'].Value    = $note
            if ($note) { $dgv.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
            $totalItems += $r.Items
        }
        if ($script:LastCounts.Count -eq 0) {
            $lblTotals.Text = 'Nothing found. If the message is recent it may not be indexed yet -- wait a few minutes and search again.'
        } else {
            $t = ("{0} item(s) in {1} mailbox(es)." -f $totalItems, $script:LastCounts.Count)
            if ($capped -gt 0) {
                $t += ("  {0} mailbox(es) hold more than 10 -- purge removes 10 per mailbox per run, so run it again for the rest." -f $capped)
            }
            $lblTotals.Text = $t
        }
        & $syncPurgeGate
    }

    $rbList.Add_CheckedChanged({ $txtMbx.Enabled = $rbList.Checked })
    $txtConfirm.Add_TextChanged($syncPurgeGate)

    $cboType.Add_SelectedIndexChanged({
        if ("$($cboType.SelectedItem)" -eq 'HardDelete') {
            $lblTypeNote.Text = 'Permanent. Nothing recovers these once the mailbox is processed.'
            $lblTypeNote.ForeColor = [System.Drawing.Color]::Firebrick
        } else {
            $lblTypeNote.Text = 'Recoverable by the user until the retention window expires.'
            $lblTypeNote.ForeColor = [System.Drawing.Color]::DimGray
        }
    })

    $refreshConnection = {
        $form.Cursor = 'WaitCursor'
        $btnConn.Enabled = $false
        try {
            $c = Get-ConnectionSummary
            $lblConn.Text = $c.Text
            $lblConn.ForeColor = $(if ($c.Ok) { [System.Drawing.Color]::ForestGreen }
                                   else { [System.Drawing.Color]::Firebrick })
            $btnSearch.Enabled = $c.Ok
            $lblStatus.Text = $(if ($c.Ok) { 'Ready.' } else { 'Connect first.' })
        } finally {
            $btnConn.Enabled = $true
            $form.Cursor = 'Default'
        }
    }

    $btnConn.Add_Click($refreshConnection)

    $btnStop.Add_Click({ $script:Cancel = $true })

    $btnSearch.Add_Click({
        $mbx = @()
        if ($rbList.Checked) {
            $mbx = @("$($txtMbx.Text)" -split '[,;\r\n]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
        $query = New-ContentQuery -From $txtFrom.Text -Subject $txtSubj.Text `
                                    -DateFrom $txtDf.Text -DateTo $txtDt.Text
        $scope = $(if ($rbList.Checked) { 'List' } else { 'All' })

        $chk = Test-QueryReady -Query $query -Scope $scope -Mailboxes $mbx
        if (-not $chk.Ok) { $txtOut.Text = $chk.Msg; return }

        # A new search invalidates the old results. Clearing them here rather
        # than after means a failed search cannot leave yesterday's counts on
        # screen with the purge button still live against them.
        $script:LastCounts = @()
        $script:SearchName = ''
        $script:Cancel = $false
        & $renderCounts

        $btnSearch.Enabled = $false
        $btnStop.Enabled = $true
        $form.Cursor = 'WaitCursor'
        $name = New-PurgeSearchName
        $txtOut.Text = ("Search : {0}{1}Query  : {2}{1}Scope  : {3}" -f `
                        $name, [Environment]::NewLine, $query,
                        $(if ($scope -eq 'All') { 'All mailboxes' } else { "$($mbx.Count) mailbox(es)" }))
        try {
            $res = Invoke-ContentSearch -Name $name -Query $query -Scope $scope `
                                        -Mailboxes $mbx -Status $lblStatus
            if ($null -eq $res) { return }

            $script:SearchName = $name
            if ("$($res.Status)" -match '(?i)Failed') {
                $lblStatus.Text = 'The search failed.'
                $txtOut.Text += ([Environment]::NewLine + "Search failed: $($res.Errors)")
                Write-FailedLog -Action 'Search' -Target $query -Detail "Search '$name' failed"
                return
            }

            $script:LastCounts = ConvertFrom-SuccessResults -Raw "$($res.SuccessResults)"
            $lblStatus.Text = 'Search complete.'
            & $renderCounts

            Write-CompletedLog -Action 'Search' -Target $query -Result 'Success' `
                -Detail ("Search '{0}' found {1} item(s) in {2} mailbox(es)" -f `
                         $name, (($script:LastCounts | Measure-Object -Property Items -Sum).Sum), $script:LastCounts.Count)
        } catch {
            $lblStatus.Text = 'The search did not run.'
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Search'
            Write-FailedLog -Action 'Search' -Target $query -ErrorRecord $_ -Detail 'Content search failed'
        } finally {
            $btnSearch.Enabled = $true
            $btnStop.Enabled = $false
            $form.Cursor = 'Default'
            & $syncPurgeGate
        }
    })

    $btnRecheck.Add_Click({
        if (-not $script:SearchName) { return }
        $btnRecheck.Enabled = $false
        $btnStop.Enabled = $true
        $script:Cancel = $false
        $form.Cursor = 'WaitCursor'
        try {
            $lblStatus.Text = 'Re-running the search...'
            Invoke-SccCommand -Name 'Start-ComplianceSearch' -Parameters @{
                Identity = $script:SearchName; ErrorAction = 'Stop' } | Out-Null
            $res = Wait-ComplianceJob -Kind 'Search' -Name $script:SearchName `
                                      -Status $lblStatus -Verb 'Re-checking'
            if ($null -eq $res) { return }
            $script:LastCounts = ConvertFrom-SuccessResults -Raw "$($res.SuccessResults)"
            & $renderCounts
            $lblStatus.Text = $(if ($script:LastCounts.Count -eq 0) {
                'Re-check complete -- nothing left.'
            } else { 'Re-check complete -- still finding hits.' })
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Re-check'
        } finally {
            $btnRecheck.Enabled = $true
            $btnStop.Enabled = $false
            $form.Cursor = 'Default'
        }
    })

    $btnPurge.Add_Click({
        $type = "$($cboType.SelectedItem)"
        $total = (($script:LastCounts | Measure-Object -Property Items -Sum).Sum)
        $boxes = $script:LastCounts.Count

        $first = [System.Windows.Forms.MessageBox]::Show(
            ("$type will be applied to $total item(s) across $boxes mailbox(es)." + "`r`n`r`n" +
             "Search: $($script:SearchName)" + "`r`n`r`n" + "Continue?"),
            'Confirm purge', 'YesNo', 'Warning')
        if ($first -ne 'Yes') { return }

        # The second gate exists only for HardDelete. SoftDelete is undoable by
        # the user; this one is not undoable by anyone, and that difference
        # deserves a different amount of friction.
        if ($type -eq 'HardDelete') {
            $second = [System.Windows.Forms.MessageBox]::Show(
                ("HardDelete is permanent." + "`r`n`r`n" +
                 "These items will not be in Recoverable Items and cannot be restored " +
                 "by the user, by you, or by support." + "`r`n`r`n" +
                 "Proceed with HardDelete?"),
                'This cannot be undone', 'YesNo', 'Stop')
            if ($second -ne 'Yes') { return }
        }

        $btnPurge.Enabled = $false
        $btnStop.Enabled = $true
        $script:Cancel = $false
        $form.Cursor = 'WaitCursor'
        try {
            $act = Invoke-PurgeAction -Name $script:SearchName -PurgeType $type -Status $lblStatus
            if ($null -eq $act) { return }

            if ("$($act.Status)" -match '(?i)Failed') {
                $lblStatus.Text = 'The purge failed.'
                $txtOut.Text = "Purge failed: $($act.Errors)"
                Write-FailedLog -Action 'Purge' -Target $script:SearchName `
                    -Detail "$type purge failed: $($act.Errors)"
                return
            }

            $lblStatus.Text = 'Purge complete.'
            $txtOut.Text = ("PURGED" + [Environment]::NewLine + [Environment]::NewLine +
                            "Type    : $type" + [Environment]::NewLine +
                            "Search  : $($script:SearchName)" + [Environment]::NewLine +
                            "Scope   : $total item(s) across $boxes mailbox(es)" + [Environment]::NewLine +
                            [Environment]::NewLine +
                            "Use Re-check to confirm. Mailboxes that held more than 10 items " +
                            "still have the remainder.")
            Write-CompletedLog -Action 'Purge' -Target $script:SearchName -Result 'Success' `
                -Detail ("$type purge of $total item(s) across $boxes mailbox(es)")
        } catch {
            $lblStatus.Text = 'The purge did not run.'
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Purge'
            Write-FailedLog -Action 'Purge' -Target $script:SearchName -ErrorRecord $_ `
                -Detail "$type purge failed"
        } finally {
            $btnStop.Enabled = $false
            $form.Cursor = 'Default'
            & $syncPurgeGate
        }
    })

    $form.Add_Shown($refreshConnection)

    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Search and Purge' | Out-Null

    [void]$form.ShowDialog()
    $form.Dispose()
}

#endregion

Show-SearchPurgeForm -SeedSender $Sender -SeedSubject $Subject -SeedFrom $DateFrom -SeedTo $DateTo
