#Requires -Version 5.1
<#
    MessageTrace.ps1
    Standalone Message Trace / Quarantine Release tool -- Layer 2 tool
    script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# Follow a message through Exchange Online -- where it went, whether it was
# delivered, and what happened if it was not.
#
# Search by sender, recipient, subject and date range, and filter by status so
# you can show only failures or only quarantined items. Results can be filtered
# further with the search box without re-running the trace.
#
# Quarantined messages can be released from here.
#
# Microsoft only keeps trace data for a limited window -- roughly the last ten
# days for the detailed view. For anything older you need a historical search,
# which Microsoft runs as a background job.
#
# WHEN THIS IS THE WRONG TOOL
#
# Message trace answers "what happened to this message" -- where it went, and
# why it did not arrive. It does not do threat hunting.
#
# For URL and attachment verdicts, campaign views, saved queries or bulk
# remediation, use THREAT EXPLORER instead. The button at the bottom opens it
# in your admin browser profile. It needs Defender for Office 365 Plan 2.
#
# Threat Explorer does not open carrying your filter. Select a row here, use
# Copy Sender or Copy Message-ID, and paste that into the portal's own filter.
#
# Read-only apart from releasing quarantine.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Actions - message trace (moved unchanged from Mailbox-Group-Inspector.ps1 lines 3100-3651)
function Invoke-MessageTrace {
    param(
        [string]$SenderAddress,
        [string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [string]$Subject,
        [string]$Status
    )
    $params = @{
        StartDate   = $StartDate
        EndDate     = $EndDate
        ResultSize  = 5000
        ErrorAction = 'Stop'
    }
    if ($SenderAddress)    { $params.SenderAddress    = $SenderAddress }
    if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }
    if ($Subject) {
        $params.Subject           = $Subject
        $params.SubjectFilterType = 'Contains'
    }
    if ($Status) { $params.Status = $Status }
    Invoke-EXOCommand -Name 'Get-MessageTraceV2' -Parameters $params
}
function Invoke-MessageTraceDetail {
    param(
        [Parameter(Mandatory)][guid]$MessageTraceId,
        [Parameter(Mandatory)][string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )
    Invoke-EXOCommand -Name 'Get-MessageTraceDetailV2' -Parameters @{
        MessageTraceId   = $MessageTraceId
        RecipientAddress = $RecipientAddress
        StartDate        = $StartDate
        EndDate           = $EndDate
        ErrorAction       = 'Stop'
    }
}
function Find-QuarantineMatch {
    param(
        [string]$MessageId,
        [string]$RecipientAddress,
        [datetime]$StartReceivedDate,
        [datetime]$EndReceivedDate
    )
    if (-not $MessageId) {
        throw "This trace result has no Message-ID value to match against quarantine -- cannot look this message up automatically."
    }
    @(Invoke-EXOCommand -Name 'Get-QuarantineMessage' -Parameters @{
        MessageId         = $MessageId
        RecipientAddress  = $RecipientAddress
        StartReceivedDate = $StartReceivedDate
        EndReceivedDate   = $EndReceivedDate
        ErrorAction       = 'Stop'
    })
}
function Invoke-QuarantineRelease {
    param([Parameter(Mandatory)]$QuarantineIdentity)
    Invoke-EXOCommand -Name 'Release-QuarantineMessage' -Parameters @{
        Identity     = $QuarantineIdentity
        ReleaseToAll = $true
        Confirm      = $false
        ErrorAction  = 'Stop'
    } | Out-Null
}
$script:TraceTimeZoneIds = [ordered]@{
    'Eastern'  = 'Eastern Standard Time'
    'Central'  = 'Central Standard Time'
    'Mountain' = 'Mountain Standard Time'
    'Pacific'  = 'Pacific Standard Time'
    'UTC'      = 'UTC'
}
function Get-SelectedTraceTimeZone {
    param([string]$Label)
    $id = $script:TraceTimeZoneIds[$Label]
    if (-not $id) { $id = 'UTC' }
    try {
        return [System.TimeZoneInfo]::FindSystemTimeZoneById($id)
    } catch {
        return [System.TimeZoneInfo]::Utc
    }
}
function ConvertTo-TraceDisplayTime {
    param($UtcValue, [System.TimeZoneInfo]$TimeZone)
    if (-not $UtcValue) { return $UtcValue }
    try {
        $forced = [DateTime]::SpecifyKind([DateTime]$UtcValue, [DateTimeKind]::Utc)
        return [System.TimeZoneInfo]::ConvertTimeFromUtc($forced, $TimeZone)
    } catch {
        return $UtcValue
    }
}
function Format-TraceDetail {
    param($Details, [System.TimeZoneInfo]$TimeZone)
    $sb = New-Object System.Text.StringBuilder
    $rows = @($Details)
    if ($rows.Count -eq 0) {
        [void]$sb.AppendLine("(no detail records returned)")
        return $sb.ToString()
    }
    foreach ($d in $rows) {
        $displayDate = if ($TimeZone) { ConvertTo-TraceDisplayTime -UtcValue $d.Date -TimeZone $TimeZone } else { $d.Date }
        [void]$sb.AppendLine("$displayDate  [$($d.Event)]  $($d.Action)")
        if ($d.Detail) { [void]$sb.AppendLine("    $($d.Detail)") }
    }
    return $sb.ToString()
}
function Show-MessageTraceForm {
    param([string]$SmtpAddress, [string]$Environment)
    $traceForm = New-Object System.Windows.Forms.Form
    $traceForm.Text = if ($SmtpAddress) { "Message Trace (from: $SmtpAddress)" } else { "Message Trace" }
    $traceForm.Size = New-Object System.Drawing.Size(1014, 754)
    $traceForm.StartPosition = 'Manual'
    $traceForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblFrom = New-Object System.Windows.Forms.Label
    $lblFrom.Text = "From:"
    $lblFrom.Location = New-Object System.Drawing.Point(12, 16)
    $lblFrom.AutoSize = $true
    $traceForm.Controls.Add($lblFrom)
    $txtFrom = New-Object System.Windows.Forms.TextBox
    $txtFrom.Location = New-Object System.Drawing.Point(55, 13)
    $txtFrom.Size = New-Object System.Drawing.Size(190, 20)
    $txtFrom.Font = New-Object System.Drawing.Font("Consolas", 9)
    if ($SmtpAddress) { $txtFrom.Text = $SmtpAddress }
    $traceForm.Controls.Add($txtFrom)
    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = "To:"
    $lblTo.Location = New-Object System.Drawing.Point(12, 40)
    $lblTo.AutoSize = $true
    $traceForm.Controls.Add($lblTo)
    $txtTo = New-Object System.Windows.Forms.TextBox
    $txtTo.Location = New-Object System.Drawing.Point(55, 37)
    $txtTo.Size = New-Object System.Drawing.Size(190, 20)
    $txtTo.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtTo)
    $lblTraceNote = New-Object System.Windows.Forms.Label
    $lblTraceNote.Text = "Read-only -- no objects are modified. Leave From/To blank to trace all mail in the date range. Exchange Online only (Get-MessageTraceV2 / Get-MessageTraceDetailV2), routed through Invoke-EXOCommand."
    $lblTraceNote.Location = New-Object System.Drawing.Point(264, 12)
    $lblTraceNote.Size = New-Object System.Drawing.Size(488, 46)
    $traceForm.Controls.Add($lblTraceNote)
    $lblTZ = New-Object System.Windows.Forms.Label
    $lblTZ.Text = "Show time in:"
    $lblTZ.Location = New-Object System.Drawing.Point(766, 14)
    $lblTZ.AutoSize = $true
    $traceForm.Controls.Add($lblTZ)
    $cboTimeZone = New-Object System.Windows.Forms.ComboBox
    $cboTimeZone.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Eastern', 'Central', 'Mountain', 'Pacific', 'UTC') | ForEach-Object { [void]$cboTimeZone.Items.Add($_) }
    $cboTimeZone.SelectedItem = 'Central'
    $cboTimeZone.Location = New-Object System.Drawing.Point(766, 34)
    $cboTimeZone.Size = New-Object System.Drawing.Size(200, 20)
    $traceForm.Controls.Add($cboTimeZone)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Status:"
    $lblStatus.Location = New-Object System.Drawing.Point(510, 72)
    $lblStatus.AutoSize = $true
    $traceForm.Controls.Add($lblStatus)
    $cboStatus = New-Object System.Windows.Forms.ComboBox
    $cboStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('All Messages', 'Delivered', 'Failed', 'Quarantined', 'FilteredAsSpam', 'Pending', 'Expanded', 'GettingStatus') |
        ForEach-Object { [void]$cboStatus.Items.Add($_) }
    $cboStatus.SelectedIndex = 0
    $cboStatus.Location = New-Object System.Drawing.Point(562, 69)
    $cboStatus.Size = New-Object System.Drawing.Size(180, 20)
    $traceForm.Controls.Add($cboStatus)
    $lblSubject = New-Object System.Windows.Forms.Label
    $lblSubject.Text = "Subject contains (optional):"
    $lblSubject.Location = New-Object System.Drawing.Point(12, 100)
    $lblSubject.AutoSize = $true
    $traceForm.Controls.Add($lblSubject)
    $txtSubject = New-Object System.Windows.Forms.TextBox
    $txtSubject.Location = New-Object System.Drawing.Point(270, 97)
    $txtSubject.Size = New-Object System.Drawing.Size(220, 20)
    $txtSubject.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtSubject)
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = "Search From:"
    $lblStart.Location = New-Object System.Drawing.Point(12, 129)
    $lblStart.AutoSize = $true
    $traceForm.Controls.Add($lblStart)
    $dtpStart = New-Object System.Windows.Forms.DateTimePicker
    $dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtpStart.Location = New-Object System.Drawing.Point(100, 125)
    $dtpStart.Size = New-Object System.Drawing.Size(120, 20)
    $dtpStart.MinDate = (Get-Date).Date.AddDays(-30)
    $dtpStart.MaxDate = (Get-Date).Date
    $dtpStart.Value   = (Get-Date).Date.AddDays(-7)
    $traceForm.Controls.Add($dtpStart)
    $lblEnd = New-Object System.Windows.Forms.Label
    $lblEnd.Text = "Through:"
    $lblEnd.Location = New-Object System.Drawing.Point(240, 129)
    $lblEnd.AutoSize = $true
    $traceForm.Controls.Add($lblEnd)
    $dtpEnd = New-Object System.Windows.Forms.DateTimePicker
    $dtpEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtpEnd.Location = New-Object System.Drawing.Point(300, 125)
    $dtpEnd.Size = New-Object System.Drawing.Size(120, 20)
    $traceForm.Controls.Add($dtpEnd)
    $dtpStart.Add_ValueChanged({
        $maxEnd = $dtpStart.Value.Date.AddDays(9)
        $today  = (Get-Date).Date
        if ($maxEnd -gt $today) { $maxEnd = $today }
        $dtpEnd.MinDate = $dtpStart.Value.Date
        $dtpEnd.MaxDate = $maxEnd
        if ($dtpEnd.Value -lt $dtpStart.Value.Date) { $dtpEnd.Value = $dtpStart.Value.Date }
        if ($dtpEnd.Value -gt $maxEnd) { $dtpEnd.Value = $maxEnd }
    })
    $dtpEnd.MinDate = $dtpStart.Value.Date
    $initialMaxEnd = $dtpStart.Value.Date.AddDays(9)
    if ($initialMaxEnd -gt (Get-Date).Date) { $initialMaxEnd = (Get-Date).Date }
    $dtpEnd.MaxDate = $initialMaxEnd
    $dtpEnd.Value   = $initialMaxEnd
    $btnTraceRun = New-Object System.Windows.Forms.Button
    $btnTraceRun.Text = "Trace"
    $btnTraceRun.Location = New-Object System.Drawing.Point(440, 123)
    $btnTraceRun.Size = New-Object System.Drawing.Size(90, 26)
    $traceForm.Controls.Add($btnTraceRun)
    $lblTraceStatus = New-Object System.Windows.Forms.Label
    $lblTraceStatus.Text = ""
    $lblTraceStatus.Location = New-Object System.Drawing.Point(12, 156)
    $lblTraceStatus.Size = New-Object System.Drawing.Size(974, 40)
    $traceForm.Controls.Add($lblTraceStatus)
    $lblSearchGrid = New-Object System.Windows.Forms.Label
    $lblSearchGrid.Text = "Filter loaded results:"
    $lblSearchGrid.Location = New-Object System.Drawing.Point(12, 204)
    $lblSearchGrid.AutoSize = $true
    $traceForm.Controls.Add($lblSearchGrid)
    $txtSearchGrid = New-Object System.Windows.Forms.TextBox
    $txtSearchGrid.Location = New-Object System.Drawing.Point(150, 200)
    $txtSearchGrid.Size = New-Object System.Drawing.Size(280, 22)
    $txtSearchGrid.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtSearchGrid)
    $btnSearchGrid = New-Object System.Windows.Forms.Button
    $btnSearchGrid.Text = "Search"
    $btnSearchGrid.Location = New-Object System.Drawing.Point(440, 198)
    $btnSearchGrid.Size = New-Object System.Drawing.Size(80, 26)
    $traceForm.Controls.Add($btnSearchGrid)
    $btnClearSearchGrid = New-Object System.Windows.Forms.Button
    $btnClearSearchGrid.Text = "Clear Filter"
    $btnClearSearchGrid.Location = New-Object System.Drawing.Point(526, 198)
    $btnClearSearchGrid.Size = New-Object System.Drawing.Size(90, 26)
    $traceForm.Controls.Add($btnClearSearchGrid)
    $grdTrace = New-Object System.Windows.Forms.DataGridView
    $grdTrace.Location = New-Object System.Drawing.Point(12, 234)
    $grdTrace.Size = New-Object System.Drawing.Size(974, 260)
    $grdTrace.ReadOnly = $true
    $grdTrace.AllowUserToAddRows = $false
    $grdTrace.AllowUserToDeleteRows = $false
    $grdTrace.AllowUserToResizeRows = $false
    $grdTrace.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdTrace.MultiSelect = $false
    $grdTrace.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdTrace.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdTrace.Columns.Add('Received', 'Received')
    [void]$grdTrace.Columns.Add('Sender', 'Sender Address')
    [void]$grdTrace.Columns.Add('Recipient', 'Recipient Address')
    [void]$grdTrace.Columns.Add('Subject', 'Subject')
    [void]$grdTrace.Columns.Add('Status', 'Status')
    [void]$grdTrace.Columns.Add('Key', 'Key')
    $grdTrace.Columns['Received'].FillWeight  = 15
    $grdTrace.Columns['Sender'].FillWeight    = 20
    $grdTrace.Columns['Recipient'].FillWeight = 20
    $grdTrace.Columns['Subject'].FillWeight   = 30
    $grdTrace.Columns['Status'].FillWeight    = 15
    $grdTrace.Columns['Received'].ValueType   = [DateTime]
    $grdTrace.Columns['Key'].Visible          = $false
    foreach ($col in $grdTrace.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
    $traceForm.Controls.Add($grdTrace)
    $lblDetail = New-Object System.Windows.Forms.Label
    $lblDetail.Text = "Detail (double-click a row above to load it here):"
    $lblDetail.Location = New-Object System.Drawing.Point(12, 502)
    $lblDetail.AutoSize = $true
    $traceForm.Controls.Add($lblDetail)
    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.Multiline = $true
    $txtDetail.ReadOnly = $true
    $txtDetail.ScrollBars = "Vertical"
    $txtDetail.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtDetail.Location = New-Object System.Drawing.Point(12, 522)
    $txtDetail.Size = New-Object System.Drawing.Size(974, 130)
    $txtDetail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtDetail)
    $btnRelease = New-Object System.Windows.Forms.Button
    $btnRelease.Text = "Release from Quarantine"
    $btnRelease.Location = New-Object System.Drawing.Point(12, 662)
    $btnRelease.Size = New-Object System.Drawing.Size(160, 28)
    $btnRelease.Enabled = $false
    $traceForm.Controls.Add($btnRelease)
    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = "Export to CSV"
    $btnExportCsv.Location = New-Object System.Drawing.Point(180, 662)
    $btnExportCsv.Size = New-Object System.Drawing.Size(110, 28)
    $btnExportCsv.Enabled = $false
    $traceForm.Controls.Add($btnExportCsv)
    # --- Threat Explorer hand-off -----------------------------------------
    #
    # Message trace answers "what happened to this message". It does not do
    # threat hunting -- URL and attachment verdicts, campaign views, saved
    # queries, bulk remediation. Those live in Threat Explorer, which needs
    # Defender for Office 365 Plan 2.
    #
    # Ported from the paused browser suite, which explained the distinction on
    # screen. Knowing which of the two tools answers your question is most of
    # the work, and it was the one useful thing that version had over this one.
    $btnExplorer = New-Object System.Windows.Forms.Button
    $btnExplorer.Text = "Threat Explorer"
    $btnExplorer.Location = New-Object System.Drawing.Point(300, 662)
    $btnExplorer.Size = New-Object System.Drawing.Size(150, 28)
    $traceForm.Controls.Add($btnExplorer)

    # Copy, rather than a pre-filled URL, and this is deliberate.
    #
    # The browser suite had a "Prepopulated" flag and a message reading "Threat
    # Explorer opened with your filter applied". It could never fire: the URL
    # template was the bare portal address with no substitution tokens, and the
    # function that would have replaced it was exported and never called. The
    # copy buttons were what actually worked there, so they are what is ported.
    #
    # Threat Explorer v3's filter parameters are not documented in a form worth
    # depending on. Copy and paste is unglamorous and correct.
    $btnCopySender = New-Object System.Windows.Forms.Button
    $btnCopySender.Text = "Copy Sender"
    $btnCopySender.Location = New-Object System.Drawing.Point(458, 662)
    $btnCopySender.Size = New-Object System.Drawing.Size(110, 28)
    $btnCopySender.Enabled = $false
    $traceForm.Controls.Add($btnCopySender)

    $btnCopyMsgId = New-Object System.Windows.Forms.Button
    $btnCopyMsgId.Text = "Copy Message-ID"
    $btnCopyMsgId.Location = New-Object System.Drawing.Point(576, 662)
    $btnCopyMsgId.Size = New-Object System.Drawing.Size(130, 28)
    $btnCopyMsgId.Enabled = $false
    $traceForm.Controls.Add($btnCopyMsgId)

    $lblCopyNote = New-Object System.Windows.Forms.Label
    $lblCopyNote.Tag = 'NoTheme'
    $lblCopyNote.Location = New-Object System.Drawing.Point(714, 666)
    $lblCopyNote.Size = New-Object System.Drawing.Size(166, 20)
    $lblCopyNote.AutoEllipsis = $true
    $lblCopyNote.ForeColor = [System.Drawing.Color]::DimGray
    $traceForm.Controls.Add($lblCopyNote)

    $tipTrace = New-Object System.Windows.Forms.ToolTip
    $tipTrace.SetToolTip($btnExplorer, @"
Opens Threat Explorer in the Microsoft 365 Defender portal.

Use it instead of message trace for: URL and attachment verdicts,
campaign views, saved queries, and bulk remediation.

Needs Defender for Office 365 Plan 2.

It does NOT open pre-filtered -- copy the sender or Message-ID
first and paste it into the portal's own filter.
"@)
    $tipTrace.SetToolTip($btnCopySender, 'Copy the selected row''s sender, to paste into a portal filter.')
    $tipTrace.SetToolTip($btnCopyMsgId,  'Copy the selected row''s Message-ID, to paste into a portal filter.')

    $btnTraceClose = New-Object System.Windows.Forms.Button
    $btnTraceClose.Text = "Close"
    $btnTraceClose.Location = New-Object System.Drawing.Point(886, 662)
    $btnTraceClose.Size = New-Object System.Drawing.Size(100, 28)
    $traceForm.Controls.Add($btnTraceClose)
    $script:CurrentTraceResults        = @()
    $script:CurrentTraceDisplayResults = @()
    $script:CurrentTraceDetailRaw       = $null
    $script:TraceSearchStart            = $null
    $script:TraceSearchEnd              = $null
    function Get-FilteredTraceResults {
        param([string]$SearchText)
        if ([string]::IsNullOrWhiteSpace($SearchText)) { return $script:CurrentTraceResults }
        $needle = "*$($SearchText.Trim())*"
        $script:CurrentTraceResults | Where-Object {
            ($_.SenderAddress    -and $_.SenderAddress    -like $needle) -or
            ($_.RecipientAddress -and $_.RecipientAddress -like $needle) -or
            ($_.Subject          -and $_.Subject          -like $needle) -or
            ($_.Status           -and $_.Status           -like $needle)
        }
    }
    function Update-TraceGridDisplay {
        $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
        $grdTrace.Columns['Received'].HeaderText = "Received ($($cboTimeZone.SelectedItem))"
        $grdTrace.Rows.Clear()
        foreach ($r in $script:CurrentTraceDisplayResults) {
            $rowKey = "$($r.MessageTraceId)|$($r.RecipientAddress)"
            $displayReceived = ConvertTo-TraceDisplayTime -UtcValue $r.Received -TimeZone $tz
            [void]$grdTrace.Rows.Add($displayReceived, $r.SenderAddress, $r.RecipientAddress, $r.Subject, $r.Status, $rowKey)
        }
    }
    $cboTimeZone.Add_SelectedIndexChanged({
        Update-TraceGridDisplay
        if ($script:CurrentTraceDetailRaw) {
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $txtDetail.Text = Format-TraceDetail -Details $script:CurrentTraceDetailRaw -TimeZone $tz
        }
    })
    $btnTraceRun.Add_Click({
        $btnTraceRun.Enabled = $false
        $btnRelease.Enabled = $false
        $btnExportCsv.Enabled = $false
        $txtDetail.Text = ""
        $txtSearchGrid.Text = ""
        $grdTrace.Rows.Clear()
        $script:CurrentTraceResults = @()
        $script:CurrentTraceDisplayResults = @()
        $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
        $lblTraceStatus.Text = "Running trace..."
        try {
            $fromAddr = $txtFrom.Text.Trim()
            $toAddr   = $txtTo.Text.Trim()
            $subject = $txtSubject.Text.Trim()
            $sd = $dtpStart.Value.Date
            $ed = $dtpEnd.Value.Date.AddDays(1).AddSeconds(-1)
            $traceParams = @{ StartDate = $sd; EndDate = $ed }
            if ($fromAddr) { $traceParams.SenderAddress    = $fromAddr }
            if ($toAddr)   { $traceParams.RecipientAddress = $toAddr }
            if ($subject) { $traceParams.Subject = $subject }
            if ($cboStatus.SelectedItem -and $cboStatus.SelectedItem -ne 'All Messages') {
                $traceParams.Status = $cboStatus.SelectedItem
            }
            $results = @(Invoke-MessageTrace @traceParams)
            $script:CurrentTraceResults = $results
            $script:CurrentTraceDisplayResults = $results
            $script:CurrentTraceDetailRaw = $null
            $script:TraceSearchStart = $sd
            $script:TraceSearchEnd   = $ed
            Update-TraceGridDisplay
            $btnExportCsv.Enabled = ($results.Count -gt 0)
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $statusNote = if ($traceParams.Status) { "  Status: $($traceParams.Status)." } else { "" }
            $lblTraceStatus.Text = "$($results.Count) result(s). $sd through $ed.$statusNote"
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceV2'
        } finally {
            $btnTraceRun.Enabled = $true
        }
    })
    $btnSearchGrid.Add_Click({
        $script:CurrentTraceDisplayResults = @(Get-FilteredTraceResults -SearchText $txtSearchGrid.Text)
        Update-TraceGridDisplay
        $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
        if ([string]::IsNullOrWhiteSpace($txtSearchGrid.Text)) {
            $lblTraceStatus.Text = "$($script:CurrentTraceResults.Count) result(s)."
        } else {
            $lblTraceStatus.Text = "$($script:CurrentTraceDisplayResults.Count) of $($script:CurrentTraceResults.Count) result(s) match '$($txtSearchGrid.Text.Trim())'."
        }
    })
    $txtSearchGrid.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $btnSearchGrid.PerformClick()
            $e.SuppressKeyPress = $true
        }
    })
    $btnClearSearchGrid.Add_Click({
        $txtSearchGrid.Text = ''
        $script:CurrentTraceDisplayResults = $script:CurrentTraceResults
        Update-TraceGridDisplay
        $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
        $lblTraceStatus.Text = "$($script:CurrentTraceResults.Count) result(s)."
    })
    function Find-TraceResultByKey {
        param([string]$RowKey)
        $script:CurrentTraceResults | Where-Object { "$($_.MessageTraceId)|$($_.RecipientAddress)" -eq $RowKey } | Select-Object -First 1
    }
    $grdTrace.Add_SelectionChanged({
        $has = ($grdTrace.SelectedRows.Count -gt 0)
        $btnRelease.Enabled   = $has
        $btnCopySender.Enabled = $has
        # Message-ID is not always present on a trace row, and a button that
        # copies an empty string is worse than one that is greyed: the first
        # looks like it worked.
        $btnCopyMsgId.Enabled = $false
        if ($has) {
            $rec = Find-TraceResultByKey -RowKey $grdTrace.SelectedRows[0].Cells['Key'].Value
            if ($rec -and "$($rec.MessageId)".Trim()) { $btnCopyMsgId.Enabled = $true }
        }
        $lblCopyNote.Text = ''
    })

    function Copy-TraceValue {
        param([string]$Value, [string]$Label)
        if (-not "$Value".Trim()) {
            $lblCopyNote.Text = "No $Label on that row."
            return
        }
        try {
            [System.Windows.Forms.Clipboard]::SetText("$Value")
            $lblCopyNote.Text = "$Label copied."
        } catch {
            # Clipboard access fails if another process holds it open. Say so
            # rather than leaving the user to wonder why nothing pasted.
            $lblCopyNote.Text = "Copy failed -- clipboard busy."
        }
    }

    $btnCopySender.Add_Click({
        if ($grdTrace.SelectedRows.Count -eq 0) { return }
        $rec = Find-TraceResultByKey -RowKey $grdTrace.SelectedRows[0].Cells['Key'].Value
        if (-not $rec) { $lblCopyNote.Text = 'Row no longer matches -- re-run Trace.'; return }
        Copy-TraceValue -Value $rec.SenderAddress -Label 'Sender'
    })

    $btnCopyMsgId.Add_Click({
        if ($grdTrace.SelectedRows.Count -eq 0) { return }
        $rec = Find-TraceResultByKey -RowKey $grdTrace.SelectedRows[0].Cells['Key'].Value
        if (-not $rec) { $lblCopyNote.Text = 'Row no longer matches -- re-run Trace.'; return }
        Copy-TraceValue -Value $rec.MessageId -Label 'Message-ID'
    })

    $btnExplorer.Add_Click({
        # Admin browser profile, same as every other admin URL in the toolkit.
        # Threat Explorer needs the elevated account; opening it in a personal
        # profile produces a sign-in prompt that looks like a permissions fault.
        $url = 'https://security.microsoft.com/threatexplorerv3'
        try {
            $st = Start-UrlInBrowserProfile -Url $url -Purpose 'Admin'
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $lblTraceStatus.Text = "Threat Explorer opened ($st). It does not carry your filter -- use the copy buttons and paste into its own filter."
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = "Could not open Threat Explorer: $($_.Exception.Message)"
        }
    })
    $grdTrace.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $rowKey = $grdTrace.Rows[$e.RowIndex].Cells['Key'].Value
        $rec = Find-TraceResultByKey -RowKey $rowKey
        if (-not $rec) {
            $txtDetail.Text = "Could not match that row back to a trace result -- run Trace again."
            return
        }
        $txtDetail.Text = "Loading..."
        try {
            $details = Invoke-MessageTraceDetail -MessageTraceId $rec.MessageTraceId -RecipientAddress $rec.RecipientAddress `
                -StartDate $script:TraceSearchStart -EndDate $script:TraceSearchEnd
            $script:CurrentTraceDetailRaw = $details
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $txtDetail.Text = Format-TraceDetail -Details $details -TimeZone $tz
        } catch {
            $script:CurrentTraceDetailRaw = $null
            $txtDetail.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceDetailV2'
        }
    })
    $btnRelease.Add_Click({
        if ($grdTrace.SelectedRows.Count -eq 0) { return }
        $rowKey = $grdTrace.SelectedRows[0].Cells['Key'].Value
        $rec = Find-TraceResultByKey -RowKey $rowKey
        if (-not $rec) {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = "Could not match that row back to a trace result -- run Trace again."
            return
        }
        $confirmSearch = [System.Windows.Forms.MessageBox]::Show(
            "Look up quarantine entries for this message?`nSender: $($rec.SenderAddress)`nRecipient: $($rec.RecipientAddress)`nSubject: $($rec.Subject)",
            "Confirm Quarantine Lookup", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirmSearch -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnRelease.Enabled = $false
        try {
            $matches = @(Find-QuarantineMatch -MessageId $rec.MessageId -RecipientAddress $rec.RecipientAddress `
                -StartReceivedDate $script:TraceSearchStart -EndReceivedDate $script:TraceSearchEnd)
            if ($matches.Count -eq 0) {
                $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
                $lblTraceStatus.Text = "No quarantine entry found for this message (it may not have been quarantined, may already be released, or may have expired -- quarantine data is only retained for a limited time)."
                return
            }
            foreach ($qmsg in $matches) {
                $confirmRelease = [System.Windows.Forms.MessageBox]::Show(
                    "Quarantine match found:`nType: $($qmsg.Type)   Policy: $($qmsg.PolicyName)`nReceived: $($qmsg.ReceivedTime)   Release status: $($qmsg.ReleaseStatus)`n`nRelease this message to ALL original recipients?",
                    "Confirm Release", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($confirmRelease -ne [System.Windows.Forms.DialogResult]::Yes) { continue }
                Invoke-QuarantineRelease -QuarantineIdentity $qmsg.Identity
                Write-CompletedLog -Task "Released QuarantineMessage $($rec.RecipientAddress) ($Environment; sender: $($rec.SenderAddress); subject: $($rec.Subject))"
            }
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $lblTraceStatus.Text = "$($matches.Count) quarantine match(es) processed."
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Quarantine release'
        } finally {
            $btnRelease.Enabled = $true
        }
    })
    $btnExportCsv.Add_Click({
        if ($script:CurrentTraceDisplayResults.Count -eq 0) { $lblTraceStatus.Text = "No results to export."; return }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $filePart = if ($SmtpAddress) { ($SmtpAddress -replace '[^a-zA-Z0-9]','_') } else { 'All' }
        $sfd.FileName = "MessageTrace_${filePart}_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($traceForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $script:CurrentTraceDisplayResults |
                Select-Object @{Name = "Received ($($cboTimeZone.SelectedItem))"; Expression = { ConvertTo-TraceDisplayTime -UtcValue $_.Received -TimeZone $tz }},
                              SenderAddress, RecipientAddress, Subject, Status, MessageTraceId |
                Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $lblTraceStatus.Text = "Exported $($script:CurrentTraceDisplayResults.Count) row(s) to $($sfd.FileName)."
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = "Export failed: $($_.Exception.Message)"
        }
    })
    $btnTraceClose.Add_Click({ $traceForm.Close() })
    $traceForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $traceForm
    Set-FormScreenCorner -Form $traceForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $traceForm -ScriptPath $PSCommandPath -Title 'Message Trace' | Out-Null

    [void]$traceForm.ShowDialog()
}
#endregion

Show-MessageTraceForm -SmtpAddress $SmtpAddress -Environment $Environment
