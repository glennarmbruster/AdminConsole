#Requires -Version 5.1
<#
    MessageTrace.ps1
    Standalone Message Trace / Quarantine Release tool -- Layer 2 tool
    script.

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    2026-09-16  The quarantine match list is $qmatches, not $matches, which is
                an automatic variable any -match in scope would overwrite.
    2026-09-16  The allow-sender prompt is gated on the same spam-only test as
                false-positive reporting, so releasing a phish hold no longer
                offers a standing allow for its sender.

    2026-09-14  Release can also report the message to Microsoft as a false
                positive (Release-QuarantineMessage -ReportFalsePositive).
                Offered only for spam-type quarantine entries, which is the
                only type the switch is documented for. A second prompt on
                release offers -AllowSender (stop quarantining this sender).
#>
# TOOLKIT-HELP-BEGIN
# Follow a message through Exchange Online -- where it went, whether it was
# delivered, and what happened if it was not.
#
# Search by sender, recipient, subject and date range, and filter by status so
# you can show only failures or only quarantined items. Results can be filtered
# further with the search box without re-running the trace.
#
# Quarantined messages can be released from here. When the entry is spam-type,
# the release prompt also offers to report it to Microsoft as a false positive
# (good mail wrongly quarantined), which feeds the spam filter. Phish and
# malware holds cannot be reported this way -- use the Submissions page in the
# Defender portal for those. A second prompt offers to ALLOW THE SENDER from
# now on (Release-QuarantineMessage -AllowSender): a standing change that stops
# future spam-quarantining of that address. It does not override a transport
# rule or a blocked-sender entry. Default is No.
#
# DOUBLE-CLICK A RESULT for the full event timeline of that one message --
# every hop, and any TRANSPORT RULES that acted on it, named. Rule hits are
# listed at the top of the pane and marked with >> in the timeline.
#
# That answers "why did this get redirected / stamped / blocked", which the
# result grid alone cannot: the grid shows the outcome, the detail shows the
# cause. Transport Rules then shows that rule's priority and what it sets --
# priority being what decides which of several matching rules won.
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

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

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
    # 5000 is a deliberate cap, so Exchange warns "There are more results
    # available" whenever a busy window fills it. Silenced because that warning
    # goes to the PowerShell stream, where nobody using this window will see it
    # -- the honest signal is the count on screen, and exactly 5000 means there
    # was more.
    #
    # WarningAction sits NEXT TO ResultSize deliberately. Comment stripping
    # blanks lines rather than removing them, so an explanation in between puts
    # the two eight lines apart and the check that looks +/-3 lines never sees
    # it. A rule you satisfy has to be satisfied where it looks.
    $params = @{
        StartDate   = $StartDate
        EndDate     = $EndDate
        ResultSize    = 5000
        WarningAction = 'SilentlyContinue'
        ErrorAction   = 'Stop'
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
function Test-QuarantineReportable {
    <#
        Can this quarantine entry be reported to Microsoft as a false positive
        on release? Release-QuarantineMessage -ReportFalsePositive is
        documented as "available only for quarantined spam messages", so it is
        offered for the spam types (Spam, HighConfSpam, Bulk) and not for
        phish, malware or transport-rule holds. Those go through the
        Submissions page in the Defender portal, where an analyst looks at
        them; the switch is a lighter signal that only feeds the spam filter.
    #>
    param([Parameter(Mandatory)]$QuarantineMessage)
    $t = "$($QuarantineMessage.Type)"
    return ($t -match '(?i)spam|bulk')
}
function Invoke-QuarantineRelease {
    param(
        [Parameter(Mandatory)]$QuarantineIdentity,
        [switch]$ReportFalsePositive,
        [switch]$AllowSender
    )
    $p = @{
        Identity     = $QuarantineIdentity
        ReleaseToAll = $true
        Confirm      = $false
        ErrorAction  = 'Stop'
    }
    # BOTH SWITCHES ARE SEPARATE DECISIONS from the release. ReportFalsePositive
    # is one way: once reported, Microsoft has the sample. AllowSender is a
    # STANDING change: future mail from that sender skips the spam quarantine,
    # which outlives this message and this ticket. Each is only ever added
    # when the operator said yes to its own question, never as a default.
    if ($ReportFalsePositive) { $p['ReportFalsePositive'] = $true }
    if ($AllowSender)         { $p['AllowSender']         = $true }
    Invoke-EXOCommand -Name 'Release-QuarantineMessage' -Parameters $p | Out-Null
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
function Get-TraceRuleHits {
    <#
        The transport rule NAMES out of one message-trace detail row.

        THE NAME IS IN Data, NOT IN Detail, AND Data IS XML. Get-MessageTraceDetailV2
        returns Event / Action / Detail as plain text and puts the structured payload
        in Data -- so a formatter that prints the first three shows THAT a rule fired
        and never says WHICH ONE. That was this pane for its whole life.

        PARSED THREE WAYS, DELIBERATELY. The attribute carrying the name has not been
        stable across message-trace versions, and this runs against whatever the
        service returns today rather than against a documented shape. Well-formed XML
        is tried first, then a regex over the raw text, and anything still
        unrecognised is SHOWN rather than swallowed -- an unparsed payload is a thing
        to look at, not a thing to hide. If you see a DATA (unparsed) line mentioning
        a rule, the attribute name has changed and this function needs the new one.

        Never throws. A malformed payload costs the rule name on one row, and the
        timeline is still worth reading without it.
    #>
    param($DetailRow)

    $raw = "$($DetailRow.Data)"
    if (-not $raw.Trim()) { return @() }
    $names = New-Object System.Collections.Generic.List[string]

    try {
        $xml = [xml]$raw
        foreach ($node in $xml.SelectNodes('//*')) {
            foreach ($attr in @($node.Attributes)) {
                if ($attr.Name -match '(?i)(ETRName|TransportRuleName|RuleName)' -and "$($attr.Value)".Trim()) {
                    $names.Add("$($attr.Value)".Trim())
                }
            }
        }
    } catch {
        # Not well-formed. The regex below is the fallback, not an error path.
    }

    if ($names.Count -eq 0) {
        foreach ($m in [regex]::Matches($raw, '(?i)(?:ETRName|TransportRuleName|RuleName)\s*=\s*"([^"]+)"')) {
            if ($m.Groups[1].Value.Trim()) { $names.Add($m.Groups[1].Value.Trim()) }
        }
    }

    # .ToArray() FIRST. A Generic.List piped straight through @() can throw on 5.1.
    return @($names.ToArray() | Select-Object -Unique)
}

function Format-TraceDetail {
    param($Details, [System.TimeZoneInfo]$TimeZone)
    $sb = New-Object System.Text.StringBuilder
    $rows = @($Details)
    if ($rows.Count -eq 0) {
        [void]$sb.AppendLine("(no detail records returned)")
        return $sb.ToString()
    }

    # RULES FIRST, TIMELINE SECOND. "Which rule did this to my message" is the
    # question this pane gets opened to answer, and finding it by reading a
    # timeline means reading the timeline. Summarise, then show the working.
    $allRules = New-Object System.Collections.Generic.List[string]
    foreach ($d in $rows) {
        foreach ($n in @(Get-TraceRuleHits -DetailRow $d)) { $allRules.Add($n) }
    }
    $uniqueRules = @($allRules.ToArray() | Select-Object -Unique)

    [void]$sb.AppendLine("=== Transport rules that touched this message ===")
    if ($uniqueRules.Count -eq 0) {
        [void]$sb.AppendLine("  None reported.")
    } else {
        foreach ($r in $uniqueRules) { [void]$sb.AppendLine("  * $r") }
        [void]$sb.AppendLine("  (Transport Rules shows each one's priority and what it sets --")
        [void]$sb.AppendLine("   priority is what decides which rule won.)")
    }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("=== Timeline ===")

    foreach ($d in $rows) {
        $displayDate = if ($TimeZone) { ConvertTo-TraceDisplayTime -UtcValue $d.Date -TimeZone $TimeZone } else { $d.Date }
        $hits   = @(Get-TraceRuleHits -DetailRow $d)
        $marker = if ($hits.Count -gt 0) { '>>' } else { '  ' }
        [void]$sb.AppendLine("$marker $displayDate  [$($d.Event)]  $($d.Action)")
        if ($d.Detail) { [void]$sb.AppendLine("        $($d.Detail)") }
        foreach ($h in $hits) { [void]$sb.AppendLine("        RULE: $h") }

        # SHOWN, NOT SWALLOWED. A payload that mentions a rule and parsed to
        # nothing means the attribute name moved -- see Get-TraceRuleHits.
        if ($hits.Count -eq 0 -and "$($d.Data)".Trim() -and "$($d.Data)" -match '(?i)rule') {
            [void]$sb.AppendLine("        DATA (unparsed): $("$($d.Data)".Trim())")
        }
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
    $lblDetail.Text = "Detail (double-click a row above) -- event timeline, and any transport rules that touched the message:"
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
        $fromAddr = $txtFrom.Text.Trim()
        $toAddr   = $txtTo.Text.Trim()
        $subject = $txtSubject.Text.Trim()
        $sd = $dtpStart.Value.Date
        $ed = $dtpEnd.Value.Date.AddDays(1).AddSeconds(-1)
        # Same parameters Invoke-MessageTrace builds (kept above for the
        # capability scanner and for anything else that calls it); assembled
        # here because the block below travels to the background worker as
        # TEXT and sees only what it is handed.
        $traceParams = @{
            StartDate = $sd; EndDate = $ed
            ResultSize = 5000; WarningAction = 'SilentlyContinue'; ErrorAction = 'Stop'
        }
        if ($fromAddr) { $traceParams.SenderAddress    = $fromAddr }
        if ($toAddr)   { $traceParams.RecipientAddress = $toAddr }
        if ($subject)  { $traceParams.Subject = $subject; $traceParams.SubjectFilterType = 'Contains' }
        if ($cboStatus.SelectedItem -and $cboStatus.SelectedItem -ne 'All Messages') {
            $traceParams.Status = $cboStatus.SelectedItem
        }
        $script:TraceSearchStart = $sd
        $script:TraceSearchEnd   = $ed
        $statusNote = if ($traceParams.Status) { "  Status: $($traceParams.Status)." } else { "" }

        # OFF THE UI THREAD when the process allows it (Start-ToolJob in
        # Common.ps1 explains the test and the fallback). A trace over a busy
        # window is the longest read in the toolkit, and this window used to
        # look hung for the duration.
        Start-ToolJob -Form $traceForm -Text "Tracing $sd to $ed ..." -Script {
            param($p)
            @(Invoke-EXOCommand -Name 'Get-MessageTraceV2' -Parameters $p)
        } -ArgumentList $traceParams -State @{ Sd = $sd; Ed = $ed; Note = $statusNote } -OnDone {
            param($r, $st)
            # Controls and $script: state are reachable here (the form's
            # message loop is still inside Show-MessageTraceForm); the click
            # handler's own locals are not, hence -State.
            try {
                if ($r.Failed) {
                    $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
                    $lblTraceStatus.Text = "Get-MessageTraceV2 failed: $($r.Error)"
                    return
                }
                $results = @($r.Output)
                $script:CurrentTraceResults = $results
                $script:CurrentTraceDisplayResults = $results
                $script:CurrentTraceDetailRaw = $null
                Update-TraceGridDisplay
                $btnExportCsv.Enabled = ($results.Count -gt 0)
                $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
                $lblTraceStatus.Text = "$($results.Count) result(s). $($st.Sd) through $($st.Ed).$($st.Note)"
                Set-ToolStatus -Form $traceForm -Text "$($results.Count) result(s)$(if ($r.Inline) { '' } else { ' (background)' })"
            } finally {
                $btnTraceRun.Enabled = $true
            }
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
        $confirmSearch = Show-ToolDialog(
            "Look up quarantine entries for this message?`nSender: $($rec.SenderAddress)`nRecipient: $($rec.RecipientAddress)`nSubject: $($rec.Subject)",
            "Confirm Quarantine Lookup", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirmSearch -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnRelease.Enabled = $false
        try {
            # NOT $matches: that is a PowerShell AUTOMATIC variable, rewritten by
            # any -match anywhere in this scope. Nothing between here and the last
            # read uses -match today, so it worked -- but it is a trap for the next
            # edit to this handler, and the rename costs nothing.
            $qmatches = @(Find-QuarantineMatch -MessageId $rec.MessageId -RecipientAddress $rec.RecipientAddress `
                -StartReceivedDate $script:TraceSearchStart -EndReceivedDate $script:TraceSearchEnd)
            if ($qmatches.Count -eq 0) {
                $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
                $lblTraceStatus.Text = "No quarantine entry found for this message (it may not have been quarantined, may already be released, or may have expired -- quarantine data is only retained for a limited time)."
                return
            }
            $released = 0; $reported = 0; $notReportable = 0; $allowed = 0
            foreach ($qmsg in $qmatches) {
                $head = "Quarantine match found:`nType: $($qmsg.Type)   Policy: $($qmsg.PolicyName)`nReceived: $($qmsg.ReceivedTime)   Release status: $($qmsg.ReleaseStatus)`n`n"
                $report = $false
                if (Test-QuarantineReportable -QuarantineMessage $qmsg) {
                    # ONE QUESTION, THREE ANSWERS, so 'release' and 'release and
                    # tell Microsoft it was wrong' are the same click count as
                    # before and the report is never a surprise default.
                    $a = Show-ToolDialog(
                        ($head + "Release this message to ALL original recipients?`n`n" +
                         "YES     release AND report it to Microsoft as a false positive`n" +
                         "            (good mail wrongly quarantined -- trains the spam filter)`n" +
                         "NO      release only`n" +
                         "CANCEL  leave it in quarantine"),
                        "Confirm Release", 'YesNoCancel', 'Warning', 'Button2')
                    if ($a -eq [System.Windows.Forms.DialogResult]::Cancel) { continue }
                    $report = ($a -eq [System.Windows.Forms.DialogResult]::Yes)
                } else {
                    $notReportable++
                    $a = Show-ToolDialog(
                        ($head + "Release this message to ALL original recipients?`n`n" +
                         "(False-positive reporting is not available for type '$($qmsg.Type)' -- Microsoft only accepts it " +
                         "for spam. Report phish or malware misses through the Submissions page in the Defender portal.)"),
                        "Confirm Release", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    if ($a -ne [System.Windows.Forms.DialogResult]::Yes) { continue }
                }
                # ALLOW THE SENDER? Asked separately and after the release
                # decision, because it is the one answer here that outlives the
                # message: it is a standing allow, not part of this release.
                # Default is No.
                #
                # SPAM TYPES ONLY, the same gate as false-positive reporting.
                # This used to be asked for every type, so releasing something
                # Defender held as PHISH offered a standing allow for its
                # sender on the same screen -- the one context where a standing
                # allow is least defensible, and the reason the reportable test
                # exists. Not offering it is not a limitation of the tool: an
                # allow for a phish or malware sender belongs in the Defender
                # portal, deliberately, with the rest of the tenant allow list.
                $allow = $false
                if (Test-QuarantineReportable -QuarantineMessage $qmsg) {
                    $sender = "$($qmsg.SenderAddress)"; if (-not $sender) { $sender = "$($rec.SenderAddress)" }
                    $a2 = Show-ToolDialog(
                        ("Also ALLOW this sender from now on?`n`n    $sender`n`n" +
                         "Future messages from this address will no longer be quarantined as spam.`n" +
                         "This is a standing change, not part of this one release. It does not override " +
                         "a transport rule or a blocked-sender entry -- those can still hold the sender's mail."),
                        "Allow sender", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question, 'Button2')
                    $allow = ($a2 -eq [System.Windows.Forms.DialogResult]::Yes)
                }

                Invoke-QuarantineRelease -QuarantineIdentity $qmsg.Identity -ReportFalsePositive:$report -AllowSender:$allow
                $released++
                if ($report) { $reported++ }
                if ($allow)  { $allowed++ }
                Write-CompletedLog -Task ("Released QuarantineMessage $($rec.RecipientAddress) ($Environment; sender: $($rec.SenderAddress); subject: $($rec.Subject)" +
                                          $(if ($report) { '; reported to Microsoft as false positive' } else { '' }) +
                                          $(if ($allow)  { "; sender $sender allowed from now on" } else { '' }) + ')')
            }
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $summary = "$($qmatches.Count) quarantine match(es): $released released"
            if ($reported -gt 0) { $summary += ", $reported reported to Microsoft as false positive" }
            if ($allowed -gt 0)  { $summary += ", $allowed sender(s) allowed from now on" }
            if ($notReportable -gt 0) { $summary += " ($notReportable not spam-type, so not reportable from here)" }
            $lblTraceStatus.Text = $summary + '.'
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
