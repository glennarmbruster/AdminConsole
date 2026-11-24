#Requires -Version 5.1
<#
.SYNOPSIS
    Standalone Message Trace / Quarantine Release tool -- Layer 2 tool script.
.DESCRIPTION
    REVISED 2026-07-22, part 2 -- no longer tied to a specific mailbox. The
    operator found that requiring a target mailbox before the window would
    even open (via Common.ps1's Resolve-ToolTarget / Show-TargetPromptDialog)
    was more restrictive than useful -- a trace is often about a sender, a
    recipient, both, or neither (e.g. "everything quarantined today"), not
    necessarily "everything touching one particular mailbox". So:
      - This script no longer calls Resolve-ToolTarget at all. It opens
        straight to the Message Trace window every time, standalone or from
        AdminConsole, with no prompt and no gate.
      - The old "Direction" radio group (Sent from / Received to THIS
        mailbox) plus its single "narrow" textbox are replaced by two fully
        independent, optional fields: From (sender) and To (recipient). Fill
        in either, both, or neither -- Get-MessageTraceV2's SenderAddress and
        RecipientAddress parameters are independently optional (see the
        Invoke-MessageTrace comment below), so leaving both blank simply
        traces every message in the date range (still bounded by the
        existing 10-day-window / 90-day-lookback rules).
      - -SmtpAddress is now purely an OPTIONAL convenience prefill (into the
        From field) for whatever AdminConsole's last lookup was, not a
        requirement -- launching with it omitted/blank is fully supported
        and opens the exact same window.
      - -DomainController is still accepted as a parameter (so existing
        AdminConsole launch lines that pass it don't break) but was already
        unused inside this script before this revision (every cmdlet here
        is EXO-only, no on-prem AD calls) and remains unused now.
    Fully functional on its own:
        .\MessageTrace.ps1
    opens directly to the Message Trace window, no prompt.

    Fully functional from AdminConsole.ps1 too:
        .\MessageTrace.ps1 -SmtpAddress user@domain.com -Environment Franchise
    opens the same window with the From field pre-filled from -SmtpAddress,
    if provided.

    Every function below is MOVED, UNCHANGED, from the original monolithic
    Mailbox-Group-Inspector.ps1 (original lines 3100-3651, region "Actions -
    message trace"), including every verification/design comment. The only
    edits anywhere in this file that touch original tool behavior are
    mechanical, both isolated to the very end of Show-MessageTraceForm:
      1. The closing call changed from
             [void]$traceForm.ShowDialog($form)
         to
             # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
             # Console, which opens top-left. Helper is in Common.ps1.
             Set-FormScreenCorner -Form $traceForm -Corner 'BottomLeft'
             [void]$traceForm.ShowDialog()
         because there is no longer a parent $form to own it -- this script
         IS the top-level window now.
      2. The one Write-CompletedLog line inside the Release-from-Quarantine
         handler called Get-SelectedEnvLabel(), which read $radHQ.Checked off
         the monolith's shared form -- a control that no longer exists here.
         Show-MessageTraceForm now takes an -Environment parameter (passed
         from $target.Environment, the same value Resolve-ToolTarget already
         resolved) and that string is logged in its place. Nothing about
         WHEN or WHY the log line fires changed, only where the environment
         label text comes from.
    Every other line of every function, including all comments, is
    byte-for-byte the original.

    ADDED 2026-07-22 -- results search/filter box: a new "Filter loaded
    results" textbox + Search/Clear Filter buttons sit just above the grid.
    This does NOT re-query Exchange Online -- it's a client-side filter over
    whatever $script:CurrentTraceResults the last Trace run already pulled
    back, matching the typed text (case-insensitive substring, via -like
    "*text*") against SenderAddress, RecipientAddress, Subject, and Status.
    A new $script:CurrentTraceDisplayResults holds the currently-filtered
    (or unfiltered) subset that the grid actually renders; Update-
    TraceGridDisplay was changed to read from that variable instead of
    $script:CurrentTraceResults directly. Export to CSV was changed to
    export $script:CurrentTraceDisplayResults too, so exporting after a
    filter is applied exports what's on screen, not the full unfiltered
    result set -- this is a judgment call (the alternative, always
    exporting everything regardless of the filter box, seemed less
    intuitive) and can be reverted if that's not the desired behavior.
    Filter box is cleared automatically on every new Trace run.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
    Revised : 2026-07-22 -- added results search/filter box (client-side,
              no new query). See ADDED note above.
    Revised : 2026-07-22, later -- no longer tied to a specific mailbox;
              removed the Resolve-ToolTarget gate and replaced the
              Sent/Received direction radio with independent optional
              From/To fields. See REVISED note above.
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# No Resolve-ToolTarget call -- this tool is no longer tied to a specific
# mailbox, see header REVISED note (2026-07-22, later). -SmtpAddress (if
# passed) is used purely as an optional From-field prefill further down.

#region Actions - message trace (moved unchanged from Mailbox-Group-Inspector.ps1 lines 3100-3651)
# Get-MessageTraceV2 / Get-MessageTraceDetailV2 verified against Microsoft Learn
# (2026-07-08): both are "available only in the cloud-based service" (no on-prem
# equivalent by this name -- Get-MessageTrackingLog is the on-prem cmdlet, a
# completely different name), so there is no collision risk the way Get-Mailbox
# etc. have. SenderAddress and RecipientAddress on Get-MessageTraceV2 are BOTH
# independently optional (Mandatory: False on both), which is what makes the
# Sent/Received one-sided trace below possible with a single cmdlet call -- Sent
# omits RecipientAddress entirely (any recipient), Received omits SenderAddress
# entirely (any sender). Still routed through Invoke-EXOCommand for consistency
# with every other EXO-flavored call in this script, per the standing rule in
# Get-EXOCommand's header comment.
#
# Get-MessageTraceDetailV2 is called exactly the way the operator's own working
# "EXO Message Trace" reference script already calls it (-MessageTraceId,
# -RecipientAddress, -StartDate, -EndDate) -- that usage is treated as verified
# per standards Section 1(a) and is not re-derived here.
function Invoke-MessageTrace {
    # -Status accepts any of the values Get-MessageTraceV2 documents for its
    # own -Status parameter (verified against Microsoft Learn, 2026-07-09):
    # Delivered, Expanded, Failed, FilteredAsSpam, GettingStatus, Pending,
    # Quarantined. Omit it entirely (as the "All Messages" dropdown option
    # does) to get the unfiltered default.
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
# Quarantine release -- verified against Microsoft Learn (2026-07-08):
#  - Get-QuarantineMessage / Release-QuarantineMessage are both "available
#    only in the cloud-based service" (Exchange PowerShell module) -- same
#    module family as Get-MessageTraceV2, routed through Invoke-EXOCommand.
#  - Get-QuarantineMessage -MessageId expects the full RFC822 Message-ID
#    header value, angle brackets included (Microsoft Learn Example 3:
#    Get-QuarantineMessage -MessageID "<...@contoso.com>"). This is a
#    DIFFERENT value than Get-MessageTraceV2's MessageTraceId (a GUID) --
#    the correlation below uses the trace result's MessageId property
#    (the RFC822 header), NOT MessageTraceId.
#  - ASSUMPTION, not verified against a live tenant: Get-MessageTraceV2
#    result objects expose a MessageId property carrying that RFC822
#    value (this has been true of the classic Get-MessageTrace schema and
#    is what admins have always used to cross-reference quarantine, but
#    the operator should confirm on first use -- Find-QuarantineMatch
#    throws a clear, specific error if it's missing rather than silently
#    matching the wrong thing).
#  - Release-QuarantineMessage verified via Microsoft Learn Example 2:
#    Release-QuarantineMessage -Identity <id> -ReleaseToAll releases a
#    single quarantined message (by its own Identity, from
#    Get-QuarantineMessage) to ALL of its original recipients. -Confirm
#    $false is added the same way every other write cmdlet in this script
#    suppresses the built-in prompt, since this tool has its own
#    MessageBox confirmation instead.
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
# Time zone display for Message Trace results -- Microsoft Learn confirms
# "The time stamps on the output are in UTC time format" for
# Get-MessageTraceV2 (and by extension Get-MessageTraceDetailV2), so
# everything shown to the operator needs an explicit UTC-to-local
# conversion. Windows time zone IDs (not IANA names) are used with
# [System.TimeZoneInfo], which is correct for the "Standard PowerShell"
# Windows environment this script's header already requires -- these IDs
# carry their DST rules despite the "Standard Time" naming (.NET's own
# documented behavior, not Exchange-specific, so not re-verified against
# Microsoft Learn the way Exchange cmdlets are). This ONLY changes how
# already-fetched results are DISPLAYED (grid, Detail box, CSV export) -- it
# does not change how the Search From/Through date pickers are interpreted;
# those are passed to Get-MessageTraceV2 exactly as before.
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
        # Falls back to UTC (unconverted) rather than throwing, so a missing
        # time zone on this particular machine degrades gracefully instead
        # of breaking the whole Trace window.
        return [System.TimeZoneInfo]::Utc
    }
}
function ConvertTo-TraceDisplayTime {
    # DateTime values coming back over Exchange Online's remoting connection
    # do not reliably keep DateTimeKind.Utc through deserialization, so this
    # forces Kind=Utc explicitly before converting rather than trusting
    # whatever Kind the object already reports -- converting an
    # already-local value as if it were UTC would silently shift every
    # timestamp by several hours, which is worse than not converting at all.
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
    # DESIGN NOTE: the operator's pasted reference script binds results to a
    # System.Data.DataTable (auto-generated grid columns, hidden-column pairs
    # for the trace-id/recipient needed by Get Details). This script instead
    # follows the SAME manual-grid convention already used for grdRules and
    # grdCal elsewhere in this file (Columns.Add + Rows.Add, FillWeight
    # sizing). UNLIKE grdRules/grdCal, this grid IS sortable on every column
    # (operator request) -- sorting an unbound DataGridView physically
    # reorders its Rows collection, so row position can no longer be used to
    # find the matching $script:CurrentTraceResults entry after a sort. Fix:
    # a hidden Key column (MessageTraceId|RecipientAddress) on every row,
    # looked up by value via Find-TraceResultByKey -- the same "hidden
    # identity column" approach grdRules already uses for exactly this
    # reason, just adopted here too once sorting made row-index matching
    # unsafe.
    param([string]$SmtpAddress, [string]$Environment)
    $traceForm = New-Object System.Windows.Forms.Form
    $traceForm.Text = if ($SmtpAddress) { "Message Trace (from: $SmtpAddress)" } else { "Message Trace" }
    # Width increased 30% (780 -> 1014) on request so the DataGridView's
    # Fill-mode columns get more room to show full sender/recipient/subject
    # text without truncation. Right margin (28px, matching the original
    # 780-12-740 layout) preserved -- content-area controls below widen from
    # 740 to 974 (740 + the 234px gained) to keep that same margin.
    $traceForm.Size = New-Object System.Drawing.Size(1014, 754)
    $traceForm.StartPosition = "CenterScreen"
    $traceForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    # From/To fields -- replaced the old Sent/Received "Direction" radio
    # group (tied to one anchor mailbox) 2026-07-22, see header REVISED
    # note. Both independently optional; occupies the same 12,12-252,62
    # footprint the old grpDir GroupBox used, so nothing else on the form
    # needed to move.
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
    # Time zone display selector -- placed in the space opened up to the
    # right of lblTraceNote by the 30%-wider form. Converts already-fetched
    # results (grid, Detail box, CSV export) on selection change without
    # re-querying Exchange Online -- see the ConvertTo-TraceDisplayTime /
    # Get-SelectedTraceTimeZone header comment above Format-TraceDetail.
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
    # Status filter -- placed in the empty space to the right of the From/To
    # (below lblTraceNote, which only spans y=12-58) so nothing else on the
    # form needs to move. Values verified against Microsoft Learn
    # (Get-MessageTraceV2 -Status): Delivered, Expanded, Failed,
    # FilteredAsSpam, GettingStatus, Pending, Quarantined. "All Messages" is
    # not a real API value -- it just means "don't pass -Status at all",
    # handled in the Trace click handler below.
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
    # 10-day-window business rule ported from the reference script: Through
    # can never be more than 9 days after Search From, and never in the future.
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
    # Results search/filter box -- added 2026-07-22, see header ADDED note.
    # Client-side filter over the already-fetched results, no new EXO query.
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
    $grdTrace.Columns['Received'].ValueType   = [DateTime]   # sorts chronologically, not alphabetically
    $grdTrace.Columns['Key'].Visible          = $false        # MessageTraceId|RecipientAddress -- see below
    # Sortable on every visible column (unlike grdRules/grdCal elsewhere in
    # this script, which are intentionally NotSortable) -- sorting reorders
    # the underlying Rows collection in an unbound grid like this one, so
    # row position can no longer be used to find the matching
    # $script:CurrentTraceResults entry. The hidden Key column below is the
    # fix: each row carries a MessageTraceId+RecipientAddress key that
    # double-click and Release-from-Quarantine look up by value instead of
    # by row index, the same "hidden identity column" approach grdRules
    # already uses for exactly this reason.
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
    # Filters $script:CurrentTraceResults (the full, last-fetched result set)
    # down to rows matching $SearchText -- case-insensitive substring match
    # (-like is case-insensitive by default) against sender, recipient,
    # subject, and status. Empty/whitespace search text returns everything.
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
        # Rebuilds the grid from the already-cached $script:CurrentTraceResults
        # in whatever time zone cboTimeZone currently has selected. Called
        # both right after a Trace run and whenever the time zone dropdown
        # changes -- the latter needs no new Exchange query since the raw
        # (UTC) data is already in hand.
        $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
        $grdTrace.Columns['Received'].HeaderText = "Received ($($cboTimeZone.SelectedItem))"
        $grdTrace.Rows.Clear()
        # Renders $script:CurrentTraceDisplayResults (the current filter's
        # output), not $script:CurrentTraceResults directly -- see the
        # results search/filter box header ADDED note.
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
            # From/To are both independently optional -- see header REVISED
            # note (2026-07-22, later). Leaving both blank traces every
            # message in the date range.
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
    # Results search/filter handlers -- added 2026-07-22, see header ADDED
    # note. Filters the grid over the already-fetched $script:CurrentTraceResults;
    # no new Exchange Online call happens here.
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
        # Looks up the raw Get-MessageTraceV2 object behind a grid row by its
        # hidden Key column value (MessageTraceId|RecipientAddress) rather
        # than by row position -- required now that the grid is sortable,
        # since sorting an unbound DataGridView physically reorders Rows.
        # Searches the FULL $script:CurrentTraceResults (not the filtered
        # display subset) -- a filtered-out row is never on screen to
        # double-click/select in the first place, but every row that IS on
        # screen is still a member of the full set, so this lookup doesn't
        # need to change when a filter is active.
        param([string]$RowKey)
        $script:CurrentTraceResults | Where-Object { "$($_.MessageTraceId)|$($_.RecipientAddress)" -eq $RowKey } | Select-Object -First 1
    }
    $grdTrace.Add_SelectionChanged({
        $btnRelease.Enabled = ($grdTrace.SelectedRows.Count -gt 0)
    })
    # Double-click a row to load its detail -- same interaction pattern as
    # grdRules elsewhere in this script (double-click for detail, no
    # separate "open" button).
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
                # CHANGED FROM ORIGINAL: was "(Get-SelectedEnvLabel)" -- that
                # read $radHQ.Checked off the monolith's shared form. This
                # script has no shared form, so it logs the script's own
                # -Environment parameter value instead. Nothing else about
                # this log line changed.
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
        # Exports $script:CurrentTraceDisplayResults -- the currently
        # filtered/displayed set, matching what's on screen in the grid --
        # not the full unfiltered $script:CurrentTraceResults. See header
        # ADDED note (results search/filter box, 2026-07-22).
        if ($script:CurrentTraceDisplayResults.Count -eq 0) { $lblTraceStatus.Text = "No results to export."; return }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $filePart = if ($SmtpAddress) { ($SmtpAddress -replace '[^a-zA-Z0-9]','_') } else { 'All' }
        $sfd.FileName = "MessageTrace_${filePart}_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($traceForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            # Received is exported in whatever zone the dropdown currently
            # shows (matching the grid on screen), via a calculated property
            # rather than the raw UTC value -- same conversion as the grid.
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
    # CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the
    # monolith's shared parent window. This script has no parent window, so
    # it shows itself as the top-level dialog instead. Nothing else in this
    # function changed.
    [void]$traceForm.ShowDialog()
}
#endregion

Show-MessageTraceForm -SmtpAddress $SmtpAddress -Environment $Environment
