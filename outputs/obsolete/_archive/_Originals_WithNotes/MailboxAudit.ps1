#Requires -Version 5.1
<#
.SYNOPSIS
    Mailbox Audit -- traces Exchange Online mailbox activity (Search-UnifiedAuditLog)
    for a single mailbox, across owner/delegate/admin actions, over a date range.
.DESCRIPTION
    Layer 2 tool script, same conventions as the other ~10 tools in this suite:
    dot-sources Common.ps1, uses Resolve-ToolTarget (prompts for a target itself
    when launched with no -SmtpAddress, same as every other tool), Tier2 entry in
    Tools.json, ObjectTypes = ['Mailbox'].

    -Environment / -DomainController are accepted ONLY for calling-convention
    consistency with AdminConsole.ps1, which passes the same three parameters to
    every tool it launches. This tool never uses either value for anything -- it
    is pure Exchange Online / Microsoft Purview audit, no on-prem AD dependency
    at all (unlike DLManager.ps1, GroupEditor.ps1's on-prem branch, etc.).

    VERIFIED against Microsoft Learn, July 2026 (Search-UnifiedAuditLog reference,
    "Search the audit log for mailbox activities in specific mailboxes," "Audit
    log activities," and "Learn about properties in Exchange audit records" --
    all fetched and read directly while building this script, not recalled from
    training data):

      - SCOPING TO "EVERYTHING THAT HAPPENED IN THIS MAILBOX": Search-UnifiedAuditLog's
        -UserIds parameter is generically documented as "the account of the user
        who performed the action" -- but Microsoft's own mailbox-activity search
        guidance explicitly recommends putting the AFFECTED MAILBOX (not a
        suspected actor) into that field for a regular/user mailbox search. That
        is exactly what this script does: -UserIds is set to the resolved
        mailbox's UPN, which (per that same guidance) returns activity performed
        by the owner, a delegate, or an admin against that mailbox -- matching
        the locked design goal of "owner + anyone who touched it," achieved
        server-side, no client-side actor filtering needed.
      - THIS IS DIFFERENT FROM SHARED MAILBOX AUDIT (not built yet, a separate
        planned script): Microsoft's guidance explicitly says shared-mailbox
        searches should put the mailbox's SMTP address in the FreeText/Keywords
        field instead, NOT -UserIds. Do not reuse this script's -UserIds approach
        for that one without re-reading that guidance first -- it's a different
        parameter for a reason.
      - RETENTION: Microsoft's current documented default audit retention is 180
        days (not the 90-day placeholder used earlier in design discussion),
        unless Audit (Premium) is licensed or a custom retention policy extends
        it. This script does NOT default the date range to 180 days -- see
        "DEFAULT DATE RANGE" below -- 180 days is only the outer boundary of
        what's available at all, surfaced in the warning label and in the
        zero-results status message, not used as the picker's starting value.
      - DEFAULT DATE RANGE is last 7 days, not the full 180-day retention window
        -- defaulting to the full window would likely be slow and return far
        more than anyone opening this tool actually wants on first load. 7 days
        is a "what just happened" default; the date pickers allow going back
        further (soft default, not a hard limit).
      - OPERATIONS LIST: the 31 selectable operations below are every Operation
        value with a documented, unambiguous mapping in Microsoft's "Exchange
        mailbox activities" table as of this writing. Three rows from that table
        are deliberately excluded, not overlooked:
          * "Label message as a record" has NO documented Operation value in
            Microsoft's own table -- nothing to map it to.
          * "Priority cleanup label applied on Exchange item" maps to the same
            Operation value (ApplyPriorityCleanup) as "Priority cleanup label
            applied" -- listing both would just be a duplicate checkbox.
          * "Perform search query" maps to SharepointSearchQueryInitiated --
            despite being listed in the Exchange mailbox activities table, that
            Operation value is plainly a SharePoint operation, not an Exchange
            mailbox one. Excluded as very likely a documentation placement
            quirk rather than something that belongs in a mailbox-audit list.
      - PERMISSIONS: running this requires the operator's account to have the
        "View-Only Audit Logs" or "Audit Logs" role assigned in Microsoft
        Purview role groups, in addition to the usual "connect to Exchange
        Online" assumption every tool in this suite makes. This script can't
        grant or cheaply verify that role -- it's surfaced as a warning label,
        not silently assumed.
      - INGESTION DELAY: an action happening in Exchange doesn't mean its audit
        record is searchable yet -- Microsoft's own docs note it can take up to
        30 minutes or more for an entry to appear. A just-performed action
        legitimately might not show up in a search run immediately after.
      - PAGING: uses -SessionCommand ReturnLargeSet with a fresh SessionId per
        search, repeatedly calling Search-UnifiedAuditLog until it returns zero
        additional records or the documented ReturnLargeSet ceiling of 50,000
        records is reached. If that ceiling is hit, the status message says so
        explicitly rather than silently presenting a possibly-incomplete result
        set as complete.
      - SUBJECT FILTER: Search-UnifiedAuditLog has no dedicated "subject" search
        parameter. The Subject box here is a CLIENT-SIDE partial-match filter
        applied after parsing each record's AuditData -- the operator-facing
        "N results" count is post-filter, but how long the search takes (and
        whether it hits the paging ceiling) is governed by the date range and
        Operations filter only, not by what's typed in Subject.
      - AuditData PARSING (v1 scope): each record's nested AuditData JSON is
        parsed for the fields Microsoft's own reference documents as common
        across operation types -- Operation, CreationTime, UserId, ClientIP,
        ClientInfoString, MailboxOwnerUPN, ResultStatus, Item.Subject,
        Item.ParentFolder.Name. Several operation types carry additional,
        operation-specific fields (for example, rule-change records carry a
        Parameters array showing exactly what changed) that are NOT yet broken
        out into their own grid columns -- double-click any row to see that
        record's full raw AuditData JSON, so nothing is actually hidden, just
        not all reformatted into columns yet.
      - Write-CompletedLog fires once per search that's actually run (not once
        per result row) -- an audit-the-auditor entry, since running a mailbox
        audit search is itself worth a record of who ran it and against what.
        This still only logs Task+DateTime, same as every other tool -- it does
        NOT solve the "no operator identity in the log" gap flagged earlier;
        that remains a suite-wide fix for a future revision, not something
        bolted on to just this one script.
      - CONFIDENCE SCORING / CANDIDATE-POOL ATTRIBUTION is explicitly OUT OF
        SCOPE for this build -- discussed as a deliberately later, separately-
        designed phase (needs additional log sources like Entra sign-in data,
        not just mailbox audit data) and is not attempted here. Every result
        row shows exactly what Search-UnifiedAuditLog/AuditData reports, at
        face value.
      - Not yet built: self-validating that the resolved target is actually a
        Mailbox (the way GroupEditor.ps1 validates Type for groups). Given a
        mismatched object here just returns an empty/irrelevant result set
        rather than a raw cmdlet exception, this was judged lower priority than
        it would be for a tool that performs writes -- flagged as a cheap
        future addition, not done in this pass.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-18
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'Mailbox Audit -- Enter Target'
if (-not $target) { return }   # operator cancelled the prompt

#region Constants ----------------------------------------------------------------
# Verified against Microsoft Learn "Audit log activities" (Exchange mailbox
# activities table), July 2026 -- FriendlyName -> Operation value, in the same
# order as the source table. See header for the three deliberately excluded rows.
$script:MailboxOperations = [ordered]@{
    'Accessed mailbox attachments'                                  = 'AttachmentAccess'
    'Accessed Mailbox folder'                                       = 'FolderBind'
    'Accessed mailbox items'                                        = 'MailItemsAccessed'
    'Accessed message'                                              = 'MessageBind'
    'Added delegate mailbox permissions'                            = 'Add-MailboxPermission'
    'Added or removed user with delegate access to calendar folder' = 'UpdateCalendarDelegation'
    'Added permissions to folder'                                   = 'AddFolderPermissions'
    'Copied messages to another folder'                             = 'Copy'
    'Created mailbox item'                                          = 'Create'
    'Created new inbox rule in Outlook web app'                     = 'New-InboxRule'
    'Delete priority cleanup label item'                            = 'PriorityCleanupDelete'
    'Deleted messages from Deleted Items folder'                    = 'SoftDelete'
    'Labeled item as a record'                                      = 'ApplyRecord'
    'Labeled message as a record'                                   = 'ApplyRecordLabel'
    'Modified folder permission'                                    = 'ModifyFolderPermissions'
    'Modified inbox rule from Outlook web app'                      = 'Set-InboxRule'
    'Moved messages to another folder'                              = 'Move'
    'Moved messages to Deleted Items folder'                        = 'MoveToDeletedItems'
    'Preserve Mailbox item'                                         = 'PreservedMailItemProactively'
    'Priority cleanup label applied'                                = 'ApplyPriorityCleanup'
    'Purged messages from the mailbox'                              = 'HardDelete'
    'Removed delegate mailbox permissions'                          = 'Remove-MailboxPermission'
    'Removed permissions from folder'                               = 'RemoveFolderPermissions'
    'Search items in a mailbox'                                     = 'SearchQueryInitiated'
    'Sent message'                                                  = 'Send'
    'Sent message using Send As permissions'                        = 'SendAs'
    'Sent message using Send On Behalf permissions'                 = 'SendOnBehalf'
    'Update item label'                                             = 'UpdateComplianceTag'
    'Updated inbox rules from Outlook client'                       = 'UpdateInboxRules'
    'Updated message'                                               = 'Update'
    'User signed in to mailbox'                                     = 'MailboxLogin'
}
# Reverse lookup (Operation value -> FriendlyName) for display in the results grid.
$script:OperationFriendlyNames = @{}
foreach ($key in $script:MailboxOperations.Keys) {
    $script:OperationFriendlyNames[$script:MailboxOperations[$key]] = $key
}
# Documented ReturnLargeSet ceiling (Search-UnifiedAuditLog reference) -- used to
# detect and warn on a possibly-truncated result set rather than presenting it
# silently as complete.
$script:AuditResultCeiling = 50000
$script:AuditPageSize = 5000
#endregion

#region Data layer -----------------------------------------------------------------
# Pages through Search-UnifiedAuditLog for one mailbox until either it returns no
# further records or the documented 50,000-record ReturnLargeSet ceiling is hit.
# Returns a PSCustomObject: Records (raw Search-UnifiedAuditLog rows, AuditData
# still a JSON string) + HitCeiling (bool).
function Invoke-MailboxAuditSearch {
    # Cancellation checks $script:AuditCancelRequested directly (set by the
    # Cancel button's click handler) rather than a [ref] parameter -- matches
    # the established convention elsewhere in this suite (DLManager.ps1's
    # member-count scan loop checks $script:CountScanCancelRequested the same
    # way) instead of introducing a new pattern just for this script.
    param(
        [string]$Mailbox,
        [datetime]$StartUtc,
        [datetime]$EndUtc,
        [string[]]$Operations,
        [scriptblock]$ProgressCallback
    )
    $sessionId = "MailboxAudit_$([Guid]::NewGuid().ToString())"
    $all = New-Object System.Collections.Generic.List[object]
    $hitCeiling = $false
    $iteration = 0
    # 15 iterations x 5,000/page = 75,000 -- comfortably past the documented
    # 50,000 ReturnLargeSet ceiling, purely as a defensive loop guard in case a
    # future service change ever behaves unexpectedly; the ceiling check below
    # is what actually governs normal completion.
    while ($iteration -lt 15) {
        $iteration++
        $params = @{
            StartDate      = $StartUtc
            EndDate        = $EndUtc
            UserIds        = $Mailbox
            SessionId      = $sessionId
            SessionCommand = 'ReturnLargeSet'
            ResultSize     = $script:AuditPageSize
            ErrorAction    = 'Stop'
        }
        if ($Operations -and $Operations.Count -gt 0) { $params['Operations'] = $Operations }
        $page = @(Search-UnifiedAuditLog @params)
        if ($page.Count -eq 0) { break }
        foreach ($row in $page) { $all.Add($row) }
        if ($ProgressCallback) { & $ProgressCallback $all.Count }
        if ($script:AuditCancelRequested) { break }
        if ($all.Count -ge $script:AuditResultCeiling) { $hitCeiling = $true; break }
        if ($page.Count -lt $script:AuditPageSize) { break }   # short page -- no more data
    }
    return [PSCustomObject]@{
        Records    = $all.ToArray()
        HitCeiling = $hitCeiling
    }
}

# Parses one raw Search-UnifiedAuditLog row's AuditData JSON into a flat, grid-
# ready object. Field names verified against Microsoft's "Learn about properties
# in Exchange audit records" reference -- see header.
function ConvertTo-MailboxAuditRecord {
    param($RawRecord)
    $parsed = $null
    try { $parsed = $RawRecord.AuditData | ConvertFrom-Json -ErrorAction Stop } catch { $parsed = $null }
    $opValue = if ($parsed -and $parsed.Operation) { "$($parsed.Operation)" } else { "$($RawRecord.Operations)" }
    $friendly = if ($script:OperationFriendlyNames.ContainsKey($opValue)) { $script:OperationFriendlyNames[$opValue] } else { $opValue }
    $creationUtc = if ($parsed -and $parsed.CreationTime) {
        try { [datetime]::Parse($parsed.CreationTime, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal) } catch { $RawRecord.CreationDate }
    } else { $RawRecord.CreationDate }
    [PSCustomObject]@{
        TimeLocal        = $creationUtc.ToLocalTime()
        Operation         = $opValue
        OperationFriendly = $friendly
        Actor             = if ($parsed) { "$($parsed.UserId)" } else { "$($RawRecord.UserIds)" }
        MailboxOwnerUPN   = if ($parsed) { "$($parsed.MailboxOwnerUPN)" } else { '' }
        ClientIP          = if ($parsed) { "$($parsed.ClientIP)" } else { '' }
        ClientInfo        = if ($parsed) { "$($parsed.ClientInfoString)" } else { '' }
        Subject           = if ($parsed -and $parsed.Item -and $parsed.Item.Subject) { "$($parsed.Item.Subject)" } else { '' }
        Folder            = if ($parsed -and $parsed.Item -and $parsed.Item.ParentFolder -and $parsed.Item.ParentFolder.Name) { "$($parsed.Item.ParentFolder.Name)" } else { '' }
        ResultStatus      = if ($parsed) { "$($parsed.ResultStatus)" } else { '' }
        RawJson           = if ($parsed) { ($parsed | ConvertTo-Json -Depth 8) } else { "$($RawRecord.AuditData)" }
    }
}
#endregion

#region Main window ----------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mailbox Audit: $($target.Smtp)"
$form.Size = New-Object System.Drawing.Size(1080, 760)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Requires an active Exchange Online connection AND the 'View-Only Audit Logs' or 'Audit Logs' role assigned to your account in Microsoft Purview. Audit entries can take 30+ minutes to appear after an action happens -- a just-performed action may not show up yet. Default retention is 180 days unless Audit (Premium) or a custom retention policy extends it."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(1040, 48)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "Mailbox: $($target.Smtp)"
$lblTarget.Location = New-Object System.Drawing.Point(12, 62)
$lblTarget.AutoSize = $true
$lblTarget.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTarget)

$grpRange = New-Object System.Windows.Forms.GroupBox
$grpRange.Text = "Date Range (local time -- converted to UTC for the search)"
$grpRange.Location = New-Object System.Drawing.Point(12, 88)
$grpRange.Size = New-Object System.Drawing.Size(400, 66)
$form.Controls.Add($grpRange)
$lblStart = New-Object System.Windows.Forms.Label
$lblStart.Text = "Start:"
$lblStart.Location = New-Object System.Drawing.Point(10, 26)
$lblStart.AutoSize = $true
$grpRange.Controls.Add($lblStart)
$dtpStart = New-Object System.Windows.Forms.DateTimePicker
$dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpStart.CustomFormat = 'yyyy-MM-dd HH:mm'
$dtpStart.Value = (Get-Date).AddDays(-7)
$dtpStart.Location = New-Object System.Drawing.Point(56, 22)
$dtpStart.Size = New-Object System.Drawing.Size(150, 22)
$grpRange.Controls.Add($dtpStart)
$lblEnd = New-Object System.Windows.Forms.Label
$lblEnd.Text = "End:"
$lblEnd.Location = New-Object System.Drawing.Point(220, 26)
$lblEnd.AutoSize = $true
$grpRange.Controls.Add($lblEnd)
$dtpEnd = New-Object System.Windows.Forms.DateTimePicker
$dtpEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpEnd.CustomFormat = 'yyyy-MM-dd HH:mm'
$dtpEnd.Value = Get-Date
$dtpEnd.Location = New-Object System.Drawing.Point(258, 22)
$dtpEnd.Size = New-Object System.Drawing.Size(150, 22)
$grpRange.Controls.Add($dtpEnd)

$lblSubject = New-Object System.Windows.Forms.Label
$lblSubject.Text = "Subject contains (optional, applied after retrieval):"
$lblSubject.Location = New-Object System.Drawing.Point(424, 96)
$lblSubject.AutoSize = $true
$form.Controls.Add($lblSubject)
$txtSubject = New-Object System.Windows.Forms.TextBox
$txtSubject.Location = New-Object System.Drawing.Point(424, 116)
$txtSubject.Size = New-Object System.Drawing.Size(300, 22)
$txtSubject.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtSubject)

$grpOps = New-Object System.Windows.Forms.GroupBox
$grpOps.Text = "Operations (none checked = all operations)"
$grpOps.Location = New-Object System.Drawing.Point(742, 88)
$grpOps.Size = New-Object System.Drawing.Size(310, 260)
$form.Controls.Add($grpOps)
$clbOps = New-Object System.Windows.Forms.CheckedListBox
$clbOps.Location = New-Object System.Drawing.Point(8, 20)
$clbOps.Size = New-Object System.Drawing.Size(292, 200)
$clbOps.CheckOnClick = $true
foreach ($friendlyName in $script:MailboxOperations.Keys) { [void]$clbOps.Items.Add($friendlyName) }
$grpOps.Controls.Add($clbOps)
$btnOpsAll = New-Object System.Windows.Forms.Button
$btnOpsAll.Text = "Select All"
$btnOpsAll.Location = New-Object System.Drawing.Point(8, 226)
$btnOpsAll.Size = New-Object System.Drawing.Size(90, 26)
$grpOps.Controls.Add($btnOpsAll)
$btnOpsNone = New-Object System.Windows.Forms.Button
$btnOpsNone.Text = "Clear All"
$btnOpsNone.Location = New-Object System.Drawing.Point(102, 226)
$btnOpsNone.Size = New-Object System.Drawing.Size(90, 26)
$grpOps.Controls.Add($btnOpsNone)
$btnOpsAll.Add_Click({ for ($i = 0; $i -lt $clbOps.Items.Count; $i++) { $clbOps.SetItemChecked($i, $true) } })
$btnOpsNone.Add_Click({ for ($i = 0; $i -lt $clbOps.Items.Count; $i++) { $clbOps.SetItemChecked($i, $false) } })

$btnAuditNow = New-Object System.Windows.Forms.Button
$btnAuditNow.Text = "Audit Now"
$btnAuditNow.Location = New-Object System.Drawing.Point(424, 150)
$btnAuditNow.Size = New-Object System.Drawing.Size(150, 32)
$btnAuditNow.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnAuditNow)
$btnCancelSearch = New-Object System.Windows.Forms.Button
$btnCancelSearch.Text = "Cancel"
$btnCancelSearch.Location = New-Object System.Drawing.Point(424, 150)
$btnCancelSearch.Size = New-Object System.Drawing.Size(150, 32)
$btnCancelSearch.Visible = $false
$form.Controls.Add($btnCancelSearch)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Set a date range and click 'Audit Now.'"
$lblStatus.Location = New-Object System.Drawing.Point(12, 358)
$lblStatus.Size = New-Object System.Drawing.Size(1040, 36)
$lblStatus.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblStatus)

$grdAudit = New-Object System.Windows.Forms.DataGridView
$grdAudit.Location = New-Object System.Drawing.Point(12, 398)
$grdAudit.Size = New-Object System.Drawing.Size(1040, 260)
$grdAudit.ReadOnly = $true
$grdAudit.AllowUserToAddRows = $false
$grdAudit.AllowUserToDeleteRows = $false
$grdAudit.AllowUserToResizeRows = $false
$grdAudit.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grdAudit.MultiSelect = $false
$grdAudit.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grdAudit.Font = New-Object System.Drawing.Font("Consolas", 8)
[void]$grdAudit.Columns.Add('TimeLocal', 'Time (Local)')
[void]$grdAudit.Columns.Add('OperationFriendly', 'Activity')
[void]$grdAudit.Columns.Add('Actor', 'Actor')
[void]$grdAudit.Columns.Add('Subject', 'Subject')
[void]$grdAudit.Columns.Add('Folder', 'Folder')
[void]$grdAudit.Columns.Add('ClientInfo', 'Client')
[void]$grdAudit.Columns.Add('ClientIP', 'Client IP')
[void]$grdAudit.Columns.Add('ResultStatus', 'Result')
$grdAudit.Columns['TimeLocal'].FillWeight        = 14
$grdAudit.Columns['OperationFriendly'].FillWeight = 20
$grdAudit.Columns['Actor'].FillWeight            = 16
$grdAudit.Columns['Subject'].FillWeight          = 20
$grdAudit.Columns['Folder'].FillWeight           = 10
$grdAudit.Columns['ClientInfo'].FillWeight       = 12
$grdAudit.Columns['ClientIP'].FillWeight         = 10
$grdAudit.Columns['ResultStatus'].FillWeight     = 8
foreach ($col in $grdAudit.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
$form.Controls.Add($grdAudit)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Results to CSV"
$btnExport.Location = New-Object System.Drawing.Point(12, 668)
$btnExport.Size = New-Object System.Drawing.Size(170, 28)
$form.Controls.Add($btnExport)
$lblRowHint = New-Object System.Windows.Forms.Label
$lblRowHint.Text = "Double-click a row to see its full raw audit record (fields not yet broken into columns are still in there)."
$lblRowHint.Location = New-Object System.Drawing.Point(196, 674)
$lblRowHint.Size = New-Object System.Drawing.Size(700, 20)
$lblRowHint.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
$lblRowHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblRowHint)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(952, 668)
$btnClose.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnClose)

$script:AuditResults = @()
$script:AuditCancelRequested = $false

# Popup showing one row's full parsed AuditData as pretty-printed JSON -- the
# escape hatch for every field this script doesn't yet break into its own
# column (see header "AuditData PARSING (v1 scope)").
function Show-AuditRawDetail {
    param([string]$Json, [string]$Title)
    $detailForm = New-Object System.Windows.Forms.Form
    $detailForm.Text = $Title
    $detailForm.Size = New-Object System.Drawing.Size(700, 560)
    $detailForm.StartPosition = "CenterParent"
    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.Multiline = $true
    $txtDetail.ReadOnly = $true
    $txtDetail.ScrollBars = "Both"
    $txtDetail.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtDetail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtDetail.Text = $Json
    $detailForm.Controls.Add($txtDetail)
    [void]$detailForm.ShowDialog()
}

$btnAuditNow.Add_Click({
    if ($dtpEnd.Value -le $dtpStart.Value) {
        [void][System.Windows.Forms.MessageBox]::Show("End date/time must be after the start date/time.", 'Invalid Range', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $selectedOps = @($clbOps.CheckedItems | ForEach-Object { $script:MailboxOperations[$_] })
    $btnAuditNow.Visible = $false
    $btnCancelSearch.Visible = $true
    $script:AuditCancelRequested = $false
    $grdAudit.Rows.Clear()
    $script:AuditResults = @()
    $lblStatus.Text = "Searching..."
    $form.Refresh()
    try {
        $startUtc = $dtpStart.Value.ToUniversalTime()
        $endUtc = $dtpEnd.Value.ToUniversalTime()
        $progress = {
            param($count)
            $lblStatus.Text = "Fetched $count record(s) so far..."
            [System.Windows.Forms.Application]::DoEvents()
        }
        $searchResult = Invoke-MailboxAuditSearch -Mailbox $target.Smtp -StartUtc $startUtc -EndUtc $endUtc -Operations $selectedOps -ProgressCallback $progress
        $parsed = @($searchResult.Records | ForEach-Object { ConvertTo-MailboxAuditRecord -RawRecord $_ })
        $subjectFilter = $txtSubject.Text.Trim()
        if ($subjectFilter -ne '') {
            $parsed = @($parsed | Where-Object { $_.Subject -like "*$subjectFilter*" })
        }
        $script:AuditResults = $parsed
        foreach ($rec in $parsed) {
            [void]$grdAudit.Rows.Add($rec.TimeLocal.ToString('yyyy-MM-dd HH:mm:ss'), $rec.OperationFriendly, $rec.Actor, $rec.Subject, $rec.Folder, $rec.ClientInfo, $rec.ClientIP, $rec.ResultStatus)
        }
        $statusMsg = "$($parsed.Count) result(s)."
        if ($script:AuditCancelRequested) { $statusMsg += " Search cancelled -- results may be incomplete." }
        if ($searchResult.HitCeiling) { $statusMsg += " Hit the 50,000-record search ceiling -- narrow the date range or Operations filter to be sure you have everything in this window." }
        if ($parsed.Count -eq 0 -and -not $script:AuditCancelRequested) {
            $statusMsg += " Zero results can mean nothing happened, OR that this falls outside the retention window, OR that mailbox auditing isn't capturing this action for this mailbox's role -- it doesn't necessarily mean 'confirmed nothing happened.'"
        }
        $lblStatus.Text = $statusMsg
        Write-CompletedLog -Task "Ran MailboxAudit search $($target.Smtp) ($($dtpStart.Value.ToString('yyyy-MM-dd HH:mm')) to $($dtpEnd.Value.ToString('yyyy-MM-dd HH:mm')); $($selectedOps.Count) operation(s) selected; $($parsed.Count) result(s))"
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Search-UnifiedAuditLog'
    } finally {
        $btnAuditNow.Visible = $true
        $btnCancelSearch.Visible = $false
    }
})
$btnCancelSearch.Add_Click({ $script:AuditCancelRequested = $true })

$grdAudit.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.RowIndex -ge $script:AuditResults.Count) { return }
    $rec = $script:AuditResults[$e.RowIndex]
    Show-AuditRawDetail -Json $rec.RawJson -Title "Raw Audit Record: $($rec.OperationFriendly) at $($rec.TimeLocal)"
})

$btnExport.Add_Click({
    if ($script:AuditResults.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("Nothing to export -- run a search first.", 'Export Results', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    try {
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $sfd.FileName = "$($target.Smtp -replace '[^a-zA-Z0-9]','_')_audit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $script:AuditResults | Select-Object TimeLocal, OperationFriendly, Operation, Actor, MailboxOwnerUPN, Subject, Folder, ClientInfo, ClientIP, ResultStatus |
            Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
        $lblStatus.Text = "Exported $($script:AuditResults.Count) result(s) to $($sfd.FileName)."
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export audit results'
    }
})

$btnClose.Add_Click({ $form.Close() })

# ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin

# Console, which opens top-left. Helper is in Common.ps1.

Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
#endregion
