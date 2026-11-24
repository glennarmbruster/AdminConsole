#Requires -Version 5.1
<#
    AuditUser.ps1
    Audit settings for a user mailbox.

    Author   : Glenn Armbruster
    Revised  : 2026-07-28
#>
# TOOLKIT-HELP-RUNS: searches the Microsoft 365 audit log. Changes nothing.
# TOOLKIT-HELP-BEGIN
# Searches the audit log for what actually HAPPENED to a mailbox -- who opened
# it, what they did, from where, and when.
#
# THIS IS NOT A SETTINGS REPORT. It does not tell you what is being recorded;
# it tells you what was recorded. "Who read that message", "did anyone else
# send as this account", "what did the delegate do on Tuesday".
#
# Pick a date range, optionally narrow to particular operations, and search.
#
# THREE THINGS THAT MAKE IT LOOK BROKEN WHEN IT IS NOT:
#
#   * Entries can take 30+ MINUTES to appear. An action performed a moment ago
#     legitimately will not be there yet, which is exactly when somebody tests
#     it and concludes the tool does not work.
#   * Default retention is 180 days, unless Audit (Premium) or a custom
#     retention policy extends it. Older than that is gone, not hidden.
#   * It needs the 'View-Only Audit Logs' or 'Audit Logs' role in Microsoft
#     Purview, on top of an Exchange Online connection.
#
# A wide date range on a busy mailbox takes minutes. It pages through the
# result set and can be stopped part way -- you keep whatever came back, and
# the window says so rather than pretending that was all of it.
#
# Each row keeps its raw JSON alongside the parsed columns. The columns answer
# most questions; the JSON answers the one they do not.
#
# Read-only. For a shared mailbox use Audit SMB, which does the same search and
# also lists who holds Full Access.
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

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. Resolve-ToolTarget is not used any more:
# its job was to decide which DIRECTORY an address belongs to, and this tool
# reads the unified audit log in Microsoft 365, which does not need to be told.
#
# An empty address is fine. The window opens and asks.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Constants ----------------------------------------------------------------
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
$script:OperationFriendlyNames = @{}
foreach ($key in $script:MailboxOperations.Keys) {
    $script:OperationFriendlyNames[$script:MailboxOperations[$key]] = $key
}
$script:AuditResultCeiling = 50000
$script:AuditPageSize = 5000
#endregion

#region Data layer -----------------------------------------------------------------
# THIS FILE'S OWN COPY, DELIBERATELY NAMED DIFFERENTLY FROM Common.ps1's.
#
# It used to be called Invoke-MailboxAuditSearch, the same as the function in
# Common.ps1 -- which this file dot-sources. The local definition wins, so
# the calls below reached THIS one and always had; but the shadowing was a
# trap with two edges. A fix made in Common never arrived here, and anyone
# tidying up an apparent duplicate would have deleted this and left the
# calls binding silently to Common's version, which does not take the same
# parameters. That is a run-time failure at the click, on a Saturday.
#
# The two are genuinely different functions that happened to share a name,
# so they now have different names. Renamed 2026-09-12; behaviour unchanged.
function Invoke-AuditUserSearch {
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
        if ($page.Count -lt $script:AuditPageSize) { break }
    }
    return [PSCustomObject]@{
        Records    = $all.ToArray()
        HitCeiling = $hitCeiling
    }
}

# THIS FILE'S OWN COPY, DELIBERATELY NAMED DIFFERENTLY FROM Common.ps1's.
#
# It used to be called ConvertTo-MailboxAuditRecord, the same as the function in
# Common.ps1 -- which this file dot-sources. The local definition wins, so
# the calls below reached THIS one and always had; but the shadowing was a
# trap with two edges. A fix made in Common never arrived here, and anyone
# tidying up an apparent duplicate would have deleted this and left the
# calls binding silently to Common's version, which does not take the same
# parameters. That is a run-time failure at the click, on a Saturday.
#
# The two are genuinely different functions that happened to share a name,
# so they now have different names. Renamed 2026-09-12; behaviour unchanged.
function ConvertTo-AuditUserRecord {
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
$form.Text = 'Mailbox Audit'
$form.ClientSize = New-Object System.Drawing.Size(1064, 708)
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
$lblTarget.Text = "Mailbox: $script:Target"
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
    if (-not "$script:Target".Trim()) {
        $lblStatus.Text = 'Enter a mailbox address above and press Look Up first.'
        return
    }
    if ($dtpEnd.Value -le $dtpStart.Value) {
        $null = Show-ToolDialog("End date/time must be after the start date/time.", 'Invalid Range', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
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
        $searchResult = Invoke-AuditUserSearch -Mailbox $script:Target -StartUtc $startUtc -EndUtc $endUtc -Operations $selectedOps -ProgressCallback $progress
        $parsed = @($searchResult.Records | ForEach-Object { ConvertTo-AuditUserRecord -RawRecord $_ })
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
        Write-CompletedLog -Task "Ran MailboxAudit search $script:Target ($($dtpStart.Value.ToString('yyyy-MM-dd HH:mm')) to $($dtpEnd.Value.ToString('yyyy-MM-dd HH:mm')); $($selectedOps.Count) operation(s) selected; $($parsed.Count) result(s))"
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Search-UnifiedAuditLog'
    } finally {
        $btnAuditNow.Visible = $true
        $btnCancelSearch.Visible = $false
    }
})
$btnCancelSearch.Add_Click({ $script:AuditCancelRequested = $true })

function Set-AuditTarget {
    # THE SEARCH IS NOT RE-RUN. It can take minutes and hit a 50,000-record
    # ceiling, so Look Up retargets and clears; 'Audit Now' stays the deliberate
    # act it always was. Clearing the grid matters more than it looks: results
    # left on screen under a new mailbox name are results about the wrong
    # person.
    $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
    $script:AuditResults = @()
    $grdAudit.Rows.Clear()
    $form.Text = if ("$script:Target".Trim()) { "Mailbox Audit: $script:Target" } else { 'Mailbox Audit' }
    $lblTarget.Text = "Mailbox: $script:Target"
    $lblStatus.Text = if ("$script:Target".Trim()) {
        "Set a date range and click 'Audit Now'."
    } else {
        'Enter a mailbox address above and press Look Up.'
    }
}

$grdAudit.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.RowIndex -ge $script:AuditResults.Count) { return }
    $rec = $script:AuditResults[$e.RowIndex]
    Show-AuditRawDetail -Json $rec.RawJson -Title "Raw Audit Record: $($rec.OperationFriendly) at $($rec.TimeLocal)"
})

$btnExport.Add_Click({
    if ($script:AuditResults.Count -eq 0) {
        $null = Show-ToolDialog("Nothing to export -- run a search first.", 'Export Results', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    try {
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $sfd.FileName = "$($script:Target -replace '[^a-zA-Z0-9]','_')_audit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $script:AuditResults | Select-Object TimeLocal, OperationFriendly, Operation, Actor, MailboxOwnerUPN, Subject, Folder, ClientInfo, ClientIP, ResultStatus |
            Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
        $lblStatus.Text = "Exported $($script:AuditResults.Count) result(s) to $($sfd.FileName)."
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export audit results'
    }
})

$btnClose.Add_Click({ $form.Close() })

# BEFORE Add-ToolHelpButton: the bar shifts every existing control down and
# grows the form, and the help button places itself from ClientSize at run
# time. Before Set-FormScreenCorner too, which positions from the size.
$bar = Add-ToolTargetBar -Form $form -InitialValue $script:Target
$bar.Button.Add_Click({ Set-AuditTarget })
Set-AuditTarget

$form.StartPosition = 'Manual'
Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Audit UserMB' | Out-Null

[void]$form.ShowDialog()
#endregion
