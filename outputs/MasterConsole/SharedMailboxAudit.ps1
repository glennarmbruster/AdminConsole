#Requires -Version 5.1
<#
    SharedMailboxAudit.ps1
    Shared MB Audit -- traces Exchange Online activity on a shared mailbox,
    cross-referenced against its current Full Access holder list.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'Shared MB Audit -- Enter Target'
if (-not $target) { return }

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
function Get-FullAccessHolders {
    param([string]$Smtp)
    $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
        Where-Object {
            $_.User -notlike 'NT AUTHORITY\SELF' -and
            $_.IsInherited -eq $false -and
            $_.Deny -eq $false -and
            ($_.AccessRights -contains 'FullAccess')
        }
    return @($fa | ForEach-Object { $_.User.ToString() })
}

function Invoke-SharedMailboxAuditSearch {
    param(
        [string]$Mailbox,
        [datetime]$StartUtc,
        [datetime]$EndUtc,
        [string[]]$Operations,
        [scriptblock]$ProgressCallback
    )
    $sessionId = "SharedMailboxAudit_$([Guid]::NewGuid().ToString())"
    $all = New-Object System.Collections.Generic.List[object]
    $hitCeiling = $false
    $iteration = 0
    while ($iteration -lt 15) {
        $iteration++
        $params = @{
            StartDate      = $StartUtc
            EndDate        = $EndUtc
            FreeText       = $Mailbox
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

function ConvertTo-SharedMailboxAuditRecord {
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
        HolderStatus      = ''
        RawJson           = if ($parsed) { ($parsed | ConvertTo-Json -Depth 8) } else { "$($RawRecord.AuditData)" }
    }
}
#endregion

#region Main window ----------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Shared MB Audit: $($target.Smtp)"
$form.Size = New-Object System.Drawing.Size(1120, 800)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Requires an active Exchange Online connection AND the 'View-Only Audit Logs' or 'Audit Logs' role assigned to your account in Microsoft Purview. Audit entries can take 30+ minutes to appear after an action happens. Default retention is 180 days unless Audit (Premium) or a custom retention policy extends it."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(1080, 48)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "Shared Mailbox: $($target.Smtp)"
$lblTarget.Location = New-Object System.Drawing.Point(12, 62)
$lblTarget.AutoSize = $true
$lblTarget.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTarget)

$grpHolders = New-Object System.Windows.Forms.GroupBox
$grpHolders.Text = "Current Full Access Holders (used to flag results below)"
$grpHolders.Location = New-Object System.Drawing.Point(12, 88)
$grpHolders.Size = New-Object System.Drawing.Size(300, 138)
$form.Controls.Add($grpHolders)
$lstHolders = New-Object System.Windows.Forms.ListBox
$lstHolders.Location = New-Object System.Drawing.Point(8, 20)
$lstHolders.Size = New-Object System.Drawing.Size(284, 80)
$lstHolders.Font = New-Object System.Drawing.Font("Consolas", 8)
$grpHolders.Controls.Add($lstHolders)
$btnRefreshHolders = New-Object System.Windows.Forms.Button
$btnRefreshHolders.Text = "Refresh Holders"
$btnRefreshHolders.Location = New-Object System.Drawing.Point(8, 104)
$btnRefreshHolders.Size = New-Object System.Drawing.Size(120, 26)
$grpHolders.Controls.Add($btnRefreshHolders)

$grpRange = New-Object System.Windows.Forms.GroupBox
$grpRange.Text = "Date Range (local time -- converted to UTC for the search)"
$grpRange.Location = New-Object System.Drawing.Point(324, 88)
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
$lblSubject.Location = New-Object System.Drawing.Point(324, 156)
$lblSubject.AutoSize = $true
$form.Controls.Add($lblSubject)
$txtSubject = New-Object System.Windows.Forms.TextBox
$txtSubject.Location = New-Object System.Drawing.Point(324, 176)
$txtSubject.Size = New-Object System.Drawing.Size(300, 22)
$txtSubject.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtSubject)

$btnAuditNow = New-Object System.Windows.Forms.Button
$btnAuditNow.Text = "Audit Now"
$btnAuditNow.Location = New-Object System.Drawing.Point(324, 206)
$btnAuditNow.Size = New-Object System.Drawing.Size(150, 32)
$btnAuditNow.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnAuditNow)
$btnCancelSearch = New-Object System.Windows.Forms.Button
$btnCancelSearch.Text = "Cancel"
$btnCancelSearch.Location = New-Object System.Drawing.Point(324, 206)
$btnCancelSearch.Size = New-Object System.Drawing.Size(150, 32)
$btnCancelSearch.Visible = $false
$form.Controls.Add($btnCancelSearch)

$grpOps = New-Object System.Windows.Forms.GroupBox
$grpOps.Text = "Operations (none checked = all operations)"
$grpOps.Location = New-Object System.Drawing.Point(742, 88)
$grpOps.Size = New-Object System.Drawing.Size(350, 260)
$form.Controls.Add($grpOps)
$clbOps = New-Object System.Windows.Forms.CheckedListBox
$clbOps.Location = New-Object System.Drawing.Point(8, 20)
$clbOps.Size = New-Object System.Drawing.Size(332, 200)
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

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Loading Full Access holders..."
$lblStatus.Location = New-Object System.Drawing.Point(12, 358)
$lblStatus.Size = New-Object System.Drawing.Size(1080, 36)
$lblStatus.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblStatus)

$grdAudit = New-Object System.Windows.Forms.DataGridView
$grdAudit.Location = New-Object System.Drawing.Point(12, 398)
$grdAudit.Size = New-Object System.Drawing.Size(1080, 300)
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
[void]$grdAudit.Columns.Add('HolderStatus', 'vs. Full Access List')
[void]$grdAudit.Columns.Add('Subject', 'Subject')
[void]$grdAudit.Columns.Add('Folder', 'Folder')
[void]$grdAudit.Columns.Add('ClientInfo', 'Client')
[void]$grdAudit.Columns.Add('ResultStatus', 'Result')
$grdAudit.Columns['TimeLocal'].FillWeight        = 13
$grdAudit.Columns['OperationFriendly'].FillWeight = 17
$grdAudit.Columns['Actor'].FillWeight            = 15
$grdAudit.Columns['HolderStatus'].FillWeight     = 16
$grdAudit.Columns['Subject'].FillWeight          = 17
$grdAudit.Columns['Folder'].FillWeight           = 8
$grdAudit.Columns['ClientInfo'].FillWeight       = 10
$grdAudit.Columns['ResultStatus'].FillWeight     = 7
foreach ($col in $grdAudit.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
$form.Controls.Add($grdAudit)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Results to CSV"
$btnExport.Location = New-Object System.Drawing.Point(12, 708)
$btnExport.Size = New-Object System.Drawing.Size(170, 28)
$form.Controls.Add($btnExport)
$lblRowHint = New-Object System.Windows.Forms.Label
$lblRowHint.Text = "Double-click a row to see its full raw audit record. 'NOT a current holder' can mean revoked access, an admin/owner action, or a permission type other than Full Access."
$lblRowHint.Location = New-Object System.Drawing.Point(196, 714)
$lblRowHint.Size = New-Object System.Drawing.Size(770, 20)
$lblRowHint.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
$lblRowHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblRowHint)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(992, 708)
$btnClose.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnClose)

$script:AuditResults = @()
$script:AuditCancelRequested = $false
$script:CurrentHolders = @()

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

function Update-HolderList {
    $lstHolders.Items.Clear()
    try {
        $script:CurrentHolders = Get-FullAccessHolders -Smtp $target.Smtp
        if ($script:CurrentHolders.Count -eq 0) {
            [void]$lstHolders.Items.Add('(none found)')
        } else {
            foreach ($h in $script:CurrentHolders) { [void]$lstHolders.Items.Add($h) }
        }
        $lblStatus.Text = "$($script:CurrentHolders.Count) Full Access holder(s) loaded. Set a date range and click 'Audit Now.'"
    } catch {
        $script:CurrentHolders = @()
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access holders)'
    }
}
$btnRefreshHolders.Add_Click({ Update-HolderList })
Update-HolderList

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
        $searchResult = Invoke-SharedMailboxAuditSearch -Mailbox $target.Smtp -StartUtc $startUtc -EndUtc $endUtc -Operations $selectedOps -ProgressCallback $progress
        $parsed = @($searchResult.Records | ForEach-Object { ConvertTo-SharedMailboxAuditRecord -RawRecord $_ })
        $subjectFilter = $txtSubject.Text.Trim()
        if ($subjectFilter -ne '') {
            $parsed = @($parsed | Where-Object { $_.Subject -like "*$subjectFilter*" })
        }
        foreach ($rec in $parsed) {
            $isHolder = $false
            foreach ($h in $script:CurrentHolders) {
                if ($h -and $rec.Actor -and $h.ToLowerInvariant() -eq $rec.Actor.ToLowerInvariant()) { $isHolder = $true; break }
            }
            $rec.HolderStatus = if ($isHolder) { 'Full Access holder' } else { 'NOT a current holder' }
        }
        $script:AuditResults = $parsed
        foreach ($rec in $parsed) {
            [void]$grdAudit.Rows.Add($rec.TimeLocal.ToString('yyyy-MM-dd HH:mm:ss'), $rec.OperationFriendly, $rec.Actor, $rec.HolderStatus, $rec.Subject, $rec.Folder, $rec.ClientInfo, $rec.ResultStatus)
        }
        $statusMsg = "$($parsed.Count) result(s)."
        if ($script:AuditCancelRequested) { $statusMsg += " Search cancelled -- results may be incomplete." }
        if ($searchResult.HitCeiling) { $statusMsg += " Hit the 50,000-record search ceiling -- narrow the date range or Operations filter to be sure you have everything in this window." }
        if ($parsed.Count -eq 0 -and -not $script:AuditCancelRequested) {
            $statusMsg += " Zero results can mean nothing happened, OR that this falls outside the retention window, OR that mailbox auditing isn't capturing this action -- it doesn't necessarily mean 'confirmed nothing happened.'"
        }
        $lblStatus.Text = $statusMsg
        Write-CompletedLog -Task "Ran SharedMailboxAudit search $($target.Smtp) ($($dtpStart.Value.ToString('yyyy-MM-dd HH:mm')) to $($dtpEnd.Value.ToString('yyyy-MM-dd HH:mm')); $($selectedOps.Count) operation(s) selected; $($script:CurrentHolders.Count) known holder(s); $($parsed.Count) result(s))"
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
        $sfd.FileName = "$($target.Smtp -replace '[^a-zA-Z0-9]','_')_sharedaudit_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $script:AuditResults | Select-Object TimeLocal, OperationFriendly, Operation, Actor, HolderStatus, MailboxOwnerUPN, Subject, Folder, ClientInfo, ClientIP, ResultStatus |
            Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
        $lblStatus.Text = "Exported $($script:AuditResults.Count) result(s) to $($sfd.FileName)."
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export audit results'
    }
})

$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
#endregion
