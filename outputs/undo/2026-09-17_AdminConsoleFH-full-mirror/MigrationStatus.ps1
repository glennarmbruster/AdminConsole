#Requires -Version 5.1
<#
    MigrationStatus.ps1
    Shows hybrid migration batches and move requests in Exchange Online, and lets the operator nudge a stuck move.

    Author   : Glenn Armbruster
    Revised  : 2026-09-12
#>
# TOOLKIT-HELP-RUNS: the Exchange Online migration service. Reads batches and move requests; WRITES only when you press Complete, Start, Stop, Resume, Suspend, Remove or Set bad-item limit and answer Yes.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# The state of hybrid mailbox moves, in two tabs.
#
#   Batches       -- every migration batch with its counts. "Batch details"
#                    shows the batch report. Start, Stop and Complete act on
#                    the selected batch after a preview and a Yes.
#   Move requests -- the move for one mailbox (or every move if the box is
#                    blank). Statistics are fetched per SELECTED row, because
#                    that call is slow and the question is nearly always about
#                    one mailbox. "Details" shows the last 40 report entries.
#                    Resume, Suspend and Set bad-item limit act on the
#                    selected move. Exchange Online retired -BadItemLimit; on
#                    a build without it the same button offers to approve the
#                    skipped items instead, which is what unblocks a cloud move.
#
# Refuses to remove a move request unless its Status is Completed.
# Every write previews, asks, and is written to the audit log.
# "Save CSV" writes the visible grid to the Reports folder.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# FIRST THING AFTER Common LOADS -- see Test-ConsoleLaunch in Common.ps1.
if (-not (Test-ConsoleLaunch)) { return }

#region Configuration --------------------------------------------------------
$ReportDir  = Get-ReportPath
$ReportTail = 40     # report entries shown by the Details buttons
$AutoStats  = 5      # statistics are fetched automatically for this many rows or fewer
#endregion

#region Helpers --------------------------------------------------------------
function Assert-ExoCmdlet {
    # A plain sentence instead of "the term is not recognized" -- the fix is
    # always the same one, so say it.
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Command -Name $Name -ErrorAction SilentlyContinue)) {
        throw "$Name is not available in this session. Connect to Exchange Online first."
    }
}

function Format-Stamp {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return '' }
    try { return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm') } catch { return "$Value" }
}

function ConvertTo-IntOrNull {
    # DBNull, not 0: a count the service did not return is not a count of zero.
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return [System.DBNull]::Value }
    try { return [int]$Value } catch { return [System.DBNull]::Value }
}

function Format-ReportEntries {
    param($Report, [int]$Tail = 40)
    $entries = @()
    if ($Report -and $Report.Entries) { $entries = @($Report.Entries) }
    if ($entries.Count -eq 0) { return '(no report entries)' }
    $shown = $entries
    if ($entries.Count -gt $Tail) { $shown = @($entries[($entries.Count - $Tail)..($entries.Count - 1)]) }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Report entries: $($entries.Count) total, last $(@($shown).Count) shown, oldest first.")
    [void]$sb.AppendLine('')
    foreach ($e in $shown) {
        [void]$sb.AppendLine(("{0}  {1}" -f (Format-Stamp -Value $e.CreationTime), ("$($e.Message)" -replace '\s+', ' ')))
    }
    return $sb.ToString()
}

function Confirm-Write {
    param([string]$Text, [string]$Title)
    $r = Show-ToolDialog($Text, $Title, 'YesNo', 'Warning', 'Button2')
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Get-SingleSelectedRow {
    # The DataRow behind the one selected grid row. Actions are one-at-a-time
    # on purpose: this tool exists for one final move, not a fleet.
    param([Parameter(Mandatory)]$Grid, [string]$What = 'row')
    if ($Grid.SelectedRows.Count -ne 1) { throw "Select exactly one $What first." }
    $item = $Grid.SelectedRows[0].DataBoundItem
    if (-not $item) { throw "Select a $What first." }
    return $item.Row
}

function Export-GridCsv {
    param([Parameter(Mandatory)]$Grid, [Parameter(Mandatory)][string]$Prefix)
    $t = $Grid.DataSource
    if (-not $t -or $t.Rows.Count -eq 0) { throw 'Nothing to save -- the grid is empty.' }
    $rows = foreach ($r in $t.Rows) {
        $o = [ordered]@{}
        foreach ($c in $t.Columns) {
            if ($c.ColumnName -eq 'Key') { continue }
            $v = $r[$c]
            if ($v -is [System.DBNull]) { $v = $null }
            $o[$c.ColumnName] = $v
        }
        [PSCustomObject]$o
    }
    if (-not (Test-Path -LiteralPath $ReportDir)) {
        New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
    }
    $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $csv   = Join-Path $ReportDir "$($Prefix)_$stamp.csv"
    $rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
    return $csv
}
#endregion

#region Data -- batches ------------------------------------------------------
function Get-BatchTable {
    param([scriptblock]$Progress)
    Assert-ExoCmdlet -Name 'Get-MigrationBatch'
    & $Progress 'Reading migration batches...' 10

    # ResultSize defaults to 1000 in the cloud; only passed where declared.
    $params   = @{ ErrorAction = 'Stop' }
    $declared = Get-CmdletParameterNames -Name 'Get-MigrationBatch'
    if ($declared -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
    $batches = @(Invoke-EXOCommand -Name 'Get-MigrationBatch' -Parameters $params)

    & $Progress "Building the list: $($batches.Count) batch(es)..." 70

    $t = New-Object System.Data.DataTable
    foreach ($c in @('Identity', 'Status')) { [void]$t.Columns.Add($c) }
    foreach ($c in @('Total', 'Active', 'Synced', 'Finalized', 'Failed')) { [void]$t.Columns.Add($c, [int]) }
    foreach ($c in @('Created', 'LastSynced')) { [void]$t.Columns.Add($c) }

    foreach ($b in $batches) {
        [void]$t.Rows.Add(
            "$($b.Identity)", "$($b.Status)",
            (ConvertTo-IntOrNull -Value $b.TotalCount),
            (ConvertTo-IntOrNull -Value $b.ActiveCount),
            (ConvertTo-IntOrNull -Value $b.SyncedCount),
            (ConvertTo-IntOrNull -Value $b.FinalizedCount),
            (ConvertTo-IntOrNull -Value $b.FailedCount),
            (Format-Stamp -Value $b.CreationDateTime),
            (Format-Stamp -Value $b.LastSyncedDateTime))
    }
    return ,$t
}

function Get-BatchDetailText {
    param([Parameter(Mandatory)][string]$Identity)
    Assert-ExoCmdlet -Name 'Get-MigrationBatch'
    $b = Invoke-EXOCommand -Name 'Get-MigrationBatch' -Parameters @{
        Identity = $Identity; IncludeReport = $true; ErrorAction = 'Stop' }
    if (-not $b) { throw "Batch '$Identity' was not returned." }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Batch        : $($b.Identity)")
    [void]$sb.AppendLine("Status       : $($b.Status)   State: $($b.State)")
    [void]$sb.AppendLine("Type         : $($b.MigrationType)   Direction: $($b.BatchDirection)")
    [void]$sb.AppendLine("Endpoint     : $($b.SourceEndpoint)")
    [void]$sb.AppendLine("Counts       : total $($b.TotalCount), active $($b.ActiveCount), synced $($b.SyncedCount), finalized $($b.FinalizedCount), failed $($b.FailedCount), stopped $($b.StoppedCount)")
    [void]$sb.AppendLine("Created      : $(Format-Stamp -Value $b.CreationDateTime)   Started: $(Format-Stamp -Value $b.StartDateTime)")
    [void]$sb.AppendLine("Last synced  : $(Format-Stamp -Value $b.LastSyncedDateTime)   Completed: $(Format-Stamp -Value $b.CompletedDateTime)")
    [void]$sb.AppendLine("Complete after: $(Format-Stamp -Value $b.CompleteAfter)")
    if ("$($b.Message)".Trim()) { [void]$sb.AppendLine("Message      : $($b.Message)") }
    [void]$sb.AppendLine('')
    [void]$sb.Append((Format-ReportEntries -Report $b.Report -Tail $ReportTail))
    return $sb.ToString()
}
#endregion

#region Data -- move requests ------------------------------------------------
function New-MoveTable {
    $t = New-Object System.Data.DataTable
    foreach ($c in @('Key', 'DisplayName', 'Status', 'StatusDetail')) { [void]$t.Columns.Add($c) }
    [void]$t.Columns.Add('PercentComplete', [int])
    foreach ($c in @('BytesTransferred', 'BadItems', 'LargeItems', 'TotalMailboxSize',
                     'Started', 'LastUpdate', 'Message', 'BatchName', 'Alias')) {
        [void]$t.Columns.Add($c)
    }
    return ,$t
}

function Get-MoveTable {
    param([string]$Filter, [scriptblock]$Progress)
    Assert-ExoCmdlet -Name 'Get-MoveRequest'

    $params   = @{ ErrorAction = 'Stop' }
    $declared = Get-CmdletParameterNames -Name 'Get-MoveRequest'
    $who = "$Filter".Trim()
    if ($who) {
        & $Progress "Looking for the move request for $who..." 10
        $params['Identity'] = $who
    } else {
        & $Progress 'Reading every move request in the tenant...' 10
        if ($declared -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
    }
    $moves = @(Invoke-EXOCommand -Name 'Get-MoveRequest' -Parameters $params)

    & $Progress "Building the list: $($moves.Count) move request(s)..." 60

    $t = New-MoveTable
    foreach ($m in $moves) {
        # The GUID is the identity that survives a rename; fall back to the
        # object's own Identity when the build does not expose it.
        $key = "$($m.ExchangeGuid)"
        if (-not $key -or $key -eq ([guid]::Empty).ToString()) { $key = "$($m.Identity)" }

        $row = $t.NewRow()
        $row['Key']         = $key
        $row['DisplayName'] = "$($m.DisplayName)"
        $row['Status']      = "$($m.Status)"
        $row['BatchName']   = "$($m.BatchName)"
        $row['Alias']       = "$($m.Alias)"
        $t.Rows.Add($row)
    }
    return ,$t
}

function Update-MoveRow {
    # One statistics call for ONE row. Deliberately not done for the whole
    # grid: Get-MoveRequestStatistics is the slow call, and a tenant-wide
    # refresh would take minutes to answer a question about one mailbox.
    param([Parameter(Mandatory)][System.Data.DataRow]$Row)
    Assert-ExoCmdlet -Name 'Get-MoveRequestStatistics'
    $s = Invoke-EXOCommand -Name 'Get-MoveRequestStatistics' -Parameters @{
        Identity = "$($Row['Key'])"; ErrorAction = 'Stop' }
    if (-not $s) { throw "No statistics returned for $($Row['DisplayName'])." }

    $Row['Status']           = "$($s.Status)"
    $Row['StatusDetail']     = "$($s.StatusDetail)"
    $Row['PercentComplete']  = ConvertTo-IntOrNull -Value $s.PercentComplete
    $Row['BytesTransferred'] = "$($s.BytesTransferred)"
    $Row['BadItems']         = "$($s.BadItemsEncountered)"
    $Row['LargeItems']       = "$($s.LargeItemsEncountered)"
    $Row['TotalMailboxSize'] = "$($s.TotalMailboxSize)"
    $Row['Started']          = Format-Stamp -Value $s.StartTimestamp
    $Row['LastUpdate']       = Format-Stamp -Value $s.LastUpdateTimestamp
    $Row['Message']          = ("$($s.Message)" -replace '\s+', ' ')
    if ("$($s.BatchName)".Trim()) { $Row['BatchName'] = "$($s.BatchName)" }
}

function Get-MoveDetailText {
    param([Parameter(Mandatory)][string]$Key, [switch]$Diagnostics)
    Assert-ExoCmdlet -Name 'Get-MoveRequestStatistics'
    $params = @{ Identity = $Key; IncludeReport = $true; ErrorAction = 'Stop' }
    $wantDiag = $false
    if ($Diagnostics) {
        $declared = Get-CmdletParameterNames -Name 'Get-MoveRequestStatistics'
        if ($declared -contains 'DiagnosticInfo') { $params['DiagnosticInfo'] = 'Verbose'; $wantDiag = $true }
    }
    $s = Invoke-EXOCommand -Name 'Get-MoveRequestStatistics' -Parameters $params
    if (-not $s) { throw "No statistics returned for '$Key'." }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Mailbox        : $($s.DisplayName)   ($($s.Alias))")
    [void]$sb.AppendLine("Status         : $($s.Status) / $($s.StatusDetail)   Sync stage: $($s.SyncStage)")
    [void]$sb.AppendLine("Progress       : $($s.PercentComplete)%   $($s.BytesTransferred) of $($s.TotalMailboxSize)   items $($s.ItemsTransferred) of $($s.TotalMailboxItemCount)")
    [void]$sb.AppendLine("Bad / large    : $($s.BadItemsEncountered) bad, $($s.LargeItemsEncountered) large   (limits: bad $($s.BadItemLimit), large $($s.LargeItemLimit))")
    [void]$sb.AppendLine("Data score     : $($s.DataConsistencyScore)")
    [void]$sb.AppendLine("Batch          : $($s.BatchName)   Complete after: $(Format-Stamp -Value $s.CompleteAfter)")
    [void]$sb.AppendLine("Started        : $(Format-Stamp -Value $s.StartTimestamp)   Last update: $(Format-Stamp -Value $s.LastUpdateTimestamp)   Last success: $(Format-Stamp -Value $s.LastSuccessfulSyncTimestamp)")
    [void]$sb.AppendLine("Duration       : overall $($s.OverallDuration), in progress $($s.TotalInProgressDuration), queued $($s.TotalQueuedDuration), stalled $($s.TotalStalledDueToHADuration)")
    if ("$($s.FailureType)".Trim()) {
        [void]$sb.AppendLine("Failure        : $($s.FailureType) on $($s.FailureSide) at $(Format-Stamp -Value $s.FailureTimestamp)")
    }
    if ("$($s.Message)".Trim()) { [void]$sb.AppendLine("Message        : $($s.Message)") }
    [void]$sb.AppendLine('')
    [void]$sb.Append((Format-ReportEntries -Report $s.Report -Tail $ReportTail))
    if ($wantDiag) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('--- DiagnosticInfo (Verbose) ---')
        [void]$sb.AppendLine("$($s.DiagnosticInfo)")
    }
    return $sb.ToString()
}
#endregion

#region Writes ---------------------------------------------------------------
function Invoke-BatchWrite {
    # Preview, ask, act, log. Returns $true only if the cmdlet ran.
    param(
        [Parameter(Mandatory)][System.Data.DataRow]$Row,
        [Parameter(Mandatory)][string]$Cmdlet,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Verb,
        [string]$Note = ''
    )
    Assert-ExoCmdlet -Name $Cmdlet
    $id = "$($Row['Identity'])"
    $preview = "$Verb migration batch:" + "`r`n`r`n" +
               "  Batch  : $id" + "`r`n" +
               "  Status : $($Row['Status'])" + "`r`n" +
               "  Counts : total $($Row['Total']), active $($Row['Active']), synced $($Row['Synced']), finalized $($Row['Finalized']), failed $($Row['Failed'])" + "`r`n`r`n" +
               "Command: $Cmdlet -Identity '$id'"
    if ($Note) { $preview += "`r`n`r`n$Note" }
    if (-not (Confirm-Write -Text $preview -Title "Confirm -- $Verb batch")) { return $false }

    try {
        Invoke-EXOCommand -Name $Cmdlet -Parameters @{ Identity = $id; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
        Write-CompletedLog -Action $Action -Target $id -Result 'Success' -Detail "$Cmdlet; status before: $($Row['Status'])"
        return $true
    } catch {
        Write-FailedLog -Action $Action -Target $id -Detail $Cmdlet -ErrorRecord $_
        throw
    }
}

function Invoke-MoveWrite {
    param(
        [Parameter(Mandatory)][System.Data.DataRow]$Row,
        [Parameter(Mandatory)][string]$Cmdlet,
        [Parameter(Mandatory)][string]$Action,
        [Parameter(Mandatory)][string]$Verb,
        [hashtable]$Extra = @{},
        [string]$Note = ''
    )
    Assert-ExoCmdlet -Name $Cmdlet
    $key    = "$($Row['Key'])"
    $name   = "$($Row['DisplayName'])"
    $target = Get-LogTarget "$($Row['Alias'])" $name

    $extraText = ''
    foreach ($k in @($Extra.Keys | Sort-Object)) { $extraText += " -$k $($Extra[$k])" }

    $preview = "$Verb the move request for:" + "`r`n`r`n" +
               "  Mailbox : $name" + "`r`n" +
               "  Status  : $($Row['Status']) $($Row['StatusDetail'])" + "`r`n" +
               "  Batch   : $($Row['BatchName'])" + "`r`n" +
               "  Progress: $($Row['PercentComplete'])%   bad items $($Row['BadItems']), large items $($Row['LargeItems'])" + "`r`n`r`n" +
               "Command: $Cmdlet -Identity '$key'$extraText"
    if ($Note) { $preview += "`r`n`r`n$Note" }
    if (-not (Confirm-Write -Text $preview -Title "Confirm -- $Verb move")) { return $false }

    $p = @{ Identity = $key; Confirm = $false; ErrorAction = 'Stop' }
    foreach ($k in $Extra.Keys) { $p[$k] = $Extra[$k] }
    try {
        Invoke-EXOCommand -Name $Cmdlet -Parameters $p | Out-Null
        Write-CompletedLog -Action $Action -Target $target -Result 'Success' `
            -Detail "$Cmdlet$extraText; status before: $($Row['Status']) $($Row['StatusDetail'])"
        return $true
    } catch {
        Write-FailedLog -Action $Action -Target $target -Detail "$Cmdlet$extraText" -ErrorRecord $_
        throw
    }
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Migration Status -- batches and move requests'
$form.Size            = New-Object System.Drawing.Size(1000, 700)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Exchange Online migration batches and the move requests inside them. Select a row, then use the buttons. Every write asks first.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(960, 18)
$form.Controls.Add($lblInfo)

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, 36)
$tabs.Size     = New-Object System.Drawing.Size(960, 470)
$form.Controls.Add($tabs)

$pgBatches      = New-Object System.Windows.Forms.TabPage
$pgBatches.Text = 'Batches'
$pgMoves        = New-Object System.Windows.Forms.TabPage
$pgMoves.Text   = 'Move requests'
[void]$tabs.TabPages.Add($pgBatches)
[void]$tabs.TabPages.Add($pgMoves)

# ---- Batches tab -----------------------------------------------------------
$btnBatchRefresh          = New-Object System.Windows.Forms.Button
$btnBatchRefresh.Text     = 'Refresh'
$btnBatchRefresh.Location = New-Object System.Drawing.Point(12, 12)
$btnBatchRefresh.Size     = New-Object System.Drawing.Size(110, 30)
$pgBatches.Controls.Add($btnBatchRefresh)

$btnBatchDetails          = New-Object System.Windows.Forms.Button
$btnBatchDetails.Text     = 'Batch details'
$btnBatchDetails.Location = New-Object System.Drawing.Point(134, 12)
$btnBatchDetails.Size     = New-Object System.Drawing.Size(120, 30)
$pgBatches.Controls.Add($btnBatchDetails)

$btnBatchStart          = New-Object System.Windows.Forms.Button
$btnBatchStart.Text     = 'Start batch'
$btnBatchStart.Location = New-Object System.Drawing.Point(266, 12)
$btnBatchStart.Size     = New-Object System.Drawing.Size(110, 30)
$pgBatches.Controls.Add($btnBatchStart)

$btnBatchStop          = New-Object System.Windows.Forms.Button
$btnBatchStop.Text     = 'Stop batch'
$btnBatchStop.Location = New-Object System.Drawing.Point(388, 12)
$btnBatchStop.Size     = New-Object System.Drawing.Size(110, 30)
$pgBatches.Controls.Add($btnBatchStop)

$btnBatchComplete          = New-Object System.Windows.Forms.Button
$btnBatchComplete.Text     = 'Complete batch'
$btnBatchComplete.Location = New-Object System.Drawing.Point(510, 12)
$btnBatchComplete.Size     = New-Object System.Drawing.Size(130, 30)
$pgBatches.Controls.Add($btnBatchComplete)

$grdBatches                       = New-Object System.Windows.Forms.DataGridView
$grdBatches.Location              = New-Object System.Drawing.Point(12, 54)
$grdBatches.Size                  = New-Object System.Drawing.Size(928, 170)
$grdBatches.ReadOnly              = $true
$grdBatches.AllowUserToAddRows    = $false
$grdBatches.AllowUserToDeleteRows = $false
$grdBatches.RowHeadersVisible     = $false
$grdBatches.MultiSelect           = $false
$grdBatches.SelectionMode         = 'FullRowSelect'
$grdBatches.AutoSizeColumnsMode   = 'AllCells'
$pgBatches.Controls.Add($grdBatches)

$txtBatch            = New-Object System.Windows.Forms.TextBox
$txtBatch.Multiline  = $true
$txtBatch.ReadOnly   = $true
$txtBatch.ScrollBars = 'Both'
$txtBatch.WordWrap   = $false
$txtBatch.MaxLength  = 2147483647
$txtBatch.Location   = New-Object System.Drawing.Point(12, 232)
$txtBatch.Size       = New-Object System.Drawing.Size(928, 200)
$txtBatch.Font       = New-Object System.Drawing.Font('Consolas', 9)
$pgBatches.Controls.Add($txtBatch)

# ---- Move requests tab -----------------------------------------------------
$lblFilter          = New-Object System.Windows.Forms.Label
$lblFilter.Text     = 'Mailbox (blank = every move request):'
$lblFilter.Location = New-Object System.Drawing.Point(12, 18)
$lblFilter.Size     = New-Object System.Drawing.Size(230, 18)
$pgMoves.Controls.Add($lblFilter)

$txtFilter          = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(246, 15)
$txtFilter.Size     = New-Object System.Drawing.Size(300, 24)
$txtFilter.Text     = "$SmtpAddress"
$pgMoves.Controls.Add($txtFilter)

$btnMoveRefresh          = New-Object System.Windows.Forms.Button
$btnMoveRefresh.Text     = 'Refresh'
$btnMoveRefresh.Location = New-Object System.Drawing.Point(558, 12)
$btnMoveRefresh.Size     = New-Object System.Drawing.Size(110, 30)
$pgMoves.Controls.Add($btnMoveRefresh)

$btnMoveStats          = New-Object System.Windows.Forms.Button
$btnMoveStats.Text     = 'Load statistics'
$btnMoveStats.Location = New-Object System.Drawing.Point(680, 12)
$btnMoveStats.Size     = New-Object System.Drawing.Size(120, 30)
$pgMoves.Controls.Add($btnMoveStats)

$btnMoveDetails          = New-Object System.Windows.Forms.Button
$btnMoveDetails.Text     = 'Details'
$btnMoveDetails.Location = New-Object System.Drawing.Point(12, 50)
$btnMoveDetails.Size     = New-Object System.Drawing.Size(110, 30)
$pgMoves.Controls.Add($btnMoveDetails)

$chkVerbose          = New-Object System.Windows.Forms.CheckBox
$chkVerbose.Text     = 'Verbose diagnostics'
$chkVerbose.Location = New-Object System.Drawing.Point(134, 55)
$chkVerbose.Size     = New-Object System.Drawing.Size(150, 20)
$pgMoves.Controls.Add($chkVerbose)

$btnMoveResume          = New-Object System.Windows.Forms.Button
$btnMoveResume.Text     = 'Resume'
$btnMoveResume.Location = New-Object System.Drawing.Point(296, 50)
$btnMoveResume.Size     = New-Object System.Drawing.Size(110, 30)
$pgMoves.Controls.Add($btnMoveResume)

$btnMoveSuspend          = New-Object System.Windows.Forms.Button
$btnMoveSuspend.Text     = 'Suspend'
$btnMoveSuspend.Location = New-Object System.Drawing.Point(418, 50)
$btnMoveSuspend.Size     = New-Object System.Drawing.Size(110, 30)
$pgMoves.Controls.Add($btnMoveSuspend)

$btnMoveRemove          = New-Object System.Windows.Forms.Button
$btnMoveRemove.Text     = 'Remove completed'
$btnMoveRemove.Location = New-Object System.Drawing.Point(540, 50)
$btnMoveRemove.Size     = New-Object System.Drawing.Size(130, 30)
$pgMoves.Controls.Add($btnMoveRemove)

$btnMoveBadItems          = New-Object System.Windows.Forms.Button
$btnMoveBadItems.Text     = 'Set bad-item limit'
$btnMoveBadItems.Location = New-Object System.Drawing.Point(682, 50)
$btnMoveBadItems.Size     = New-Object System.Drawing.Size(130, 30)
$pgMoves.Controls.Add($btnMoveBadItems)

$grdMoves                       = New-Object System.Windows.Forms.DataGridView
$grdMoves.Location              = New-Object System.Drawing.Point(12, 92)
$grdMoves.Size                  = New-Object System.Drawing.Size(928, 160)
$grdMoves.ReadOnly              = $true
$grdMoves.AllowUserToAddRows    = $false
$grdMoves.AllowUserToDeleteRows = $false
$grdMoves.RowHeadersVisible     = $false
$grdMoves.MultiSelect           = $true
$grdMoves.SelectionMode         = 'FullRowSelect'
$grdMoves.AutoSizeColumnsMode   = 'AllCells'
$pgMoves.Controls.Add($grdMoves)

$txtMove            = New-Object System.Windows.Forms.TextBox
$txtMove.Multiline  = $true
$txtMove.ReadOnly   = $true
$txtMove.ScrollBars = 'Both'
$txtMove.WordWrap   = $false
$txtMove.MaxLength  = 2147483647
$txtMove.Location   = New-Object System.Drawing.Point(12, 260)
$txtMove.Size       = New-Object System.Drawing.Size(928, 172)
$txtMove.Font       = New-Object System.Drawing.Font('Consolas', 9)
$pgMoves.Controls.Add($txtMove)

# ---- Below the tabs --------------------------------------------------------
$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 514)
$lblStatus.Size     = New-Object System.Drawing.Size(960, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 534 -Width 960 -Height 12

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'MigrationStatus' `
                -X 12 -Y 554 -Width 960 -StatusLabel $lblStatus

$btnSave          = New-Object System.Windows.Forms.Button
$btnSave.Text     = 'Save CSV'
$btnSave.Location = New-Object System.Drawing.Point(12, 592)
$btnSave.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnSave)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(134, 592)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(862, 592)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$btnOpen.Add_Click({
    try { Open-ReportFolder -Path $ReportDir }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Open Reports folder' }
})
#endregion

#region Handlers -- batches --------------------------------------------------
$btnBatchRefresh.Add_Click({
    $btnBatchRefresh.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Migration batches'
        $t = Get-BatchTable -Progress $progress
        $grdBatches.DataSource = $t
        Complete-ProgressReporter -Activity 'Migration batches' -Bar $pbar
        $lblStatus.Text = "$($t.Rows.Count) migration batch(es)."
        if ($t.Rows.Count -eq 0) { $txtBatch.Text = 'No migration batches exist in this tenant.' }
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtBatch.Text  = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MigrationBatch'
    } finally {
        Complete-ProgressReporter -Activity 'Migration batches'
        $form.Cursor = $prev
        $btnBatchRefresh.Enabled = $true
    }
})

$btnBatchDetails.Add_Click({
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $row = Get-SingleSelectedRow -Grid $grdBatches -What 'batch'
        $lblStatus.Text = "Reading the report for $($row['Identity'])..."
        $txtBatch.Text  = Get-BatchDetailText -Identity "$($row['Identity'])"
        $lblStatus.Text = "Report for $($row['Identity'])."
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtBatch.Text  = Format-ErrorDetail -ErrorRecord $_ -Context 'Batch details'
    } finally {
        $form.Cursor = $prev
    }
})

$btnBatchStart.Add_Click({
    try {
        $row = Get-SingleSelectedRow -Grid $grdBatches -What 'batch'
        $note = 'Starts a Created batch, resumes a Stopped one, or retries the failed users of a Failed, SyncedWithErrors or CompletedWithErrors batch.'
        if (Invoke-BatchWrite -Row $row -Cmdlet 'Start-MigrationBatch' -Action 'StartBatch' -Verb 'Start' -Note $note) {
            $lblStatus.Text = "Start requested for $($row['Identity']). Refresh to see the new status."
            $btnBatchRefresh.PerformClick()
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Start batch' }
})

$btnBatchStop.Add_Click({
    try {
        $row = Get-SingleSelectedRow -Grid $grdBatches -What 'batch'
        $note = 'Mailboxes still syncing stop immediately. Mailboxes already migrated are not affected. The service refuses if every user in the batch is already completed or failed.'
        if (Invoke-BatchWrite -Row $row -Cmdlet 'Stop-MigrationBatch' -Action 'StopBatch' -Verb 'Stop' -Note $note) {
            $lblStatus.Text = "Stop requested for $($row['Identity']). Refresh to see the new status."
            $btnBatchRefresh.PerformClick()
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Stop batch' }
})

$btnBatchComplete.Add_Click({
    try {
        $row = Get-SingleSelectedRow -Grid $grdBatches -What 'batch'
        $note = 'This is the cut-over: a final incremental sync, then the source mailbox becomes a mail user and Outlook is repointed. ' +
                'In Exchange Online it sets CompleteAfter to now; a second request within 8 hours may be ignored by the service. ' +
                'Expected only when Status is Synced or SyncedWithErrors.'
        $st = "$($row['Status'])"
        if ($st -notin @('Synced', 'SyncedWithErrors', 'SyncedwithErrors')) {
            $note = "*** Status is '$st', not Synced. The service will most likely reject this. ***`r`n`r`n" + $note
        }
        if (Invoke-BatchWrite -Row $row -Cmdlet 'Complete-MigrationBatch' -Action 'CompleteBatch' -Verb 'Complete' -Note $note) {
            $lblStatus.Text = "Completion requested for $($row['Identity']). Refresh to watch it finalize."
            $btnBatchRefresh.PerformClick()
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Complete batch' }
})
#endregion

#region Handlers -- move requests --------------------------------------------
function Update-SelectedMoveRows {
    # Statistics for the selected rows only. Called by the button and, when
    # the list is short, by Refresh -- one final move should not need a
    # second click to show its progress.
    param([Parameter(Mandatory)]$Grid, [System.Windows.Forms.Label]$Label, $Bar, [switch]$AllRows)
    $t = $Grid.DataSource
    if (-not $t) { throw 'Refresh first.' }
    $rows = @()
    if ($AllRows) {
        foreach ($r in $t.Rows) { $rows += $r }
    } else {
        foreach ($gr in $Grid.SelectedRows) { if ($gr.DataBoundItem) { $rows += $gr.DataBoundItem.Row } }
    }
    if ($rows.Count -eq 0) { throw 'Select one or more move requests first.' }

    $progress = New-ProgressReporter -Label $Label -Bar $Bar -Activity 'Move statistics'
    $failed = @()
    $n = 0
    foreach ($r in $rows) {
        $n++
        & $progress "Statistics: $n of $($rows.Count) -- $($r['DisplayName'])..." ([int](100.0 * $n / $rows.Count))
        try { Update-MoveRow -Row $r }
        catch { $failed += "$($r['DisplayName']) : $($_.Exception.Message)" }
    }
    Complete-ProgressReporter -Activity 'Move statistics' -Bar $Bar
    $Grid.Refresh()
    return $failed
}

$btnMoveRefresh.Add_Click({
    $btnMoveRefresh.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtMove.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Move requests'
        $t = Get-MoveTable -Filter $txtFilter.Text -Progress $progress
        $grdMoves.DataSource = $t
        if ($grdMoves.Columns.Contains('Key'))     { $grdMoves.Columns['Key'].Visible = $false }
        if ($grdMoves.Columns.Contains('Alias'))   { $grdMoves.Columns['Alias'].Visible = $false }
        if ($grdMoves.Columns.Contains('Message')) {
            $grdMoves.Columns['Message'].AutoSizeMode = 'None'
            $grdMoves.Columns['Message'].Width = 320
        }
        Complete-ProgressReporter -Activity 'Move requests' -Bar $pbar

        if ($t.Rows.Count -eq 0) {
            $lblStatus.Text = 'No move request found.'
            $txtMove.Text   = 'No move request matched. A mailbox with no move request has either not been added to a batch yet, or its completed request has already been removed.'
            return
        }

        $lblStatus.Text = "$($t.Rows.Count) move request(s)."
        if ($t.Rows.Count -le $AutoStats) {
            $failed = @(Update-SelectedMoveRows -Grid $grdMoves -Label $lblStatus -Bar $pbar -AllRows)
            $lblStatus.Text = "$($t.Rows.Count) move request(s), statistics loaded."
            if ($failed.Count -gt 0) {
                $txtMove.Text = "Statistics could not be read for:`r`n  " + ($failed -join "`r`n  ")
            }
        } else {
            $lblStatus.Text = "$($t.Rows.Count) move request(s). Select rows and press Load statistics -- not fetched for all, it is slow."
        }
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtMove.Text   = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MoveRequest'
    } finally {
        Complete-ProgressReporter -Activity 'Move requests'
        $form.Cursor = $prev
        $btnMoveRefresh.Enabled = $true
    }
})

$btnMoveStats.Add_Click({
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $failed = @(Update-SelectedMoveRows -Grid $grdMoves -Label $lblStatus -Bar $pbar)
        $lblStatus.Text = 'Statistics loaded for the selected row(s).'
        if ($failed.Count -gt 0) {
            $txtMove.Text = "Statistics could not be read for:`r`n  " + ($failed -join "`r`n  ")
        }
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtMove.Text   = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MoveRequestStatistics'
    } finally {
        Complete-ProgressReporter -Activity 'Move statistics'
        $form.Cursor = $prev
    }
})

$btnMoveDetails.Add_Click({
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $row = Get-SingleSelectedRow -Grid $grdMoves -What 'move request'
        $lblStatus.Text = "Reading the report for $($row['DisplayName'])..."
        $txtMove.Text   = Get-MoveDetailText -Key "$($row['Key'])" -Diagnostics:$chkVerbose.Checked
        $lblStatus.Text = "Report for $($row['DisplayName'])."
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtMove.Text   = Format-ErrorDetail -ErrorRecord $_ -Context 'Move request details'
    } finally {
        $form.Cursor = $prev
    }
})

$btnMoveResume.Add_Click({
    try {
        $row  = Get-SingleSelectedRow -Grid $grdMoves -What 'move request'
        $note = 'Resumes a Suspended, AutoSuspended or Failed move. A move that failed on bad items will fail again unless the items are dealt with first.'
        if (Invoke-MoveWrite -Row $row -Cmdlet 'Resume-MoveRequest' -Action 'ResumeMove' -Verb 'Resume' -Note $note) {
            Update-MoveRow -Row $row
            $grdMoves.Refresh()
            $lblStatus.Text = "Resumed $($row['DisplayName']). Now $($row['Status']) $($row['StatusDetail'])."
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Resume move request' }
})

$btnMoveSuspend.Add_Click({
    try {
        $row  = Get-SingleSelectedRow -Grid $grdMoves -What 'move request'
        $note = 'Pauses the move where it is. Not possible once the move has reached CompletionInProgress.'
        if (Invoke-MoveWrite -Row $row -Cmdlet 'Suspend-MoveRequest' -Action 'SuspendMove' -Verb 'Suspend' -Note $note) {
            Update-MoveRow -Row $row
            $grdMoves.Refresh()
            $lblStatus.Text = "Suspended $($row['DisplayName']). Now $($row['Status']) $($row['StatusDetail'])."
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Suspend move request' }
})

$btnMoveRemove.Add_Click({
    try {
        $row = Get-SingleSelectedRow -Grid $grdMoves -What 'move request'
        # THE REFUSAL. Remove-MoveRequest on anything not finished CANCELS the
        # move; on a finished one it only clears the record. Read live, not
        # from the grid, so a stale row cannot talk us into a cancel.
        Update-MoveRow -Row $row
        $grdMoves.Refresh()
        $st = "$($row['Status'])"
        if ($st -notlike 'Completed*') {
            $null = Show-ToolDialog(
                ("Refusing to remove the move request for $($row['DisplayName']).`r`n`r`n" +
                 "Its Status is '$st $($row['StatusDetail'])', not Completed. Removing it now would CANCEL the move, not tidy it up.`r`n`r`n" +
                 'Wait for it to complete, or suspend it if you need it to stop.'),
                'Not completed', 'OK', 'Warning')
            $lblStatus.Text = "Not removed -- $($row['DisplayName']) is $st."
            return
        }
        $note = 'The move is finished. This only clears the completed request record; the mailbox is not touched.'
        if (Invoke-MoveWrite -Row $row -Cmdlet 'Remove-MoveRequest' -Action 'RemoveMove' -Verb 'Remove the completed' -Note $note) {
            $lblStatus.Text = "Removed the completed move request for $($row['DisplayName'])."
            $btnMoveRefresh.PerformClick()
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove move request' }
})

$btnMoveBadItems.Add_Click({
    try {
        $row = Get-SingleSelectedRow -Grid $grdMoves -What 'move request'
        Assert-ExoCmdlet -Name 'Set-MoveRequest'
        $declared = Get-CmdletParameterNames -Name 'Set-MoveRequest'

        if ($declared -contains 'BadItemLimit') {
            $answer = Show-InputDialog -Title 'Bad-item limit' -Default '10' `
                -Prompt "New bad-item limit for $($row['DisplayName']) (currently $($row['BadItems']) bad item(s) encountered). Whole number, or 'unlimited':"
            if (-not $answer) { return }
            $answer = $answer.Trim()
            $limit = $null
            if ($answer -match '^\d+$') { $limit = [int]$answer }
            elseif ($answer -ieq 'unlimited') { $limit = 'Unlimited' }
            else { throw "'$answer' is not a whole number or 'unlimited'." }
            $note = 'The move will skip up to this many corrupt or missing items rather than fail. Resume the move afterwards if it is Failed.'
            if (Invoke-MoveWrite -Row $row -Cmdlet 'Set-MoveRequest' -Action 'SetBadItemLimit' -Verb 'Set the bad-item limit on' `
                    -Extra @{ BadItemLimit = $limit } -Note $note) {
                Update-MoveRow -Row $row
                $grdMoves.Refresh()
                $lblStatus.Text = "Bad-item limit set to $limit for $($row['DisplayName'])."
            }
            return
        }

        # Exchange Online retired -BadItemLimit in favour of the Data
        # Consistency Score. On that build the way to unstick a move that is
        # holding on skipped items is to APPROVE them, which is what this
        # falls through to -- clearly labelled, because it is not what the
        # button says and the operator must know which one just ran.
        $note = 'This Exchange Online build does not declare -BadItemLimit (retired in favour of the Data Consistency Score). ' +
                'The cloud equivalent is to approve the skipped items found so far, so the move can complete without them. ' +
                'Those items are LEFT BEHIND. Review Details first.'
        if (Invoke-MoveWrite -Row $row -Cmdlet 'Set-MoveRequest' -Action 'ApproveSkippedItems' -Verb 'Approve skipped items on' `
                -Extra @{ SkippedItemApprovalTime = (Get-Date).ToUniversalTime() } -Note $note) {
            Update-MoveRow -Row $row
            $grdMoves.Refresh()
            $lblStatus.Text = "Skipped items approved for $($row['DisplayName']). Resume the move if it is not already running."
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Set bad-item limit' }
})
#endregion

#region Handlers -- shared ---------------------------------------------------
$btnSave.Add_Click({
    try {
        if ($tabs.SelectedTab -eq $pgBatches) {
            $csv = Export-GridCsv -Grid $grdBatches -Prefix 'MigrationStatus_Batches'
            $what = 'batches'
        } else {
            $csv = Export-GridCsv -Grid $grdMoves -Prefix 'MigrationStatus_Moves'
            $what = 'move requests'
        }
        $files.Refresh()
        $files.SelectPath($csv)
        $lblStatus.Text = "Saved $what to $csv"
        Write-CompletedLog -Task "Migration status snapshot ($what)" -Result 'Success'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Save CSV' }
})

# Launched against a mailbox: go straight to its move request.
$form.Add_Shown({
    try {
        if ("$SmtpAddress".Trim()) {
            $tabs.SelectedTab = $pgMoves
            $btnMoveRefresh.PerformClick()
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Initial refresh' }
})
#endregion

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Migration Status' | Out-Null

[void]$form.ShowDialog()
