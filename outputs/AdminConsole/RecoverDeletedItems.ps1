#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    RecoverDeletedItems.ps1
    Find and restore items a user has deleted, including ones already purged from Deleted Items. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-08-27
#>
# TOOLKIT-HELP-RUNS: searches one mailbox's deleted items. Restore puts selected items back where they came from. Nothing is deleted.
# TOOLKIT-HELP-BEGIN
# "I deleted an email and now it's gone."
#
# Searches three places, in order of how far gone the item is:
#
#   Deleted Items       still in the folder -- the user could get it back themselves
#   Recoverable Items   emptied from Deleted Items, or shift-deleted. Outlook's
#                       "Recover Deleted Items" sees this too, but not everybody
#                       knows it exists
#   Purged Items        deleted from Recoverable Items as well. Only a hold or a
#                       retention policy keeps anything here, and only an admin
#                       can see it
#
# Narrow by subject, by date, by item type. Then tick the ones to bring back and
# press Restore. Each item goes BACK TO THE FOLDER IT WAS DELETED FROM -- Exchange
# remembers, and that is what LastParentPath in the grid shows you.
#
# RESTORE IS PER ITEM, NEVER PER FILTER. Restore-RecoverableItems will happily
# restore everything that matches a subject, and "everything matching 'Invoice'"
# on a finance mailbox is not a thing to do by accident. This tool only ever
# restores the exact items you ticked, by their EntryID.
#
# The date filter is on LAST MODIFIED, not on received. For a deleted item that
# is usually the deletion time, which is what people actually remember.
#
# Exchange Online only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

#region Data ------------------------------------------------------------------

function Find-DeletedItem {
    <#
        Learn 2026-08-27 (Get-RecoverableItems): -Identity, -SourceFolder
        (DeletedItems | RecoverableItems | PurgedItems), -SubjectContains,
        -FilterItemType, -FilterStartTime, -FilterEndTime, -ResultSize.
        Returns EntryID, Subject, ItemClass, LastModifiedTime, LastParentPath.

        ResultSize explicit (C3). A mailbox with years of deleted mail and no
        filter is exactly the call that silently stops at the default.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$Source,
        [string]$Subject,
        [string]$ItemType,
        # NULLABLE, NOT [datetime]. This is the bug the first version shipped
        # with: [datetime] is a value type and cannot hold $null, so passing an
        # unchecked date box through a [datetime] parameter fails at BIND time
        # with "Cannot convert null to type System.DateTime" -- before a single
        # line of the function runs, on every search with a date box unticked,
        # which is most of them. [Nullable[datetime]] accepts $null and means
        # exactly what the checkbox means: "no filter".
        [Nullable[datetime]]$From,
        [Nullable[datetime]]$To,
        [int]$Max = 5000
    )

    $p = @{
        Identity     = $Smtp
        SourceFolder = $Source
        ResultSize   = $Max
        ErrorAction  = 'Stop'
    }
    if ("$Subject".Trim())  { $p.SubjectContains = "$Subject".Trim() }
    if ("$ItemType".Trim()) { $p.FilterItemType  = "$ItemType".Trim() }
    if ($null -ne $From)    { $p.FilterStartTime = $From.Value }
    if ($null -ne $To)      { $p.FilterEndTime   = $To.Value }

    return @(Invoke-EXOCommand -Name 'Get-RecoverableItems' -Parameters $p)
}

function Restore-DeletedItem {
    <#
        Learn 2026-08-27 (Restore-RecoverableItems): -Identity <mailbox>,
        -EntryID <string>, -SourceFolder. EntryID is unique within the mailbox
        and is the ONLY selector this function uses -- see the help block on
        why restoring by filter is not offered.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$EntryId,
        [Parameter(Mandatory)][string]$Source
    )
    return Invoke-EXOCommand -Name 'Restore-RecoverableItems' -Parameters @{
        Identity     = $Smtp
        EntryID      = $EntryId
        SourceFolder = $Source
        ErrorAction  = 'Stop'
    }
}
#endregion

#region UI --------------------------------------------------------------------

function Show-RecoverForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Recover Deleted Items'
    $form.ClientSize = New-Object System.Drawing.Size(1000, 680)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblMbx = New-Object System.Windows.Forms.Label
    $lblMbx.Text = 'Mailbox:'
    $lblMbx.Location = New-Object System.Drawing.Point(12, ($y + 3))
    $lblMbx.Size = New-Object System.Drawing.Size(70, 18)
    $form.Controls.Add($lblMbx)

    $txtMbx = New-Object System.Windows.Forms.TextBox
    $txtMbx.Location = New-Object System.Drawing.Point(84, $y)
    $txtMbx.Size = New-Object System.Drawing.Size(320, 24)
    $txtMbx.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtMbx.Text = "$Seed"
    $form.Controls.Add($txtMbx)

    $lblSrc = New-Object System.Windows.Forms.Label
    $lblSrc.Text = 'Look in:'
    $lblSrc.Location = New-Object System.Drawing.Point(420, ($y + 3))
    $lblSrc.Size = New-Object System.Drawing.Size(60, 18)
    $form.Controls.Add($lblSrc)

    $cboSrc = New-Object System.Windows.Forms.ComboBox
    $cboSrc.DropDownStyle = 'DropDownList'
    @('RecoverableItems', 'DeletedItems', 'PurgedItems') | ForEach-Object { [void]$cboSrc.Items.Add($_) }
    $cboSrc.SelectedIndex = 0
    $cboSrc.Location = New-Object System.Drawing.Point(482, $y)
    $cboSrc.Size = New-Object System.Drawing.Size(160, 24)
    $form.Controls.Add($cboSrc)

    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Text = 'Type:'
    $lblType.Location = New-Object System.Drawing.Point(658, ($y + 3))
    $lblType.Size = New-Object System.Drawing.Size(44, 18)
    $form.Controls.Add($lblType)

    $cboType = New-Object System.Windows.Forms.ComboBox
    $cboType.DropDownStyle = 'DropDownList'
    # The blank entry means "no filter". ItemClass values are the MAPI ones.
    @('', 'IPM.Note', 'IPM.Appointment', 'IPM.Contact', 'IPM.Task', 'IPM.StickyNote') |
        ForEach-Object { [void]$cboType.Items.Add($_) }
    $cboType.SelectedIndex = 0
    $cboType.Location = New-Object System.Drawing.Point(704, $y)
    $cboType.Size = New-Object System.Drawing.Size(150, 24)
    $form.Controls.Add($cboType)
    $y += 34

    $lblSubj = New-Object System.Windows.Forms.Label
    $lblSubj.Text = 'Subject contains:'
    $lblSubj.Location = New-Object System.Drawing.Point(12, ($y + 3))
    $lblSubj.Size = New-Object System.Drawing.Size(110, 18)
    $form.Controls.Add($lblSubj)

    $txtSubj = New-Object System.Windows.Forms.TextBox
    $txtSubj.Location = New-Object System.Drawing.Point(124, $y)
    $txtSubj.Size = New-Object System.Drawing.Size(280, 24)
    $form.Controls.Add($txtSubj)

    $chkFrom = New-Object System.Windows.Forms.CheckBox
    $chkFrom.Text = 'Modified from'
    $chkFrom.Location = New-Object System.Drawing.Point(420, ($y + 2))
    $chkFrom.Size = New-Object System.Drawing.Size(110, 22)
    $form.Controls.Add($chkFrom)

    $dtFrom = New-Object System.Windows.Forms.DateTimePicker
    $dtFrom.Format = 'Short'
    $dtFrom.Value = (Get-Date).AddDays(-30)
    $dtFrom.Location = New-Object System.Drawing.Point(532, $y)
    $dtFrom.Size = New-Object System.Drawing.Size(110, 24)
    $form.Controls.Add($dtFrom)

    $chkTo = New-Object System.Windows.Forms.CheckBox
    $chkTo.Text = 'to'
    $chkTo.Location = New-Object System.Drawing.Point(658, ($y + 2))
    $chkTo.Size = New-Object System.Drawing.Size(40, 22)
    $form.Controls.Add($chkTo)

    $dtTo = New-Object System.Windows.Forms.DateTimePicker
    $dtTo.Format = 'Short'
    $dtTo.Location = New-Object System.Drawing.Point(704, $y)
    $dtTo.Size = New-Object System.Drawing.Size(110, 24)
    $form.Controls.Add($dtTo)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = 'Search'
    $btnSearch.Location = New-Object System.Drawing.Point(858, ($y - 2))
    $btnSearch.Size = New-Object System.Drawing.Size(130, 28)
    $form.Controls.Add($btnSearch)
    $y += 34

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, $y)
    $lblStatus.Size = New-Object System.Drawing.Size(976, 18)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text = 'Enter a mailbox and press Search. Recoverable Items is where emptied and shift-deleted mail goes.'
    $form.Controls.Add($lblStatus)
    $y += 22

    $pb = Add-ReportProgressBar -Form $form -X 12 -Y $y -Width 976 -Height 14
    $y += 24

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, $y)
    $grid.Size = New-Object System.Drawing.Size(976, 480)
    $grid.ReadOnly = $false
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $true
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowHeadersVisible = $false
    $grid.Font = New-Object System.Drawing.Font('Consolas', 9)

    $colPick = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colPick.HeaderText = ''; $colPick.Name = 'Pick'; $colPick.FillWeight = 4
    [void]$grid.Columns.Add($colPick)
    foreach ($c in @(@('Subject', 40), @('ItemClass', 12), @('LastModified', 16), @('LastParentPath', 28))) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c[0]; $col.HeaderText = $c[0]; $col.FillWeight = $c[1]; $col.ReadOnly = $true
        [void]$grid.Columns.Add($col)
    }
    $colId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colId.Name = 'EntryID'; $colId.Visible = $false
    [void]$grid.Columns.Add($colId)
    $form.Controls.Add($grid)
    $y += 490

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = 'Tick all'
    $btnAll.Location = New-Object System.Drawing.Point(12, $y)
    $btnAll.Size = New-Object System.Drawing.Size(100, 28)
    $form.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = 'Untick all'
    $btnNone.Location = New-Object System.Drawing.Point(120, $y)
    $btnNone.Size = New-Object System.Drawing.Size(100, 28)
    $form.Controls.Add($btnNone)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Text = 'Restore ticked'
    $btnRestore.Location = New-Object System.Drawing.Point(240, $y)
    $btnRestore.Size = New-Object System.Drawing.Size(150, 28)
    $btnRestore.Enabled = $false
    $form.Controls.Add($btnRestore)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Location = New-Object System.Drawing.Point(858, $y)
    $btnClose.Size = New-Object System.Drawing.Size(130, 28)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy'   -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Smtp'   -NotePropertyValue ''     -Force
    $form | Add-Member -NotePropertyName 'Source' -NotePropertyValue ''     -Force

    $btnAll.Add_Click({  foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = $true } })
    $btnNone.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = $false } })

    $btnSearch.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $grid.Rows.Clear()
            $btnRestore.Enabled = $false

            $s = Get-TargetSmtp -Text $txtMbx.Text
            if (-not $s) { $lblStatus.Text = 'Enter a mailbox address.'; return }
            $form.Smtp = $s
            $form.Source = "$($cboSrc.SelectedItem)"

            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Searching'
            & $p "Searching $($form.Source) on $s..." -1

            $from = $null; $to = $null
            if ($chkFrom.Checked) { $from = $dtFrom.Value.Date }
            if ($chkTo.Checked)   { $to   = $dtTo.Value.Date.AddDays(1).AddSeconds(-1) }

            $items = Find-DeletedItem -Smtp $s -Source $form.Source `
                        -Subject $txtSubj.Text -ItemType "$($cboType.SelectedItem)" `
                        -From $from -To $to

            $grid.SuspendLayout()
            foreach ($it in $items) {
                $lm = ''
                try { $lm = ([datetime]$it.LastModifiedTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { $lm = "$($it.LastModifiedTime)" }
                [void]$grid.Rows.Add($false, "$($it.Subject)", "$($it.ItemClass)", $lm, "$($it.LastParentPath)", "$($it.EntryID)")
            }
            $grid.ResumeLayout()

            & $p "$($items.Count) item(s) in $($form.Source). Tick the ones to restore." 100
            $btnRestore.Enabled = ($items.Count -gt 0)
            if ($items.Count -ge 5000) {
                $lblStatus.Text = "5,000 shown -- the cap. Narrow the search; there may be more."
            }
        } catch {
            $lblStatus.Text = 'Search failed -- see the error dialog.'
            Show-HandlerError -ErrorRecord $_ -Context 'Search (Get-RecoverableItems)'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    $btnRestore.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $picked = @($grid.Rows | Where-Object { $_.Cells['Pick'].Value -eq $true })
            if ($picked.Count -eq 0) { $lblStatus.Text = 'Nothing ticked.'; return }

            $a = Show-ToolDialog(
                ("Restore $($picked.Count) item(s) to $($form.Smtp)?" + "`r`n`r`n" +
                 "Each goes back to the folder it was deleted from." + "`r`n" +
                 "From: $($form.Source)"),
                'Confirm', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnRestore.Enabled = $false; $btnSearch.Enabled = $false
            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Restoring'

            $ok = 0; $failed = 0; $i = 0; $errs = @()
            foreach ($r in $picked) {
                $i++
                $subj = "$($r.Cells['Subject'].Value)"
                & $p "Restoring $i of $($picked.Count): $subj" ([int](100 * $i / $picked.Count))
                # ONE ITEM PER CALL, EACH IN ITS OWN TRY. A batch that stops at
                # the first failure leaves the user with half their mail back and
                # no list of which half.
                try {
                    Restore-DeletedItem -Smtp $form.Smtp -EntryId "$($r.Cells['EntryID'].Value)" -Source $form.Source | Out-Null
                    $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200, 240, 200)
                    $r.Cells['Pick'].Value = $false
                    $ok++
                } catch {
                    $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 200, 200)
                    $failed++
                    $errs += "$subj -- $($_.Exception.Message)"
                }
            }

            $lblStatus.Text = "$ok restored, $failed failed." + $(if ($failed) { ' Failed rows are red.' } else { '' })
            $result = if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }
            Write-CompletedLog -Action 'RestoreDeletedItems' -Target (Get-LogTarget $form.Smtp) `
                -Result $result -Detail "$ok restored, $failed failed from $($form.Source)$(if ($errs) { '; ' + ($errs -join ' | ') })"
        } catch {
            $lblStatus.Text = 'Restore failed -- see the error dialog.'
            Show-HandlerError -ErrorRecord $_ -Context 'Restore (Restore-RecoverableItems)'
            Write-FailedLog -Action 'RestoreDeletedItems' -Target "$($form.Smtp)" -ErrorRecord $_
        } finally {
            $btnRestore.Enabled = $true; $btnSearch.Enabled = $true
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Recover Deleted Items' | Out-Null
    [void]$form.ShowDialog()
}
#endregion

Show-RecoverForm -Seed (Get-TargetSmtp -Text $SmtpAddress)
