#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    InboxRules.ps1
    Inbox rules -- list, flag the risky ones, enable / disable / delete in bulk.

    Author   : Glenn Armbruster
    Revised  : 2026-09-09  (risk flags, tick-many, detail pane, hidden rules, export)
#>
# TOOLKIT-HELP-RUNS: reads the mailbox's inbox rules. Enable, Disable and Delete change them and ask first.
# TOOLKIT-HELP-BEGIN
# View, disable, enable and delete the inbox rules on someone's mailbox.
#
# The usual reasons to look: mail vanishing into a folder nobody remembers
# creating, mail being forwarded somewhere it should not be, or a rule left
# behind by a compromise. The RISK column does that first read for you:
#
#   EXTERNAL FORWARD   forwards, redirects or forwards-as-attachment to an
#                      address outside the accepted domains. The classic
#                      compromise rule. Red.
#   DELETES            deletes the message, or moves it to Deleted Items,
#                      Junk, RSS or Conversation History -- the folders
#                      nobody looks in. Red.
#   HIDES              marks as read and stops processing, with no visible
#                      effect -- the quiet half of a compromise. Amber.
#   forward (internal) forwards to somebody in the company. Plain.
#
# Tick the rules you want and press Enable / Disable / Delete. Delete cannot
# be undone; Disable can, so prefer it while you are still working out
# whether a rule is wanted. Click a row to read its full description in the
# pane below.
#
# 'Include hidden' also lists the rules Outlook and other clients create for
# themselves -- they are normally noise, but a compromise can hide there too.
#
# Export writes the grid to the reports folder as <mailbox>_InboxRules.csv.
#
# Exchange Online only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Data ------------------------------------------------------------------
function Get-RuleList {
    # Learn 2026-09-09 (Get-InboxRule): -Mailbox; -IncludeHidden lists the
    # client-created rules that are hidden by default.
    param([string]$Smtp, [bool]$IncludeHidden)
    $p = @{ Mailbox = $Smtp; ErrorAction = 'Stop' }
    if ($IncludeHidden) { $p.IncludeHidden = $true }
    Write-ToolStep "Reading inbox rules of $Smtp ..."
    return @(Invoke-EXOCommand -Name 'Get-InboxRule' -Parameters $p)
}

function Set-RuleState {
    # Learn 2026-09-09 (Enable-/Disable-/Remove-InboxRule): -Identity takes the
    # rule identity ("mailbox\RuleId"); -Confirm:$false suppresses the prompt.
    param([string]$Identity, [ValidateSet('Enable', 'Disable', 'Remove')][string]$Verb)
    $p = @{ Identity = $Identity; Confirm = $false; ErrorAction = 'Stop' }
    # Literal cmdlet names, not "$Verb-InboxRule": the capability scanner and
    # the audit-coverage test read the source for the names.
    switch ($Verb) {
        'Enable'  { Invoke-EXOCommand -Name 'Enable-InboxRule'  -Parameters $p | Out-Null }
        'Disable' { Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters $p | Out-Null }
        'Remove'  { Invoke-EXOCommand -Name 'Remove-InboxRule'  -Parameters $p | Out-Null }
    }
}

function Get-RuleRecipients {
    # ForwardTo and friends come back as "Name [SMTP:addr]" strings or as
    # recipient objects; either way, the addresses are what matter.
    param([object]$Value)
    $out = @()
    foreach ($v in @($Value)) {
        $s = "$v"
        $m = [regex]::Matches($s, '[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}')
        if ($m.Count) { foreach ($x in $m) { $out += $x.Value.ToLower() } } elseif ($s) { $out += $s }
    }
    return @($out | Select-Object -Unique)
}

function Get-RuleAssessment {
    <#
        One line each for Actions, Conditions and Risk, from the rule's own
        properties rather than its Description -- the Description is prose and
        localised, the properties are not.
    #>
    param([object]$Rule, [string[]]$AcceptedDomains)

    $acts = @(); $risk = @(); $level = 0     # 0 plain, 1 amber, 2 red
    $fwd = @(); foreach ($p in 'ForwardTo', 'ForwardAsAttachmentTo', 'RedirectTo') { if ($Rule.$p) { $fwd += Get-RuleRecipients $Rule.$p } }
    if ($fwd.Count) {
        $ext = @($fwd | Where-Object { $d = Get-SmtpDomain $_; $d -and $AcceptedDomains -notcontains $d })
        $acts += "forward to $($fwd -join ', ')"
        if ($ext.Count) { $risk += "EXTERNAL FORWARD: $($ext -join ', ')"; $level = 2 }
    }
    if ($Rule.DeleteMessage) { $acts += 'delete'; $risk += 'DELETES the message'; $level = 2 }
    if ($Rule.MoveToFolder) {
        $f = "$($Rule.MoveToFolder)"
        $acts += "move to $f"
        if ($f -match 'Deleted Items|Junk|RSS|Conversation History|Archive') { $risk += "moves to $f"; $level = [Math]::Max($level, 2) }
    }
    if ($Rule.CopyToFolder)          { $acts += "copy to $($Rule.CopyToFolder)" }
    if ($Rule.MarkAsRead)            { $acts += 'mark read' }
    if ($Rule.StopProcessingRules)   { $acts += 'stop processing' }
    if ($Rule.MarkImportance)        { $acts += "importance $($Rule.MarkImportance)" }
    if ($Rule.ApplyCategory)         { $acts += "category $(@($Rule.ApplyCategory) -join '/')" }
    if ($Rule.SendTextMessageNotificationTo) { $acts += 'SMS notify' }
    if ($Rule.MarkAsRead -and $Rule.StopProcessingRules -and $level -eq 0 -and -not $Rule.MoveToFolder -and -not $fwd.Count) {
        $risk += 'HIDES: marks read and stops, no visible effect'; $level = 1
    }
    if (-not "$($Rule.Name)".Trim()) { $risk += 'no name'; $level = [Math]::Max($level, 1) }

    $conds = @()
    foreach ($p in 'From', 'SentTo', 'SubjectContainsWords', 'SubjectOrBodyContainsWords', 'BodyContainsWords', 'FromAddressContainsWords', 'HeaderContainsWords', 'RecipientAddressContainsWords', 'HasAttachment', 'MyNameInToBox', 'MyNameInCcBox', 'SentOnlyToMe', 'WithImportance', 'MessageTypeMatches', 'HasClassification', 'WithinSizeRangeMaximum', 'ReceivedAfterDate', 'ReceivedBeforeDate') {
        $v = $Rule.$p
        if ($null -eq $v -or "$v" -eq '' -or "$v" -eq 'False') { continue }
        if ($v -is [bool]) { $conds += $p } else { $conds += "$p=$(@($v | ForEach-Object { "$_" }) -join '/')" }
    }
    if ($conds.Count -eq 0) { $conds += '(all mail)' }

    return [PSCustomObject]@{
        Actions    = ($acts -join '; ')
        Conditions = ($conds -join '; ')
        Risk       = ($risk -join '; ')
        Level      = $level
    }
}
#endregion

#region Form ------------------------------------------------------------------
function Show-InboxRulesForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Inbox Rules'
    $form.ClientSize = New-Object System.Drawing.Size(1000, 620)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, $y); $lblStatus.Size = New-Object System.Drawing.Size(700, 18)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray; $lblStatus.Text = 'Enter a mailbox and press Look Up.'
    $form.Controls.Add($lblStatus)
    $chkHidden = New-Object System.Windows.Forms.CheckBox
    $chkHidden.Text = 'Include hidden'; $chkHidden.Location = New-Object System.Drawing.Point(740, ($y - 2)); $chkHidden.AutoSize = $true
    $form.Controls.Add($chkHidden)
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = 'Refresh'; $btnRefresh.Location = New-Object System.Drawing.Point(888, ($y - 4)); $btnRefresh.Size = New-Object System.Drawing.Size(100, 26)
    $form.Controls.Add($btnRefresh)
    $y += 26

    $pb = Add-ReportProgressBar -Form $form -X 12 -Y $y -Width 976 -Height 12
    $y += 20

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, $y); $grid.Size = New-Object System.Drawing.Size(976, 300)
    $grid.ReadOnly = $false; $grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false; $grid.AllowUserToResizeRows = $false
    $grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = 'Fill'; $grid.RowHeadersVisible = $false
    $grid.Font = New-Object System.Drawing.Font('Consolas', 9)
    $colPick = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colPick.HeaderText = ''; $colPick.Name = 'Pick'; $colPick.FillWeight = 3
    [void]$grid.Columns.Add($colPick)
    foreach ($c in @(@('Name', 18), @('On', 4), @('Pri', 4), @('Risk', 22), @('Actions', 24), @('Conditions', 25))) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c[0]; $col.HeaderText = $c[0]; $col.FillWeight = $c[1]; $col.ReadOnly = $true
        $col.SortMode = 'NotSortable'
        [void]$grid.Columns.Add($col)
    }
    $colId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colId.Name = 'Identity'; $colId.Visible = $false
    [void]$grid.Columns.Add($colId)
    $form.Controls.Add($grid)
    $y += 308

    $lblDet = New-Object System.Windows.Forms.Label
    $lblDet.Text = 'Rule detail (click a row):'; $lblDet.Location = New-Object System.Drawing.Point(12, $y); $lblDet.AutoSize = $true
    $form.Controls.Add($lblDet)
    $y += 20
    $txtDet = New-Object System.Windows.Forms.TextBox
    $txtDet.Multiline = $true; $txtDet.ReadOnly = $true; $txtDet.ScrollBars = 'Vertical'; $txtDet.WordWrap = $true
    $txtDet.Location = New-Object System.Drawing.Point(12, $y); $txtDet.Size = New-Object System.Drawing.Size(976, 170)
    $txtDet.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtDet)
    $y += 180

    $btnAll = New-Object System.Windows.Forms.Button;  $btnAll.Text = 'Tick all';       $btnAll.Location = New-Object System.Drawing.Point(12, $y);  $btnAll.Size = New-Object System.Drawing.Size(90, 28);  $form.Controls.Add($btnAll)
    $btnNone = New-Object System.Windows.Forms.Button; $btnNone.Text = 'Untick all';    $btnNone.Location = New-Object System.Drawing.Point(108, $y); $btnNone.Size = New-Object System.Drawing.Size(90, 28); $form.Controls.Add($btnNone)
    $btnRisky = New-Object System.Windows.Forms.Button; $btnRisky.Text = 'Tick risky';  $btnRisky.Location = New-Object System.Drawing.Point(204, $y); $btnRisky.Size = New-Object System.Drawing.Size(90, 28); $form.Controls.Add($btnRisky)
    $btnEnable = New-Object System.Windows.Forms.Button;  $btnEnable.Text = 'Enable ticked';   $btnEnable.Location = New-Object System.Drawing.Point(330, $y);  $btnEnable.Size = New-Object System.Drawing.Size(120, 28);  $btnEnable.Enabled = $false;  $form.Controls.Add($btnEnable)
    $btnDisable = New-Object System.Windows.Forms.Button; $btnDisable.Text = 'Disable ticked'; $btnDisable.Location = New-Object System.Drawing.Point(456, $y); $btnDisable.Size = New-Object System.Drawing.Size(120, 28); $btnDisable.Enabled = $false; $form.Controls.Add($btnDisable)
    $btnDelete = New-Object System.Windows.Forms.Button;  $btnDelete.Text = 'Delete ticked';   $btnDelete.Location = New-Object System.Drawing.Point(582, $y);  $btnDelete.Size = New-Object System.Drawing.Size(120, 28);  $btnDelete.Enabled = $false;  $form.Controls.Add($btnDelete)
    $btnExport = New-Object System.Windows.Forms.Button;  $btnExport.Text = 'Export CSV';      $btnExport.Location = New-Object System.Drawing.Point(744, $y);  $btnExport.Size = New-Object System.Drawing.Size(110, 28);  $btnExport.Enabled = $false;  $form.Controls.Add($btnExport)
    $btnClose = New-Object System.Windows.Forms.Button;   $btnClose.Text = 'Close';            $btnClose.Location = New-Object System.Drawing.Point(888, $y);   $btnClose.Size = New-Object System.Drawing.Size(100, 28);   $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy'  -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Rules' -NotePropertyValue @{}    -Force   # Identity -> @{ Rule; Assess }

    $grid.Add_CurrentCellDirtyStateChanged({ if ($grid.IsCurrentCellDirty) { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) | Out-Null } })
    $btnAll.Add_Click({  foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = $true };  $grid.EndEdit() | Out-Null })
    $btnNone.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = $false }; $grid.EndEdit() | Out-Null })
    $btnRisky.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = ($form.Rules["$($r.Cells['Identity'].Value)"].Assess.Level -ge 1) }; $grid.EndEdit() | Out-Null })

    $load = {
        $grid.Rows.Clear(); $txtDet.Text = ''
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false; $btnExport.Enabled = $false
        $form.Rules = @{}
        if (-not "$script:Target".Trim()) { $form.Text = 'Inbox Rules'; $lblStatus.Text = 'Enter a mailbox and press Look Up.'; return }
        $form.Text = "Inbox Rules: $script:Target"
        $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Reading'
        & $p 'Reading rules...' -1
        $domains = @(Get-AcceptedDomainList)
        $rules = Get-RuleList -Smtp $script:Target -IncludeHidden $chkHidden.Checked
        $nRed = 0; $nAmber = 0
        foreach ($r in ($rules | Sort-Object Priority)) {
            $a = Get-RuleAssessment -Rule $r -AcceptedDomains $domains
            $id = "$($r.Identity)"
            $form.Rules[$id] = @{ Rule = $r; Assess = $a }
            $n = $grid.Rows.Add($false, "$($r.Name)", $(if ($r.Enabled) { 'on' } else { 'off' }), "$($r.Priority)", $a.Risk, $a.Actions, $a.Conditions, $id)
            if ($a.Level -ge 2) { $grid.Rows[$n].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 205, 200); $nRed++ }
            elseif ($a.Level -eq 1) { $grid.Rows[$n].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 240, 200); $nAmber++ }
            if (-not $r.Enabled) { $grid.Rows[$n].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray }
        }
        & $p "$($rules.Count) rule(s)$(if ($nRed -or $nAmber) { " -- $nRed red, $nAmber amber" } else { ', none flagged' }). Tick rows to act on them." 100
        $live = ($rules.Count -gt 0)
        $btnEnable.Enabled = $live; $btnDisable.Enabled = $live; $btnDelete.Enabled = $live; $btnExport.Enabled = $live
    }
    $doLoad = {
        if ($form.Busy) { return }
        $form.Busy = $true
        try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; & $load }
        catch { $lblStatus.Text = 'Load failed -- see the error dialog.'; Show-HandlerError -ErrorRecord $_ -Context 'Get-InboxRule' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    }
    $btnRefresh.Add_Click($doLoad)
    $chkHidden.Add_CheckedChanged({ if ("$script:Target".Trim()) { & $doLoad } })

    $grid.Add_SelectionChanged({
        if ($grid.SelectedRows.Count -ne 1) { return }
        $e = $form.Rules["$($grid.SelectedRows[0].Cells['Identity'].Value)"]
        if (-not $e) { return }
        $r = $e.Rule; $a = $e.Assess
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("$($r.Name)   [priority $($r.Priority); $(if ($r.Enabled) { 'ENABLED' } else { 'disabled' })$(if ($r.InError) { '; IN ERROR' })]")
        if ($a.Risk) { [void]$sb.AppendLine("RISK: $($a.Risk)") }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Conditions : $($a.Conditions)")
        [void]$sb.AppendLine("Actions    : $($a.Actions)")
        if ($r.ExceptIfFrom -or $r.ExceptIfSubjectContainsWords -or $r.ExceptIfSentTo) { [void]$sb.AppendLine("Exceptions : from $(@($r.ExceptIfFrom) -join '/'); subject $(@($r.ExceptIfSubjectContainsWords) -join '/'); to $(@($r.ExceptIfSentTo) -join '/')") }
        [void]$sb.AppendLine('')
        # Parenthesised: inside a method call the comma in -replace's argument
        # list reads as a SECOND method argument, and AppendLine has no
        # two-argument overload.
        $desc = ("$($r.Description)" -replace '\r?\n\s*\r?\n', "`r`n")
        [void]$sb.AppendLine($desc)
        $txtDet.Text = $sb.ToString()
    })

    $act = {
        param([string]$Verb, [string]$Word)
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $grid.EndEdit() | Out-Null
            $picked = @($grid.Rows | Where-Object { $_.Cells['Pick'].Value -eq $true })
            if ($picked.Count -eq 0) { $lblStatus.Text = 'Tick one or more rules first.'; return }
            $names = @($picked | ForEach-Object { "$($_.Cells['Name'].Value)" })
            $warn = $(if ($Verb -eq 'Remove') { "`r`n`r`nDelete cannot be undone. Disable is reversible." } else { '' })
            $a = Show-ToolDialog(("$Word $($picked.Count) rule(s) on $script:Target" + "?`r`n`r`n  " + ($names -join "`r`n  ") + $warn),
                'Confirm', 'OKCancel', $(if ($Verb -eq 'Remove') { 'Warning' } else { 'Question' }))
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity $Word
            $ok = 0; $failed = 0; $i = 0; $errs = @()
            foreach ($row in $picked) {
                $i++
                $nm = "$($row.Cells['Name'].Value)"
                & $p "$Word $i of $($picked.Count): $nm" ([int](100 * $i / $picked.Count))
                try { Set-RuleState -Identity "$($row.Cells['Identity'].Value)" -Verb $Verb; $ok++ }
                catch { $failed++; $errs += "$nm -- $($_.Exception.Message)"; Write-FailedLog -Action "${Verb}InboxRule" -Target (Get-LogTarget $script:Target) -ErrorRecord $_ }
            }
            Write-CompletedLog -Action "${Verb}InboxRule" -Target (Get-LogTarget $script:Target) -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) `
                -Detail "$ok $($Word.ToLower())d, $failed failed: $($names -join ', ')"
            if ($failed) { Show-ToolDialog(("$failed of $($picked.Count) failed:`r`n`r`n  " + ($errs -join "`r`n  ")), 'Partial', 'OK', 'Warning') | Out-Null }
            & $load
            $lblStatus.Text = "$ok $($Word.ToLower())d, $failed failed."
        } catch { $lblStatus.Text = "$Word failed -- see the error dialog."; Show-HandlerError -ErrorRecord $_ -Context "$Verb-InboxRule" }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    }
    $btnEnable.Add_Click({  & $act 'Enable'  'Enable' })
    $btnDisable.Add_Click({ & $act 'Disable' 'Disable' })
    $btnDelete.Add_Click({  & $act 'Remove'  'Delete' })

    $btnExport.Add_Click({
        try {
            $local = ("$script:Target" -split '@')[0]
            $path = Join-Path (Get-ReportPath) "${local}_InboxRules.csv"
            $rows = foreach ($r in $grid.Rows) {
                [PSCustomObject]@{ Mailbox = $script:Target; Name = $r.Cells['Name'].Value; Enabled = $r.Cells['On'].Value; Priority = $r.Cells['Pri'].Value
                    Risk = $r.Cells['Risk'].Value; Actions = $r.Cells['Actions'].Value; Conditions = $r.Cells['Conditions'].Value
                    Description = "$($form.Rules["$($r.Cells['Identity'].Value)"].Rule.Description)" -replace '\s+', ' '
                    ExportedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }
            }
            $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
            Write-CompletedLog -Action 'ExportInboxRules' -Target (Get-LogTarget $script:Target) -Result 'Success' -Detail "$(@($rows).Count) rule(s) -> $path"
            $lblStatus.Text = "Exported $(@($rows).Count) rule(s) to $path"
        } catch { Show-HandlerError -ErrorRecord $_ -Context 'Export CSV' }
    })

    $bar = Add-ToolTargetBar -Form $form -InitialValue $script:Target
    $bar.Button.Add_Click({ $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text; & $doLoad })
    $form.Add_Shown({ if ("$script:Target".Trim()) { & $doLoad } })

    Set-ConsoleTheme -Form $form
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Inbox Rules' | Out-Null
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    [void]$form.ShowDialog()
}
#endregion

Show-InboxRulesForm
