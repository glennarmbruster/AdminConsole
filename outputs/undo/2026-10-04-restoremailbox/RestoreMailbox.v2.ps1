#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    RestoreMailbox.ps1
    Bring back a soft-deleted mailbox, or pull its contents into another one. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-10-04  (password on restore; shared stays shared)
#>
# TOOLKIT-HELP-RUNS: lists soft-deleted mailboxes. Both restore actions change things and ask first.
# TOOLKIT-HELP-BEGIN
# A deleted mailbox is not gone. Exchange Online keeps it for 30 DAYS as
# "soft-deleted", and this tool is the way back.
#
# Load with the filter blank and EVERY soft-deleted mailbox is listed; type
# part of a name or address to narrow it. Tick the rows you want. Restore in
# place takes every tick at once; Copy wants exactly one.
#
# TWO WAYS BACK, AND WHICH ONE DEPENDS ON WHETHER THE PERSON STILL EXISTS:
#
#   Restore in place   The mailbox comes back as itself, attached to its user.
#                      This needs the USER ACCOUNT to still exist -- or to be
#                      restorable from the Entra recycle bin. In a hybrid
#                      directory that means the on-premises AD account. If
#                      Disable-User deleted the AD account and a sync has run,
#                      the user is gone from Entra and this will fail. It says
#                      so rather than pretending.
#
#                      A user that is in the Entra RECYCLE BIN (deleted, not
#                      yet purged -- which is what a directory sync does
#                      when an account leaves its scope) comes back WITH the
#                      mailbox. Microsoft Learn says that restore needs a new
#                      password, so one is set: random, never shown, and the
#                      account stays sign-in blocked. A mailbox that was
#                      shared before and comes back as a user mailbox is
#                      converted back to shared.
#
#   Copy into another  The CONTENTS go into a mailbox you name -- a manager's,
#                      a shared mailbox, a fresh account for the returning
#                      person. The soft-deleted mailbox is left alone. This
#                      works whether or not the original user exists, and it
#                      is the answer to "we need what was in there, not the
#                      account".
#
# The copy is a RESTORE REQUEST: Exchange queues it and works through it in the
# background. Big mailboxes take hours. The Requests panel shows progress; you
# do not need to keep this window open.
#
# THE 30-DAY CLOCK IS THE ONE THAT MATTERS. The grid shows days remaining. A
# mailbox at 2 days is a today problem.
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

function Get-SoftDeletedMailboxList {
    <#
        Learn 2026-08-27 (Get-Mailbox): -SoftDeletedMailbox switch returns only
        soft-deleted mailboxes. -ResultSize Unlimited (C3): a tenant that has
        been offboarding for a month can have more than the default.
    #>
    return @(Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
        SoftDeletedMailbox = $true; ResultSize = 'Unlimited'; ErrorAction = 'Stop'
    })
}

function New-RestorePassword {
    # Random, never shown. The account is a mailbox holder, not a sign-in.
    $chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789!#$%&*+-=?@'
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $b = New-Object byte[] 24; $rng.GetBytes($b)
    $plain = -join ($b | ForEach-Object { $chars[$_ % $chars.Length] })
    return (ConvertTo-SecureString -String $plain -AsPlainText -Force)
}

function Restore-SoftDeletedInPlace {
    # Learn 2026-10-04 (Undo-SoftDeletedMailbox): -SoftDeletedObject <alias or
    # address>; -Password <SecureString>. "If the Microsoft account wasn't
    # deleted when the mailbox was deleted, you have to specify a new
    # Microsoft account and password" -- in practice the restore of a user
    # from the recycle bin wants the password too, so it is always passed.
    # Fails if the associated user was permanently deleted.
    param([Parameter(Mandatory)][string]$Smtp)
    return Invoke-EXOCommand -Name 'Undo-SoftDeletedMailbox' -Parameters @{
        SoftDeletedObject = $Smtp; Password = (New-RestorePassword); ErrorAction = 'Stop'
    }
}

function Start-ContentRestore {
    <#
        Learn 2026-08-27 (New-MailboxRestoreRequest): -SourceMailbox is best
        given as the soft-deleted mailbox's GUID; -TargetMailbox; -Name;
        -AllowLegacyDNMismatch; -TargetRootFolder.

        -AllowLegacyDNMismatch is ALWAYS passed. Source and target are
        different people, so their LegacyExchangeDNs never match, and without
        the switch the request is refused for exactly that reason.

        -TargetRootFolder puts everything under one named folder in the target
        rather than merging into their Inbox. Merging a departed colleague's
        mail into a manager's Inbox is the complaint that arrives an hour after
        the restore.
    #>
    param(
        [Parameter(Mandatory)][string]$SourceGuid,
        [Parameter(Mandatory)][string]$Target,
        [Parameter(Mandatory)][string]$RootFolder,
        [Parameter(Mandatory)][string]$Name
    )
    return Invoke-EXOCommand -Name 'New-MailboxRestoreRequest' -Parameters @{
        SourceMailbox         = $SourceGuid
        TargetMailbox         = $Target
        TargetRootFolder      = $RootFolder
        Name                  = $Name
        AllowLegacyDNMismatch = $true
        ErrorAction           = 'Stop'
    }
}

function Get-RestoreRequestList {
    # Learn 2026-08-27 (Get-MailboxRestoreRequest): returns Name, Status,
    # TargetMailbox, RequestQueue. Statistics adds PercentComplete.
    $reqs = @(Invoke-EXOCommand -Name 'Get-MailboxRestoreRequest' -Parameters @{ ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
    $out = @()
    foreach ($r in $reqs) {
        # Statistics is one round trip PER REQUEST, so it is only asked for the
        # ones still moving. A completed or failed request has a status that
        # says so and a percentage nobody needs; a tenant with months of old
        # requests would otherwise pay a call each on every Refresh.
        $pct = ''
        if ("$($r.Status)" -match 'InProgress|Queued|Suspended|AutoSuspended') {
            try {
                $s = Invoke-EXOCommand -Name 'Get-MailboxRestoreRequestStatistics' -Parameters @{ Identity = $r.Identity; ErrorAction = 'Stop' }
                $pct = "$($s.PercentComplete)%"
            } catch { }
        }
        $out += [PSCustomObject]@{ Name = "$($r.Name)"; Target = "$($r.TargetMailbox)"; Status = "$($r.Status)"; Percent = $pct }
    }
    return $out
}
#endregion

#region UI --------------------------------------------------------------------

function Show-RestoreForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Restore Mailbox -- soft-deleted'
    $form.ClientSize = New-Object System.Drawing.Size(1060, 700)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblFind = New-Object System.Windows.Forms.Label
    $lblFind.Text = 'Filter (name or address):'
    $lblFind.Location = New-Object System.Drawing.Point(12, ($y + 3))
    $lblFind.Size = New-Object System.Drawing.Size(160, 18)
    $form.Controls.Add($lblFind)

    $txtFind = New-Object System.Windows.Forms.TextBox
    $txtFind.Location = New-Object System.Drawing.Point(176, $y)
    $txtFind.Size = New-Object System.Drawing.Size(300, 24)
    $txtFind.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtFind.Text = "$Seed"
    $form.Controls.Add($txtFind)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Load soft-deleted'
    $btnLoad.Location = New-Object System.Drawing.Point(488, ($y - 2))
    $btnLoad.Size = New-Object System.Drawing.Size(150, 28)
    $form.Controls.Add($btnLoad)
    $y += 34

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, $y)
    $lblStatus.Size = New-Object System.Drawing.Size(1036, 18)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text = 'Press Load. The filter narrows the list; leave it blank for everything.'
    $form.Controls.Add($lblStatus)
    $y += 22

    $pb = Add-ReportProgressBar -Form $form -X 12 -Y $y -Width 1036 -Height 14
    $y += 24

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, $y)
    $grid.Size = New-Object System.Drawing.Size(1036, 250)
    # Grid is editable ONLY so the Pick checkbox works; every text column is
    # ReadOnly individually. Ticks are the selection: Restore in place acts on
    # every ticked row, Copy wants exactly one tick.
    $grid.ReadOnly = $false
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowHeadersVisible = $false
    $grid.Font = New-Object System.Drawing.Font('Consolas', 9)
    $colPick = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colPick.HeaderText = ''; $colPick.Name = 'Pick'; $colPick.FillWeight = 4
    [void]$grid.Columns.Add($colPick)
    foreach ($c in @(@('DisplayName', 22), @('Address', 30), @('Deleted', 14), @('DaysLeft', 8), @('Type', 12))) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c[0]; $col.HeaderText = $c[0]; $col.FillWeight = $c[1]; $col.ReadOnly = $true
        [void]$grid.Columns.Add($col)
    }
    $colG = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colG.Name = 'Guid'; $colG.Visible = $false
    [void]$grid.Columns.Add($colG)
    $form.Controls.Add($grid)
    $y += 260

    # --- actions ---
    $grpA = New-Object System.Windows.Forms.GroupBox
    $grpA.Text = 'Restore in place -- tick one or many; each comes back as itself (needs its user account to still exist)'
    $grpA.Location = New-Object System.Drawing.Point(12, $y)
    $grpA.Size = New-Object System.Drawing.Size(1036, 52)
    $form.Controls.Add($grpA)

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = 'Tick all'
    $btnAll.Location = New-Object System.Drawing.Point(12, 18)
    $btnAll.Size = New-Object System.Drawing.Size(100, 26)
    $grpA.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = 'Untick all'
    $btnNone.Location = New-Object System.Drawing.Point(120, 18)
    $btnNone.Size = New-Object System.Drawing.Size(100, 26)
    $grpA.Controls.Add($btnNone)

    $btnUndo = New-Object System.Windows.Forms.Button
    $btnUndo.Text = 'Restore ticked in place'   # one or many
    $btnUndo.Location = New-Object System.Drawing.Point(240, 18)
    $btnUndo.Size = New-Object System.Drawing.Size(200, 26)
    $btnUndo.Enabled = $false
    $grpA.Controls.Add($btnUndo)
    $y += 60

    $grpB = New-Object System.Windows.Forms.GroupBox
    $grpB.Text = 'Copy contents into another mailbox -- tick exactly ONE; works whether or not the user still exists'
    $grpB.Location = New-Object System.Drawing.Point(12, $y)
    $grpB.Size = New-Object System.Drawing.Size(1036, 52)
    $form.Controls.Add($grpB)

    $lblTgt = New-Object System.Windows.Forms.Label
    $lblTgt.Text = 'Into mailbox:'
    $lblTgt.Location = New-Object System.Drawing.Point(12, 22)
    $lblTgt.Size = New-Object System.Drawing.Size(90, 18)
    $grpB.Controls.Add($lblTgt)

    $txtTgt = New-Object System.Windows.Forms.TextBox
    $txtTgt.Location = New-Object System.Drawing.Point(104, 18)
    $txtTgt.Size = New-Object System.Drawing.Size(300, 24)
    $txtTgt.Font = New-Object System.Drawing.Font('Consolas', 9)
    $grpB.Controls.Add($txtTgt)

    $lblRoot = New-Object System.Windows.Forms.Label
    $lblRoot.Text = 'under folder:'
    $lblRoot.Location = New-Object System.Drawing.Point(416, 22)
    $lblRoot.Size = New-Object System.Drawing.Size(90, 18)
    $grpB.Controls.Add($lblRoot)

    $txtRoot = New-Object System.Windows.Forms.TextBox
    $txtRoot.Location = New-Object System.Drawing.Point(508, 18)
    $txtRoot.Size = New-Object System.Drawing.Size(240, 24)
    $txtRoot.Font = New-Object System.Drawing.Font('Consolas', 9)
    $grpB.Controls.Add($txtRoot)

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Start copy'
    $btnCopy.Location = New-Object System.Drawing.Point(760, 17)
    $btnCopy.Size = New-Object System.Drawing.Size(130, 26)
    $btnCopy.Enabled = $false
    $grpB.Controls.Add($btnCopy)
    $y += 60

    # --- requests ---
    $lblReq = New-Object System.Windows.Forms.Label
    $lblReq.Text = 'Restore requests in progress:'
    $lblReq.Location = New-Object System.Drawing.Point(12, $y)
    $lblReq.Size = New-Object System.Drawing.Size(300, 18)
    $form.Controls.Add($lblReq)

    $btnReq = New-Object System.Windows.Forms.Button
    $btnReq.Text = 'Refresh'
    $btnReq.Location = New-Object System.Drawing.Point(918, ($y - 4))
    $btnReq.Size = New-Object System.Drawing.Size(130, 26)
    $form.Controls.Add($btnReq)
    $y += 24

    $gridReq = New-Object System.Windows.Forms.DataGridView
    $gridReq.Location = New-Object System.Drawing.Point(12, $y)
    $gridReq.Size = New-Object System.Drawing.Size(1036, 90)
    $gridReq.ReadOnly = $true; $gridReq.AllowUserToAddRows = $false
    $gridReq.AutoSizeColumnsMode = 'Fill'; $gridReq.RowHeadersVisible = $false
    $gridReq.Font = New-Object System.Drawing.Font('Consolas', 9)
    foreach ($c in @('Name', 'Target', 'Status', 'Percent')) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c; $col.HeaderText = $c
        [void]$gridReq.Columns.Add($col)
    }
    $form.Controls.Add($gridReq)
    $y += 100

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Location = New-Object System.Drawing.Point(918, $y)
    $btnClose.Size = New-Object System.Drawing.Size(130, 28)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force

    $refreshReq = {
        $gridReq.Rows.Clear()
        try {
            foreach ($r in (Get-RestoreRequestList)) { [void]$gridReq.Rows.Add($r.Name, $r.Target, $r.Status, $r.Percent) }
        } catch { [void]$gridReq.Rows.Add('(could not read requests)', '', "$($_.Exception.Message)", '') }
    }

    $btnLoad.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $grid.Rows.Clear(); $btnUndo.Enabled = $false; $btnCopy.Enabled = $false
            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Reading'
            & $p 'Reading soft-deleted mailboxes...' -1
            $all = Get-SoftDeletedMailboxList
            $q = "$($txtFind.Text)".Trim().ToLower()
            $shown = 0
            foreach ($m in $all) {
                $hay = ("$($m.DisplayName) $($m.PrimarySmtpAddress) $($m.Alias)").ToLower()
                if ($q -and $hay.IndexOf($q) -lt 0) { continue }
                $del = ''; $left = ''
                try {
                    $d = [datetime]$m.WhenSoftDeleted
                    $del = $d.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
                    $left = [Math]::Max(0, [int](30 - ((Get-Date).ToUniversalTime() - $d.ToUniversalTime()).TotalDays))
                } catch { }
                $row = $grid.Rows.Add($false, "$($m.DisplayName)", "$($m.PrimarySmtpAddress)", $del, "$left", "$($m.RecipientTypeDetails)", "$($m.ExchangeGuid)")
                if ($left -ne '' -and [int]$left -le 5) { $grid.Rows[$row].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 210, 200) }
                $shown++
            }
            & $p "$shown of $($all.Count) soft-deleted mailbox(es) shown. Tick the ones to restore. Rows in red expire within 5 days." 100
            $btnUndo.Enabled = ($shown -gt 0); $btnCopy.Enabled = ($shown -gt 0)
            & $refreshReq
        } catch {
            $lblStatus.Text = 'Load failed -- see the error dialog.'
            Show-HandlerError -ErrorRecord $_ -Context 'Load (Get-Mailbox -SoftDeletedMailbox)'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    $btnReq.Add_Click({ if (-not $form.Busy) { & $refreshReq } })

    $btnAll.Add_Click({  foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = $true };  $grid.EndEdit() | Out-Null })
    $btnNone.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = $false }; $grid.EndEdit() | Out-Null })

    # A checkbox click leaves the cell in edit mode until focus moves; commit
    # it at once so a tick followed straight by a button press counts.
    $grid.Add_CurrentCellDirtyStateChanged({ if ($grid.IsCurrentCellDirty) { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) | Out-Null } })

    $grid.Add_SelectionChanged({
        if ($grid.SelectedRows.Count -eq 1) {
            $n = "$($grid.SelectedRows[0].Cells['DisplayName'].Value)"
            if (-not "$($txtRoot.Text)".Trim() -or $txtRoot.Tag -eq 'auto') {
                $txtRoot.Text = "Restored - $n"; $txtRoot.Tag = 'auto'
            }
        }
    })
    $txtRoot.Add_KeyPress({ $txtRoot.Tag = 'manual' })

    $btnUndo.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            # MANY AT ONCE. Restore in place is one call per mailbox with no
            # shared inputs, so a batch is just the single case in a loop. Copy
            # stays one-at-a-time -- it needs a target and a folder name per
            # source, and those are decisions, not a selection.
            $grid.EndEdit() | Out-Null
            $sel = @($grid.Rows | Where-Object { $_.Cells['Pick'].Value -eq $true })
            if ($sel.Count -eq 0) { $lblStatus.Text = 'Tick one or more mailboxes.'; return }
            $names = @($sel | ForEach-Object { "$($_.Cells['DisplayName'].Value)  <$($_.Cells['Address'].Value)>" })
            $a = Show-ToolDialog(
                ("Restore $($sel.Count) mailbox(es) in place?" + "`r`n`r`n  " + ($names -join "`r`n  ") + "`r`n`r`n" +
                 "Each needs its user account to still exist in Entra. Any that were deleted and synced away will fail and be marked red -- use Copy for those."),
                'Confirm', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnUndo.Enabled = $false; $btnCopy.Enabled = $false; $btnLoad.Enabled = $false
            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Restoring'
            $ok = 0; $failed = 0; $i = 0; $errs = @()

            foreach ($r in $sel) {
                $i++
                $smtp = "$($r.Cells['Address'].Value)"
                & $p "Restoring $i of $($sel.Count): $smtp" ([int](100 * $i / $sel.Count))
                # EACH IN ITS OWN TRY. One mailbox whose user is gone must not
                # stop the nine whose users are still there.
                try {
                    Restore-SoftDeletedInPlace -Smtp $smtp | Out-Null
                    # SHARED STAYS SHARED. A shared mailbox restored with its
                    # user can come back as a UserMailbox; Learn (Set-Mailbox)
                    # -Type Shared puts it back. Only when it WAS shared.
                    if ("$($r.Cells['Type'].Value)" -eq 'SharedMailbox') {
                        try {
                            $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $smtp; ErrorAction = 'Stop' }
                            if ($mb -and "$($mb.RecipientTypeDetails)" -eq 'UserMailbox') {
                                Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{ Identity = $smtp; Type = 'Shared'; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                                $r.Cells['Type'].Value = 'SharedMailbox (converted back)'
                            }
                        } catch { $errs += "$smtp -- restored, but could not convert back to Shared: $($_.Exception.Message)" }
                    }
                    $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200, 240, 200)
                    $r.Cells['Pick'].Value = $false
                    $ok++
                } catch {
                    $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 200, 200)
                    $failed++
                    $msg = "$($_.Exception.Message)"
                    $why = if ($msg -match 'not found|does not exist|deleted|no such') { 'user account gone -- use Copy' } else { $msg }
                    $errs += "$smtp -- $why"
                }
            }

            $result = if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }
            Write-CompletedLog -Action 'RestoreMailboxInPlace' -Target $(if ($sel.Count -eq 1) { Get-LogTarget $sel[0].Cells['Address'].Value } else { "$($sel.Count) mailboxes" }) `
                -Result $result -Detail "$ok restored, $failed failed (Undo-SoftDeletedMailbox)$(if ($errs) { ': ' + ($errs -join ' | ') })"
            $lblStatus.Text = "$ok restored, $failed failed." + $(if ($failed) { ' Failed rows are red.' } else { ' Allow a few minutes before they answer.' })
            if ($failed) {
                Show-ToolDialog(("$failed of $($sel.Count) did not restore:" + "`r`n`r`n  " + ($errs -join "`r`n  ") +
                    "`r`n`r`nA mailbox whose user account is gone cannot be restored in place. Use 'Copy contents into another mailbox' for those."),
                    'Partial', 'OK', 'Warning') | Out-Null
            }
        } catch {
            $lblStatus.Text = 'Restore failed -- see the error dialog.'
            Show-HandlerError -ErrorRecord $_ -Context 'Restore in place'
        } finally {
            $btnUndo.Enabled = $true; $btnCopy.Enabled = $true; $btnLoad.Enabled = $true
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    $btnCopy.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $grid.EndEdit() | Out-Null
            $picked = @($grid.Rows | Where-Object { $_.Cells['Pick'].Value -eq $true })
            # One tick, or no ticks and one highlighted row -- either names the source.
            if ($picked.Count -eq 0 -and $grid.SelectedRows.Count -eq 1) { $picked = @($grid.SelectedRows[0]) }
            if ($picked.Count -ne 1) { $lblStatus.Text = "Copy is one at a time -- $($picked.Count) ticked. Each copy needs its own target and folder."; return }
            $r = $picked[0]
            $tgt = Get-TargetSmtp -Text $txtTgt.Text
            if (-not $tgt) { $lblStatus.Text = 'Enter the mailbox to copy INTO.'; return }
            $root = "$($txtRoot.Text)".Trim()
            if (-not $root) { $lblStatus.Text = 'Enter a folder name to put the copy under.'; return }
            $tmb = $null
            try { $tmb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $tgt; ErrorAction = 'Stop' } } catch { }
            if (-not $tmb) { $lblStatus.Text = "Target mailbox '$tgt' was not found."; return }

            $src = "$($r.Cells['DisplayName'].Value)"
            $a = Show-ToolDialog(
                ("Copy the contents of:" + "`r`n  $src  ($($r.Cells['Address'].Value))" + "`r`n`r`ninto:" + "`r`n  $($tmb.PrimarySmtpAddress)" +
                 "`r`n`r`nunder a folder called:" + "`r`n  $root" + "`r`n`r`nThe soft-deleted mailbox is not changed. This runs in the background and can take hours."),
                'Confirm', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $name = "Restore $src -> $($tmb.Alias) $(Get-Date -Format 'yyyyMMdd-HHmm')"
            Start-ContentRestore -SourceGuid "$($r.Cells['Guid'].Value)" -Target "$($tmb.PrimarySmtpAddress)" -RootFolder $root -Name $name | Out-Null
            Write-CompletedLog -Action 'RestoreMailboxContents' -Target (Get-LogTarget $tmb.PrimarySmtpAddress) -Result 'Success' `
                -Detail "From soft-deleted $($r.Cells['Address'].Value) under folder '$root'; request '$name'"
            $lblStatus.Text = "Restore request queued: $name. Press Refresh to watch it."
            & $refreshReq
        } catch {
            $lblStatus.Text = 'Copy failed -- see the error dialog.'
            Show-HandlerError -ErrorRecord $_ -Context 'Copy (New-MailboxRestoreRequest)'
            Write-FailedLog -Action 'RestoreMailboxContents' -Target "$($txtTgt.Text)" -ErrorRecord $_
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Restore Mailbox' | Out-Null
    [void]$form.ShowDialog()
}
#endregion

Show-RestoreForm -Seed (Get-TargetSmtp -Text $SmtpAddress)
