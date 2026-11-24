#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    6-CLOUD-Restore.ps1
    Copied from AdminConsole\RestoreMailbox.ps1 and rewired to read Step 5's
    output instead of a filter box. The console tool is untouched.

    Author   : Glenn Armbruster
    Revised  : 2026-09-08
#>
# TOOLKIT-HELP-RUNS: reads the soft-deleted list and the recipient directory. Restore creates and changes things and asks first.
# TOOLKIT-HELP-BEGIN
# STEP 6 of the Migration project. Brings back, in the cloud, everything
# Step 5 pushed out of sync scope -- as CLOUD-ONLY objects.
#
# Reads the newest Migration-Moved_*.csv and works out where each object is
# now:
#
#   SHARED MAILBOX   The sync deleted the Entra user, which put the mailbox
#                    in "soft-deleted" for 30 days. Undo-SoftDeletedMailbox
#                    restores the USER AND THE MAILBOX in one call -- the
#                    PowerShell form of admin center > Deleted users >
#                    Restore -- with a password set at the same time, so
#                    there is no prompt. It comes back with its addresses,
#                    permissions and contents. If it came back as a user
#                    mailbox it is converted to shared.
#
#   DISTRIBUTION LIST  Lists have NO recycle bin: the sync deleted it
#                    outright. It is RECREATED from the Step 1 export --
#                    name, alias, address, type, owners, join/depart rules
#                    -- then its members are added one by one. Step 7 stamps
#                    the rest (proxies, X500, restrictions, custom attributes).
#
# THE STATE COLUMN IS THE GATE. "ready" rows are ticked for you. "still
# synced" means the sync has not run yet -- wait and press Refresh. "already
# restored" means a previous run did it. Nothing happens to a row that is
# not ticked, and nothing happens at all until you confirm.
#
# MAILBOXES FIRST, THEN LISTS, always -- a list's members include mailboxes
# from the same batch, and a member that does not exist yet is skipped.
#
# THE PASSWORD. $DefaultPassword at the top of this script. Leave it blank
# and each mailbox gets a random one nobody knows, which is the right answer
# for a shared mailbox (its account stays sign-in blocked either way). Fill
# it in if the team has a standard.
#
# Output lands next to this script:
#     Migration-Restored_<date>_<time>.csv
#
# Needs Exchange Online connected in this window. Run Step 7 next.
# TOOLKIT-HELP-END

# ---- The one setting. Blank = random per mailbox, never shown. ----
$DefaultPassword = ''

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$toolkitRoot = $PSScriptRoot
while ($toolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $toolkitRoot 'Common.ps1'))) {
    $toolkitRoot = Split-Path -Parent $toolkitRoot
}
if (-not $toolkitRoot) { throw "Common.ps1 was not found in any folder above '$PSScriptRoot'." }
. (Join-Path $toolkitRoot 'Common.ps1')
. (Join-Path $PSScriptRoot '_Shared.ps1')

$basePath = $PSScriptRoot

#region Data ------------------------------------------------------------------

function exo { param([string]$n, [hashtable]$p) Invoke-EXOCommand -Name $n -Parameters $p }

function New-RestorePassword {
    param([string]$Fixed)
    $plain = "$Fixed"
    if (-not $plain) {
        # 24 characters from a set with no look-alikes. Never printed.
        $chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789!#$%&*+-=?@'
        $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
        $b = New-Object byte[] 24; $rng.GetBytes($b)
        $plain = -join ($b | ForEach-Object { $chars[$_ % $chars.Length] })
    }
    return (ConvertTo-SecureString -String $plain -AsPlainText -Force)
}

function Get-SoftDeletedMailboxList {
    # Learn (Get-Mailbox): -SoftDeletedMailbox returns only soft-deleted ones.
    return @(exo 'Get-Mailbox' @{ SoftDeletedMailbox = $true; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
}

function Restore-SoftDeletedInPlace {
    <#
        Learn 2026-09-08 (Undo-SoftDeletedMailbox): -SoftDeletedObject <alias
        or address>; -Password <SecureString> "specifies a new password for
        the mailbox". User and mailbox come back together. No prompt.
    #>
    param([Parameter(Mandatory)][string]$Smtp, [Parameter(Mandatory)][securestring]$Password)
    return (exo 'Undo-SoftDeletedMailbox' @{ SoftDeletedObject = $Smtp; Password = $Password; ErrorAction = 'Stop' })
}

function Get-ObjectState {
    <#
        Where is this object right now? Returns State, Ready, Detail, and for
        a soft-deleted mailbox its identity for Undo-.
    #>
    param([Parameter(Mandatory)][object]$Row, [object[]]$SoftDeleted)
    $smtp = "$($Row.PrimarySmtpAddress)"
    if ($Row.Type -eq 'SMB') {
        $sd = $SoftDeleted | Where-Object { "$($_.ExchangeGuid)" -eq "$($Row.ExchangeGuid)" -or "$($_.PrimarySmtpAddress)" -ieq $smtp } | Select-Object -First 1
        if ($sd) {
            $left = ''
            try { $left = [Math]::Max(0, [int](30 - ((Get-Date).ToUniversalTime() - ([datetime]$sd.WhenSoftDeleted).ToUniversalTime()).TotalDays)) } catch { }
            return [PSCustomObject]@{ State = 'soft-deleted -- ready'; Ready = $true; Detail = "$left day(s) left"; Identity = "$($sd.PrimarySmtpAddress)"; DaysLeft = "$left" }
        }
        $mb = $null
        try { $mb = exo 'Get-Mailbox' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
        if ($mb) {
            if ("$($mb.IsDirSynced)" -eq 'True') { return [PSCustomObject]@{ State = 'still synced -- wait'; Ready = $false; Detail = 'sync has not removed it yet'; Identity = ''; DaysLeft = '' } }
            return [PSCustomObject]@{ State = 'already restored'; Ready = $false; Detail = "$($mb.RecipientTypeDetails), cloud-only"; Identity = ''; DaysLeft = '' }
        }
        return [PSCustomObject]@{ State = 'NOT FOUND'; Ready = $false; Detail = 'not active, not soft-deleted -- check Deleted users in the admin center'; Identity = ''; DaysLeft = '' }
    }
    # DL
    $dg = $null
    try { $dg = exo 'Get-DistributionGroup' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
    if ($dg) {
        if ("$($dg.IsDirSynced)" -eq 'True') { return [PSCustomObject]@{ State = 'still synced -- wait'; Ready = $false; Detail = 'sync has not removed it yet'; Identity = ''; DaysLeft = '' } }
        return [PSCustomObject]@{ State = 'already recreated'; Ready = $false; Detail = 'cloud-only list exists'; Identity = ''; DaysLeft = '' }
    }
    $other = $null
    try { $other = exo 'Get-Recipient' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
    if ($other) { return [PSCustomObject]@{ State = 'ADDRESS IN USE'; Ready = $false; Detail = "$($other.RecipientTypeDetails) $($other.Name) holds this address"; Identity = ''; DaysLeft = '' } }
    if (-not $Row.Export) { return [PSCustomObject]@{ State = 'NO STEP 1 EXPORT'; Ready = $false; Detail = 'gone, but no Migration-DL-ONPREM row to rebuild from'; Identity = ''; DaysLeft = '' } }
    return [PSCustomObject]@{ State = 'gone -- ready to recreate'; Ready = $true; Detail = "$($Row.Export.MemberCount) member(s) in export"; Identity = ''; DaysLeft = '' }
}

function New-CloudListFromExport {
    <#
        Learn 2026-09-08 (New-DistributionGroup, EXO): -Name, -Alias,
        -DisplayName, -PrimarySmtpAddress, -Type Distribution|Security,
        -ManagedBy, -MemberJoinRestriction, -MemberDepartRestriction,
        -RequireSenderAuthenticationEnabled, -IgnoreNamingPolicy, -Notes.
        Members are added afterwards, one at a time, so one missing member
        does not fail the create.
    #>
    param([Parameter(Mandatory)][object]$OnPrem, [object]$Cloud)
    $type = if ("$($OnPrem.GroupType)" -match 'SecurityEnabled') { 'Security' } else { 'Distribution' }
    $owners = @(ConvertFrom-CsvList $OnPrem.ManagedBy)
    if ($Cloud -and "$($Cloud.ManagedBy)") { $owners = @($owners + (ConvertFrom-CsvList $Cloud.ManagedBy) | Select-Object -Unique) }
    $p = @{
        Name                               = "$($OnPrem.Name)"
        Alias                              = "$($OnPrem.Alias)"
        DisplayName                        = "$($OnPrem.DisplayName)"
        PrimarySmtpAddress                 = "$($OnPrem.PrimarySmtpAddress)"
        Type                               = $type
        RequireSenderAuthenticationEnabled = ("$($OnPrem.RequireSenderAuthenticationEnabled)" -ne 'False')
        IgnoreNamingPolicy                 = $true
        ErrorAction                        = 'Stop'
    }
    if ("$($OnPrem.MemberJoinRestriction)")   { $p.MemberJoinRestriction   = "$($OnPrem.MemberJoinRestriction)" }
    if ("$($OnPrem.MemberDepartRestriction)") { $p.MemberDepartRestriction = "$($OnPrem.MemberDepartRestriction)" }
    if ("$($OnPrem.Notes)")                   { $p.Notes = "$($OnPrem.Notes)" }
    # Owners that resolve go on at create; the rest are Step 7's problem and
    # are reported.
    $good = @(); $badOwners = @()
    foreach ($o in $owners) { try { [void](exo 'Get-Recipient' @{ Identity = $o; ErrorAction = 'Stop' }); $good += $o } catch { $badOwners += $o } }
    if ($good.Count) { $p.ManagedBy = $good }
    $dg = exo 'New-DistributionGroup' $p
    # Members: cloud export first (it knows cloud-only recipients), on-prem
    # as the fallback, union of both.
    $members = @(ConvertFrom-CsvList $OnPrem.Members)
    if ($Cloud -and "$($Cloud.Members)") { $members = @($members + (ConvertFrom-CsvList $Cloud.Members) | Select-Object -Unique) }
    $added = 0; $missing = @()
    foreach ($m in $members) {
        try {
            exo 'Add-DistributionGroupMember' @{ Identity = "$($dg.PrimarySmtpAddress)"; Member = $m; BypassSecurityGroupManagerCheck = $true; ErrorAction = 'Stop' } | Out-Null
            $added++
        } catch { $missing += $m }
    }
    return [PSCustomObject]@{ Group = $dg; Added = $added; Missing = $missing; BadOwners = $badOwners }
}
#endregion

#region UI --------------------------------------------------------------------

function Show-RestoreForm {
    param([Parameter(Mandatory)][object]$MovedFile, [object[]]$Plan)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "6 - CLOUD Restore  --  $($MovedFile.Name)"
    $form.ClientSize = New-Object System.Drawing.Size(1100, 560)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, $y)
    $lblStatus.Size = New-Object System.Drawing.Size(940, 18)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text = "Press Refresh to read the cloud state of the $($Plan.Count) object(s) in $($MovedFile.Name)."
    $form.Controls.Add($lblStatus)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Refresh'
    $btnLoad.Location = New-Object System.Drawing.Point(958, ($y - 4))
    $btnLoad.Size = New-Object System.Drawing.Size(130, 28)
    $form.Controls.Add($btnLoad)
    $y += 26

    $pb = Add-ReportProgressBar -Form $form -X 12 -Y $y -Width 1076 -Height 14
    $y += 24

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, $y)
    $grid.Size = New-Object System.Drawing.Size(1076, 380)
    $grid.ReadOnly = $false
    $grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = 'Fill'; $grid.RowHeadersVisible = $false
    $grid.Font = New-Object System.Drawing.Font('Consolas', 9)
    $colPick = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colPick.HeaderText = ''; $colPick.Name = 'Pick'; $colPick.FillWeight = 4
    [void]$grid.Columns.Add($colPick)
    foreach ($c in @(@('Type', 5), @('Name', 18), @('Address', 26), @('State', 18), @('Detail', 25), @('DaysLeft', 6))) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c[0]; $col.HeaderText = $c[0]; $col.FillWeight = $c[1]; $col.ReadOnly = $true
        [void]$grid.Columns.Add($col)
    }
    $colIx = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colIx.Name = 'Ix'; $colIx.Visible = $false
    [void]$grid.Columns.Add($colIx)
    $form.Controls.Add($grid)
    $y += 390

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = 'Tick all ready'; $btnAll.Location = New-Object System.Drawing.Point(12, $y); $btnAll.Size = New-Object System.Drawing.Size(120, 28)
    $form.Controls.Add($btnAll)
    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = 'Untick all'; $btnNone.Location = New-Object System.Drawing.Point(140, $y); $btnNone.Size = New-Object System.Drawing.Size(100, 28)
    $form.Controls.Add($btnNone)
    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = 'Restore ticked'; $btnGo.Location = New-Object System.Drawing.Point(260, $y); $btnGo.Size = New-Object System.Drawing.Size(160, 28); $btnGo.Enabled = $false
    $form.Controls.Add($btnGo)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(958, $y); $btnClose.Size = New-Object System.Drawing.Size(130, 28)
    $form.Controls.Add($btnClose); $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy'  -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Plan'  -NotePropertyValue $Plan  -Force
    $form | Add-Member -NotePropertyName 'State' -NotePropertyValue @{}    -Force

    $btnAll.Add_Click({  foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = [bool]($form.State["$($r.Cells['Ix'].Value)"].Ready) }; $grid.EndEdit() | Out-Null })
    $btnNone.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells['Pick'].Value = $false }; $grid.EndEdit() | Out-Null })
    $grid.Add_CurrentCellDirtyStateChanged({ if ($grid.IsCurrentCellDirty) { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) | Out-Null } })

    $btnLoad.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $grid.Rows.Clear(); $btnGo.Enabled = $false
            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Reading'
            & $p 'Reading soft-deleted mailboxes...' -1
            $sd = Get-SoftDeletedMailboxList
            $ready = 0; $i = 0
            foreach ($row in $form.Plan) {
                $i++
                & $p "Checking $i of $($form.Plan.Count): $($row.PrimarySmtpAddress)" ([int](100 * $i / $form.Plan.Count))
                $st = Get-ObjectState -Row $row -SoftDeleted $sd
                $form.State["$($row.Ix)"] = $st
                $n = $grid.Rows.Add([bool]$st.Ready, "$($row.Type)", "$($row.Name)", "$($row.PrimarySmtpAddress)", $st.State, $st.Detail, $st.DaysLeft, "$($row.Ix)")
                if ($st.Ready) { $ready++; $grid.Rows[$n].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(225, 245, 225) }
                elseif ($st.State -match 'already') { $grid.Rows[$n].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray }
                elseif ($st.State -match 'wait') { $grid.Rows[$n].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 205) }
                else { $grid.Rows[$n].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 210, 200) }
            }
            & $p "$ready of $($form.Plan.Count) ready and ticked. Yellow = sync not through yet (wait, Refresh). Red = needs a look." 100
            $btnGo.Enabled = ($ready -gt 0)
        } catch {
            $lblStatus.Text = 'Refresh failed -- see the error dialog.'
            Show-HandlerError -ErrorRecord $_ -Context 'Refresh (Get-Mailbox -SoftDeletedMailbox)'
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    $btnGo.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $grid.EndEdit() | Out-Null
            $picked = @($grid.Rows | Where-Object { $_.Cells['Pick'].Value -eq $true })
            if ($picked.Count -eq 0) { $lblStatus.Text = 'Tick one or more rows.'; return }
            $notReady = @($picked | Where-Object { -not $form.State["$($_.Cells['Ix'].Value)"].Ready })
            if ($notReady.Count) { $lblStatus.Text = "$($notReady.Count) ticked row(s) are not ready -- untick them (Refresh re-reads the state)."; return }
            # Mailboxes first, then lists -- list members include the mailboxes.
            $picked = @($picked | Sort-Object { if ("$($_.Cells['Type'].Value)" -eq 'SMB') { 0 } else { 1 } })
            $nS = @($picked | Where-Object { "$($_.Cells['Type'].Value)" -eq 'SMB' }).Count
            $nD = $picked.Count - $nS
            $a = [System.Windows.Forms.MessageBox]::Show(
                ("Restore $nS mailbox(es) and recreate $nD list(s) as cloud-only objects?" + "`r`n`r`n  " +
                 (($picked | ForEach-Object { "$($_.Cells['Type'].Value)  $($_.Cells['Address'].Value)" }) -join "`r`n  ") +
                 "`r`n`r`nMailboxes come back with their permissions and contents. Lists are rebuilt from the Step 1 export; Step 7 finishes them."),
                'Confirm', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnGo.Enabled = $false; $btnLoad.Enabled = $false
            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Restoring'
            $ok = 0; $failed = 0; $i = 0; $errs = @(); $out = @()
            foreach ($r in $picked) {
                $i++
                $row = $form.Plan | Where-Object { "$($_.Ix)" -eq "$($r.Cells['Ix'].Value)" } | Select-Object -First 1
                $st = $form.State["$($row.Ix)"]
                $smtp = "$($row.PrimarySmtpAddress)"
                & $p "$i of $($picked.Count): $smtp" ([int](100 * $i / $picked.Count))
                $detail = ''
                try {
                    if ($row.Type -eq 'SMB') {
                        Restore-SoftDeletedInPlace -Smtp $st.Identity -Password (New-RestorePassword -Fixed $DefaultPassword) | Out-Null
                        $detail = 'restored (user + mailbox)'
                        # A mailbox restored as UserMailbox becomes Shared.
                        # Learn (Set-Mailbox): -Type Shared.
                        $mb = $null
                        try { $mb = exo 'Get-Mailbox' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
                        if ($mb -and "$($mb.RecipientTypeDetails)" -eq 'UserMailbox') {
                            exo 'Set-Mailbox' @{ Identity = $smtp; Type = 'Shared'; ErrorAction = 'Stop' } | Out-Null
                            $detail += ', converted to Shared'
                        }
                        Write-CompletedLog -Action 'MigrationRestoreMailbox' -Target (Get-LogTarget $smtp) -Result 'Success' -Detail "Undo-SoftDeletedMailbox; $detail"
                    } else {
                        $cloudRow = $row.CloudExport
                        $res = New-CloudListFromExport -OnPrem $row.Export -Cloud $cloudRow
                        $detail = "recreated; $($res.Added) member(s) added"
                        if ($res.Missing.Count)   { $detail += "; $($res.Missing.Count) member(s) not found: $($res.Missing -join ', ')" }
                        if ($res.BadOwners.Count) { $detail += "; owner(s) not found: $($res.BadOwners -join ', ')" }
                        Write-CompletedLog -Action 'MigrationRecreateList' -Target (Get-LogTarget $smtp) -Result $(if ($res.Missing.Count -or $res.BadOwners.Count) { 'Partial' } else { 'Success' }) -Detail "New-DistributionGroup; $detail"
                    }
                    $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200, 240, 200)
                    $r.Cells['State'].Value = 'done'; $r.Cells['Detail'].Value = $detail; $r.Cells['Pick'].Value = $false
                    $form.State["$($row.Ix)"] = [PSCustomObject]@{ State = 'done'; Ready = $false; Detail = $detail; Identity = ''; DaysLeft = '' }
                    $ok++
                } catch {
                    $failed++
                    $detail = "FAILED: $($_.Exception.Message)"
                    $r.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 200, 200)
                    $r.Cells['Detail'].Value = $detail
                    $errs += "$smtp -- $($_.Exception.Message)"
                    Write-FailedLog -Action $(if ($row.Type -eq 'SMB') { 'MigrationRestoreMailbox' } else { 'MigrationRecreateList' }) -Target (Get-LogTarget $smtp) -ErrorRecord $_
                }
                $out += [PSCustomObject]@{ Type = $row.Type; Name = $row.Name; PrimarySmtpAddress = $smtp; Result = $(if ($detail -like 'FAILED*') { 'Failed' } else { 'Success' }); Detail = $detail; RestoredUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }
            }
            $outFile = Join-Path $basePath "Migration-Restored_$(Get-MigrationTimestamp).csv"
            $out | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
            $lblStatus.Text = "$ok done, $failed failed. Written to $([System.IO.Path]::GetFileName($outFile)). Next: Step 7."
            if ($failed) {
                [System.Windows.Forms.MessageBox]::Show(("$failed of $($picked.Count) failed:" + "`r`n`r`n  " + ($errs -join "`r`n  ")), 'Partial', 'OK', 'Warning') | Out-Null
            }
        } catch {
            $lblStatus.Text = 'Restore failed -- see the error dialog.'
            Show-HandlerError -ErrorRecord $_ -Context 'Restore ticked'
        } finally {
            $btnGo.Enabled = $true; $btnLoad.Enabled = $true
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title '6 - CLOUD Restore' | Out-Null
    [void]$form.ShowDialog()
}
#endregion

# ---- Inputs: newest Moved file, newest DL exports (for the recreate) ----
$movedFile = Get-NewestExport -Folder $basePath -Pattern 'Migration-Moved_*.csv'
if (-not $movedFile) {
    [System.Windows.Forms.MessageBox]::Show('No Migration-Moved_*.csv found next to this script. Run Step 5 first.', '6 - CLOUD Restore', 'OK', 'Warning') | Out-Null
    return
}
$moved = @(Import-Csv -LiteralPath $movedFile.FullName | Where-Object { $_.Result -eq 'Moved' -or $_.Result -eq 'AlreadyThere' })
if ($moved.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("$($movedFile.Name) has no moved rows.", '6 - CLOUD Restore', 'OK', 'Warning') | Out-Null; return }

$dlOn = @(); $dlCl = @()
$f = Get-NewestExport -Folder $basePath -Pattern 'Migration-DL-ONPREM_*.csv'; if ($f) { $dlOn = @(Import-Csv -LiteralPath $f.FullName) }
$f = Get-NewestExport -Folder $basePath -Pattern 'Migration-DL-CLOUD_*.csv';  if ($f) { $dlCl = @(Import-Csv -LiteralPath $f.FullName) }

$plan = @(); $ix = 0
foreach ($m in $moved) {
    $ix++
    $smtp = "$($m.PrimarySmtpAddress)"
    $exp = $null; $cexp = $null
    if ($m.Type -eq 'DL') {
        $exp  = $dlOn | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $smtp -and -not "$($_.Error)" } | Select-Object -First 1
        $cexp = $dlCl | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $smtp -and -not "$($_.Error)" } | Select-Object -First 1
    }
    $plan += [PSCustomObject]@{ Ix = $ix; Type = "$($m.Type)"; Name = "$($m.Name)"; PrimarySmtpAddress = $smtp; ExchangeGuid = "$($m.ExchangeGuid)"; Export = $exp; CloudExport = $cexp }
}

Show-RestoreForm -MovedFile $movedFile -Plan $plan
