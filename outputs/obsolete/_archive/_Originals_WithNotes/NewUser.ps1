#Requires -Version 5.1
<#
.SYNOPSIS
    Enables on-prem Remote Mailboxes (with archive) in Exchange Hybrid, with
    automatic HQ vs Franchise detection.
.DESCRIPTION
    REWRITTEN 2026-07-24 from the original two-box version. What changed, and
    why:

    ONE INPUT BOX INSTEAD OF TWO. The original made the operator know in
    advance whether each account was HQ or Franchise and type it into the
    matching box. That is no longer needed: every account is looked up in HQ
    first, then Franchise, and whichever forest answers is where it acts.
    Mixed batches are fine -- HQ and Franchise accounts can sit in the same
    list, in any order.

    DETECTION IS BY sAMAccountName, NOT BY EMAIL -- and this distinction is
    load-bearing, so do not "simplify" it to reuse Find-TargetEnvironment.
    That existing function resolves an object by its MAIL ATTRIBUTE, which is
    right for every other tool in this suite because they all act on things
    that are already mail-enabled. The accounts THIS tool targets are not
    mail-enabled yet -- that is the entire reason it is being run -- so their
    mail attribute is empty, and a mail-based lookup would report "not found"
    for every valid account. This uses Common.ps1's Find-UserEnvironmentBySam
    instead, which resolves by sAMAccountName: an attribute that exists from
    the moment the account is created.

    "ACCOUNT DOES NOT EXIST YET" IS NOW A FIRST-CLASS OUTCOME rather than a
    confusing Exchange error mid-run. Every account is resolved BEFORE
    anything is changed; any that exist in neither forest are collected and
    reported in a single popup, and nothing is attempted for them. One popup
    listing them all, not one per account -- this is a batch tool.

    RESULTS STAY IN THE WINDOW. The original hid the form and wrote to the
    console, where output scrolled away. There is a results box now, and the
    form no longer hides.

    THE FRANCHISE DC DROPDOWN REMAINS, but is now consulted only for accounts
    that actually resolve to Franchise. HQ accounts still run forest-wide with
    no DC pinned. Note detection itself uses Common.ps1's own
    $script:FranchiseAutoDetectDC; the dropdown governs the WRITE, so the
    operator can still aim changes at a specific DC and verify against that
    same DC afterwards.

    UNCHANGED FROM THE ORIGINAL, deliberately:
      - Uses only *-RemoteMailbox cmdlets, which are on-prem-only and
        unambiguous. It never calls Get-Mailbox / Set-Mailbox, so it is immune
        to the on-prem/cloud cmdlet collision that makes the DL tools risky to
        run alongside an Exchange Online session.
      - Enables the archive with a second Enable-RemoteMailbox call, and
        treats an archive failure as a WARNING on an otherwise-successful
        enable rather than failing the whole operation.
      - Pre-checks with Get-RemoteMailbox and SKIPS accounts that already have
        one, showing their existing state.
      - RemoteRoutingAddress is <sam>@expresspros.mail.onmicrosoft.com.
      - HQ uses Set-ADServerSettings -ViewEntireForest; Franchise pins
        -PreferredServer and passes -DomainController.
      - Does NOT touch CN, UPN, sAMAccountName, display name, email address
        policy, or licensing. GUID stamping still intentionally omitted (the
        cloud mailbox is not provisioned quickly enough at enable time).
      - Logs one CompletedTasks row per verified enable.
.NOTES
    Author          : Glenn Armbruster
    Created         : 2026-06-17
    Revised         : 2026-06-24 0900
    Revised         : 2026-07-24 -- single box + HQ/Franchise auto-detect by
                       sAMAccountName; missing accounts reported explicitly;
                       results shown in-window. See REWRITTEN note above.
    Revised         : 2026-07-27 -- archive state now derived from ArchiveGuid
                       instead of raw ArchiveStatus, which read "None" on
                       archives that had in fact been created (the status is
                       cloud-side and had not replicated back yet). See
                       Format-ArchiveState. Also: an archive failure is now
                       recorded in the CompletedTasks row, not just in the
                       results box.
    Run from        : On-prem Exchange Management Shell
    Required modules: ActiveDirectory (via EMS host); no ExchangeOnlineManagement
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Common.ps1 supplies Find-UserEnvironmentBySam, Get-ADServerString,
# Format-ErrorDetail, Write-CompletedLog, $script:FranchiseDCs and
# $script:FranchiseAutoDetectDC -- so this tool no longer keeps its own copies
# of the DC list or the logging helper.
. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------
# Specific to this tool, so it stays local rather than moving to Common.ps1.
$RemoteRoutingDomain = 'expresspros.mail.onmicrosoft.com'
$DefaultFranchiseDC  = 'fran-vm-dc03.franchises.expresspersonnel.com'
#endregion

#region Parsing --------------------------------------------------------------
function Get-AccountList {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }
    return $Raw -split "[`r`n,;]+" |
           ForEach-Object { $_.Trim() } |
           Where-Object   { $_ -ne '' } |
           Select-Object  -Unique
}
#endregion

#region Actions --------------------------------------------------------------
function Get-OnPremState {
    # ArchiveGuid is selected alongside ArchiveStatus deliberately -- see
    # Format-ArchiveState below for why ArchiveStatus alone is misleading here.
    param([string]$Id, [string]$DC)   # empty $DC = HQ (no preferred server)
    $params = @{ Identity = $Id; ErrorAction = 'Stop' }
    if ($DC) { $params['DomainController'] = $DC }
    Get-RemoteMailbox @params |
        Select-Object Name, PrimarySmtpAddress, RemoteRoutingAddress, ArchiveStatus, ArchiveGuid
}

function Format-ArchiveState {
    <#
        Renders the archive line for a RemoteMailbox.

        WHY THIS EXISTS: reading raw ArchiveStatus here produced a false
        negative -- it showed "None" immediately after a successful
        Enable-RemoteMailbox -Archive, on accounts whose archive had in fact
        been created. ArchiveStatus is a CLOUD-side property; on an on-prem
        RemoteMailbox object it reflects what has replicated back from
        Exchange Online, which lags the enable by minutes to hours, so
        checking it seconds later will essentially always read "None". This is
        the same behaviour Microsoft KB4486848 describes.

        ArchiveGuid is the reliable indicator, because Enable-RemoteMailbox
        -Archive stamps it on the on-prem AD object as part of the enable
        itself -- nothing has to round-trip to the cloud for it to be true.
        Every other tool in this suite (MailboxHealth, MailboxPermissions,
        Mailbox-Group-Inspector, M365Logic) already derives ArchiveEnabled
        this way; this tool was the last one reading ArchiveStatus raw.

        So: GUID present = archive is enabled, full stop. ArchiveStatus is
        still shown in parentheses, but only as replication detail -- never as
        the verdict.
    #>
    param($State)
    $hasGuid = ($State.ArchiveGuid -and $State.ArchiveGuid -ne [Guid]::Empty)
    $status  = if ($State.ArchiveStatus) { $State.ArchiveStatus } else { 'None' }

    if (-not $hasGuid) {
        return ("      Archive              : Not enabled  (ArchiveStatus: {0})" -f $status)
    }
    if ($status -eq 'None') {
        # The exact case that prompted this: enabled on-prem, cloud not caught up.
        return "      Archive              : Enabled -- provisioning; may not display immediately"
    }
    return ("      Archive              : Enabled  (ArchiveStatus: {0})" -f $status)
}

function Format-OnPremState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("    --- $Heading ---")
    [void]$sb.AppendLine(("      Name                 : {0}" -f $State.Name))
    [void]$sb.AppendLine(("      PrimarySmtpAddress   : {0}" -f $State.PrimarySmtpAddress))
    [void]$sb.AppendLine(("      RemoteRoutingAddress : {0}" -f $State.RemoteRoutingAddress))
    [void]$sb.AppendLine((Format-ArchiveState -State $State))
    return $sb.ToString()
}

function Invoke-MailboxEnable {
    <#
        Enables one account's remote mailbox + archive. The environment is
        already resolved by the caller, so this simply acts on what it's told.
        Returns report text; the caller accumulates it.
    #>
    param(
        [string]$Account,
        [string]$DC,          # '' = HQ (forest-wide); populated = Franchise
        [string]$Mode         # 'HQ' or 'Franchise', for display + logging
    )
    $out = New-Object System.Text.StringBuilder
    $mb  = $Account.Trim()
    $rra = "$mb@$RemoteRoutingDomain"

    [void]$out.AppendLine("----- $mb  [$Mode] -----")

    # Pin the DC for Franchise writes; go forest-wide for HQ. Unchanged from
    # the original -- see its note about allowing replication before verifying
    # HQ changes.
    try {
        if ($DC) {
            Set-ADServerSettings -PreferredServer $DC -ErrorAction Stop
        } else {
            Set-ADServerSettings -PreferredServer $null -ViewEntireForest $true -ErrorAction Stop
        }
    } catch {
        [void]$out.AppendLine("  Could not set AD server scope: $($_.Exception.Message)")
        return $out.ToString()
    }

    # Pre-check: already mail-enabled? Doubles as the BEFORE state.
    $before = $null
    try { $before = Get-OnPremState -Id $mb -DC $DC } catch { $before = $null }
    if ($before) {
        [void]$out.AppendLine("  SKIPPED -- already has a remote mailbox.")
        [void]$out.Append((Format-OnPremState -Heading 'EXISTING (on-prem)' -State $before))
        return $out.ToString()
    }

    # 1. Enable the remote mailbox.
    $enableParams = @{
        Identity             = $mb
        Alias                = $mb
        RemoteRoutingAddress = $rra
        Confirm              = $false
        ErrorAction          = 'Stop'
    }
    if ($DC) { $enableParams['DomainController'] = $DC }
    try {
        Enable-RemoteMailbox @enableParams | Out-Null
        [void]$out.AppendLine("  Mailbox enabled.  ($rra)")
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'does not support recipients of this type' -or $msg -match 'already') {
            [void]$out.AppendLine("  SKIPPED -- already enabled, or not eligible for a remote mailbox.")
        } else {
            [void]$out.AppendLine("  FAILED to enable: $msg")
        }
        return $out.ToString()
    }

    # 2. Enable the archive. A failure here is a warning, not a failure of the
    #    enable above -- the mailbox exists either way.
    $archiveParams = @{ Identity = $mb; Archive = $true; ErrorAction = 'Stop' }
    if ($DC) { $archiveParams['DomainController'] = $DC }
    $archiveOk = $false
    try {
        Enable-RemoteMailbox @archiveParams | Out-Null
        $archiveOk = $true
        [void]$out.AppendLine("  Archive enabled.")
    } catch {
        [void]$out.AppendLine("  WARNING -- mailbox enabled but archive failed: $($_.Exception.Message)")
    }

    # 3. Verify, then log only on a verified success.
    #    The archive outcome is carried into the log line: verification passes
    #    on the mailbox alone, so without this an archive failure would produce
    #    a CSV row that reads as a clean success once the window is closed.
    try {
        $after = Get-OnPremState -Id $mb -DC $DC
        [void]$out.AppendLine("  Verified.")
        [void]$out.Append((Format-OnPremState -Heading 'AFTER (on-prem)' -State $after))
        $logTask = if ($archiveOk) { "Enabled RemoteMailbox $mb ($Mode)" }
                   else            { "Enabled RemoteMailbox $mb ($Mode) -- ARCHIVE FAILED" }
        Write-CompletedLog -Task $logTask
        [void]$out.AppendLine("  Logged to CompletedTasks.csv")
    } catch {
        [void]$out.AppendLine("  Could not verify after enable: $($_.Exception.Message)")
        [void]$out.AppendLine("  NOT logged (verification failed).")
    }
    return $out.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Enable Remote Mailbox'
$form.Size            = New-Object System.Drawing.Size(760, 700)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Text      = 'Connect to on-prem Exchange before running.'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size      = New-Object System.Drawing.Size(430, 20)
$lblWarn.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

# Forest context, mirroring the tenant label the other consoles now carry.
$lblCtx           = New-Object System.Windows.Forms.Label
$lblCtx.Text      = "HQ forest: $($script:HQDomain)"
$lblCtx.Location  = New-Object System.Drawing.Point(450, 10)
$lblCtx.Size      = New-Object System.Drawing.Size(282, 20)
$lblCtx.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblCtx.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblCtx)

$lblDC          = New-Object System.Windows.Forms.Label
$lblDC.Text     = 'Franchise DC (used only for accounts found in Franchise):'
$lblDC.Location = New-Object System.Drawing.Point(12, 40)
$lblDC.AutoSize = $true
$form.Controls.Add($lblDC)

$cmbDC = New-Object System.Windows.Forms.ComboBox
$cmbDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbDC.Location = New-Object System.Drawing.Point(360, 37)
$cmbDC.Size     = New-Object System.Drawing.Size(372, 22)
$script:FranchiseDCs | ForEach-Object { [void]$cmbDC.Items.Add($_) }
$idx = $cmbDC.Items.IndexOf($DefaultFranchiseDC)
if ($idx -ge 0) { $cmbDC.SelectedIndex = $idx } else { $cmbDC.SelectedIndex = 0 }
$form.Controls.Add($cmbDC)

# ONE box now -- see header. Mixed HQ/Franchise batches are fine.
$lblAcc          = New-Object System.Windows.Forms.Label
$lblAcc.Text     = 'Account name(s) -- one per line (sAMAccountName). HQ and Franchise may be mixed; each is detected automatically.'
$lblAcc.Location = New-Object System.Drawing.Point(12, 70)
$lblAcc.AutoSize = $true
$form.Controls.Add($lblAcc)

$txtAcc               = New-Object System.Windows.Forms.TextBox
$txtAcc.Multiline     = $true
$txtAcc.ScrollBars    = 'Vertical'
$txtAcc.AcceptsReturn = $true
$txtAcc.Location      = New-Object System.Drawing.Point(12, 90)
$txtAcc.Size          = New-Object System.Drawing.Size(720, 160)
$txtAcc.Font          = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtAcc)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run'
$btnRun.Location = New-Object System.Drawing.Point(12, 260)
$btnRun.Size     = New-Object System.Drawing.Size(100, 28)
$btnRun.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text     = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(120, 260)
$btnClear.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClear)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(642, 260)
$btnClose.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)

$lblResult          = New-Object System.Windows.Forms.Label
$lblResult.Text     = 'Result:'
$lblResult.Location = New-Object System.Drawing.Point(12, 296)
$lblResult.Size     = New-Object System.Drawing.Size(720, 20)
$lblResult.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblResult)

# Results stay in the window -- the original hid the form and wrote to the
# console, where output scrolled away.
$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(12, 318)
$txtOut.Size       = New-Object System.Drawing.Size(720, 330)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.BackColor  = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($txtOut)

# AcceptButton intentionally NOT set -- Enter inserts newlines in the input box.

$btnRun.Add_Click({
    $btnRun.Enabled   = $false
    $btnClose.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $accounts = Get-AccountList -Raw $txtAcc.Text
        if ($accounts.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                'No account names entered.', 'Nothing to do',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        $franDC = $cmbDC.SelectedItem

        # --- Pass 1: resolve EVERY account before changing anything. Two
        # --- reasons: the confirmation can then state exactly what will happen
        # --- and where, and missing accounts surface up front instead of
        # --- halfway through a batch.
        $lblResult.Text = 'Resolving accounts...'
        $form.Refresh()

        $resolved    = @()
        $missing     = @()
        $probeErrors = @()
        foreach ($a in $accounts) {
            $envInfo = Find-UserEnvironmentBySam -SamAccountName $a
            if ($envInfo.Errors.Count -gt 0) { $probeErrors += $envInfo.Errors }
            if ($envInfo.Environment) {
                # HQ writes go forest-wide (''); Franchise writes are pinned to
                # the DC the OPERATOR picked, not the detection DC.
                $writeDC = ''
                if ($envInfo.Environment -eq 'Franchise') { $writeDC = $franDC }
                $resolved += [PSCustomObject]@{
                    Account     = $a
                    Environment = $envInfo.Environment
                    DC          = $writeDC
                }
            } else {
                $missing += $a
            }
        }

        # --- The requested popup: accounts that exist in neither forest.
        if ($missing.Count -gt 0) {
            $extra = ''
            if ($resolved.Count -gt 0) {
                $extra = "`r`n`r`nThe other $($resolved.Count) account(s) can still be processed."
            }
            [void][System.Windows.Forms.MessageBox]::Show(
                ("Account does not yet exist:`r`n`r`n{0}`r`n`r`nNot found in HQ or Franchise AD. Create the account first, then run this again.{1}" -f `
                    ($missing -join "`r`n"), $extra),
                'Account does not yet exist',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        }

        if ($resolved.Count -eq 0) {
            $lblResult.Text = 'Nothing to do -- no account resolved.'
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('No account was found in HQ or Franchise AD:')
            foreach ($m in $missing) { [void]$sb.AppendLine("  $m") }
            if ($probeErrors.Count -gt 0) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('--- Lookup errors ---')
                foreach ($e in ($probeErrors | Select-Object -Unique)) { [void]$sb.AppendLine($e) }
            }
            $txtOut.Text = $sb.ToString()
            return
        }

        # --- Confirm, naming the environment each account resolved to, so a
        # --- wrong detection is visible BEFORE anything is written.
        $hqN   = @($resolved | Where-Object { $_.Environment -eq 'HQ' }).Count
        $franN = @($resolved | Where-Object { $_.Environment -eq 'Franchise' }).Count
        $bits  = @()
        if ($hqN)   { $bits += "$hqN HQ" }
        if ($franN) { $bits += "$franN Franchise (on $franDC)" }

        $preview = ($resolved | ForEach-Object { "  $($_.Account)  ->  $($_.Environment)" }) -join "`r`n"
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            ("Enable remote mailbox + archive for {0} account(s)?`r`n`r`n{1}`r`n`r`n{2}" -f `
                $resolved.Count, $preview, ($bits -join '   |   ')),
            'Confirm',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $lblResult.Text = 'Cancelled.'
            $txtOut.Text = 'Cancelled. No changes made.'
            return
        }

        # --- Pass 2: act. HQ first, then Franchise -- same order as the
        # --- original, keeping DC-pinned Franchise writes grouped together.
        $out = New-Object System.Text.StringBuilder
        [void]$out.AppendLine("===== $(Get-Date -Format 'yyyy-MM-dd HH:mm') =====")
        if ($missing.Count -gt 0) {
            [void]$out.AppendLine('')
            [void]$out.AppendLine('DOES NOT EXIST YET (skipped, nothing attempted):')
            foreach ($m in $missing) { [void]$out.AppendLine("  $m") }
        }

        $n = 0
        foreach ($grp in @('HQ', 'Franchise')) {
            $set = @($resolved | Where-Object { $_.Environment -eq $grp })
            if ($set.Count -eq 0) { continue }
            $onDC = ''
            if ($grp -eq 'Franchise') { $onDC = " on $franDC" }
            [void]$out.AppendLine('')
            [void]$out.AppendLine("===== $grp batch ($($set.Count) account(s))$onDC =====")
            foreach ($r in $set) {
                $n++
                $lblResult.Text = "Processing $n of $($resolved.Count): $($r.Account) [$($r.Environment)] ..."
                $form.Refresh()
                [void]$out.Append((Invoke-MailboxEnable -Account $r.Account -DC $r.DC -Mode $r.Environment))
                [void]$out.AppendLine('')
            }
        }

        if ($probeErrors.Count -gt 0) {
            [void]$out.AppendLine('--- Lookup errors (the accounts above were still processed) ---')
            foreach ($e in ($probeErrors | Select-Object -Unique)) { [void]$out.AppendLine($e) }
        }

        $lblResult.Text = "Done -- $($resolved.Count) processed, $($missing.Count) missing."
        $txtOut.Text = $out.ToString()
    }
    catch {
        $lblResult.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Enable Remote Mailbox'
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnRun.Enabled   = $true
        $btnClose.Enabled = $true
    }
})

$btnClear.Add_Click({
    $txtAcc.Text = ''
    $txtOut.Text = ''
    $lblResult.Text = 'Result:'
})

$btnClose.Add_Click({ $form.Close() })

# ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin

# Console, which opens top-left. Helper is in Common.ps1.

Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
$form.Dispose()
#endregion
