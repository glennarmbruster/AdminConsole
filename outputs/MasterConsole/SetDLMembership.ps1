#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- writes to on-premises Exchange; must not
#                                 share a process with an Exchange Online session,
#                                 because the two share cmdlet names.
<#
    SetDLMembership.ps1
    Bulk add or remove users across distribution lists -- every user in the
    user block, against every DL in the DL block.

    On-premises Exchange only. Reasoning lives in
    ScriptNotes/SetDLMembership.ps1.notes.md.

    Author   : Glenn Armbruster
    Created  : 2026-08-03
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Adds or removes a set of people across a set of distribution lists in one
# go. Built for the DL cloud migration, where membership moves in batches.
#
# Paste distribution lists in one box and users in the other. The operation
# is every user against every list: ten users across four lists is forty
# changes, and the preview tells you that number before anything runs.
#
#   * Check  -- resolves everything and shows exactly what will happen
#   * Run    -- does it, after a confirmation
#
# Anything that will not resolve is named individually and skipped, rather
# than failing the whole batch. Results are written to the audit log, one
# row per person.
#
# There is no undo. A removal is recoverable only by adding people back, and
# the audit log is what tells you who was in what.
#
# On-premises only, and it opens in its own window so it cannot be confused
# with the cloud.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Parsing ----------------------------------------------------------------
function ConvertTo-EntryList {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

    $seen = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    $kept = New-Object System.Collections.ArrayList
    foreach ($line in ($Text -split "`r`n|`r|`n|,|;")) {
        $v = "$line".Trim()
        if (-not $v) { continue }
        if ($seen.Add($v)) { [void]$kept.Add($v) }
    }
    return @($kept.ToArray())
}
#endregion

#region Resolution -------------------------------------------------------------
function Resolve-DLEntry {
    param([Parameter(Mandatory)][string]$Entry)
    try {
        $g = Get-DistributionGroup -Identity $Entry -ErrorAction Stop

        # DO NOT round-trip $g.Identity back into -Identity.
        #
        # 2026-08-06: resolving m365team@expresspros.com gave Identity
        # "# E-Mail Team" -- the group's AD CN, with no domain and no path.
        # Passing that back is not a bind, it is a FRESH forest-wide name
        # lookup, and it can land on a different object than the one just
        # resolved. The same cmdlet with the same member and the same switch
        # succeeded by hand with the address and failed here with the name.
        #
        # It surfaced as "out of the current user's write scope", which reads
        # as a permissions problem and is not one.
        #
        # Guid is unique forest-wide, is never ambiguous, and is accepted by
        # both on-premises and Exchange Online. DN and primary address are the
        # fallbacks; Identity is last resort only because it is better than
        # failing outright.
        $key = $null; $keyKind = $null
        foreach ($p in 'Guid', 'DistinguishedName', 'PrimarySmtpAddress') {
            $prop = $g.PSObject.Properties[$p]
            if ($prop -and $prop.Value) { $key = "$($prop.Value)"; $keyKind = $p; break }
        }
        if (-not $key) { $key = "$($g.Identity)"; $keyKind = 'Identity (ambiguous -- last resort)' }

        return [PSCustomObject]@{
            Ok = $true; Input = $Entry
            Id = $key; KeyKind = $keyKind
            Display = "$($g.DisplayName) <$($g.PrimarySmtpAddress)>"
            Name = "$($g.Name)"
            Error = $null
        }
    } catch {
        return [PSCustomObject]@{
            Ok = $false; Input = $Entry; Id = $null; KeyKind = $null
            Display = $null; Name = $null
            Error = "$($_.Exception.Message)"
        }
    }
}

function Resolve-UserEntry {
    param([Parameter(Mandatory)][string]$Entry)

    $tries = @($Entry)
    if ($Entry -notlike '*@*') {
        $completed = Get-TargetSmtp -Text $Entry
        if ($completed -and $completed -ne $Entry) { $tries += $completed }
    }

    $lastErr = $null
    foreach ($t in $tries) {
        try {
            $r = Get-Recipient -Identity $t -ErrorAction Stop
            return [PSCustomObject]@{
                Ok = $true; Input = $Entry
                Id = "$($r.PrimarySmtpAddress)"
                Display = "$($r.DisplayName) <$($r.PrimarySmtpAddress)>"
                Error = $null
            }
        } catch { $lastErr = "$($_.Exception.Message)" }
    }
    return [PSCustomObject]@{
        Ok = $false; Input = $Entry; Id = $null; Display = $null; Error = $lastErr
    }
}

function Get-Preflight {
    param([string[]]$DLs, [string[]]$Users, [scriptblock]$Progress)

    $rDLs   = New-Object System.Collections.ArrayList
    $rUsers = New-Object System.Collections.ArrayList

    $i = 0
    foreach ($d in $DLs) {
        $i++
        if ($Progress) { & $Progress "Checking DL $i of $($DLs.Count): $d" }
        [void]$rDLs.Add((Resolve-DLEntry -Entry $d))
    }
    $i = 0
    foreach ($u in $Users) {
        $i++
        if ($Progress) { & $Progress "Checking user $i of $($Users.Count): $u" }
        [void]$rUsers.Add((Resolve-UserEntry -Entry $u))
    }

    return [PSCustomObject]@{
        DLs   = @($rDLs.ToArray())
        Users = @($rUsers.ToArray())
    }
}

function Format-Preflight {
    param($Pre, [string]$Action)

    $okDLs   = @($Pre.DLs   | Where-Object { $_.Ok })
    $okUsers = @($Pre.Users | Where-Object { $_.Ok })
    $badDLs  = @($Pre.DLs   | Where-Object { -not $_.Ok })
    $badUsers= @($Pre.Users | Where-Object { -not $_.Ok })

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("Action     : $(if ($Action) { $Action } else { '(not selected -- choose Add or Remove)' })")
    [void]$lines.Add("DLs        : $($okDLs.Count) resolved, $($badDLs.Count) not found")
    [void]$lines.Add("Users      : $($okUsers.Count) resolved, $($badUsers.Count) not found")
    [void]$lines.Add("Operations : $($okDLs.Count * $okUsers.Count)  ($($okUsers.Count) user(s) x $($okDLs.Count) DL(s))")
    [void]$lines.Add('')

    if ($okDLs.Count -gt 0) {
        [void]$lines.Add('Distribution lists:')
        foreach ($d in $okDLs) { [void]$lines.Add("    $($d.Display)") }
        [void]$lines.Add('')
    }
    if ($okUsers.Count -gt 0) {
        [void]$lines.Add('Users:')
        foreach ($u in $okUsers) { [void]$lines.Add("    $($u.Display)") }
        [void]$lines.Add('')
    }
    if ($badDLs.Count -gt 0) {
        [void]$lines.Add("*** $($badDLs.Count) DL(s) NOT FOUND -- these will be skipped: ***")
        foreach ($d in $badDLs) { [void]$lines.Add("    $($d.Input)  --  $($d.Error)") }
        [void]$lines.Add('')
    }
    if ($badUsers.Count -gt 0) {
        [void]$lines.Add("*** $($badUsers.Count) user(s) NOT FOUND -- these will be skipped: ***")
        foreach ($u in $badUsers) { [void]$lines.Add("    $($u.Input)  --  $($u.Error)") }
    }
    return [string[]]@($lines.ToArray())
}
#endregion

#region Actions ----------------------------------------------------------------
function Set-DLMember {
    param(
        [Parameter(Mandatory)][string]$DLIdentity,
        [Parameter(Mandatory)][string]$Member,
        [Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Action
    )

    $params = @{
        Identity                        = $DLIdentity
        Member                          = $Member
        BypassSecurityGroupManagerCheck = $true
        ErrorAction                     = 'Stop'
    }
    try {
        if ($Action -eq 'Add') {
            Add-DistributionGroupMember @params
            return [PSCustomObject]@{ State = 'Done'; Text = 'ADDED' }
        }
        $params['Confirm'] = $false
        Remove-DistributionGroupMember @params
        return [PSCustomObject]@{ State = 'Done'; Text = 'REMOVED' }
    } catch {
        $msg = "$($_.Exception.Message)"
        if ($msg -match 'already a member|not a member of the group') {
            return [PSCustomObject]@{ State = 'Skipped'; Text = 'SKIP (already in desired state)' }
        }
        # "write scope" reads like a rights problem and usually is not one. It
        # is what Exchange says when the session is bound to a controller that
        # is not authoritative for the object's domain -- HQ object, franchise
        # DC. Naming that here saves the hour it cost to find the first time.
        if ($msg -match 'write scope') {
            $msg += "  [Despite the wording this is usually NOT a permissions problem. It normally means the identity passed matched a different object than intended, or the session is bound to a domain controller that cannot write this object's domain. Check the bind key logged above.]"
        }
        return [PSCustomObject]@{ State = 'Failed'; Text = "ERROR: $msg"; Error = $msg }
    }
}

function Invoke-DLMembershipChange {
    param(
        [Parameter(Mandatory)]$Pre,
        [Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Action,
        [scriptblock]$Progress
    )

    # NO -PreferredServer, deliberately.
    #
    # Pinning the shell to a franchise DC was the 2026-08-06 failure: the
    # franchises domain is a CHILD of expresspersonnel.com, so a franchise
    # controller is not authoritative for an HQ object. Reads still worked --
    # the DL resolved by name -- and only the write failed, with "out of the
    # current user's write scope", which reads like a permissions problem and
    # is not one.
    #
    # This tool has to reach HQ and Franchise in the same batch, so it lets
    # Exchange bind to a controller that can actually serve each object.
    #
    # -ViewEntireForest is asserted rather than assumed: Connect-ExSE sets it,
    # but a session established some other way may not have, and without it a
    # lookup silently sees only the local domain -- which reads as "no such
    # group" rather than as a scoping problem.
    $forestWarning = $null
    try {
        Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
    } catch {
        # Non-fatal: some RBAC roles cannot set it. Say so rather than stopping.
        $forestWarning = "Could not set -ViewEntireForest: $($_.Exception.Message). Objects outside the local domain may not resolve."
    }

    $dls   = @($Pre.DLs   | Where-Object { $_.Ok })
    $users = @($Pre.Users | Where-Object { $_.Ok })

    $log = New-Object System.Text.StringBuilder
    [void]$log.AppendLine("Action  : $Action")
    [void]$log.AppendLine("Binding : forest-wide, no domain controller pinned")
    if ($forestWarning) { [void]$log.AppendLine("WARNING : $forestWarning") }
    [void]$log.AppendLine("Started : $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    [void]$log.AppendLine("Scope   : $($users.Count) user(s) x $($dls.Count) DL(s) = $($users.Count * $dls.Count) operations")
    [void]$log.AppendLine('')
    [void]$log.AppendLine('Resolved distribution lists:')
    foreach ($d in $dls) {
        [void]$log.AppendLine("    $($d.Input)")
        [void]$log.AppendLine("        -> $($d.Display)")
        [void]$log.AppendLine("        -> AD name : $($d.Name)")
        [void]$log.AppendLine("        -> Bound by: $($d.KeyKind) = $($d.Id)")
        # An address can be a SECONDARY proxy on some other group. Exchange
        # resolves it happily and acts on that other group. Showing the input,
        # the display name and the Identity on separate lines makes a mismatch
        # impossible to read past.
        if ($d.Name -and $d.Display -and $d.Display -notlike "*$($d.Name)*") {
            [void]$log.AppendLine("        ** NOTE: AD name and display name differ. Harmless -- the bind uses $($d.KeyKind), not the name.")
        }
    }
    [void]$log.AppendLine('---------------------------------------------')

    $done = 0; $skip = 0; $fail = 0
    $n = 0
    $total = $users.Count * $dls.Count

    foreach ($u in $users) {
        [void]$log.AppendLine('')
        [void]$log.AppendLine("$($u.Display)")

        $states = New-Object System.Collections.ArrayList
        $errors = New-Object System.Collections.ArrayList

        foreach ($d in $dls) {
            $n++
            if ($Progress) { & $Progress "$n of $total -- $($u.Input) -> $($d.Input)" ([int](100 * $n / [Math]::Max(1, $total))) }

            $r = Set-DLMember -DLIdentity $d.Id -Member $u.Id -Action $Action
            [void]$log.AppendLine("    $($r.Text.PadRight(34)) $($d.Display)")
            [void]$states.Add($r.State)
            if ($r.State -eq 'Failed') { [void]$errors.Add("$($d.Input): $($r.Error)") }

            switch ($r.State) {
                'Done'    { $done++ }
                'Skipped' { $skip++ }
                'Failed'  { $fail++ }
            }
        }

        $result =
            if ($states -contains 'Failed')   { 'Failed' }
            elseif ($states -contains 'Done') { if ($states -contains 'Skipped') { 'Partial' } else { 'Success' } }
            else { 'Success' }

        $detail = "DLs: $((@($dls | ForEach-Object { $_.Input })) -join ', ')"
        if ($errors.Count -gt 0) { $detail += " -- " + ($errors -join '; ') }

        Write-CompletedLog -Action $Action -Target $u.Id -Result $result `
            -Detail $detail -Environment 'HQ'
    }

    [void]$log.AppendLine('')
    [void]$log.AppendLine('---------------------------------------------')
    [void]$log.AppendLine("Done  Succeeded: $done   Skipped: $skip   Errors: $fail")

    return [PSCustomObject]@{ Output = $log.ToString(); Failed = $fail }
}
#endregion

#region WinForms UI ------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'DL Membership -- bulk add / remove'
$form.Size          = New-Object System.Drawing.Size(840, 780)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Tag       = 'NoTheme'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size      = New-Object System.Drawing.Size(800, 18)
$lblWarn.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblWarn)

# Two questions, not one.
#
# "Is EMS available" was the only check here until 2026-08-06, and it is not
# sufficient. EMS can be available while Exchange Online is ALSO loaded in the
# same process, and the two share cmdlet names -- Get-DistributionGroup,
# Remove-DistributionGroupMember and the rest. Last import wins, so this banner
# read "Connected: on-premises Exchange" in green while every call went to the
# cloud, which then refused to modify a directory-synchronised group and said
# "out of the current user's write scope".
#
# A tool that writes on-premises has to refuse to run in a process where the
# cloud might answer instead.
$emsState = Test-EmsAvailable
$exoState = Test-ExoInProcess
$script:Blocked = $false

if ($exoState.Present) {
    $lblWarn.Text      = 'BLOCKED: Exchange Online is connected in this process. Run this tool with "Own window" ticked.'
    $lblWarn.ForeColor = [System.Drawing.Color]::Red
    $script:Blocked    = $true
} elseif ($emsState.Available) {
    $lblWarn.Text      = 'Connected: on-premises Exchange only. Safe to run.'
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
} else {
    $lblWarn.Text      = 'NOT connected to on-premises Exchange -- nothing can run. Use Connect ExSE.'
    $lblWarn.ForeColor = [System.Drawing.Color]::Red
    $script:Blocked    = $true
}

$lblDLs          = New-Object System.Windows.Forms.Label
$lblDLs.Text     = 'Distribution lists  (name or address, one per line):'
$lblDLs.Location = New-Object System.Drawing.Point(12, 36)
$lblDLs.AutoSize = $true
$form.Controls.Add($lblDLs)

$txtDLs            = New-Object System.Windows.Forms.TextBox
$txtDLs.Location   = New-Object System.Drawing.Point(12, 54)
$txtDLs.Size       = New-Object System.Drawing.Size(800, 120)
$txtDLs.Multiline  = $true
$txtDLs.ScrollBars = 'Vertical'
$txtDLs.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtDLs)

$lblUsers          = New-Object System.Windows.Forms.Label
$lblUsers.Text     = "Users  (alias or address -- @$($script:MailDomain) added if omitted, one per line):"
$lblUsers.Location = New-Object System.Drawing.Point(12, 182)
$lblUsers.AutoSize = $true
$form.Controls.Add($lblUsers)

$txtUsers            = New-Object System.Windows.Forms.TextBox
$txtUsers.Location   = New-Object System.Drawing.Point(12, 200)
$txtUsers.Size       = New-Object System.Drawing.Size(800, 120)
$txtUsers.Multiline  = $true
$txtUsers.ScrollBars = 'Vertical'
$txtUsers.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtUsers)

$grpAction          = New-Object System.Windows.Forms.GroupBox
$grpAction.Text     = 'Action  (required -- select one before running)'
$grpAction.Location = New-Object System.Drawing.Point(12, 328)
$grpAction.Size     = New-Object System.Drawing.Size(360, 48)
$form.Controls.Add($grpAction)

$radAdd          = New-Object System.Windows.Forms.RadioButton
$radAdd.Text     = 'Add to all listed DLs'
$radAdd.Location = New-Object System.Drawing.Point(14, 18)
$radAdd.AutoSize = $true
$radAdd.Checked  = $false
$grpAction.Controls.Add($radAdd)

$radRemove          = New-Object System.Windows.Forms.RadioButton
$radRemove.Text     = 'Remove from all listed DLs'
$radRemove.Location = New-Object System.Drawing.Point(180, 18)
$radRemove.AutoSize = $true
$radRemove.Checked  = $false
$grpAction.Controls.Add($radRemove)

$btnCheck          = New-Object System.Windows.Forms.Button
$btnCheck.Text     = 'Check'
$btnCheck.Location = New-Object System.Drawing.Point(390, 336)
$btnCheck.Size     = New-Object System.Drawing.Size(100, 32)
$form.Controls.Add($btnCheck)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run'
$btnRun.Location = New-Object System.Drawing.Point(498, 336)
$btnRun.Size     = New-Object System.Drawing.Size(100, 32)
$btnRun.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(712, 336)
$btnClose.Size     = New-Object System.Drawing.Size(100, 32)
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 384)
$lblStatus.Size     = New-Object System.Drawing.Size(800, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 404 -Width 800 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(12, 424)
$txtOut.Size       = New-Object System.Drawing.Size(800, 306)
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
# WordWrap ON, deliberately. It was $false, and that is what hid
# "...because the object is being synchronized from your on-premises
# organization" off the right edge of the window for three rounds of
# diagnosis. A log you have to scroll sideways to read is a log nobody reads
# sideways.
$txtOut.WordWrap   = $true
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$script:Pre     = $null
$script:PreFrom = $null

function Get-SelectedAction {
    if ($radAdd.Checked)    { return 'Add' }
    if ($radRemove.Checked) { return 'Remove' }
    return $null
}

function Get-InputSignature { "$($txtDLs.Text)" + [char]1 + "$($txtUsers.Text)" }

function Update-Preflight {
    $dls   = ConvertTo-EntryList -Text $txtDLs.Text
    $users = ConvertTo-EntryList -Text $txtUsers.Text
    if ($dls.Count -eq 0 -or $users.Count -eq 0) {
        $script:Pre = $null
        $txtOut.Text = 'Enter at least one distribution list and at least one user.'
        return $false
    }

    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $cb = { param([string]$m) $lblStatus.Text = $m; [System.Windows.Forms.Application]::DoEvents() }
        $script:Pre     = Get-Preflight -DLs $dls -Users $users -Progress $cb
        $script:PreFrom = Get-InputSignature
        $txtOut.Lines   = Format-Preflight -Pre $script:Pre -Action (Get-SelectedAction)
        $lblStatus.Text = 'Checked.'
    } finally {
        $form.Cursor = $prev
    }
    return $true
}

$btnCheck.Add_Click({
    if ($script:Blocked) {
        [System.Windows.Forms.MessageBox]::Show($lblWarn.Text, 'Cannot run in this process', 'OK', 'Error') | Out-Null
        return
    }
    try { [void](Update-Preflight) }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Check DL membership input' }
})

$btnRun.Add_Click({
    # Refuse rather than warn. A tool that can silently write to the wrong
    # Exchange must not be one confirmation dialog away from doing it.
    if ($script:Blocked) {
        [System.Windows.Forms.MessageBox]::Show(
            "$($lblWarn.Text)`r`n`r`n$($exoState.Detail)`r`n`r`nExchange Online and on-premises Exchange share cmdlet names. With both loaded in one process there is no way to be certain which one answers, so this tool will not write.`r`n`r`nRelaunch it from the console with 'Own window' ticked, which gives it a clean process connected only to on-premises Exchange.",
            'Cannot run in this process', 'OK', 'Error') | Out-Null
        return
    }
    $btnRun.Enabled = $false; $btnCheck.Enabled = $false
    try {
        $action = Get-SelectedAction
        if (-not $action) {
            [void][System.Windows.Forms.MessageBox]::Show(
                'Select Add or Remove before running.', 'Action Required',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $ems = Test-EmsAvailable
        if (-not $ems.Available) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Not connected to on-premises Exchange.`n`n$($ems.Detail)", 'Not Connected',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if (-not $script:Pre -or (Get-InputSignature) -ne $script:PreFrom) {
            if (-not (Update-Preflight)) { return }
        }

        $okDLs   = @($script:Pre.DLs   | Where-Object { $_.Ok })
        $okUsers = @($script:Pre.Users | Where-Object { $_.Ok })
        if ($okDLs.Count -eq 0 -or $okUsers.Count -eq 0) {
            $txtOut.Text = 'Nothing resolved -- see the check above. Nothing was written.'
            return
        }

        $ops     = $okDLs.Count * $okUsers.Count
        $badDLs  = @($script:Pre.DLs   | Where-Object { -not $_.Ok }).Count
        $badUsrs = @($script:Pre.Users | Where-Object { -not $_.Ok }).Count

        $msg = "$action $($okUsers.Count) user(s) across $($okDLs.Count) DL(s) = $ops operations."
        if ($badDLs -or $badUsrs) {
            $msg += "`n`nSkipping $badDLs unresolved DL(s) and $badUsrs unresolved user(s)."
        }
        if ($ops -ge 100) { $msg += "`n`nThat is a large batch. Check the list above before continuing." }
        $msg += "`n`nContinue?"

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            $msg, "Confirm $action",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $txtOut.Text = 'Cancelled.'
            return
        }

        $prev = $form.Cursor
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'DL membership'
            $result = Invoke-DLMembershipChange -Pre $script:Pre -Action $action -Progress $progress
            Complete-ProgressReporter -Activity 'DL membership' -Bar $pbar
            $txtOut.Text = $result.Output
            $lblStatus.Text = if ($result.Failed -gt 0) { "Finished with $($result.Failed) error(s)." } else { 'Finished.' }
        } finally {
            $form.Cursor = $prev
            Complete-ProgressReporter -Activity 'DL membership'
        }
    } catch {
        Show-HandlerError -ErrorRecord $_ -Context 'DL membership change'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'DL membership change'
    } finally {
        $btnRun.Enabled = $true; $btnCheck.Enabled = $true
    }
})

$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Consistent '?' in the top-right of every tool. Added after the theme so
# it is not restyled into a normal button, and BringToFront'd inside the
# helper because WinForms draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Membership' | Out-Null

[void]$form.ShowDialog()
#endregion
