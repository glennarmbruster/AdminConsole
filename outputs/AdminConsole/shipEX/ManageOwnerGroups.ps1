#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- writes to on-premises Exchange; must not
#                                 share a process with an Exchange Online session,
#                                 because the two share cmdlet names.
<#
    ManageOwnerGroups.ps1
    Adds or removes franchise owners from the FranchiseOwners distribution
    groups -- the US group plus the alphabetical sub-group for their surname.

    On-premises Exchange only (Franchise). Reasoning lives in
    ScriptNotes/ManageOwnerGroups.ps1.notes.md.

    Author   : Glenn Armbruster
    Created  : 2026-07-01
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Adds or removes franchise owners from the FranchiseOwners distribution groups.
#
# Each owner belongs in two: the US-wide group, and the alphabetical sub-group
# for their surname. Doing one and forgetting the other is the mistake this
# exists to prevent, so both are handled together.
#
# On-premises Exchange only, franchise groups only. It opens in its own window
# so it cannot be confused with the cloud.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
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

#region Constants --------------------------------------------------------------
$script:OwnerGroupStemUS = 'FranchiseOwnersUS'
$script:OwnerGroupBands  = @(
    [PSCustomObject]@{ Stem = 'FranchiseOwnersA-F'; Pattern = '^[A-F]$' }
    [PSCustomObject]@{ Stem = 'FranchiseOwnersG-L'; Pattern = '^[G-L]$' }
    [PSCustomObject]@{ Stem = 'FranchiseOwnersM-R'; Pattern = '^[M-R]$' }
    [PSCustomObject]@{ Stem = 'FranchiseOwnersS-Z'; Pattern = '^[S-Z]$' }
)

$script:MaxMembersPerRun = 10
#endregion

#region Parsing ----------------------------------------------------------------
function Get-OwnerGroupAddress {
    param([Parameter(Mandatory)][string]$Stem)
    return (Get-TargetSmtp -Text $Stem)
}

function Get-AlphaGroupStem {
    param([string]$LocalPart)

    $parts = "$LocalPart".Split('.')
    if ($parts.Count -lt 2) { return $null }
    $lastName = $parts[-1]
    if (-not $lastName) { return $null }

    $firstLetter = $lastName[0].ToString().ToUpper()
    foreach ($band in $script:OwnerGroupBands) {
        if ($firstLetter -match $band.Pattern) { return $band.Stem }
    }
    # Non A-Z initial: an accented surname, or a local part starting with a
    # digit. There is no band for it, and inventing one would be wrong.
    return $null
}

function ConvertTo-OwnerList {
    param([string]$Text)

    $result = [PSCustomObject]@{ Members = @(); Dropped = @() }
    if ([string]::IsNullOrWhiteSpace($Text)) { return $result }

    $lines = @($Text -split "`r`n|`r|`n|,|;" | Where-Object { $_.Trim() -ne '' })
    $seen  = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)

    $kept    = New-Object System.Collections.ArrayList
    $dropped = New-Object System.Collections.ArrayList

    foreach ($line in $lines) {
        $local = ($line.Trim() -replace '@.*$', '').ToLower().Trim()
        if (-not $local) { continue }
        if (-not $seen.Add($local)) { continue }
        if ($kept.Count -ge $script:MaxMembersPerRun) {
            [void]$dropped.Add($local)
            continue
        }
        [void]$kept.Add($local)
    }

    $result.Members = @($kept.ToArray())
    $result.Dropped = @($dropped.ToArray())
    return $result
}

function Get-ParsedPreview {
    param($Parsed, [string]$Action)

    $members = @($Parsed.Members)
    $lines   = New-Object System.Collections.ArrayList

    $actionDisplay = if ($Action) { $Action } else { '(not selected -- choose Add or Remove)' }
    [void]$lines.Add("Action  : $actionDisplay")
    [void]$lines.Add("Members : $($members.Count)")
    [void]$lines.Add("Domain  : @$($(Get-MailDomain))")
    [void]$lines.Add('')

    $usGroup = Get-OwnerGroupAddress -Stem $script:OwnerGroupStemUS

    foreach ($m in $members) {
        $stem = Get-AlphaGroupStem -LocalPart $m
        $agDisplay = if ($stem) {
            Get-OwnerGroupAddress -Stem $stem
        } else {
            "UNKNOWN -- cannot determine surname from '$m'"
        }
        [void]$lines.Add((Get-TargetSmtp -Text $m))
        [void]$lines.Add("    US group   : $usGroup")
        [void]$lines.Add("    Alpha group: $agDisplay")
    }

    if ($members.Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add("Total operations: $($members.Count * 2)  ($($members.Count) member(s) x 2 groups each)")

        $noAlpha = @($members | Where-Object { -not (Get-AlphaGroupStem -LocalPart $_) })
        if ($noAlpha.Count -gt 0) {
            [void]$lines.Add("Note: $($noAlpha.Count) member(s) have no alpha group -- only the US group will be attempted.")
        }
    }

    if (@($Parsed.Dropped).Count -gt 0) {
        [void]$lines.Add('')
        [void]$lines.Add("*** $(@($Parsed.Dropped).Count) entr(y/ies) BEYOND THE LIMIT OF $($script:MaxMembersPerRun) AND WILL NOT BE PROCESSED: ***")
        foreach ($d in @($Parsed.Dropped)) { [void]$lines.Add("    $d") }
        [void]$lines.Add('Run them as a second batch.')
    }

    return [string[]]@($lines.ToArray())
}
#endregion

#region Actions ----------------------------------------------------------------
function Set-OwnerGroupMembership {
    param(
        [Parameter(Mandatory)][string]$GroupAddress,
        [Parameter(Mandatory)][string]$MemberAddress,
        [Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Action
    )

    $params = @{
        Identity                        = $GroupAddress
        Member                          = $MemberAddress
        BypassSecurityGroupManagerCheck = $true
        ErrorAction                     = 'Stop'
    }

    try {
        if ($Action -eq 'Add') {
            Add-DistributionGroupMember @params
            return [PSCustomObject]@{ State = 'Done'; Text = "ADDED   -> $GroupAddress" }
        }
        $params['Confirm'] = $false
        Remove-DistributionGroupMember @params
        return [PSCustomObject]@{ State = 'Done'; Text = "REMOVED -> $GroupAddress" }
    } catch {
        Write-FailedLog -Action "$Action DistributionGroupMember" -Target $GroupAddress -ErrorRecord $_ -Detail "member: $MemberAddress"
        $msg = "$($_.Exception.Message)"
        if ($msg -match 'already a member|not a member of the group') {
            return [PSCustomObject]@{ State = 'Skipped'; Text = "SKIP    -> $GroupAddress (already in desired state)" }
        }
        return [PSCustomObject]@{ State = 'Failed'; Text = "ERROR   -> $GroupAddress : $msg"; Error = $msg }
    }
}

function Invoke-OwnerGroupAction {
    param(
        [Parameter(Mandatory)][string[]]$Members,
        [Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Action,
        [string]$DC
    )

    # Pin only a value Tenant.json names as a controller. A domain reaches
    # here whenever the directory has more than one DC, and -PreferredServer
    # rejects it with error 0x52.
    $pinDC = Get-PinnableDomainController -Candidate $DC
    if ($pinDC) { Set-ADServerSettings -PreferredServer $pinDC -ErrorAction Stop }

    $usGroup = Get-OwnerGroupAddress -Stem $script:OwnerGroupStemUS

    $log = New-Object System.Text.StringBuilder
    [void]$log.AppendLine("Action  : $Action")
    [void]$log.AppendLine("DC      : $(if ($DC) { $DC } else { '(nearest -- not pinned)' })")
    [void]$log.AppendLine("Started : $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    [void]$log.AppendLine('---------------------------------------------')

    $doneCount = 0; $skipCount = 0; $failCount = 0

    foreach ($member in $Members) {
        $email  = Get-TargetSmtp -Text $member
        $stem   = Get-AlphaGroupStem -LocalPart $member

        [void]$log.AppendLine('')
        [void]$log.AppendLine("Member: $email")

        $states = New-Object System.Collections.ArrayList
        $errors = New-Object System.Collections.ArrayList

        $targets = @($usGroup)
        if ($stem) {
            $targets += (Get-OwnerGroupAddress -Stem $stem)
        } else {
            [void]$log.AppendLine("  SKIP    -> alpha group (cannot determine surname from '$member')")
            [void]$states.Add('Skipped')
        }

        foreach ($grp in $targets) {
            $r = Set-OwnerGroupMembership -GroupAddress $grp -MemberAddress $email -Action $Action
            [void]$log.AppendLine("  $($r.Text)")
            [void]$states.Add($r.State)
            if ($r.State -eq 'Failed') { [void]$errors.Add($r.Error) }
        }

        $memberResult =
            if ($states -contains 'Failed')  { 'Failed' }
            elseif ($states -contains 'Done') { if ($states -contains 'Skipped') { 'Partial' } else { 'Success' } }
            else { 'Success' }   # everything already in the desired state

        $detailParts = @("groups: $($targets -join ', ')")
        if (-not $stem) { $detailParts += 'no alpha group resolved' }
        if ($errors.Count -gt 0) { $detailParts += ($errors -join '; ') }

        Write-CompletedLog -Action $Action -Target $email -Result $memberResult `
            -Detail ("Franchise Owner groups -- " + ($detailParts -join ' -- ')) `
            -Environment 'Franchise'

        foreach ($s in $states) {
            switch ($s) {
                'Done'    { $doneCount++ }
                'Skipped' { $skipCount++ }
                'Failed'  { $failCount++ }
            }
        }
    }

    [void]$log.AppendLine('')
    [void]$log.AppendLine('---------------------------------------------')
    [void]$log.AppendLine("Done  Succeeded: $doneCount   Skipped: $skipCount   Errors: $failCount")

    return [PSCustomObject]@{
        Output = $log.ToString()
        Failed = $failCount
        Done   = $doneCount
    }
}
#endregion

#region WinForms UI ------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Franchise Owner Group Manager'
$form.Size          = New-Object System.Drawing.Size(760, 720)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

# NoTheme-tagged: this is a live connection fact, and its red/green has to
# survive whatever palette is active.
$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Tag       = 'NoTheme'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size      = New-Object System.Drawing.Size(720, 18)
$lblWarn.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblWarn)

# EMS being available is not the same as EMS being the one that answers.
# Exchange Online shares these cmdlet names, so both loaded in one process
# means the cloud may win. See Test-ExoInProcess in Common.ps1.
$emsState = Test-EmsAvailable
$exoState = Test-ExoInProcess
if ($exoState.Present) {
    $lblWarn.Text      = 'BLOCKED: Exchange Online is connected in this process. Relaunch from the console -- it runs this tool in its own window.'
    $lblWarn.ForeColor = [System.Drawing.Color]::Red
} elseif ($emsState.Available) {
    $lblWarn.Text      = 'Connected: on-premises Exchange only. Franchise groups only -- Exchange Online is not used by this tool.'
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
} else {
    $lblWarn.Text      = 'NOT connected to on-premises Exchange -- open an Exchange Management Shell. Nothing can run.'
    $lblWarn.ForeColor = [System.Drawing.Color]::Red
}

$lblDC          = New-Object System.Windows.Forms.Label
$lblDC.Text     = 'Domain Controller:'
$lblDC.Location = New-Object System.Drawing.Point(12, 38)
$lblDC.AutoSize = $true
$form.Controls.Add($lblDC)

$cmbDC               = New-Object System.Windows.Forms.ComboBox
$cmbDC.Location      = New-Object System.Drawing.Point(12, 56)
$cmbDC.Size          = New-Object System.Drawing.Size(460, 22)
$cmbDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
# Sourced from Common.ps1, not a second copy. The duplicated list was already
# drifting -- a DC added to the console would not have appeared here.
foreach ($dc in (Get-TenantDomainController)) { [void]$cmbDC.Items.Add($dc) }
if ($cmbDC.Items.Count -gt 0) {
    # Get-TenantDirectoryServer can return a DOMAIN, which is never in this
    # list of controllers -- so IndexOf gave -1 and the selection silently fell
    # back to item 0, which is whichever directory happens to sort first. Ask
    # for a real controller instead.
    $preferred = if ($DomainController) { $DomainController } else { @(Get-TenantDomainController)[0] }
    $idx = $cmbDC.Items.IndexOf($preferred)
    $cmbDC.SelectedIndex = if ($idx -ge 0) { $idx } else { 0 }
}
$lblDC.Visible    = $false
$cmbDC.Visible    = $false
# DC selector hidden 2026-08-03: the toolkit binds to the default franchise
# controller from Common.ps1 and nobody was choosing another one. The control
# is kept rather than deleted so it still supplies that default to the code
# below -- .Visible is the whole change, and reversing it is one word.
$form.Controls.Add($cmbDC)

$grpAction          = New-Object System.Windows.Forms.GroupBox
$grpAction.Text     = 'Action  (required -- select one before running)'
$grpAction.Location = New-Object System.Drawing.Point(12, 88)
$grpAction.Size     = New-Object System.Drawing.Size(460, 65)
$form.Controls.Add($grpAction)

$radAdd          = New-Object System.Windows.Forms.RadioButton
$radAdd.Text     = 'Add to groups'
$radAdd.Location = New-Object System.Drawing.Point(14, 18)
$radAdd.AutoSize = $true
$radAdd.Checked  = $false
$grpAction.Controls.Add($radAdd)

$radRemove          = New-Object System.Windows.Forms.RadioButton
$radRemove.Text     = 'Remove from groups'
$radRemove.Location = New-Object System.Drawing.Point(14, 38)
$radRemove.AutoSize = $true
$radRemove.Checked  = $false
$grpAction.Controls.Add($radRemove)

$lblMembers          = New-Object System.Windows.Forms.Label
$lblMembers.Text     = "Owners/Members  (firstname.lastname -- @$($(Get-MailDomain)) optional, one per line, up to $($script:MaxMembersPerRun)):"
$lblMembers.Location = New-Object System.Drawing.Point(12, 163)
$lblMembers.AutoSize = $true
$form.Controls.Add($lblMembers)

$txtMembers            = New-Object System.Windows.Forms.TextBox
$txtMembers.Location   = New-Object System.Drawing.Point(12, 181)
$txtMembers.Size       = New-Object System.Drawing.Size(720, 110)
$txtMembers.Multiline  = $true
$txtMembers.ScrollBars = 'Vertical'
$txtMembers.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtMembers)

$lblParsed          = New-Object System.Windows.Forms.Label
$lblParsed.Text     = 'Parsed Preview:'
$lblParsed.Location = New-Object System.Drawing.Point(12, 301)
$lblParsed.AutoSize = $true
$form.Controls.Add($lblParsed)

$txtParsed            = New-Object System.Windows.Forms.TextBox
$txtParsed.Location   = New-Object System.Drawing.Point(12, 319)
$txtParsed.Size       = New-Object System.Drawing.Size(720, 128)
$txtParsed.Multiline  = $true
$txtParsed.ScrollBars = 'Vertical'
$txtParsed.ReadOnly   = $true
$txtParsed.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtParsed)

$btnParse          = New-Object System.Windows.Forms.Button
$btnParse.Text     = 'Parse'
$btnParse.Location = New-Object System.Drawing.Point(12, 458)
$btnParse.Size     = New-Object System.Drawing.Size(80, 30)
$form.Controls.Add($btnParse)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run'
$btnRun.Location = New-Object System.Drawing.Point(100, 458)
$btnRun.Size     = New-Object System.Drawing.Size(80, 30)
$btnRun.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(188, 458)
$btnClose.Size     = New-Object System.Drawing.Size(80, 30)
$form.Controls.Add($btnClose)

$lblOut          = New-Object System.Windows.Forms.Label
$lblOut.Text     = 'Result:'
$lblOut.Location = New-Object System.Drawing.Point(12, 499)
$lblOut.AutoSize = $true
$form.Controls.Add($lblOut)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(12, 517)
$txtOut.Size       = New-Object System.Drawing.Size(720, 160)
$txtOut.Multiline  = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.ReadOnly   = $true
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$script:ParsedResult = $null
$script:ParsedFrom   = $null

function Get-SelectedAction {
    if ($radAdd.Checked)    { return 'Add' }
    if ($radRemove.Checked) { return 'Remove' }
    return $null
}

function Update-ParsedPreview {
    param([string]$Action)
    $script:ParsedResult = ConvertTo-OwnerList -Text $txtMembers.Text
    $script:ParsedFrom   = $txtMembers.Text
    $txtParsed.Lines     = Get-ParsedPreview -Parsed $script:ParsedResult -Action $Action
}

$btnParse.Add_Click({
    try {
        Update-ParsedPreview -Action (Get-SelectedAction)
    } catch {
        Show-HandlerError -ErrorRecord $_ -Context 'Parse owner list'
    }
})

$btnRun.Add_Click({
    $btnRun.Enabled   = $false
    $btnParse.Enabled = $false
    try {
        $action = Get-SelectedAction
        if (-not $action) {
            $null = Show-ToolDialog(
                'Select Add or Remove before running.',
                'Action Required',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            $txtOut.Text = 'Cancelled -- no action selected.'
            return
        }

        $exo = Test-ExoInProcess
        if ($exo.Present) {
            $null = Show-ToolDialog(
                "Exchange Online is connected in this process.`n`n$($exo.Detail)`n`nThe two share cmdlet names, so there is no way to be certain which one would answer. Relaunch this tool from the console, which gives it a clean on-premises process.",
                'Cannot run in this process',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            $txtOut.Text = 'Cancelled -- Exchange Online is connected in this process.'
            return
        }
        $ems = Test-EmsAvailable
        if (-not $ems.Available) {
            $null = Show-ToolDialog(
                "Not connected to on-premises Exchange.`n`n$($ems.Detail)",
                'Not Connected',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            $txtOut.Text = 'Cancelled -- not connected to on-premises Exchange.'
            return
        }

        if (-not $script:ParsedResult -or $txtMembers.Text -ne $script:ParsedFrom) {
            Update-ParsedPreview -Action $action
        }

        $members = @($script:ParsedResult.Members)
        if ($members.Count -eq 0) {
            $txtOut.Text = 'No members found. Enter at least one owner and try again.'
            return
        }

        $dropped     = @($script:ParsedResult.Dropped)
        $memberLines = ($members | ForEach-Object { "  $(Get-TargetSmtp -Text $_)" }) -join "`n"
        $droppedNote = if ($dropped.Count -gt 0) {
            "`n`nNOT INCLUDED (over the limit of $($script:MaxMembersPerRun)): $($dropped -join ', ')"
        } else { '' }

        $confirm = Show-ToolDialog(
            "$action $($members.Count) member(s) x 2 groups each = $($members.Count * 2) operations.`n`nMembers:`n$memberLines$droppedNote`n`nContinue?",
            "Confirm $action",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $txtOut.Text = 'Cancelled.'
            return
        }

        $prev        = $form.Cursor
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $result      = Invoke-OwnerGroupAction -Members $members -Action $action -DC $cmbDC.SelectedItem
            $txtOut.Text = $result.Output
        } finally {
            $form.Cursor = $prev
        }
    } catch {
        Show-HandlerError -ErrorRecord $_ -Context "$(Get-SelectedAction) Franchise Owner groups"
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Franchise Owner groups'
    } finally {
        $btnRun.Enabled   = $true
        $btnParse.Enabled = $true
    }
})

$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Owner Groups' | Out-Null

[void]$form.ShowDialog()
#endregion
