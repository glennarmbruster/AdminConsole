#Requires -Version 5.1
<#
    GroupEditor.ps1
    Standalone Group Editor (owner / member delegation for M365 Groups and
    on-prem synced groups) tool -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# View and change who owns and who belongs to a group -- both Microsoft 365
# groups and on-premises groups synced to the cloud.
#
# Look the group up, then add or remove owners and members. Owners matter more
# than they look: a group with no owner cannot be managed by anyone except an
# admin, and nobody gets told when the last one leaves.
#
# Changes take effect immediately and are written to the audit log.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')


$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -ToolTitle 'Group Editor -- Enter Target'
if (-not $target) { return }

$groupInfo = Resolve-Target -Smtp $target.Smtp
if ($groupInfo.Type -ne 'UnifiedGroup' -and $groupInfo.Type -ne 'DistributionGroup') {
    $detail = if ($groupInfo.ResolutionErrors.Count -gt 0) { ($groupInfo.ResolutionErrors -join "`n`n") } else { "No matching Microsoft 365 group or distribution group was found for '$($target.Smtp)'." }
    [void][System.Windows.Forms.MessageBox]::Show(
        "'$($target.Smtp)' did not resolve to a group (Type: $($groupInfo.Type)).`n`n$detail",
        'Group Editor -- Not a Group', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

#region Actions - group -------------------------------------------------------
function Get-GroupState {
    <#
        Everything the editor shows about one group, either kind.

        Microsoft 365 groups and distribution groups are read by different
        cmdlets that return different property names, so they are normalised
        here into one shape -- otherwise every caller has to branch, and the
        display code drifts apart between the two.

        Errors are collected rather than thrown: a group whose members cannot
        be listed is still worth showing with its properties, and an editor
        that refuses to open teaches nobody anything.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][ValidateSet('UnifiedGroup', 'DistributionGroup')][string]$Type
    )

    $state = [PSCustomObject]@{
        DisplayName  = ''
        PrimarySmtp  = $Smtp
        GroupType    = $Type
        Hidden       = $false
        Owners       = @()
        Members      = @()
        Errors       = @()
    }

    if ($Type -eq 'UnifiedGroup') {
        try {
            $g = Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{
                Identity = $Smtp; ErrorAction = 'Stop' }
            $state.DisplayName = "$($g.DisplayName)"
            $state.PrimarySmtp = "$($g.PrimarySmtpAddress)"
            $state.GroupType   = "$($g.RecipientTypeDetails)"
            $state.Hidden      = [bool]$g.HiddenFromAddressListsEnabled
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroup')
            return $state
        }
        foreach ($link in @('Owners', 'Members')) {
            try {
                $vals = @(Invoke-EXOCommand -Name 'Get-UnifiedGroupLinks' -Parameters @{
                    Identity = $Smtp; LinkType = $link; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
                    ForEach-Object { "$($_.DisplayName) <$($_.PrimarySmtpAddress)>" })
                if ($link -eq 'Owners') { $state.Owners = $vals } else { $state.Members = $vals }
            } catch {
                $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context "Get-UnifiedGroupLinks ($link)")
            }
        }
        return $state
    }

    try {
        $g = Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop' }
        $state.DisplayName = "$($g.DisplayName)"
        $state.PrimarySmtp = "$($g.PrimarySmtpAddress)"
        $state.GroupType   = "$($g.RecipientTypeDetails)"
        $state.Hidden      = [bool]$g.HiddenFromAddressListsEnabled

        # A directory-synced group is mastered somewhere else. Saying so up
        # front is the difference between "this tool is broken" and "this
        # change has to be made at the source".
        if ($g.IsDirSynced) {
            $state.Errors += 'This group is directory-synced (IsDirSynced True). Owner and membership changes must be made where it is mastered; Exchange Online will refuse them here.'
        }
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup')
        return $state
    }

    foreach ($o in @($g.ManagedBy)) {
        try {
            $r = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
                Identity = "$o"; ErrorAction = 'Stop' }
            $state.Owners += "$($r.DisplayName) <$($r.PrimarySmtpAddress)>"
        } catch {
            $state.Owners += "$o (could not re-resolve)"
        }
    }
    try {
        $state.Members = @(Invoke-EXOCommand -Name 'Get-DistributionGroupMember' -Parameters @{
            Identity = $Smtp; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
            ForEach-Object { "$($_.DisplayName) <$($_.PrimarySmtpAddress)>  [$($_.RecipientTypeDetails)]" })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroupMember')
    }
    return $state
}
function Format-GroupState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading (Group) ---")
    [void]$sb.AppendLine("  Display Name        : $($State.DisplayName)")
    [void]$sb.AppendLine("  Primary SMTP        : $($State.PrimarySmtp)")
    [void]$sb.AppendLine("  Group Type          : $($State.GroupType)")
    [void]$sb.AppendLine("  Hidden From GAL     : $($State.Hidden)")
    [void]$sb.AppendLine("  Owner(s)            :")
    if ($State.Owners.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.Owners | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    [void]$sb.AppendLine("  Members ($($State.Members.Count)) :")
    if ($State.Members.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.Members | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    if ($State.Errors.Count -gt 0) {
        [void]$sb.AppendLine("  --- Field errors (partial data above is still valid) ---")
        $State.Errors | ForEach-Object { [void]$sb.AppendLine("  $_") }
    }
    return $sb.ToString()
}
function Add-UnifiedOwner {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Owners'; Links = $Grantee; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-UnifiedOwner {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Owners'; Links = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Add-UnifiedMember {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Members'; Links = $Grantee; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-UnifiedMember {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-UnifiedGroupLinks' -Parameters @{
        Identity = $Smtp; LinkType = 'Members'; Links = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Set-GroupOwner {
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$Owner)
    Invoke-EXOCommand -Name 'Set-DistributionGroup' -Parameters @{
        Identity = $Identity; ManagedBy = $Owner; ErrorAction = 'Stop' } | Out-Null
}
function Add-GroupMemberCloud {
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$Member)
    Invoke-EXOCommand -Name 'Add-DistributionGroupMember' -Parameters @{
        Identity = $Identity; Member = $Member; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
}
function Remove-GroupMemberCloud {
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$Member)
    Invoke-EXOCommand -Name 'Remove-DistributionGroupMember' -Parameters @{
        Identity = $Identity; Member = $Member; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
}
#endregion

#region NEW - wrapper window for group delegation (see .DESCRIPTION above)
function Show-GroupEditorForm {
    param([Parameter(Mandatory)][string]$Smtp, [Parameter(Mandatory)][string]$Type)
    $editForm = New-Object System.Windows.Forms.Form
    $editForm.Text = "Group Editor: $Smtp"
    $editForm.Size = New-Object System.Drawing.Size(780, 420)
    $editForm.StartPosition = "CenterScreen"
    $editForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true
    $txtOut.ReadOnly = $true
    $txtOut.ScrollBars = "Vertical"
    $txtOut.Location = New-Object System.Drawing.Point(12, 12)
    $txtOut.Size = New-Object System.Drawing.Size(740, 240)
    $txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
    $editForm.Controls.Add($txtOut)

    $grpGrp = New-Object System.Windows.Forms.GroupBox
    $grpGrp.Text = "Group Delegation ($Type)"
    $grpGrp.Location = New-Object System.Drawing.Point(12, 262)
    $grpGrp.Size = New-Object System.Drawing.Size(740, 90)
    $editForm.Controls.Add($grpGrp)
    $lblGrpGrantee = New-Object System.Windows.Forms.Label
    $lblGrpGrantee.Text = "Grantee (email or logon):"
    $lblGrpGrantee.Location = New-Object System.Drawing.Point(10, 22)
    $lblGrpGrantee.AutoSize = $true
    $grpGrp.Controls.Add($lblGrpGrantee)
    $txtGrpGrantee = New-Object System.Windows.Forms.TextBox
    $txtGrpGrantee.Location = New-Object System.Drawing.Point(170, 19)
    $txtGrpGrantee.Size = New-Object System.Drawing.Size(260, 22)
    $txtGrpGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grpGrp.Controls.Add($txtGrpGrantee)
    $btnOwnerSet = New-Object System.Windows.Forms.Button
    $btnOwnerSet.Text = "Set/Add Owner"
    $btnOwnerSet.Location = New-Object System.Drawing.Point(10, 54)
    $btnOwnerSet.Size = New-Object System.Drawing.Size(140, 28)
    $grpGrp.Controls.Add($btnOwnerSet)
    $btnOwnerRemove = New-Object System.Windows.Forms.Button
    $btnOwnerRemove.Text = "Remove Owner"
    $btnOwnerRemove.Location = New-Object System.Drawing.Point(156, 54)
    $btnOwnerRemove.Size = New-Object System.Drawing.Size(140, 28)
    $grpGrp.Controls.Add($btnOwnerRemove)
    $btnMemberAdd = New-Object System.Windows.Forms.Button
    $btnMemberAdd.Text = "Add Member"
    $btnMemberAdd.Location = New-Object System.Drawing.Point(302, 54)
    $btnMemberAdd.Size = New-Object System.Drawing.Size(140, 28)
    $grpGrp.Controls.Add($btnMemberAdd)
    $btnMemberRemove = New-Object System.Windows.Forms.Button
    $btnMemberRemove.Text = "Remove Member"
    $btnMemberRemove.Location = New-Object System.Drawing.Point(448, 54)
    $btnMemberRemove.Size = New-Object System.Drawing.Size(140, 28)
    $grpGrp.Controls.Add($btnMemberRemove)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 360)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $editForm.Controls.Add($btnRefresh)
    $btnEditClose = New-Object System.Windows.Forms.Button
    $btnEditClose.Text = "Close"
    $btnEditClose.Location = New-Object System.Drawing.Point(652, 360)
    $btnEditClose.Size = New-Object System.Drawing.Size(100, 28)
    $editForm.Controls.Add($btnEditClose)

    function Update-StateDisplay {
        try {
            $state = Get-GroupState -Smtp $Smtp -Type $Type
            $txtOut.Text = Format-GroupState -Heading "CURRENT STATE" -State $state
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Group Editor'
        }
    }

    function Invoke-GroupEdit {
        param([string]$ActionLabel, [string]$Verb, [scriptblock]$Action)
        $grantee = $txtGrpGrantee.Text.Trim()
        if (-not $grantee) { $txtOut.Text = "Enter a grantee (email or logon) first."; return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "$ActionLabel for '$grantee' on group '$Smtp'?",
            "Confirm $ActionLabel",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $grpGrp.Enabled = $false
        try {
            $before = Get-GroupState -Smtp $Smtp -Type $Type
            $out = New-Object System.Text.StringBuilder
            [void]$out.Append((Format-GroupState -Heading "BEFORE" -State $before))
            & $Action
            [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
            $after = Get-GroupState -Smtp $Smtp -Type $Type
            [void]$out.Append((Format-GroupState -Heading "AFTER" -State $after))
            $txtOut.Text = $out.ToString()
            Write-CompletedLog -Task "$Verb $Smtp (grantee: $grantee)"
        }
        catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context $ActionLabel
        }
        finally {
            $grpGrp.Enabled = $true
        }
    }

    Update-StateDisplay
    $btnRefresh.Add_Click({ Update-StateDisplay })
    $btnOwnerSet.Add_Click({
        Invoke-GroupEdit -ActionLabel "Set/Add Owner" -Verb "Set Owner" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($Type -eq 'UnifiedGroup') { Add-UnifiedOwner -Smtp $Smtp -Grantee $grantee }
            else { Set-GroupOwner -Identity $Smtp -Owner $grantee }
        }
    })
    $btnOwnerRemove.Add_Click({
        Invoke-GroupEdit -ActionLabel "Remove Owner" -Verb "Removed Owner" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($Type -eq 'UnifiedGroup') { Remove-UnifiedOwner -Smtp $Smtp -Grantee $grantee }
            else { throw "Set-DistributionGroup -ManagedBy replaces the whole owner list rather than removing one entry. Use 'Set/Add Owner' with the owners you want to keep." }
        }
    })
    $btnMemberAdd.Add_Click({
        Invoke-GroupEdit -ActionLabel "Add Member" -Verb "Added Member" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($Type -eq 'UnifiedGroup') { Add-UnifiedMember -Smtp $Smtp -Grantee $grantee }
            else { Add-GroupMemberCloud -Identity $Smtp -Member $grantee }
        }
    })
    $btnMemberRemove.Add_Click({
        Invoke-GroupEdit -ActionLabel "Remove Member" -Verb "Removed Member" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($Type -eq 'UnifiedGroup') { Remove-UnifiedMember -Smtp $Smtp -Grantee $grantee }
            else { Remove-GroupMemberCloud -Identity $Smtp -Member $grantee }
        }
    })
    $btnEditClose.Add_Click({ $editForm.Close() })
    Set-ConsoleTheme -Form $editForm
    Add-ToolHelpButton -Form $editForm -ScriptPath $PSCommandPath -Title 'Group Editor' | Out-Null
    Set-FormScreenCorner -Form $editForm -Corner 'BottomLeft'
    [void]$editForm.ShowDialog()
}
#endregion

Show-GroupEditorForm -Smtp $target.Smtp -Type $groupInfo.Type
