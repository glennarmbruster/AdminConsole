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
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'Group Editor -- Enter Target'
if (-not $target) { return }

$groupInfo = Resolve-Target -Smtp $target.Smtp -ADServer $target.ADServer
if ($groupInfo.Type -ne 'UnifiedGroup' -and $groupInfo.Type -ne 'OnpremGroup') {
    $detail = if ($groupInfo.ResolutionErrors.Count -gt 0) { ($groupInfo.ResolutionErrors -join "`n`n") } else { "No matching Microsoft 365 Group or on-prem group was found for '$($target.Smtp)'." }
    [void][System.Windows.Forms.MessageBox]::Show(
        "'$($target.Smtp)' did not resolve to a group (Type: $($groupInfo.Type)).`n`n$detail",
        'Group Editor -- Not a Group', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

#region Actions - group (moved unchanged from Mailbox-Group-Inspector.ps1 lines 985-1105)
function Get-GroupState {
    param([string]$Type, [string]$Smtp, [string]$ADGroupDN, [string]$ADServer)
    $state = [ordered]@{
        DisplayName = $null
        PrimarySmtp = $Smtp
        GroupType   = $null
        Hidden      = $null
        Owners      = @()
        Members     = @()
        Errors      = @()
    }
    if ($Type -eq 'UnifiedGroup') {
        try {
            $g = Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
            $state.DisplayName = $g.DisplayName
            $state.GroupType   = 'Microsoft 365 Group (Unified)'
            $state.Hidden      = $g.HiddenFromAddressListsEnabled
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroup (core properties)')
        }
        try {
            $owners = Invoke-EXOCommand -Name 'Get-UnifiedGroupLinks' -Parameters @{ Identity = $Smtp; LinkType = 'Owners'; ErrorAction = 'Stop' }
            $state.Owners = @($owners | ForEach-Object { $_.PrimarySmtpAddress })
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroupLinks -LinkType Owners')
        }
        try {
            $members = Invoke-EXOCommand -Name 'Get-UnifiedGroupLinks' -Parameters @{ Identity = $Smtp; LinkType = 'Members'; ErrorAction = 'Stop' }
            $state.Members = @($members | ForEach-Object { $_.PrimarySmtpAddress })
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroupLinks -LinkType Members')
        }
    }
    elseif ($Type -eq 'OnpremGroup') {
        try {
            $g = Get-ADGroup -Identity $ADGroupDN -Server $ADServer `
                    -Properties mail, ManagedBy, GroupCategory, GroupScope, msExchHideFromAddressLists -ErrorAction Stop
            $state.DisplayName = $g.Name
            $state.PrimarySmtp = $g.mail
            $state.GroupType   = if ($g.GroupCategory -eq 'Security') { 'Mail-Enabled Security Group (on-prem, synced)' }
                                  else { 'Distribution Group (on-prem, synced)' }
            $state.Hidden      = [bool]$g.msExchHideFromAddressLists
            if ($g.ManagedBy) {
                try {
                    $owner = Get-ADUser -Identity $g.ManagedBy -Server $ADServer -ErrorAction Stop
                    $state.Owners = @($owner.SamAccountName)
                } catch {
                    $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADUser (resolving ManagedBy owner)')
                }
            }
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADGroup (core properties)')
        }
        try {
            $members = Get-ADGroupMember -Identity $ADGroupDN -Server $ADServer -ErrorAction Stop
            $state.Members = @($members | ForEach-Object { $_.SamAccountName })
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADGroupMember')
        }
    }
    return [PSCustomObject]$state
}
function Format-GroupState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading (Group) ---")
    [void]$sb.AppendLine("  Display Name        : $($State.DisplayName)")
    [void]$sb.AppendLine("  Primary SMTP        : $($State.PrimarySmtp)")
    [void]$sb.AppendLine("  Group Type          : $($State.GroupType)  (read-only -- see script header)")
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
function Set-OnpremOwner {
    param([string]$ADGroupDN, [string]$Grantee, [string]$ADServer)
    Set-ADGroup -Identity $ADGroupDN -ManagedBy $Grantee -Server $ADServer -ErrorAction Stop
}
function Add-OnpremMember {
    param([string]$ADGroupDN, [string]$Grantee, [string]$ADServer)
    Add-ADGroupMember -Identity $ADGroupDN -Members $Grantee -Server $ADServer -ErrorAction Stop
}
function Remove-OnpremMember {
    param([string]$ADGroupDN, [string]$Grantee, [string]$ADServer)
    Remove-ADGroupMember -Identity $ADGroupDN -Members $Grantee -Server $ADServer -Confirm:$false -ErrorAction Stop
}
#endregion

#region NEW - wrapper window for group delegation (see .DESCRIPTION above)
function Show-GroupEditorForm {
    param([string]$Smtp, [string]$Type, [string]$ADGroupDN, [string]$ADServer, [string]$Environment)
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
            $state = Get-GroupState -Type $Type -Smtp $Smtp -ADGroupDN $ADGroupDN -ADServer $ADServer
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
            $before = Get-GroupState -Type $Type -Smtp $Smtp -ADGroupDN $ADGroupDN -ADServer $ADServer
            $out = New-Object System.Text.StringBuilder
            [void]$out.Append((Format-GroupState -Heading "BEFORE" -State $before))
            & $Action
            [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
            $after = Get-GroupState -Type $Type -Smtp $Smtp -ADGroupDN $ADGroupDN -ADServer $ADServer
            [void]$out.Append((Format-GroupState -Heading "AFTER" -State $after))
            $txtOut.Text = $out.ToString()
            Write-CompletedLog -Task "$Verb $Smtp ($Environment; grantee: $grantee)"
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
            else { Set-OnpremOwner -ADGroupDN $ADGroupDN -Grantee $grantee -ADServer $ADServer }
        }
    })
    $btnOwnerRemove.Add_Click({
        Invoke-GroupEdit -ActionLabel "Remove Owner" -Verb "Removed Owner" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($Type -eq 'UnifiedGroup') { Remove-UnifiedOwner -Smtp $Smtp -Grantee $grantee }
            else { throw "On-prem groups have a single ManagedBy owner. Use 'Set/Add Owner' with a new value to replace it -- there is no separate remove action for the AD ManagedBy attribute." }
        }
    })
    $btnMemberAdd.Add_Click({
        Invoke-GroupEdit -ActionLabel "Add Member" -Verb "Added Member" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($Type -eq 'UnifiedGroup') { Add-UnifiedMember -Smtp $Smtp -Grantee $grantee }
            else { Add-OnpremMember -ADGroupDN $ADGroupDN -Grantee $grantee -ADServer $ADServer }
        }
    })
    $btnMemberRemove.Add_Click({
        Invoke-GroupEdit -ActionLabel "Remove Member" -Verb "Removed Member" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($Type -eq 'UnifiedGroup') { Remove-UnifiedMember -Smtp $Smtp -Grantee $grantee }
            else { Remove-OnpremMember -ADGroupDN $ADGroupDN -Grantee $grantee -ADServer $ADServer }
        }
    })
    $btnEditClose.Add_Click({ $editForm.Close() })
    Set-ConsoleTheme -Form $editForm
    Set-FormScreenCorner -Form $editForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $editForm -ScriptPath $PSCommandPath -Title 'Group Editor' | Out-Null

    [void]$editForm.ShowDialog()
}
#endregion

Show-GroupEditorForm -Smtp $target.Smtp -Type $groupInfo.Type -ADGroupDN $groupInfo.ADGroupDN -ADServer $target.ADServer -Environment $target.Environment
