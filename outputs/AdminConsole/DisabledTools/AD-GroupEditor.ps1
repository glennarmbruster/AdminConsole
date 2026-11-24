#Requires -Version 5.1
# TOOLKIT-VARIANT: OnPremAD
# ---------------------------------------------------------------------------
# AD- PREFIX: this is the ACTIVE DIRECTORY version of GroupEditor.
#
# It reads and writes raw AD attributes -- mailNickname, proxyAddresses,
# targetAddress -- through Get-ADUser/Get-ADGroup, and needs a reachable domain
# controller. GroupEditor.ps1 beside it does the same job against Exchange Online
# properties (Alias, EmailAddresses) and needs no directory at all.
#
# TWO FILES RATHER THAN TWO CODE PATHS. The alternative was one tool branching
# on Test-TenantCapability, which would have given the six most attribute-heavy
# tools in the toolkit a second path each, both needing testing. Prefixing the
# names keeps each file doing one thing, and Publish-Toolkit lists which set to
# strip for a given tenant.
#
# Ship AD-* where Tenant.json has OnPremAD true. Ship the unprefixed ones
# everywhere. Shipping both is not harmful -- the wrong one greys itself -- but
# it is two buttons doing one job.
# ---------------------------------------------------------------------------
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

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

# THE ADDRESS BOX IS ON THE FORM -- see Add-ToolTargetBar in Common.ps1.
#
# Resolve-ToolTarget is KEPT here, unlike in the cloud Group Editor. This one
# writes Active Directory, so it needs the directory and the controller, and an
# address it cannot place is an address this tool cannot serve. What changed is
# WHEN it runs: per lookup rather than once before the window existed. See
# Resolve-ToolTargetOnDemand in Common.ps1.
#
# FIVE VALUES MOVE TOGETHER OR NOT AT ALL. The address, its directory, its
# controller, the group's kind and its distinguished name are one answer to one
# question. Set-AdGroupTarget clears all five before it starts, so a lookup that
# fails halfway cannot leave the previous group's DN paired with the new
# address -- which would write to the wrong object while the title bar named the
# right one.
$script:Target        = ''
$script:GroupType     = $null
$script:ADGroupDN     = $null
$script:ADServer      = $null
$script:TargetEnv     = $null
$script:TargetProblem = ''

function Set-AdGroupTarget {
    param([string]$Text)

    $script:Target        = ''
    $script:GroupType     = $null
    $script:ADGroupDN     = $null
    $script:ADServer      = $null
    $script:TargetEnv     = $null
    $script:TargetProblem = ''

    if (-not "$Text".Trim()) { return }

    $t = Resolve-ToolTargetOnDemand -Text $Text -DomainController $DomainController -ToolTitle 'Group Editor -- Enter Target'
    if (-not $t) {
        # Resolve-ToolTarget has already shown its own warning box saying the
        # address is not in the on-premises directory. Recording it here as well
        # means the window still says so after the box is dismissed.
        $script:TargetProblem = "'$Text' was not found in the on-premises directory."
        return
    }

    $script:Target    = $t.Smtp
    $script:ADServer  = $t.ADServer
    $script:TargetEnv = $t.Environment

    try {
        $info = Resolve-Target -Smtp $t.Smtp -ADServer $t.ADServer
    } catch {
        $script:TargetProblem = "Could not resolve '$($t.Smtp)': $($_.Exception.Message)"
        return
    }

    if ($info.Type -ne 'UnifiedGroup' -and $info.Type -ne 'OnpremGroup') {
        # REPORTED IN THE WINDOW, NOT IN A MESSAGE BOX. The operator's next move
        # is to correct the address in the box already in front of them.
        $detail = if ($info.ResolutionErrors.Count -gt 0) {
            ($info.ResolutionErrors -join "`n`n")
        } else {
            "No Microsoft 365 Group or on-premises group was found for '$($t.Smtp)'."
        }
        $script:TargetProblem = "'$($t.Smtp)' did not resolve to a group (Type: $($info.Type)).`r`n`r`n$detail"
        return
    }

    $script:GroupType = $info.Type
    $script:ADGroupDN = $info.ADGroupDN
}

Set-AdGroupTarget -Text $SmtpAddress

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
    # NO PARAMETERS. All five values can change while the window is open, so
    # they live in the $script: variables Set-AdGroupTarget fills rather than in
    # parameters fixed the moment the form was built.
    $editForm = New-Object System.Windows.Forms.Form
    $editForm.Text = 'Group Editor'
    $editForm.ClientSize = New-Object System.Drawing.Size(764, 400)
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
    $grpGrp.Text = 'Group Delegation'
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
        # The delegation box is the only thing that WRITES. Greying it is how
        # the window says "I do not know what this is yet" without a dialog.
        $ready = [bool]$script:GroupType
        $grpGrp.Enabled = $ready
        $grpGrp.Text = if ($ready) { "Group Delegation ($script:GroupType)" } else { 'Group Delegation' }
        $editForm.Text = if ("$script:Target".Trim()) { "Group Editor: $script:Target" } else { 'Group Editor' }

        if (-not $ready) {
            $txtOut.Text = if ($script:TargetProblem) {
                $script:TargetProblem
            } else {
                'Enter a group address above and press Look Up.'
            }
            return
        }

        try {
            $state = Get-GroupState -Type $script:GroupType -Smtp $script:Target -ADGroupDN $script:ADGroupDN -ADServer $script:ADServer
            $txtOut.Text = Format-GroupState -Heading "CURRENT STATE" -State $state
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Group Editor'
        }
    }

    function Invoke-GroupEdit {
        param([string]$ActionLabel, [string]$Verb, [scriptblock]$Action)
        if (-not $script:GroupType) { $txtOut.Text = 'Look up a group first.'; return }
        $grantee = $txtGrpGrantee.Text.Trim()
        if (-not $grantee) { $txtOut.Text = "Enter a grantee (email or logon) first."; return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "$ActionLabel for '$grantee' on group '$script:Target'?",
            "Confirm $ActionLabel",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $grpGrp.Enabled = $false
        try {
            $before = Get-GroupState -Type $script:GroupType -Smtp $script:Target -ADGroupDN $script:ADGroupDN -ADServer $script:ADServer
            $out = New-Object System.Text.StringBuilder
            [void]$out.Append((Format-GroupState -Heading "BEFORE" -State $before))
            & $Action
            [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
            $after = Get-GroupState -Type $script:GroupType -Smtp $script:Target -ADGroupDN $script:ADGroupDN -ADServer $script:ADServer
            [void]$out.Append((Format-GroupState -Heading "AFTER" -State $after))
            $txtOut.Text = $out.ToString()
            Write-CompletedLog -Task "$Verb $script:Target ($script:TargetEnv; grantee: $grantee)"
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
            if ($script:GroupType -eq 'UnifiedGroup') { Add-UnifiedOwner -Smtp $script:Target -Grantee $grantee }
            else { Set-OnpremOwner -ADGroupDN $script:ADGroupDN -Grantee $grantee -ADServer $script:ADServer }
        }
    })
    $btnOwnerRemove.Add_Click({
        Invoke-GroupEdit -ActionLabel "Remove Owner" -Verb "Removed Owner" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($script:GroupType -eq 'UnifiedGroup') { Remove-UnifiedOwner -Smtp $script:Target -Grantee $grantee }
            else { throw "On-prem groups have a single ManagedBy owner. Use 'Set/Add Owner' with a new value to replace it -- there is no separate remove action for the AD ManagedBy attribute." }
        }
    })
    $btnMemberAdd.Add_Click({
        Invoke-GroupEdit -ActionLabel "Add Member" -Verb "Added Member" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($script:GroupType -eq 'UnifiedGroup') { Add-UnifiedMember -Smtp $script:Target -Grantee $grantee }
            else { Add-OnpremMember -ADGroupDN $script:ADGroupDN -Grantee $grantee -ADServer $script:ADServer }
        }
    })
    $btnMemberRemove.Add_Click({
        Invoke-GroupEdit -ActionLabel "Remove Member" -Verb "Removed Member" -Action {
            $grantee = $txtGrpGrantee.Text.Trim()
            if ($script:GroupType -eq 'UnifiedGroup') { Remove-UnifiedMember -Smtp $script:Target -Grantee $grantee }
            else { Remove-OnpremMember -ADGroupDN $script:ADGroupDN -Grantee $grantee -ADServer $script:ADServer }
        }
    })
    $btnEditClose.Add_Click({ $editForm.Close() })

    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down and
    # grows the form, and the help button places itself from ClientSize at run
    # time. Before Set-FormScreenCorner too, which positions from the size.
    $bar = Add-ToolTargetBar -Form $editForm -Label 'Group address:' -InitialValue $script:Target

    # RESOLVES ALL FIVE, THEN REDRAWS. Update-StateDisplay reads whatever
    # Set-AdGroupTarget managed to fill and reports the rest.
    $bar.Button.Add_Click({
        Set-AdGroupTarget -Text $bar.TextBox.Text
        Update-StateDisplay
    })

    Set-ConsoleTheme -Form $editForm
    Set-FormScreenCorner -Form $editForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $editForm -ScriptPath $PSCommandPath -Title 'Group Editor' | Out-Null

    [void]$editForm.ShowDialog()
}
#endregion

Show-GroupEditorForm
