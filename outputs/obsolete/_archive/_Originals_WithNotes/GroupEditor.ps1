#Requires -Version 5.1
<#
.SYNOPSIS
    Standalone Group Editor (owner / member delegation for M365 Groups and
    on-prem synced groups) tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\GroupEditor.ps1
    prompts for a group / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Group Editor window.

    Fully functional from AdminConsole.ps1 too:
        .\GroupEditor.ps1 -SmtpAddress group@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Group Editor window
    for that group.

    Get-GroupState, Format-GroupState, Add-UnifiedOwner, Remove-UnifiedOwner,
    Add-UnifiedMember, Remove-UnifiedMember, Set-OnpremOwner, Add-OnpremMember,
    and Remove-OnpremMember below are MOVED, UNCHANGED, from the original
    monolithic Mailbox-Group-Inspector.ps1 (original lines 985-1105, region
    "Actions - group").

    THIS SCRIPT IS DIFFERENT FROM EVERY OTHER TOOL SCRIPT in one respect:
    AdminConsole.ps1 only ever passes -SmtpAddress/-Environment/-DomainController
    to a tool -- it does NOT pass the object Type (UnifiedGroup vs
    OnpremGroup) or the on-prem ADGroupDN it already resolved while deciding
    which buttons to show (see Resolve-Target's own comment in Common.ps1:
    "Individual tool scripts do NOT need to call this themselves -- each one
    already assumes its own target type"). Every other tool assumes Mailbox
    and never needs to ask. Group editing genuinely cannot proceed without
    knowing Type (it changes which cmdlets run) and, for an on-prem group,
    ADGroupDN (Get-ADGroup/Get-ADGroupMember/Set-ADGroup/Add-ADGroupMember/
    Remove-ADGroupMember all key off it, not off the mail-address Identity).
    So this script calls Resolve-Target itself, immediately after
    Resolve-ToolTarget resolves Smtp/Environment/ADServer, to get exactly
    the Type/ADGroupDN it needs -- the one exception to the "tool scripts
    don't call Resolve-Target" rule, made necessary by what this specific
    tool does, not a change to that rule for anyone else.

    LIKE MailboxPermissions.ps1 and InboxRules.ps1, this tool's editable
    controls (the grantee textbox and the four Owner/Member buttons) never
    had their own Show-XForm in the monolith -- they lived directly on the
    monolith's shared main form as the "$grpGrp" group box, and
    Invoke-GroupEdit (original lines 4136-4168) referenced $grpGrp /
    $txtGrpGrantee / $txtOut / $btnRun / $script:LastSmtp / $script:Target /
    Get-SelectedADServer / Get-SelectedEnvLabel as outer-scope
    variables/functions belonging to that shared form. So this file adds ONE
    new function, Show-GroupEditorForm, that did not exist in the monolith --
    a small new wrapper window built to hold exactly the same
    grantee-textbox-plus-four-button layout the monolith drew directly on
    its main form, plus a state display box on top (populated via
    Get-GroupState / Format-GroupState, the same before/after pattern
    Invoke-GroupEdit already used).

    Invoke-GroupEdit itself is reproduced with the following NECESSARY,
    mechanical adaptations (the monolith's version cannot run unmodified
    outside its main form, since the controls/variables/functions it
    referenced no longer exist):
      - $grpGrp / $txtGrpGrantee / $txtOut are now the controls built by
        Show-GroupEditorForm below, instead of the monolith's main form
        controls of the same names.
      - $btnRun.Enabled = $true/$false (disabling the main form's Lookup
        button during the edit) is removed -- there is no main-form Lookup
        button in this process.
      - $script:LastSmtp / $script:Target are replaced with the -Smtp /
        -Type / -ADGroupDN / -ADServer parameters this script already
        resolved (Invoke-GroupEdit is now a nested function inside
        Show-GroupEditorForm and reads $Smtp/$Type/$ADGroupDN/$ADServer from
        that enclosing scope instead).
      - The "Run a Lookup first" guard is removed -- Resolve-ToolTarget plus
        this script's own Resolve-Target call already guarantee a resolved
        group before this window opens at all, so that state is no longer
        reachable here.
      - Get-SelectedADServer() calls are replaced with the enclosing
        $ADServer parameter (Get-SelectedADServer read $cboDC.SelectedItem /
        $radHQ.Checked off the monolith's main form, which no longer exists;
        Common.ps1's Get-ADServerString already produced the equivalent
        value once, during this script's own Resolve-Target call above).
      - The Write-CompletedLog line's "(Get-SelectedEnvLabel)" call is
        replaced with the -Environment parameter this script already
        resolved.
    Every other line of every function, including all comments and the
    confirm-dialog / before-after-log mechanics, is unchanged in behavior.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'Group Editor -- Enter Target'
if (-not $target) { return }   # operator cancelled the prompt

# See .DESCRIPTION above: this is the one tool script that must call
# Resolve-Target itself, since AdminConsole.ps1 doesn't hand off the object
# Type/ADGroupDN it already resolved while deciding whether to show this
# tool's button at all.
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
# Unified Group (Microsoft 365 Group) edits -- EXO, resolved via Invoke-EXOCommand
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
# On-prem group edits -- ActiveDirectory module, unambiguous cmdlets
function Set-OnpremOwner {
    param([string]$ADGroupDN, [string]$Grantee, [string]$ADServer)
    # Replaces the single AD ManagedBy value -- see header "KNOWN LIMITATION".
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
# This function did not exist in the monolith. It exists only to give
# Invoke-GroupEdit the $grpGrp/$txtGrpGrantee/$txtOut controls it already
# expected to find in outer scope, using the exact same grantee-textbox +
# four-button layout the monolith drew directly on its shared main form
# (original "$grpGrp" group box), plus a BEFORE/AFTER state display on top.
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

    # Reproduced from Invoke-GroupEdit (original lines 4136-4168) -- see the
    # .DESCRIPTION header above for the exact list of mechanical adaptations
    # (control names now resolve to this form's own controls, $btnRun
    # disabling removed, $script:LastSmtp/$script:Target replaced by the
    # enclosing $Smtp/$Type/$ADGroupDN/$ADServer, the "Run a Lookup first"
    # guard removed since Resolve-ToolTarget + this script's own
    # Resolve-Target call already guarantee a resolved group,
    # Get-SelectedADServer replaced by the enclosing $ADServer,
    # Get-SelectedEnvLabel replaced by the enclosing $Environment). The
    # confirm-dialog / BEFORE-AFTER-log mechanics themselves are unchanged.
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
    # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
    # Console, which opens top-left. Helper is in Common.ps1.
    Set-FormScreenCorner -Form $editForm -Corner 'BottomLeft'
    [void]$editForm.ShowDialog()
}
#endregion

Show-GroupEditorForm -Smtp $target.Smtp -Type $groupInfo.Type -ADGroupDN $groupInfo.ADGroupDN -ADServer $target.ADServer -Environment $target.Environment
