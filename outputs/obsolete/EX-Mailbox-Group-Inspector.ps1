#Requires -Version 5.1
# TOOLKIT-TENANT: Express
# ---------------------------------------------------------------------------
# EX- PREFIX: this script ships to ONE tenant only.
#
# Deliberately standalone -- it does not dot-source Common.ps1 and keeps its own
# copies of the domains, the domain controllers and several shared functions.
# That is why it survives being handed to someone with no toolkit, and why it
# cannot read Tenant.json.
#
# LEFT EXACTLY AS IT WAS. An automated pass over the toolkit once rewrote
# $script:MailDomain here into $(Get-MailDomain) -- on the LEFT of an
# assignment, which is not valid PowerShell and broke a 3,100-line working
# tool. EX- files are exempt from the tenant rules for a reason; they should be
# exempt from bulk edits for the same one.
# ---------------------------------------------------------------------------
<#
    Mailbox-Group-Inspector.ps1
    Looks up a mailbox or group by email address and displays/edits its
    EAC-equivalent properties (mailbox delegation + usage, or group
    delegation + membership).

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Common.ps1 -- the shared library. These three tools predate it and each
# carried private copies of its functions, including a v1 two-column logger
# that wrote to a DIFFERENT audit file with no Result column. That split the
# toolkit's history across two schemas and made any count of failures a
# statement about only part of the toolkit. See ARCHITECTURE 14n.
#
# Loading Common is cheap: it defines functions and executes nothing at the
# top level, so this adds no launch cost.
. (Join-Path $PSScriptRoot 'Common.ps1')

#region Parsing ---------------------------------------------------------------
$script:MailDomain = 'expresspros.com'
# Get-TargetSmtp now comes from Common.ps1 (was a private copy here).
#endregion
#region Helpers - error formatting --------------------------------------------
# Format-ErrorDetail now comes from Common.ps1 (was a private copy here).
# Get-EXOCommand now comes from Common.ps1 (was a private copy here).
# Invoke-EXOCommand now comes from Common.ps1 (was a private copy here).
#endregion
#region Actions - resolution ---------------------------------------------------
# Resolve-Target now comes from Common.ps1 (was a private copy here).
#endregion
#region Actions - mailbox -------------------------------------------------------
function Get-MailboxState {
    param([string]$Smtp)
    $state = [ordered]@{
        DisplayName          = $null
        PrimarySmtpAddress   = $null
        Alias                = $null
        RecipientTypeDetails = $null
        Hidden               = $null
        LitigationHold       = $null
        ArchiveEnabled       = $null
        ArchiveStatus        = $null
        MailboxSize          = $null
        MailboxItemCount     = $null
        ArchiveSize          = $null
        ArchiveItemCount     = $null
        FullAccess           = @()
        SendAs               = @()
        Errors               = @()
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.DisplayName          = $mbx.DisplayName
        $state.PrimarySmtpAddress   = $mbx.PrimarySmtpAddress
        $state.Alias                = $mbx.Alias
        $state.RecipientTypeDetails = $mbx.RecipientTypeDetails
        $state.Hidden               = $mbx.HiddenFromAddressListsEnabled
        $state.LitigationHold       = $mbx.LitigationHoldEnabled
        $state.ArchiveEnabled = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $state.ArchiveStatus  = $mbx.ArchiveStatus
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (core properties)')
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.MailboxSize      = $stats.TotalItemSize
        $state.MailboxItemCount = $stats.ItemCount
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (mailbox size)')
    }
    if ($state.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $state.ArchiveSize      = $astats.TotalItemSize
            $state.ArchiveItemCount = $astats.ItemCount
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (archive size)')
        }
    }
    try {
        $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.User -notlike 'NT AUTHORITY\SELF' -and
                $_.IsInherited -eq $false -and
                $_.Deny -eq $false -and
                ($_.AccessRights -contains 'FullAccess')
            }
        $state.FullAccess = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access)')
    }
    try {
        $sa = Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.Trustee -notlike 'NT AUTHORITY\SELF' -and
                $_.AccessControlType -eq 'Allow'
            }
        $state.SendAs = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As)')
    }
    return [PSCustomObject]$state
}
function Format-MailboxState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading (Mailbox) ---")
    [void]$sb.AppendLine("  Display Name        : $($State.DisplayName)")
    [void]$sb.AppendLine("  Primary SMTP        : $($State.PrimarySmtpAddress)")
    [void]$sb.AppendLine("  Alias               : $($State.Alias)")
    [void]$sb.AppendLine("  Mailbox Type        : $($State.RecipientTypeDetails)")
    [void]$sb.AppendLine("  Hidden From GAL     : $($State.Hidden)")
    [void]$sb.AppendLine("  Litigation Hold     : $($State.LitigationHold)")
    [void]$sb.AppendLine("  Mailbox Size        : $($State.MailboxSize)  ($($State.MailboxItemCount) items)")
    [void]$sb.AppendLine("  Archive Enabled     : $($State.ArchiveEnabled)  (ArchiveStatus: $($State.ArchiveStatus))")
    if ($State.ArchiveEnabled) {
        [void]$sb.AppendLine("  Archive Size        : $($State.ArchiveSize)  ($($State.ArchiveItemCount) items)")
    }
    [void]$sb.AppendLine("  Full Access grantees:")
    if ($State.FullAccess.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.FullAccess | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    [void]$sb.AppendLine("  Send As grantees    :")
    if ($State.SendAs.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.SendAs | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    if ($State.Errors.Count -gt 0) {
        [void]$sb.AppendLine("  --- Field errors (partial data above is still valid) ---")
        $State.Errors | ForEach-Object { [void]$sb.AppendLine("  $_") }
    }
    return $sb.ToString()
}
function Grant-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Grant-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
#endregion
#region Actions - inbox rules ---------------------------------------------------
# Get-InboxRulesForMailbox now comes from Common.ps1 (was a private copy here).
function Enable-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Enable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Disable-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Remove-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Update-RulesGrid {
    param([string]$Smtp)
    $grdRules.Rows.Clear()
    $result = Get-InboxRulesForMailbox -Smtp $Smtp
    if ($result.Error) {
        $txtOut.Text = $result.Error
        $script:CurrentRules = @()
        return
    }
    $script:CurrentRules = $result.Rules
    foreach ($r in $result.Rules) {
        $desc = if ($r.Description) { ($r.Description -replace '\s+', ' ').Trim() } else { '' }
        [void]$grdRules.Rows.Add($r.Name, $desc, $r.Enabled, $r.Identity.ToString())
    }
}
function Show-RuleDetailForm {
    param($Rule, [string]$Smtp)
    $ruleForm = New-Object System.Windows.Forms.Form
    $ruleForm.Text = "Inbox Rule: $($Rule.Name)"
    $ruleForm.Size = New-Object System.Drawing.Size(560, 430)
    $ruleForm.StartPosition = "CenterParent"
    $ruleForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = "Rule Name : $($Rule.Name)"
    $lblName.Location = New-Object System.Drawing.Point(12, 12)
    $lblName.AutoSize = $true
    $ruleForm.Controls.Add($lblName)
    $lblMailbox = New-Object System.Windows.Forms.Label
    $lblMailbox.Text = "Mailbox   : $Smtp"
    $lblMailbox.Location = New-Object System.Drawing.Point(12, 32)
    $lblMailbox.AutoSize = $true
    $ruleForm.Controls.Add($lblMailbox)
    $lblEnabled = New-Object System.Windows.Forms.Label
    $lblEnabled.Text = "Enabled   : $($Rule.Enabled)"
    $lblEnabled.Location = New-Object System.Drawing.Point(12, 52)
    $lblEnabled.AutoSize = $true
    $lblEnabled.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $ruleForm.Controls.Add($lblEnabled)
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Full Description:"
    $lblDesc.Location = New-Object System.Drawing.Point(12, 78)
    $lblDesc.AutoSize = $true
    $ruleForm.Controls.Add($lblDesc)
    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Multiline = $true
    $txtDesc.ReadOnly = $true
    $txtDesc.ScrollBars = "Vertical"
    $txtDesc.Location = New-Object System.Drawing.Point(12, 98)
    $txtDesc.Size = New-Object System.Drawing.Size(520, 200)
    $txtDesc.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtDesc.Text = $Rule.Description
    $ruleForm.Controls.Add($txtDesc)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(12, 304)
    $lblStatus.Size = New-Object System.Drawing.Size(520, 34)
    $ruleForm.Controls.Add($lblStatus)
    $btnEnable = New-Object System.Windows.Forms.Button
    $btnEnable.Text = "Enable"
    $btnEnable.Location = New-Object System.Drawing.Point(12, 344)
    $btnEnable.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnEnable)
    $btnDisable = New-Object System.Windows.Forms.Button
    $btnDisable.Text = "Disable"
    $btnDisable.Location = New-Object System.Drawing.Point(120, 344)
    $btnDisable.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnDisable)
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "Delete"
    $btnDelete.Location = New-Object System.Drawing.Point(228, 344)
    $btnDelete.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnDelete)
    $btnRuleClose = New-Object System.Windows.Forms.Button
    $btnRuleClose.Text = "Close"
    $btnRuleClose.Location = New-Object System.Drawing.Point(432, 344)
    $btnRuleClose.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnRuleClose)
    $ruleIdentity = $Rule.Identity.ToString()
    $btnEnable.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Enable Inbox rule '$($Rule.Name)' on mailbox '$Smtp'?",
            "Confirm Enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Enable-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Enabled successfully."
            $lblEnabled.Text = "Enabled   : True"
            Write-CompletedLog -Task "Enabled InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Enable-InboxRule'
        } finally {
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnDisable.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Disable Inbox rule '$($Rule.Name)' on mailbox '$Smtp'?",
            "Confirm Disable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Disable-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Disabled successfully."
            $lblEnabled.Text = "Enabled   : False"
            Write-CompletedLog -Task "Disabled InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Disable-InboxRule'
        } finally {
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnDelete.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Permanently delete Inbox rule '$($Rule.Name)' on mailbox '$Smtp'? This cannot be undone.",
            "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Remove-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            Write-CompletedLog -Task "Deleted InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
            [System.Windows.Forms.MessageBox]::Show("Rule deleted.", "Deleted", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            $ruleForm.Close()
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-InboxRule'
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnRuleClose.Add_Click({ $ruleForm.Close() })
    [void]$ruleForm.ShowDialog($form)
}
#endregion
#region Actions - group ---------------------------------------------------------
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
#region Actions - out of office ---------------------------------------------------
function Get-OutOfOfficeConfig {
    param([string]$Smtp)
    Invoke-EXOCommand -Name 'Get-MailboxAutoReplyConfiguration' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
}
function Set-OutOfOfficeConfig {
    param(
        [string]$Smtp,
        [string]$AutoReplyState,
        [string]$ExternalAudience,
        [string]$InternalMessage,
        [string]$ExternalMessage,
        [Nullable[DateTime]]$StartTime,
        [Nullable[DateTime]]$EndTime
    )
    $params = @{
        Identity         = $Smtp
        AutoReplyState   = $AutoReplyState
        ExternalAudience = $ExternalAudience
        InternalMessage  = $InternalMessage
        ExternalMessage  = $ExternalMessage
        Confirm          = $false
        ErrorAction      = 'Stop'
    }
    if ($AutoReplyState -eq 'Scheduled') {
        if ($StartTime) { $params.StartTime = $StartTime }
        if ($EndTime)   { $params.EndTime   = $EndTime }
    }
    Invoke-EXOCommand -Name 'Set-MailboxAutoReplyConfiguration' -Parameters $params | Out-Null
}
function Show-OutOfOfficeForm {
    param([string]$Smtp)
    $oofForm = New-Object System.Windows.Forms.Form
    $oofForm.Text = "Out of Office: $Smtp"
    $oofForm.Size = New-Object System.Drawing.Size(560, 560)
    $oofForm.StartPosition = "CenterParent"
    $oofForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblState = New-Object System.Windows.Forms.Label
    $lblState.Text = "State:"
    $lblState.Location = New-Object System.Drawing.Point(12, 15)
    $lblState.AutoSize = $true
    $oofForm.Controls.Add($lblState)
    $cboState = New-Object System.Windows.Forms.ComboBox
    $cboState.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Disabled', 'Enabled', 'Scheduled') | ForEach-Object { [void]$cboState.Items.Add($_) }
    $cboState.Location = New-Object System.Drawing.Point(140, 12)
    $cboState.Size = New-Object System.Drawing.Size(150, 20)
    $oofForm.Controls.Add($cboState)
    $lblAudience = New-Object System.Windows.Forms.Label
    $lblAudience.Text = "External audience:"
    $lblAudience.Location = New-Object System.Drawing.Point(300, 15)
    $lblAudience.AutoSize = $true
    $oofForm.Controls.Add($lblAudience)
    $cboAudience = New-Object System.Windows.Forms.ComboBox
    $cboAudience.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('None', 'Known', 'All') | ForEach-Object { [void]$cboAudience.Items.Add($_) }
    $cboAudience.Location = New-Object System.Drawing.Point(420, 12)
    $cboAudience.Size = New-Object System.Drawing.Size(100, 20)
    $oofForm.Controls.Add($cboAudience)
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = "Start (Scheduled only, e.g. 2026-07-10 08:00):"
    $lblStart.Location = New-Object System.Drawing.Point(12, 45)
    $lblStart.AutoSize = $true
    $oofForm.Controls.Add($lblStart)
    $txtStart = New-Object System.Windows.Forms.TextBox
    $txtStart.Location = New-Object System.Drawing.Point(280, 42)
    $txtStart.Size = New-Object System.Drawing.Size(240, 20)
    $txtStart.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtStart)
    $lblEnd = New-Object System.Windows.Forms.Label
    $lblEnd.Text = "End (Scheduled only):"
    $lblEnd.Location = New-Object System.Drawing.Point(12, 70)
    $lblEnd.AutoSize = $true
    $oofForm.Controls.Add($lblEnd)
    $txtEnd = New-Object System.Windows.Forms.TextBox
    $txtEnd.Location = New-Object System.Drawing.Point(280, 67)
    $txtEnd.Size = New-Object System.Drawing.Size(240, 20)
    $txtEnd.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtEnd)
    $lblInt = New-Object System.Windows.Forms.Label
    $lblInt.Text = "Internal message:"
    $lblInt.Location = New-Object System.Drawing.Point(12, 98)
    $lblInt.AutoSize = $true
    $oofForm.Controls.Add($lblInt)
    $txtInt = New-Object System.Windows.Forms.TextBox
    $txtInt.Multiline = $true
    $txtInt.ScrollBars = "Vertical"
    $txtInt.Location = New-Object System.Drawing.Point(12, 118)
    $txtInt.Size = New-Object System.Drawing.Size(520, 140)
    $txtInt.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtInt)
    $lblExt = New-Object System.Windows.Forms.Label
    $lblExt.Text = "External message:"
    $lblExt.Location = New-Object System.Drawing.Point(12, 264)
    $lblExt.AutoSize = $true
    $oofForm.Controls.Add($lblExt)
    $txtExt = New-Object System.Windows.Forms.TextBox
    $txtExt.Multiline = $true
    $txtExt.ScrollBars = "Vertical"
    $txtExt.Location = New-Object System.Drawing.Point(12, 284)
    $txtExt.Size = New-Object System.Drawing.Size(520, 140)
    $txtExt.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtExt)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(12, 430)
    $lblStatus.Size = New-Object System.Drawing.Size(520, 40)
    $oofForm.Controls.Add($lblStatus)
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save"
    $btnSave.Location = New-Object System.Drawing.Point(12, 474)
    $btnSave.Size = New-Object System.Drawing.Size(100, 28)
    $oofForm.Controls.Add($btnSave)
    $btnOofClose = New-Object System.Windows.Forms.Button
    $btnOofClose.Text = "Close"
    $btnOofClose.Location = New-Object System.Drawing.Point(120, 474)
    $btnOofClose.Size = New-Object System.Drawing.Size(100, 28)
    $oofForm.Controls.Add($btnOofClose)
    function Update-OofFormFromCurrent {
        try {
            $cfg = Get-OutOfOfficeConfig -Smtp $Smtp
            $cboState.SelectedItem    = $cfg.AutoReplyState.ToString()
            $cboAudience.SelectedItem = $cfg.ExternalAudience.ToString()
            $txtStart.Text = if ($cfg.StartTime) { $cfg.StartTime.ToString() } else { '' }
            $txtEnd.Text   = if ($cfg.EndTime)   { $cfg.EndTime.ToString() }   else { '' }
            $txtInt.Text   = $cfg.InternalMessage
            $txtExt.Text   = $cfg.ExternalMessage
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Loaded current settings."
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxAutoReplyConfiguration'
        }
    }
    Update-OofFormFromCurrent
    $btnSave.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Save Out of Office settings for '$Smtp'? State: $($cboState.SelectedItem)",
            "Confirm Save", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnSave.Enabled = $false
        try {
            $st = $null; $et = $null
            if ($cboState.SelectedItem -eq 'Scheduled') {
                if ($txtStart.Text.Trim()) { $st = [DateTime]::Parse($txtStart.Text.Trim()) }
                if ($txtEnd.Text.Trim())   { $et = [DateTime]::Parse($txtEnd.Text.Trim()) }
            }
            Set-OutOfOfficeConfig -Smtp $Smtp -AutoReplyState $cboState.SelectedItem -ExternalAudience $cboAudience.SelectedItem `
                -InternalMessage $txtInt.Text -ExternalMessage $txtExt.Text -StartTime $st -EndTime $et
            Write-CompletedLog -Task "Set OutOfOffice $Smtp (EXO; state: $($cboState.SelectedItem))"
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Saved successfully."
            Update-OofFormFromCurrent
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-MailboxAutoReplyConfiguration'
        } finally {
            $btnSave.Enabled = $true
        }
    })
    $btnOofClose.Add_Click({ $oofForm.Close() })
    [void]$oofForm.ShowDialog($form)
}
#endregion
#region Actions - calendar permissions --------------------------------------------
function Get-CalendarFolderIdentity {
    param([string]$Smtp)
    $folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{ Identity = $Smtp; FolderScope = 'Calendar'; ErrorAction = 'Stop' })
    if ($folders.Count -eq 0) { throw "No Calendar-scope folder found for $Smtp." }
    $primary = $folders | Where-Object { $_.FolderPath -eq '/Calendar' } | Select-Object -First 1
    if (-not $primary) { $primary = $folders | Select-Object -First 1 }
    return "$($Smtp):\$($primary.Name)"
}
function Get-CalendarPermissions {
    param([string]$Smtp)
    try {
        $calId = Get-CalendarFolderIdentity -Smtp $Smtp
        $perms = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $calId; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ CalendarIdentity = $calId; Permissions = $perms; Error = $null }
    } catch {
        return [PSCustomObject]@{ CalendarIdentity = $null; Permissions = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxFolderPermission (Calendar)') }
    }
}
function Set-CalendarPermission {
    param([string]$CalendarIdentity, [string]$Grantee, [string]$Role)
    $existing = @()
    try {
        $existing = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $CalendarIdentity; User = $Grantee; ErrorAction = 'Stop' })
    } catch {
        $existing = @()
    }
    if ($existing.Count -gt 0) {
        Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Updated'
    } else {
        Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Added'
    }
}
function Remove-CalendarPermission {
    param([string]$CalendarIdentity, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxFolderPermission' -Parameters @{
        Identity = $CalendarIdentity; User = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Show-CalendarPermissionsForm {
    param([string]$Smtp)
    $calForm = New-Object System.Windows.Forms.Form
    $calForm.Text = "Calendar Permissions: $Smtp"
    $calForm.Size = New-Object System.Drawing.Size(620, 480)
    $calForm.StartPosition = "CenterParent"
    $calForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblCalId = New-Object System.Windows.Forms.Label
    $lblCalId.Text = "Calendar folder:"
    $lblCalId.Location = New-Object System.Drawing.Point(12, 12)
    $lblCalId.AutoSize = $true
    $calForm.Controls.Add($lblCalId)
    $grdCal = New-Object System.Windows.Forms.DataGridView
    $grdCal.Location = New-Object System.Drawing.Point(12, 34)
    $grdCal.Size = New-Object System.Drawing.Size(580, 220)
    $grdCal.ReadOnly = $true
    $grdCal.AllowUserToAddRows = $false
    $grdCal.AllowUserToDeleteRows = $false
    $grdCal.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdCal.MultiSelect = $false
    $grdCal.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdCal.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdCal.Columns.Add('User', 'User')
    [void]$grdCal.Columns.Add('AccessRights', 'Access Rights')
    foreach ($c in $grdCal.Columns) { $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }
    $calForm.Controls.Add($grdCal)
    $lblGrantee = New-Object System.Windows.Forms.Label
    $lblGrantee.Text = "User (email or logon):"
    $lblGrantee.Location = New-Object System.Drawing.Point(12, 266)
    $lblGrantee.AutoSize = $true
    $calForm.Controls.Add($lblGrantee)
    $txtGrantee = New-Object System.Windows.Forms.TextBox
    $txtGrantee.Location = New-Object System.Drawing.Point(160, 263)
    $txtGrantee.Size = New-Object System.Drawing.Size(220, 20)
    $txtGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $calForm.Controls.Add($txtGrantee)
    $lblRole = New-Object System.Windows.Forms.Label
    $lblRole.Text = "Role:"
    $lblRole.Location = New-Object System.Drawing.Point(390, 266)
    $lblRole.AutoSize = $true
    $calForm.Controls.Add($lblRole)
    $cboRole = New-Object System.Windows.Forms.ComboBox
    $cboRole.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Reviewer', 'Editor', 'Owner') | ForEach-Object { [void]$cboRole.Items.Add($_) }
    $cboRole.SelectedIndex = 0
    $cboRole.Location = New-Object System.Drawing.Point(430, 263)
    $cboRole.Size = New-Object System.Drawing.Size(120, 20)
    $calForm.Controls.Add($cboRole)
    $lblCalStatus = New-Object System.Windows.Forms.Label
    $lblCalStatus.Text = ""
    $lblCalStatus.Location = New-Object System.Drawing.Point(12, 296)
    $lblCalStatus.Size = New-Object System.Drawing.Size(580, 40)
    $calForm.Controls.Add($lblCalStatus)
    $btnAddPerm = New-Object System.Windows.Forms.Button
    $btnAddPerm.Text = "Add / Update Permission"
    $btnAddPerm.Location = New-Object System.Drawing.Point(12, 340)
    $btnAddPerm.Size = New-Object System.Drawing.Size(170, 28)
    $calForm.Controls.Add($btnAddPerm)
    $btnRemovePerm = New-Object System.Windows.Forms.Button
    $btnRemovePerm.Text = "Remove Selected"
    $btnRemovePerm.Location = New-Object System.Drawing.Point(190, 340)
    $btnRemovePerm.Size = New-Object System.Drawing.Size(150, 28)
    $calForm.Controls.Add($btnRemovePerm)
    $btnCalClose = New-Object System.Windows.Forms.Button
    $btnCalClose.Text = "Close"
    $btnCalClose.Location = New-Object System.Drawing.Point(500, 340)
    $btnCalClose.Size = New-Object System.Drawing.Size(90, 28)
    $calForm.Controls.Add($btnCalClose)
    $script:CalendarIdentityForForm = $null
    function Update-CalGridFromCurrent {
        $grdCal.Rows.Clear()
        $result = Get-CalendarPermissions -Smtp $Smtp
        if ($result.Error) {
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = $result.Error
            $lblCalId.Text = "Calendar folder: (unresolved)"
            return
        }
        $script:CalendarIdentityForForm = $result.CalendarIdentity
        $lblCalId.Text = "Calendar folder: $($result.CalendarIdentity)"
        foreach ($p in $result.Permissions) {
            [void]$grdCal.Rows.Add($p.User.ToString(), ($p.AccessRights -join ', '))
        }
    }
    Update-CalGridFromCurrent
    $btnAddPerm.Add_Click({
        $grantee = $txtGrantee.Text.Trim()
        if (-not $grantee) { $lblCalStatus.Text = "Enter a user first."; return }
        if (-not $script:CalendarIdentityForForm) { $lblCalStatus.Text = "Calendar folder not resolved -- see status above."; return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Grant '$($cboRole.SelectedItem)' calendar access to '$grantee' on '$Smtp'?",
            "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnAddPerm.Enabled = $false; $btnRemovePerm.Enabled = $false
        try {
            $action = Set-CalendarPermission -CalendarIdentity $script:CalendarIdentityForForm -Grantee $grantee -Role $cboRole.SelectedItem
            Write-CompletedLog -Task "$action CalendarPermission $Smtp (EXO; user: $grantee; role: $($cboRole.SelectedItem))"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "$action permission for $grantee."
            Update-CalGridFromCurrent
        } catch {
            Write-FailedLog -Action 'GrantCalendarPermission' -Target $Smtp -ErrorRecord $_ -Detail "user: $grantee"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Calendar permission grant'
        } finally {
            $btnAddPerm.Enabled = $true; $btnRemovePerm.Enabled = $true
        }
    })
    $btnRemovePerm.Add_Click({
        if ($grdCal.SelectedRows.Count -eq 0) { $lblCalStatus.Text = "Select a row first."; return }
        $grantee = $grdCal.SelectedRows[0].Cells['User'].Value
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove all calendar permissions for '$grantee' on '$Smtp'?",
            "Confirm Remove", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnAddPerm.Enabled = $false; $btnRemovePerm.Enabled = $false
        try {
            Remove-CalendarPermission -CalendarIdentity $script:CalendarIdentityForForm -Grantee $grantee
            Write-CompletedLog -Task "Removed CalendarPermission $Smtp (EXO; user: $grantee)"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "Removed permission for $grantee."
            Update-CalGridFromCurrent
        } catch {
            Write-FailedLog -Action 'RemoveCalendarPermission' -Target $Smtp -ErrorRecord $_ -Detail "user: $grantee"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-MailboxFolderPermission (Calendar)'
        } finally {
            $btnAddPerm.Enabled = $true; $btnRemovePerm.Enabled = $true
        }
    })
    $btnCalClose.Add_Click({ $calForm.Close() })
    [void]$calForm.ShowDialog($form)
}
#endregion
#region Actions - mailbox health ---------------------------------------------------
function Get-MailboxDLMemberships {
    param(
        [string]$Smtp,
        [string]$ADServer,
        [scriptblock]$ProgressAction,
        [scriptblock]$ShouldCancel
    )
    $result = [ordered]@{
        Memberships = @()
        ScanErrors  = @()
        FatalError  = $null
        Cancelled   = $false
        NoADAccount = $false
    }
    try {
        $adUsers = @(Get-ADUser -Filter "mail -eq '$Smtp'" -Properties memberOf -Server $ADServer -ErrorAction Stop)
    } catch {
        $result.FatalError = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADUser (DL membership lookup)')
        return [PSCustomObject]$result
    }
    if ($adUsers.Count -eq 0) {
        $result.NoADAccount = $true
        return [PSCustomObject]$result
    }
    if ($adUsers.Count -gt 1) {
        $result.FatalError = "$($adUsers.Count) on-prem AD users matched mail '$Smtp'. Expected 0 or 1 -- resolve the duplicate mail attribute before using this feature."
        return [PSCustomObject]$result
    }
    $dns = @($adUsers[0].memberOf)
    $found = @()
    $i = 0
    foreach ($dn in $dns) {
        $i++
        if ($ProgressAction) {
            try { & $ProgressAction $i $dns.Count $dn } catch { }
        }
        try {
            $g = Get-ADGroup -Identity $dn -Server $ADServer -ErrorAction Stop
            $found += $g.Name
        } catch {
            $result.ScanErrors += "$($dn): $($_.Exception.Message)"
        }
        if ($ShouldCancel -and (& $ShouldCancel)) {
            $result.Cancelled = $true
            break
        }
    }
    $result.Memberships = $found
    return [PSCustomObject]$result
}
function Get-MailboxHealthSnapshot {
    param([string]$Smtp)
    $health = [ordered]@{
        LastLogonTime              = $null
        ItemCount                  = $null
        MailboxSize                = $null
        ProhibitSendQuota          = $null
        ProhibitSendReceiveQuota   = $null
        ArchiveEnabled             = $null
        ArchiveSize                = $null
        AccountDisabled            = $null
        ForwardingSmtpAddress      = $null
        ForwardingAddress          = $null
        DeliverToMailboxAndForward = $null
        LitigationHoldEnabled      = $null
        RetentionHoldEnabled       = $null
        InPlaceHoldsCount          = $null
        FullAccessDelegates        = @()
        SendAsDelegates            = @()
        RuleCount                  = $null
        SuspiciousRules            = @()
        OWAEnabled                 = $null
        ActiveSyncEnabled          = $null
        PopEnabled                 = $null
        ImapEnabled                = $null
        MAPIEnabled                = $null
        EwsEnabled                 = $null
        Errors                     = @()
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.LastLogonTime = $stats.LastLogonTime
        $health.ItemCount     = $stats.ItemCount
        $health.MailboxSize   = $stats.TotalItemSize
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (health)')
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.ArchiveEnabled             = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $health.AccountDisabled            = $mbx.AccountDisabled
        $health.ProhibitSendQuota          = $mbx.ProhibitSendQuota
        $health.ProhibitSendReceiveQuota   = $mbx.ProhibitSendReceiveQuota
        $health.ForwardingSmtpAddress      = $mbx.ForwardingSmtpAddress
        $health.ForwardingAddress          = $mbx.ForwardingAddress
        $health.DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
        $health.LitigationHoldEnabled      = $mbx.LitigationHoldEnabled
        $health.RetentionHoldEnabled       = $mbx.RetentionHoldEnabled
        $health.InPlaceHoldsCount          = if ($mbx.InPlaceHolds) { @($mbx.InPlaceHolds).Count } else { 0 }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (quota/forwarding/hold/account status)')
    }
    if ($health.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $health.ArchiveSize = $astats.TotalItemSize
        } catch {
            $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (health)')
        }
    }
    try {
        $fa = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.User -notlike 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false -and $_.Deny -eq $false -and ($_.AccessRights -contains 'FullAccess') })
        $health.FullAccessDelegates = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access delegates)')
    }
    try {
        $sa = @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.Trustee -notlike 'NT AUTHORITY\SELF' -and $_.AccessControlType -eq 'Allow' })
        $health.SendAsDelegates = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As delegates)')
    }
    try {
        $rules = Get-InboxRulesForMailbox -Smtp $Smtp
        if ($rules.Error) {
            $health.Errors += $rules.Error
        } else {
            $health.RuleCount = $rules.Rules.Count
            $health.SuspiciousRules = @(
                $rules.Rules | Where-Object {
                    $_.Enabled -and (
                        $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or
                        $_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage
                    )
                } | ForEach-Object {
                    $actions = @()
                    if ($_.ForwardTo -or $_.ForwardAsAttachmentTo) { $actions += 'Forward' }
                    if ($_.RedirectTo) { $actions += 'Redirect' }
                    if ($_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage) { $actions += 'Delete' }
                    "$($_.Name)  [$($actions -join ', ')]"
                }
            )
        }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule (rule count / suspicious rule scan)')
    }
    try {
        $cas = Invoke-EXOCommand -Name 'Get-CASMailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.OWAEnabled        = $cas.OWAEnabled
        $health.ActiveSyncEnabled = $cas.ActiveSyncEnabled
        $health.PopEnabled        = $cas.PopEnabled
        $health.ImapEnabled       = $cas.ImapEnabled
        $health.MAPIEnabled       = $cas.MAPIEnabled
        $health.EwsEnabled        = $cas.EwsEnabled
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-CASMailbox (protocol access)')
    }
    return [PSCustomObject]$health
}
function Show-MailboxHealthForm {
    param([string]$Smtp)
    $healthForm = New-Object System.Windows.Forms.Form
    $healthForm.Text = "Mailbox Health: $Smtp"
    $healthForm.Size = New-Object System.Drawing.Size(640, 660)
    $healthForm.StartPosition = "CenterParent"
    $healthForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $txtHealth = New-Object System.Windows.Forms.TextBox
    $txtHealth.Multiline = $true
    $txtHealth.ReadOnly = $true
    $txtHealth.ScrollBars = "Vertical"
    $txtHealth.Location = New-Object System.Drawing.Point(12, 12)
    $txtHealth.Size = New-Object System.Drawing.Size(600, 560)
    $txtHealth.Font = New-Object System.Drawing.Font("Consolas", 9)
    $healthForm.Controls.Add($txtHealth)
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 580)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnRefresh)
    $btnCopyHealth = New-Object System.Windows.Forms.Button
    $btnCopyHealth.Text = "Copy"
    $btnCopyHealth.Location = New-Object System.Drawing.Point(120, 580)
    $btnCopyHealth.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnCopyHealth)
    $btnHealthClose = New-Object System.Windows.Forms.Button
    $btnHealthClose.Text = "Close"
    $btnHealthClose.Location = New-Object System.Drawing.Point(228, 580)
    $btnHealthClose.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnHealthClose)
    $copyResetTimer = New-Object System.Windows.Forms.Timer
    $copyResetTimer.Interval = 1200
    $copyResetTimer.Add_Tick({
        $btnCopyHealth.Text = "Copy"
        $copyResetTimer.Stop()
    })
    function Update-HealthFromCurrent {
        $btnRefresh.Enabled = $false
        $txtHealth.Text = "Loading mailbox health..."
        $healthForm.Refresh()
        try {
            $h = Get-MailboxHealthSnapshot -Smtp $Smtp
            $sb = New-Object System.Text.StringBuilder
            $issues = @()
            if ($h.AccountDisabled -eq $true) { $issues += "Account sign-in is disabled (AccountDisabled=True) -- unverified as a Get-Mailbox field, confirm before relying on it." }
            if ($h.ForwardingSmtpAddress -or $h.ForwardingAddress) {
                $fwdTarget = if ($h.ForwardingSmtpAddress) { $h.ForwardingSmtpAddress } else { $h.ForwardingAddress }
                $fwdMode = if ($h.DeliverToMailboxAndForward) { "copy kept in mailbox AND forwarded" } else { "forwarded ONLY -- nothing kept in this mailbox" }
                $issues += "Mail forwarding is configured to '$fwdTarget' ($fwdMode)."
            }
            if ($h.LitigationHoldEnabled -eq $true) { $issues += "Litigation Hold is enabled." }
            if ($h.RetentionHoldEnabled -eq $true) { $issues += "Retention Hold is enabled." }
            if ($h.InPlaceHoldsCount -gt 0) { $issues += "$($h.InPlaceHoldsCount) In-Place/eDiscovery hold(s) applied." }
            if ($h.SuspiciousRules.Count -gt 0) { $issues += "$($h.SuspiciousRules.Count) enabled Inbox rule(s) forward, redirect, or delete mail -- see 'Suspicious Inbox Rules' below." }
            [void]$sb.AppendLine("Mailbox                : $Smtp")
            [void]$sb.AppendLine("=== Potential Issues (review before escalating) ===")
            if ($issues.Count -eq 0) { [void]$sb.AppendLine("  None of the checks below flagged anything.") }
            else { $issues | ForEach-Object { [void]$sb.AppendLine("  - $_") } }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Basics ---")
            [void]$sb.AppendLine("Last Logon             : $($h.LastLogonTime)")
            [void]$sb.AppendLine("Item Count              : $($h.ItemCount)")
            [void]$sb.AppendLine("Mailbox Size            : $($h.MailboxSize)")
            [void]$sb.AppendLine("ProhibitSendQuota       : $($h.ProhibitSendQuota)  (compare to Mailbox Size above)")
            [void]$sb.AppendLine("ProhibitSendReceiveQuota: $($h.ProhibitSendReceiveQuota)")
            [void]$sb.AppendLine("Archive Enabled         : $($h.ArchiveEnabled)")
            if ($h.ArchiveEnabled) { [void]$sb.AppendLine("Archive Size            : $($h.ArchiveSize)") }
            [void]$sb.AppendLine("Account Disabled        : $($h.AccountDisabled)  (unverified Get-Mailbox field -- see script header)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Forwarding & Holds ---")
            [void]$sb.AppendLine("Forwarding SMTP Address : $($h.ForwardingSmtpAddress)")
            [void]$sb.AppendLine("Forwarding Address      : $($h.ForwardingAddress)")
            [void]$sb.AppendLine("Deliver And Forward     : $($h.DeliverToMailboxAndForward)")
            [void]$sb.AppendLine("Litigation Hold         : $($h.LitigationHoldEnabled)")
            [void]$sb.AppendLine("Retention Hold          : $($h.RetentionHoldEnabled)")
            [void]$sb.AppendLine("In-Place Holds          : $($h.InPlaceHoldsCount)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Delegates ---")
            [void]$sb.AppendLine("Full Access grantees ($($h.FullAccessDelegates.Count)):")
            if ($h.FullAccessDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.FullAccessDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("Send As grantees ($($h.SendAsDelegates.Count)):")
            if ($h.SendAsDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.SendAsDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("  (Distribution List memberships are not scanned here -- use the 'DL Memberships' button, since a full-tenant scan is too slow to run on every Health click.)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Inbox Rules ---")
            [void]$sb.AppendLine("Inbox Rule Count        : $($h.RuleCount)")
            [void]$sb.AppendLine("Suspicious Inbox Rules (enabled, forward/redirect/delete):")
            if ($h.SuspiciousRules.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.SuspiciousRules | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Protocol Access (enabled/disabled -- not traffic volume) ---")
            [void]$sb.AppendLine("  OWA        : $($h.OWAEnabled)")
            [void]$sb.AppendLine("  ActiveSync : $($h.ActiveSyncEnabled)")
            [void]$sb.AppendLine("  POP        : $($h.PopEnabled)")
            [void]$sb.AppendLine("  IMAP       : $($h.ImapEnabled)")
            [void]$sb.AppendLine("  MAPI       : $($h.MAPIEnabled)")
            [void]$sb.AppendLine("  EWS        : $($h.EwsEnabled)")
            if ($h.Errors.Count -gt 0) {
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("--- Field errors (partial data above is still valid) ---")
                $h.Errors | ForEach-Object { [void]$sb.AppendLine($_) }
            }
            $txtHealth.Text = $sb.ToString()
        } catch {
            $txtHealth.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Health'
        } finally {
            $btnRefresh.Enabled = $true
        }
    }
    Update-HealthFromCurrent
    $btnRefresh.Add_Click({ Update-HealthFromCurrent })
    $btnCopyHealth.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($txtHealth.Text)
            $btnCopyHealth.Text = "Copied!"
            $copyResetTimer.Stop()
            $copyResetTimer.Start()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Copy to clipboard failed: $($_.Exception.Message)",
                "Copy", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
    })
    $btnHealthClose.Add_Click({ $healthForm.Close() })
    [void]$healthForm.ShowDialog($form)
}
function Show-DLMembershipsForm {
    param([string]$Smtp, [string]$ADServer)
    $dlForm = New-Object System.Windows.Forms.Form
    $dlForm.Text = "DL Memberships: $Smtp"
    $dlForm.Size = New-Object System.Drawing.Size(560, 520)
    $dlForm.StartPosition = "CenterParent"
    $dlForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads this mailbox's on-prem AD account (matched by mail attribute) and its memberOf attribute -- the same backlink EAC's Group Membership pane uses. Direct memberships only, classic DLs / mail-enabled security groups only (not Microsoft 365 Groups)."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(520, 46)
    $dlForm.Controls.Add($lblNote)
    $lblProgress = New-Object System.Windows.Forms.Label
    $lblProgress.Text = "Not started."
    $lblProgress.Location = New-Object System.Drawing.Point(12, 62)
    $lblProgress.Size = New-Object System.Drawing.Size(520, 20)
    $lblProgress.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $dlForm.Controls.Add($lblProgress)
    $txtDL = New-Object System.Windows.Forms.TextBox
    $txtDL.Multiline = $true
    $txtDL.ReadOnly = $true
    $txtDL.ScrollBars = "Vertical"
    $txtDL.Location = New-Object System.Drawing.Point(12, 88)
    $txtDL.Size = New-Object System.Drawing.Size(520, 340)
    $txtDL.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlForm.Controls.Add($txtDL)
    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "Scan"
    $btnScan.Location = New-Object System.Drawing.Point(12, 436)
    $btnScan.Size = New-Object System.Drawing.Size(100, 28)
    $dlForm.Controls.Add($btnScan)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(120, 436)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 28)
    $btnCancel.Enabled = $false
    $dlForm.Controls.Add($btnCancel)
    $btnDLClose = New-Object System.Windows.Forms.Button
    $btnDLClose.Text = "Close"
    $btnDLClose.Location = New-Object System.Drawing.Point(432, 436)
    $btnDLClose.Size = New-Object System.Drawing.Size(100, 28)
    $dlForm.Controls.Add($btnDLClose)
    $script:DLScanCancelRequested = $false
    $btnScan.Add_Click({
        $btnScan.Enabled = $false
        $btnCancel.Enabled = $true
        $txtDL.Text = ""
        $script:DLScanCancelRequested = $false
        $lblProgress.Text = "Looking up on-prem AD account..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $progressCB = {
                param($i, $total, $dn)
                $lblProgress.Text = "Resolving group $i of $total`: $dn"
                [System.Windows.Forms.Application]::DoEvents()
            }
            $cancelCB = { return $script:DLScanCancelRequested }
            $result = Get-MailboxDLMemberships -Smtp $Smtp -ADServer $ADServer -ProgressAction $progressCB -ShouldCancel $cancelCB
            if ($result.FatalError) {
                $lblProgress.Text = "Failed."
                $txtDL.Text = $result.FatalError
            } elseif ($result.NoADAccount) {
                $lblProgress.Text = "No on-prem AD account found."
                $txtDL.Text = "No on-prem AD user account has a mail attribute matching '$Smtp'. This mailbox may be cloud-only (never synced from on-prem AD) -- DL membership can't be read this way for it. If you believe this account IS synced, double-check the Environment/Domain Controller selection at the top of the main window."
            } else {
                $status = if ($result.Cancelled) { "Cancelled -- partial results shown below." } else { "Done." }
                $lblProgress.Text = $status
                $sb = New-Object System.Text.StringBuilder
                [void]$sb.AppendLine("$status")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("Memberships ($($result.Memberships.Count)):")
                if ($result.Memberships.Count -eq 0) { [void]$sb.AppendLine("  (none found)") }
                else { $result.Memberships | Sort-Object | ForEach-Object { [void]$sb.AppendLine("  $_") } }
                if ($result.ScanErrors.Count -gt 0) {
                    [void]$sb.AppendLine("")
                    [void]$sb.AppendLine("Groups that could not be resolved ($($result.ScanErrors.Count)) -- membership above may be incomplete:")
                    $result.ScanErrors | ForEach-Object { [void]$sb.AppendLine("  $_") }
                }
                $txtDL.Text = $sb.ToString()
            }
        } catch {
            $lblProgress.Text = "Failed."
            $txtDL.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'DL Membership Lookup'
        } finally {
            $btnScan.Enabled = $true
            $btnCancel.Enabled = $false
        }
    })
    $btnCancel.Add_Click({
        $script:DLScanCancelRequested = $true
        $btnCancel.Enabled = $false
        $lblProgress.Text = "$($lblProgress.Text)  (cancelling after the current group...)"
    })
    $btnDLClose.Add_Click({ $dlForm.Close() })
    [void]$dlForm.ShowDialog($form)
}
#endregion
#region Actions - mail attributes --------------------------------------------
$script:MailAttribsDisplayOrder = @(
    'Company', 'DisplayName', 'name', 'GivenName', 'Initials', 'Surname',
    'mail', 'mailNickname', 'msExchArchiveName', 'PrimarySmtpAddress',
    'proxyAddresses', 'targetAddress', 'UserPrincipalName'
)
$script:MailAttribsDisplayLabels = @{
    'Company'            = 'Office'
    'GivenName'          = 'FirstName'
    'Surname'            = 'LastName'
    'PrimarySmtpAddress' = 'Primary SMTP'
}
$script:MailAttribsEditableAttributes = @(
    'DisplayName', 'GivenName', 'Initials', 'Surname', 'mail', 'mailNickname',
    'msExchArchiveName', 'PrimarySmtpAddress', 'targetAddress'
)
$script:MailAttribsNeverBlank = @('name', 'PrimarySmtpAddress')
function Get-MailAttribsADUser {
    param(
        [Parameter(Mandatory)] [string]$Smtp,
        [string]$ADServer
    )
    $safeSmtp = $Smtp.Replace("'", "''")
    $params = @{
        Filter      = "mail -eq '$safeSmtp'"
        Properties  = @('Company', 'DisplayName', 'GivenName', 'Initials', 'Surname', 'mail', 'mailNickname',
                         'msExchArchiveName', 'Name', 'proxyAddresses', 'targetAddress',
                         'UserPrincipalName', 'DistinguishedName', 'SamAccountName')
        ErrorAction = 'Stop'
    }
    if ($ADServer) { $params['Server'] = $ADServer }
    $results = @(Get-ADUser @params)
    if ($results.Count -eq 0) {
        throw "No on-prem AD user account has a mail attribute matching '$Smtp'. This mailbox may be cloud-only (never synced from on-prem AD), or the Environment/Domain Controller selection at the top of the main window may need to change."
    }
    if ($results.Count -gt 1) {
        throw "$($results.Count) on-prem AD users matched mail '$Smtp'. Expected 0 or 1 -- resolve the duplicate mail attribute before using this feature."
    }
    return $results[0]
}
function Get-PrimarySmtpFromProxyAddresses {
    param($ProxyAddresses)
    foreach ($addr in @($ProxyAddresses)) {
        if ("$addr" -cmatch '^SMTP:(.+)$') { return $Matches[1] }
    }
    return ''
}
function Set-PrimarySmtpInProxyAddresses {
    param(
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewPrimary,
        [bool]$KeepOldPrimaryAsAlias = $true
    )
    $kept = New-Object System.Collections.Generic.List[string]
    $oldPrimary = $null
    foreach ($addr in @($CurrentProxyAddresses)) {
        $s = "$addr"
        if ($s -cmatch '^SMTP:(.+)$') {
            $oldPrimary = $Matches[1]
            continue
        }
        if ($s -match '^smtp:(.+)$' -and $Matches[1] -ieq $NewPrimary) {
            continue
        }
        [void]$kept.Add($s)
    }
    if ($KeepOldPrimaryAsAlias -and $oldPrimary -and $oldPrimary -ine $NewPrimary) {
        [void]$kept.Add("smtp:$oldPrimary")
    }
    $result = New-Object System.Collections.Generic.List[string]
    [void]$result.Add("SMTP:$NewPrimary")
    foreach ($k in $kept) { [void]$result.Add($k) }
    return $result.ToArray()
}
function Add-SecondarySmtpToProxyAddresses {
    param(
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewSecondary
    )
    foreach ($addr in @($CurrentProxyAddresses)) {
        if ("$addr" -match '^smtp:(.+)$' -and $Matches[1] -ieq $NewSecondary) {
            throw "'$NewSecondary' already exists in proxyAddresses (as $addr)."
        }
    }
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($addr in @($CurrentProxyAddresses)) { [void]$result.Add("$addr") }
    [void]$result.Add("smtp:$NewSecondary")
    return $result.ToArray()
}
function Format-AttributeValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = New-Object System.Collections.Generic.List[string]
        foreach ($item in $Value) { [void]$items.Add("$item") }
        return ($items -join '; ')
    }
    return "$Value"
}
function Format-MailAttribsTable {
    param($ADUserObject)
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Attribute', [string])
    [void]$table.Columns.Add('Value', [string])
    [void]$table.Columns.Add('_Editable', [string])
    [void]$table.Columns.Add('_AttributeKey', [string])
    foreach ($attrName in $script:MailAttribsDisplayOrder) {
        $rawValue = if ($attrName -eq 'PrimarySmtpAddress') {
            Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $ADUserObject.proxyAddresses
        } else {
            $ADUserObject.$attrName
        }
        $label = if ($script:MailAttribsDisplayLabels.ContainsKey($attrName)) { $script:MailAttribsDisplayLabels[$attrName] } else { $attrName }
        $row = $table.NewRow()
        $row['Attribute']     = $label
        $row['Value']         = Format-AttributeValue -Value $rawValue
        $row['_Editable']     = if ($script:MailAttribsEditableAttributes -contains $attrName) { 'Y' } else { 'N' }
        $row['_AttributeKey'] = $attrName
        [void]$table.Rows.Add($row)
    }
    return ,$table
}
function Set-MailAttribsAttribute {
    param(
        [Parameter(Mandatory)] [string]$SamAccountName,
        [Parameter(Mandatory)] [string]$DistinguishedName,
        [Parameter(Mandatory)] [string]$AttributeName,
        [string]$NewValue,
        $CurrentProxyAddresses,
        [bool]$KeepOldPrimaryAsAlias = $true,
        [string]$ADServer
    )
    $commonParams = @{ ErrorAction = 'Stop' }
    if ($ADServer) { $commonParams['Server'] = $ADServer }
    $isBlank = ("$NewValue".Trim() -eq '')
    switch ($AttributeName) {
        'DisplayName' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear displayName @commonParams }
            else { Set-ADUser -Identity $SamAccountName -DisplayName $NewValue @commonParams }
        }
        'GivenName' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear givenName @commonParams }
            else { Set-ADUser -Identity $SamAccountName -GivenName $NewValue @commonParams }
        }
        'Initials' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear initials @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Initials $NewValue @commonParams }
        }
        'Surname' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear sn @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Surname $NewValue @commonParams }
        }
        'mail' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear mail @commonParams }
            else { Set-ADUser -Identity $SamAccountName -EmailAddress $NewValue @commonParams }
        }
        'mailNickname' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear mailNickname @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Replace @{ mailNickname = $NewValue } @commonParams }
        }
        'msExchArchiveName' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear msExchArchiveName @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Replace @{ msExchArchiveName = $NewValue } @commonParams }
        }
        'targetAddress' {
            if ($isBlank) { Set-ADUser -Identity $SamAccountName -Clear targetAddress @commonParams }
            else { Set-ADUser -Identity $SamAccountName -Replace @{ targetAddress = $NewValue } @commonParams }
        }
        'PrimarySmtpAddress' {
            if ($isBlank) { throw "PrimarySmtpAddress cannot be blank." }
            $newProxyAddresses = Set-PrimarySmtpInProxyAddresses -CurrentProxyAddresses $CurrentProxyAddresses -NewPrimary $NewValue.Trim() -KeepOldPrimaryAsAlias $KeepOldPrimaryAsAlias
            Set-ADUser -Identity $SamAccountName -Replace @{ proxyAddresses = $newProxyAddresses } @commonParams
        }
        'name' {
            if ($isBlank) { throw "The object's Name cannot be blank." }
            Rename-ADObject -Identity $DistinguishedName -NewName $NewValue @commonParams
        }
        default { throw "Attribute '$AttributeName' is not handled by Set-MailAttribsAttribute." }
    }
}
function Add-MailAttribsSecondaryProxyAddress {
    param(
        [Parameter(Mandatory)] [string]$SamAccountName,
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewSecondary,
        [string]$ADServer
    )
    $commonParams = @{ ErrorAction = 'Stop' }
    if ($ADServer) { $commonParams['Server'] = $ADServer }
    $newProxyAddresses = Add-SecondarySmtpToProxyAddresses -CurrentProxyAddresses $CurrentProxyAddresses -NewSecondary $NewSecondary
    Set-ADUser -Identity $SamAccountName -Replace @{ proxyAddresses = $newProxyAddresses } @commonParams
}
function Show-MailAttribsEditDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$AttributeName,
        [string]$CurrentValue
    )
    $warningNote = switch ($AttributeName) {
        'name' { 'This renames the AD OBJECT itself (its Common Name / RDN in the directory) via Rename-ADObject -- a structural directory change, distinct from DisplayName. Cannot be left blank.' }
        'PrimarySmtpAddress' { 'This rewrites proxyAddresses: sets this value as the SMTP: (primary) entry. If a different primary already exists, you will be prompted next to Keep it as a secondary alias or Remove it entirely. Does NOT update the separate "mail" attribute -- edit that too if it needs to match. Cannot be left blank.' }
        default { '' }
    }
    $dlgHeight = if ($warningNote -ne '') { 330 } else { 300 }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit Attribute -- $AttributeName"
    $dlg.Size = New-Object System.Drawing.Size(520, $dlgHeight)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblAttr = New-Object System.Windows.Forms.Label
    $lblAttr.Text = "Attribute: $AttributeName"
    $lblAttr.Location = New-Object System.Drawing.Point(12, 12)
    $lblAttr.AutoSize = $true
    $lblAttr.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblAttr)
    $nextY = 42
    if ($warningNote -ne '') {
        $lblWarnNote = New-Object System.Windows.Forms.Label
        $lblWarnNote.Text = $warningNote
        $lblWarnNote.Location = New-Object System.Drawing.Point(12, $nextY)
        $lblWarnNote.Size = New-Object System.Drawing.Size(480, 40)
        $lblWarnNote.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
        $lblWarnNote.ForeColor = [System.Drawing.Color]::DarkOrange
        $dlg.Controls.Add($lblWarnNote)
        $nextY += 46
    }
    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Text = "Current Value:"
    $lblCurrent.Location = New-Object System.Drawing.Point(12, $nextY)
    $lblCurrent.AutoSize = $true
    $dlg.Controls.Add($lblCurrent)
    $nextY += 18
    $txtCurrent = New-Object System.Windows.Forms.TextBox
    $txtCurrent.Location = New-Object System.Drawing.Point(12, $nextY)
    $txtCurrent.Size = New-Object System.Drawing.Size(480, 40)
    $txtCurrent.Multiline = $true
    $txtCurrent.ReadOnly = $true
    $txtCurrent.ScrollBars = "Vertical"
    $txtCurrent.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtCurrent.Text = $CurrentValue
    $txtCurrent.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtCurrent)
    $nextY += 48
    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "New Value (blank only allowed if this field is already blank):"
    $lblNew.Location = New-Object System.Drawing.Point(12, $nextY)
    $lblNew.AutoSize = $true
    $dlg.Controls.Add($lblNew)
    $nextY += 18
    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(12, $nextY)
    $txtNew.Size = New-Object System.Drawing.Size(480, 60)
    $txtNew.Multiline = $true
    $txtNew.ScrollBars = "Vertical"
    $txtNew.Text = $CurrentValue
    $txtNew.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtNew)
    $nextY += 68
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(320, $nextY)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(408, $nextY)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    $result = $dlg.ShowDialog($Owner)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $txtNew.Text
}
function Show-KeepOrRemoveOldPrimaryDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$OldPrimary,
        [string]$NewPrimary
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Previous Primary SMTP Address"
    $dlg.Size = New-Object System.Drawing.Size(460, 210)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $dlg.Tag = 'Cancel'
    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Text = "Current primary SMTP address:`n$OldPrimary`n`nIt is being replaced by:`n$NewPrimary`n`nWhat should I do with the current PrimarySmtpAddress?"
    $lblMsg.Location = New-Object System.Drawing.Point(12, 12)
    $lblMsg.Size = New-Object System.Drawing.Size(430, 100)
    $dlg.Controls.Add($lblMsg)
    $btnKeep = New-Object System.Windows.Forms.Button
    $btnKeep.Text = "Keep as Alias"
    $btnKeep.Location = New-Object System.Drawing.Point(20, 130)
    $btnKeep.Size = New-Object System.Drawing.Size(130, 32)
    $btnKeep.Add_Click({ $dlg.Tag = 'Keep'; $dlg.Close() })
    $dlg.Controls.Add($btnKeep)
    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "Remove"
    $btnRemove.Location = New-Object System.Drawing.Point(165, 130)
    $btnRemove.Size = New-Object System.Drawing.Size(110, 32)
    $btnRemove.Add_Click({ $dlg.Tag = 'Remove'; $dlg.Close() })
    $dlg.Controls.Add($btnRemove)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(290, 130)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Add_Click({ $dlg.Tag = 'Cancel'; $dlg.Close() })
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    [void]$dlg.ShowDialog($Owner)
    return "$($dlg.Tag)"
}
function Show-MailAttribsForm {
    param([string]$Smtp, [string]$ADServer)
    $mailForm = New-Object System.Windows.Forms.Form
    $mailForm.Text = "Mail Attributes: $Smtp"
    $mailForm.Size = New-Object System.Drawing.Size(680, 700)
    $mailForm.StartPosition = "CenterParent"
    $mailForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads this mailbox's on-prem AD account (matched by mail attribute) -- the same fixed attribute set and edit mechanics as the Renaming Script tool. Double-click a row, or select it and click 'Edit Attribute' below, to edit it. 'proxyAddresses' is read-only here -- edit 'Primary SMTP' to change the primary address (you'll be prompted to Keep the old primary as an alias or Remove it), or use 'Add Secondary' below to add another address. 'name', 'UserPrincipalName', and 'Office' are read-only/context-only in this tool."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(640, 58)
    $mailForm.Controls.Add($lblNote)
    $grdMail = New-Object System.Windows.Forms.DataGridView
    $grdMail.Location = New-Object System.Drawing.Point(12, 74)
    $grdMail.Size = New-Object System.Drawing.Size(640, 340)
    $grdMail.ReadOnly = $true
    $grdMail.AllowUserToAddRows = $false
    $grdMail.AllowUserToDeleteRows = $false
    $grdMail.AllowUserToResizeRows = $false
    $grdMail.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdMail.MultiSelect = $false
    $grdMail.AutoGenerateColumns = $false
    $grdMail.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $grdMail.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $grdMail.RowHeadersVisible = $false
    $grdMail.BackgroundColor = [System.Drawing.Color]::White
    $grdMail.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $grdMail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grdMail.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
    $colMailAttribute = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailAttribute.Name = 'Attribute'
    $colMailAttribute.HeaderText = 'Attribute'
    $colMailAttribute.DataPropertyName = 'Attribute'
    $colMailAttribute.Width = 160
    $colMailAttribute.ReadOnly = $true
    [void]$grdMail.Columns.Add($colMailAttribute)
    $colMailValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailValue.Name = 'Value'
    $colMailValue.HeaderText = 'Value'
    $colMailValue.DataPropertyName = 'Value'
    $colMailValue.ReadOnly = $true
    $colMailValue.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    [void]$grdMail.Columns.Add($colMailValue)
    $colMailEditableFlag = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailEditableFlag.Name = '_Editable'
    $colMailEditableFlag.DataPropertyName = '_Editable'
    $colMailEditableFlag.Visible = $false
    [void]$grdMail.Columns.Add($colMailEditableFlag)
    $colMailAttributeKey = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailAttributeKey.Name = '_AttributeKey'
    $colMailAttributeKey.DataPropertyName = '_AttributeKey'
    $colMailAttributeKey.Visible = $false
    [void]$grdMail.Columns.Add($colMailAttributeKey)
    $mailForm.Controls.Add($grdMail)
    $grdMail.DataSource = New-Object System.Windows.Forms.BindingSource
    $mailBindingSource = $grdMail.DataSource
    $lblAddSmtp = New-Object System.Windows.Forms.Label
    $lblAddSmtp.Text = "Add secondary SMTP address (appended as smtp: -- does not replace the primary or any other address):"
    $lblAddSmtp.Location = New-Object System.Drawing.Point(12, 422)
    $lblAddSmtp.Size = New-Object System.Drawing.Size(640, 16)
    $mailForm.Controls.Add($lblAddSmtp)
    $txtAddSmtp = New-Object System.Windows.Forms.TextBox
    $txtAddSmtp.Location = New-Object System.Drawing.Point(12, 440)
    $txtAddSmtp.Size = New-Object System.Drawing.Size(430, 22)
    $txtAddSmtp.Font = New-Object System.Drawing.Font("Consolas", 9)
    $mailForm.Controls.Add($txtAddSmtp)
    $btnAddSmtp = New-Object System.Windows.Forms.Button
    $btnAddSmtp.Text = "Add Secondary"
    $btnAddSmtp.Location = New-Object System.Drawing.Point(452, 439)
    $btnAddSmtp.Size = New-Object System.Drawing.Size(110, 26)
    $mailForm.Controls.Add($btnAddSmtp)
    $lblMailLog = New-Object System.Windows.Forms.Label
    $lblMailLog.Text = "Log:"
    $lblMailLog.Location = New-Object System.Drawing.Point(12, 474)
    $lblMailLog.AutoSize = $true
    $mailForm.Controls.Add($lblMailLog)
    $txtMailLog = New-Object System.Windows.Forms.TextBox
    $txtMailLog.Multiline = $true
    $txtMailLog.ReadOnly = $true
    $txtMailLog.ScrollBars = "Vertical"
    $txtMailLog.Location = New-Object System.Drawing.Point(12, 494)
    $txtMailLog.Size = New-Object System.Drawing.Size(640, 110)
    $txtMailLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtMailLog.BackColor = [System.Drawing.Color]::WhiteSmoke
    $mailForm.Controls.Add($txtMailLog)
    $btnMailRefresh = New-Object System.Windows.Forms.Button
    $btnMailRefresh.Text = "Refresh"
    $btnMailRefresh.Location = New-Object System.Drawing.Point(12, 616)
    $btnMailRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $mailForm.Controls.Add($btnMailRefresh)
    $btnMailEdit = New-Object System.Windows.Forms.Button
    $btnMailEdit.Text = "Edit Attribute"
    $btnMailEdit.Location = New-Object System.Drawing.Point(122, 616)
    $btnMailEdit.Size = New-Object System.Drawing.Size(140, 28)
    $btnMailEdit.Enabled = $false
    $btnMailEdit.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $mailForm.Controls.Add($btnMailEdit)
    $btnMailClose = New-Object System.Windows.Forms.Button
    $btnMailClose.Text = "Close"
    $btnMailClose.Location = New-Object System.Drawing.Point(552, 616)
    $btnMailClose.Size = New-Object System.Drawing.Size(100, 28)
    $mailForm.Controls.Add($btnMailClose)
    $script:MailAttribsCurrentADUser = $null
    function Update-MailAttribsGrid {
        param($ADUserObject)
        $table = Format-MailAttribsTable -ADUserObject $ADUserObject
        $mailBindingSource.DataSource = $null
        $mailBindingSource.DataSource = $table
        $grdMail.Refresh()
    }
    function Update-MailAttribsFromCurrent {
        $btnMailRefresh.Enabled = $false
        $btnAddSmtp.Enabled = $false
        $txtMailLog.Text = "Loading mail attributes..."
        $mailForm.Refresh()
        try {
            $script:MailAttribsCurrentADUser = Get-MailAttribsADUser -Smtp $Smtp -ADServer $ADServer
            Update-MailAttribsGrid -ADUserObject $script:MailAttribsCurrentADUser
            $txtMailLog.Text = "Loaded mail attributes for SamAccountName '$($script:MailAttribsCurrentADUser.SamAccountName)'."
        } catch {
            $script:MailAttribsCurrentADUser = $null
            $mailBindingSource.DataSource = $null
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
        }
    }
    Update-MailAttribsFromCurrent
    $btnMailRefresh.Add_Click({ Update-MailAttribsFromCurrent })
    function Invoke-MailAttribsEditSelectedRow {
        if ($grdMail.SelectedRows.Count -eq 0 -or $null -eq $script:MailAttribsCurrentADUser) { return }
        $row = $grdMail.SelectedRows[0]
        $attrLabel = "$($row.Cells['Attribute'].Value)"
        $attrName  = "$($row.Cells['_AttributeKey'].Value)"
        if ($script:MailAttribsEditableAttributes -notcontains $attrName) { return }
        $btnMailRefresh.Enabled = $false
        $btnMailEdit.Enabled = $false
        $btnAddSmtp.Enabled = $false
        try {
            $oldValue = "$($row.Cells['Value'].Value)"
            $newValue = Show-MailAttribsEditDialog -Owner $mailForm -AttributeName $attrLabel -CurrentValue $oldValue
            if ($null -eq $newValue) {
                $txtMailLog.Text = "Edit cancelled."
                return
            }
            if ($script:MailAttribsNeverBlank -contains $attrName -and $newValue.Trim() -eq '') {
                [void][System.Windows.Forms.MessageBox]::Show("'$attrLabel' cannot be blank.", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $oldIsBlank = ("$oldValue".Trim() -eq '')
            $newIsBlank = $newValue.Trim() -eq ''
            if ($newIsBlank -and -not $oldIsBlank) {
                [void][System.Windows.Forms.MessageBox]::Show("'$attrLabel' currently has a value and cannot be saved blank. Enter a replacement value instead.", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $keepOldPrimaryAsAlias = $true
            if ($attrName -eq 'PrimarySmtpAddress') {
                $oldPrimaryTrim = $oldValue.Trim()
                $newPrimaryTrim = $newValue.Trim()
                if ($oldPrimaryTrim -ne '' -and -not ($oldPrimaryTrim -ieq $newPrimaryTrim)) {
                    $choice = Show-KeepOrRemoveOldPrimaryDialog -Owner $mailForm -OldPrimary $oldPrimaryTrim -NewPrimary $newPrimaryTrim
                    if ($choice -ne 'Keep' -and $choice -ne 'Remove') {
                        $txtMailLog.Text = "Edit cancelled."
                        return
                    }
                    $keepOldPrimaryAsAlias = ($choice -eq 'Keep')
                }
            }
            $displayNew = if ($newValue.Trim() -eq '') { '(cleared)' } else { $newValue.Replace("`r`n", '; ') }
            $confirmExtra = switch ($attrName) {
                'PrimarySmtpAddress' {
                    if ($keepOldPrimaryAsAlias) { "`n`nThis rewrites the proxyAddresses list: sets this as SMTP: (primary) and demotes the previous primary to smtp: (secondary alias)." }
                    else { "`n`nThis rewrites the proxyAddresses list: sets this as SMTP: (primary) and REMOVES the previous primary address from proxyAddresses entirely." }
                }
                default { '' }
            }
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Change '$attrLabel' for $Smtp ($($script:MailAttribsCurrentADUser.SamAccountName)):`n`nFrom: $oldValue`nTo:   $displayNew$confirmExtra",
                'Confirm Attribute Edit',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtMailLog.Text = "Cancelled."
                return
            }
            Set-MailAttribsAttribute -SamAccountName $script:MailAttribsCurrentADUser.SamAccountName -DistinguishedName $script:MailAttribsCurrentADUser.DistinguishedName `
                -AttributeName $attrName -NewValue $newValue -CurrentProxyAddresses $script:MailAttribsCurrentADUser.proxyAddresses `
                -KeepOldPrimaryAsAlias $keepOldPrimaryAsAlias -ADServer $ADServer
            $refreshed = Get-MailAttribsADUser -Smtp $Smtp -ADServer $ADServer
            $script:MailAttribsCurrentADUser = $refreshed
            $afterValue = if ($attrName -eq 'PrimarySmtpAddress') {
                Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $refreshed.proxyAddresses
            } else {
                Format-AttributeValue -Value ($refreshed.$attrName)
            }
            Update-MailAttribsGrid -ADUserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE ---")
            [void]$log.AppendLine("$attrLabel = $oldValue")
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine("$attrLabel = $afterValue")
            $txtMailLog.Text = $log.ToString().TrimEnd()
            $logVerb = if ($attrName -eq 'PrimarySmtpAddress') {
                "Edited Attribute $attrName ($(if($keepOldPrimaryAsAlias){'kept old primary as alias'}else{'removed old primary'}))"
            } else {
                "Edited Attribute $attrName"
            }
            Write-CompletedLog -Task "$logVerb $($script:MailAttribsCurrentADUser.SamAccountName) (Mail Attribs; $Smtp)"
        } catch {
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs edit'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
            $hasRow = ($grdMail.SelectedRows.Count -gt 0 -and $null -ne $script:MailAttribsCurrentADUser)
            $isEditable = $false
            if ($hasRow) { $isEditable = ("$($grdMail.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
            $btnMailEdit.Enabled = ($hasRow -and $isEditable)
        }
    }
    $grdMail.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $grdMail.ClearSelection()
        $grdMail.Rows[$e.RowIndex].Selected = $true
        Invoke-MailAttribsEditSelectedRow
    })
    $grdMail.Add_SelectionChanged({
        $hasRow = ($grdMail.SelectedRows.Count -gt 0 -and $null -ne $script:MailAttribsCurrentADUser)
        $isEditable = $false
        if ($hasRow) { $isEditable = ("$($grdMail.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
        $btnMailEdit.Enabled = ($hasRow -and $isEditable)
    })
    $btnMailEdit.Add_Click({ Invoke-MailAttribsEditSelectedRow })
    $btnAddSmtp.Add_Click({
        if ($null -eq $script:MailAttribsCurrentADUser) {
            [void][System.Windows.Forms.MessageBox]::Show("Load mail attributes first (click Refresh).", 'No User Loaded', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $btnMailRefresh.Enabled = $false
        $btnAddSmtp.Enabled = $false
        try {
            $newAddr = $txtAddSmtp.Text.Trim()
            if ($newAddr -eq '') {
                [void][System.Windows.Forms.MessageBox]::Show("Enter an address to add.", 'Missing Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            if ($newAddr -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                [void][System.Windows.Forms.MessageBox]::Show("'$newAddr' does not look like a valid email address (expected user@domain.tld).", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $beforeValue = Format-AttributeValue -Value $script:MailAttribsCurrentADUser.proxyAddresses
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Add 'smtp:$newAddr' as a secondary proxy address for $Smtp ($($script:MailAttribsCurrentADUser.SamAccountName))?`n`nThis does not change the primary address or remove any existing address.",
                'Confirm Add Secondary SMTP Address',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtMailLog.Text = "Cancelled."
                return
            }
            Add-MailAttribsSecondaryProxyAddress -SamAccountName $script:MailAttribsCurrentADUser.SamAccountName -CurrentProxyAddresses $script:MailAttribsCurrentADUser.proxyAddresses `
                -NewSecondary $newAddr -ADServer $ADServer
            $refreshed = Get-MailAttribsADUser -Smtp $Smtp -ADServer $ADServer
            $script:MailAttribsCurrentADUser = $refreshed
            $afterValue = Format-AttributeValue -Value $refreshed.proxyAddresses
            Update-MailAttribsGrid -ADUserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE (proxyAddresses) ---")
            [void]$log.AppendLine($beforeValue)
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine($afterValue)
            $txtMailLog.Text = $log.ToString().TrimEnd()
            Write-CompletedLog -Task "Added Secondary SMTP $($script:MailAttribsCurrentADUser.SamAccountName) (Mail Attribs; $Smtp; smtp:$newAddr)"
            $txtAddSmtp.Text = ''
        } catch {
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add Secondary SMTP'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
        }
    })
    $btnMailClose.Add_Click({ $mailForm.Close() })
    [void]$mailForm.ShowDialog($form)
}
#endregion
#region Actions - all attributes ----------------------------------------------
$script:AllAttribsEditableAttributes = @(
    'Title', 'Department', 'Description', 'Company', 'Division',
    'Office', 'OfficePhone', 'MobilePhone', 'Fax', 'HomePhone',
    'StreetAddress', 'City', 'State', 'PostalCode', 'POBox',
    'EmployeeID', 'EmployeeNumber', 'Organization', 'Initials', 'HomePage'
)
$script:AllAttribsExcludedProperties = @(
    'PropertyNames', 'AddedProperties', 'RemovedProperties',
    'ModifiedProperties', 'PropertyCount', 'nTSecurityDescriptor'
)
function Get-AllAttribsADUser {
    param(
        [Parameter(Mandatory)] [string]$Smtp,
        [string]$ADServer
    )
    $safeSmtp = $Smtp.Replace("'", "''")
    $params = @{
        Filter      = "mail -eq '$safeSmtp'"
        Properties  = '*'
        ErrorAction = 'Stop'
    }
    if ($ADServer) { $params['Server'] = $ADServer }
    $results = @(Get-ADUser @params)
    if ($results.Count -eq 0) {
        throw "No on-prem AD user account has a mail attribute matching '$Smtp'. This mailbox may be cloud-only (never synced from on-prem AD), or the Environment/Domain Controller selection at the top of the main window may need to change."
    }
    if ($results.Count -gt 1) {
        throw "$($results.Count) on-prem AD users matched mail '$Smtp'. Expected 0 or 1 -- resolve the duplicate mail attribute before using this feature."
    }
    return $results[0]
}
function Format-AllAttribsValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [byte[]]) { return '<binary data>' }
    if ($Value -is [System.Security.Principal.SecurityIdentifier]) { return $Value.Value }
    if ($Value -is [guid]) { return $Value.ToString() }
    if ($Value -is [datetime]) { return ('{0:yyyy-MM-dd HH:mm:ss}' -f $Value) }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = New-Object System.Collections.Generic.List[string]
        foreach ($item in $Value) {
            if ($item -is [byte]) { return '<binary data>' }
            [void]$items.Add("$item")
        }
        return ($items -join '; ')
    }
    return "$Value"
}
function Format-AllAttribsTable {
    param($ADUserObject)
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Attribute', [string])
    [void]$table.Columns.Add('Value', [string])
    [void]$table.Columns.Add('_Editable', [string])
    $props = $ADUserObject.PSObject.Properties |
        Where-Object { $script:AllAttribsExcludedProperties -notcontains $_.Name } |
        Sort-Object Name
    foreach ($p in $props) {
        $row = $table.NewRow()
        $row['Attribute'] = $p.Name
        $row['Value']     = Format-AllAttribsValue -Value $p.Value
        $row['_Editable'] = if ($script:AllAttribsEditableAttributes -contains $p.Name) { 'Y' } else { 'N' }
        [void]$table.Rows.Add($row)
    }
    return ,$table
}
function Set-AllAttribsAttribute {
    param(
        [Parameter(Mandatory)] [string]$SamAccountName,
        [Parameter(Mandatory)] [string]$AttributeName,
        [string]$NewValue,
        [string]$ADServer
    )
    $params = @{
        Identity    = $SamAccountName
        ErrorAction = 'Stop'
    }
    if ($ADServer) { $params['Server'] = $ADServer }
    if ("$NewValue".Trim() -eq '') {
        $params['Clear'] = $AttributeName
    } else {
        $params[$AttributeName] = $NewValue
    }
    Set-ADUser @params
}
function Show-AllAttribsEditDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$AttributeName,
        [string]$CurrentValue
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit Attribute -- $AttributeName"
    $dlg.Size = New-Object System.Drawing.Size(480, 240)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblAttr = New-Object System.Windows.Forms.Label
    $lblAttr.Text = "Attribute: $AttributeName"
    $lblAttr.Location = New-Object System.Drawing.Point(12, 12)
    $lblAttr.AutoSize = $true
    $lblAttr.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblAttr)
    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Text = "Current Value:"
    $lblCurrent.Location = New-Object System.Drawing.Point(12, 42)
    $lblCurrent.AutoSize = $true
    $dlg.Controls.Add($lblCurrent)
    $txtCurrent = New-Object System.Windows.Forms.TextBox
    $txtCurrent.Location = New-Object System.Drawing.Point(12, 60)
    $txtCurrent.Size = New-Object System.Drawing.Size(440, 22)
    $txtCurrent.ReadOnly = $true
    $txtCurrent.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtCurrent.Text = $CurrentValue
    $txtCurrent.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtCurrent)
    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "New Value (leave blank to clear the attribute):"
    $lblNew.Location = New-Object System.Drawing.Point(12, 92)
    $lblNew.AutoSize = $true
    $dlg.Controls.Add($lblNew)
    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(12, 110)
    $txtNew.Size = New-Object System.Drawing.Size(440, 22)
    $txtNew.Text = $CurrentValue
    $txtNew.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtNew)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(280, 150)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(368, 150)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    $result = $dlg.ShowDialog($Owner)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $txtNew.Text
}
function Show-AllAttribsForm {
    param([string]$Smtp, [string]$ADServer)
    $allForm = New-Object System.Windows.Forms.Form
    $allForm.Text = "All Attributes: $Smtp"
    $allForm.Size = New-Object System.Drawing.Size(680, 750)
    $allForm.StartPosition = "CenterParent"
    $allForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads every populated attribute (Get-ADUser -Properties *) on this mailbox's on-prem AD account (matched by mail attribute). Click a column header to sort. Double-click a row, or select it and click 'Edit Attribute' (enabled only for the curated editable set below), to edit it -- everything else is view-only."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(640, 40)
    $allForm.Controls.Add($lblNote)
    $grdAll = New-Object System.Windows.Forms.DataGridView
    $grdAll.Location = New-Object System.Drawing.Point(12, 58)
    $grdAll.Size = New-Object System.Drawing.Size(640, 440)
    $grdAll.ReadOnly = $true
    $grdAll.AllowUserToAddRows = $false
    $grdAll.AllowUserToDeleteRows = $false
    $grdAll.AllowUserToResizeRows = $false
    $grdAll.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdAll.MultiSelect = $false
    $grdAll.AutoGenerateColumns = $false
    $grdAll.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $grdAll.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $grdAll.RowHeadersVisible = $false
    $grdAll.BackgroundColor = [System.Drawing.Color]::White
    $grdAll.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $grdAll.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grdAll.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
    $colAllAttribute = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAllAttribute.Name = 'Attribute'
    $colAllAttribute.HeaderText = 'Attribute'
    $colAllAttribute.DataPropertyName = 'Attribute'
    $colAllAttribute.Width = 260
    $colAllAttribute.ReadOnly = $true
    [void]$grdAll.Columns.Add($colAllAttribute)
    $colAllValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAllValue.Name = 'Value'
    $colAllValue.HeaderText = 'Value'
    $colAllValue.DataPropertyName = 'Value'
    $colAllValue.ReadOnly = $true
    $colAllValue.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    [void]$grdAll.Columns.Add($colAllValue)
    $colAllEditableFlag = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAllEditableFlag.Name = '_Editable'
    $colAllEditableFlag.DataPropertyName = '_Editable'
    $colAllEditableFlag.Visible = $false
    [void]$grdAll.Columns.Add($colAllEditableFlag)
    $allForm.Controls.Add($grdAll)
    $grdAll.DataSource = New-Object System.Windows.Forms.BindingSource
    $allBindingSource = $grdAll.DataSource
    $lblDupeNote = New-Object System.Windows.Forms.Label
    $lblDupeNote.Text = "Note: some values appear twice -- once under a friendly name (e.g. Office) and once under the raw LDAP name (e.g. physicalDeliveryOfficeName). Only the friendly row is editable."
    $lblDupeNote.Location = New-Object System.Drawing.Point(12, 502)
    $lblDupeNote.Size = New-Object System.Drawing.Size(640, 16)
    $lblDupeNote.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 7.5, [System.Drawing.FontStyle]::Italic)
    $lblDupeNote.ForeColor = [System.Drawing.Color]::DimGray
    $allForm.Controls.Add($lblDupeNote)
    $btnAllRefresh = New-Object System.Windows.Forms.Button
    $btnAllRefresh.Text = "Refresh"
    $btnAllRefresh.Location = New-Object System.Drawing.Point(12, 522)
    $btnAllRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $allForm.Controls.Add($btnAllRefresh)
    $btnAllEdit = New-Object System.Windows.Forms.Button
    $btnAllEdit.Text = "Edit Attribute"
    $btnAllEdit.Location = New-Object System.Drawing.Point(122, 522)
    $btnAllEdit.Size = New-Object System.Drawing.Size(140, 28)
    $btnAllEdit.Enabled = $false
    $btnAllEdit.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $allForm.Controls.Add($btnAllEdit)
    $btnAllClose = New-Object System.Windows.Forms.Button
    $btnAllClose.Text = "Close"
    $btnAllClose.Location = New-Object System.Drawing.Point(552, 522)
    $btnAllClose.Size = New-Object System.Drawing.Size(100, 28)
    $allForm.Controls.Add($btnAllClose)
    $lblAllLog = New-Object System.Windows.Forms.Label
    $lblAllLog.Text = "Log:"
    $lblAllLog.Location = New-Object System.Drawing.Point(12, 558)
    $lblAllLog.AutoSize = $true
    $allForm.Controls.Add($lblAllLog)
    $txtAllLog = New-Object System.Windows.Forms.TextBox
    $txtAllLog.Multiline = $true
    $txtAllLog.ReadOnly = $true
    $txtAllLog.ScrollBars = "Vertical"
    $txtAllLog.Location = New-Object System.Drawing.Point(12, 578)
    $txtAllLog.Size = New-Object System.Drawing.Size(640, 110)
    $txtAllLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtAllLog.BackColor = [System.Drawing.Color]::WhiteSmoke
    $allForm.Controls.Add($txtAllLog)
    $script:AllAttribsCurrentADUser = $null
    function Update-AllAttribsGrid {
        param($ADUserObject)
        $table = Format-AllAttribsTable -ADUserObject $ADUserObject
        $allBindingSource.DataSource = $null
        $allBindingSource.DataSource = $table
        $grdAll.Refresh()
    }
    function Update-AllAttribsFromCurrent {
        $btnAllRefresh.Enabled = $false
        $btnAllEdit.Enabled = $false
        $txtAllLog.Text = "Loading all attributes..."
        $allForm.Refresh()
        try {
            $script:AllAttribsCurrentADUser = Get-AllAttribsADUser -Smtp $Smtp -ADServer $ADServer
            Update-AllAttribsGrid -ADUserObject $script:AllAttribsCurrentADUser
            $txtAllLog.Text = "Loaded attributes for SamAccountName '$($script:AllAttribsCurrentADUser.SamAccountName)'."
        } catch {
            $script:AllAttribsCurrentADUser = $null
            $allBindingSource.DataSource = $null
            $txtAllLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'All Attribs'
        } finally {
            $btnAllRefresh.Enabled = $true
        }
    }
    Update-AllAttribsFromCurrent
    $btnAllRefresh.Add_Click({ Update-AllAttribsFromCurrent })
    $grdAll.Add_SelectionChanged({
        $hasRow = ($grdAll.SelectedRows.Count -gt 0 -and $null -ne $script:AllAttribsCurrentADUser)
        $isEditable = $false
        if ($hasRow) { $isEditable = ("$($grdAll.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
        $btnAllEdit.Enabled = ($hasRow -and $isEditable)
    })
    function Invoke-AllAttribsEditSelectedRow {
        if ($grdAll.SelectedRows.Count -eq 0 -or $null -eq $script:AllAttribsCurrentADUser) { return }
        $row = $grdAll.SelectedRows[0]
        $attrName = "$($row.Cells['Attribute'].Value)"
        if ($script:AllAttribsEditableAttributes -notcontains $attrName) { return }
        $oldValue = "$($row.Cells['Value'].Value)"
        $btnAllEdit.Enabled = $false
        $btnAllRefresh.Enabled = $false
        try {
            $newValue = Show-AllAttribsEditDialog -Owner $allForm -AttributeName $attrName -CurrentValue $oldValue
            if ($null -eq $newValue) {
                $txtAllLog.Text = "Edit cancelled."
                return
            }
            $displayNew = if ($newValue.Trim() -eq '') { '(cleared)' } else { $newValue }
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Change '$attrName' for $Smtp ($($script:AllAttribsCurrentADUser.SamAccountName)):`n`nFrom: $oldValue`nTo:   $displayNew",
                'Confirm Attribute Edit',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtAllLog.Text = "Cancelled."
                return
            }
            Set-AllAttribsAttribute -SamAccountName $script:AllAttribsCurrentADUser.SamAccountName -AttributeName $attrName -NewValue $newValue -ADServer $ADServer
            $refreshed = Get-AllAttribsADUser -Smtp $Smtp -ADServer $ADServer
            $script:AllAttribsCurrentADUser = $refreshed
            $afterValue = Format-AllAttribsValue -Value ($refreshed.PSObject.Properties[$attrName].Value)
            Update-AllAttribsGrid -ADUserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE ---")
            [void]$log.AppendLine("$attrName = $oldValue")
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine("$attrName = $afterValue")
            $txtAllLog.Text = $log.ToString().TrimEnd()
            Write-CompletedLog -Task "Edited Attribute $attrName $($script:AllAttribsCurrentADUser.SamAccountName) (All Attribs; $Smtp)"
        } catch {
            $txtAllLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'All Attribs edit'
        } finally {
            $btnAllRefresh.Enabled = $true
            $hasRow = ($grdAll.SelectedRows.Count -gt 0 -and $null -ne $script:AllAttribsCurrentADUser)
            $isEditable = $false
            if ($hasRow) { $isEditable = ("$($grdAll.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
            $btnAllEdit.Enabled = ($hasRow -and $isEditable)
        }
    }
    $btnAllEdit.Add_Click({ Invoke-AllAttribsEditSelectedRow })
    $grdAll.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $grdAll.ClearSelection()
        $grdAll.Rows[$e.RowIndex].Selected = $true
        Invoke-AllAttribsEditSelectedRow
    })
    $btnAllClose.Add_Click({ $allForm.Close() })
    [void]$allForm.ShowDialog($form)
}
#endregion
#region Actions - message trace ---------------------------------------------------
function Invoke-MessageTrace {
    param(
        [string]$SenderAddress,
        [string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [string]$Subject,
        [string]$Status
    )
    # 5000 is a deliberate cap, so Exchange warns "There are more results
    # available" whenever a busy window fills it. Silenced because that warning
    # goes to the PowerShell stream, where nobody using this window will see it
    # -- the honest signal is the count on screen, and exactly 5000 means there
    # was more.
    #
    # WarningAction sits NEXT TO ResultSize deliberately. Comment stripping
    # blanks lines rather than removing them, so an explanation in between puts
    # the two eight lines apart and the check that looks +/-3 lines never sees
    # it. A rule you satisfy has to be satisfied where it looks.
    $params = @{
        StartDate   = $StartDate
        EndDate     = $EndDate
        ResultSize    = 5000
        WarningAction = 'SilentlyContinue'
        ErrorAction   = 'Stop'
    }
    if ($SenderAddress)    { $params.SenderAddress    = $SenderAddress }
    if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }
    if ($Subject) {
        $params.Subject           = $Subject
        $params.SubjectFilterType = 'Contains'
    }
    if ($Status) { $params.Status = $Status }
    Invoke-EXOCommand -Name 'Get-MessageTraceV2' -Parameters $params
}
function Invoke-MessageTraceDetail {
    param(
        [Parameter(Mandatory)][guid]$MessageTraceId,
        [Parameter(Mandatory)][string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )
    Invoke-EXOCommand -Name 'Get-MessageTraceDetailV2' -Parameters @{
        MessageTraceId   = $MessageTraceId
        RecipientAddress = $RecipientAddress
        StartDate        = $StartDate
        EndDate           = $EndDate
        ErrorAction       = 'Stop'
    }
}
function Find-QuarantineMatch {
    param(
        [string]$MessageId,
        [string]$RecipientAddress,
        [datetime]$StartReceivedDate,
        [datetime]$EndReceivedDate
    )
    if (-not $MessageId) {
        throw "This trace result has no Message-ID value to match against quarantine -- cannot look this message up automatically."
    }
    @(Invoke-EXOCommand -Name 'Get-QuarantineMessage' -Parameters @{
        MessageId         = $MessageId
        RecipientAddress  = $RecipientAddress
        StartReceivedDate = $StartReceivedDate
        EndReceivedDate   = $EndReceivedDate
        ErrorAction       = 'Stop'
    })
}
function Invoke-QuarantineRelease {
    param([Parameter(Mandatory)]$QuarantineIdentity)
    Invoke-EXOCommand -Name 'Release-QuarantineMessage' -Parameters @{
        Identity     = $QuarantineIdentity
        ReleaseToAll = $true
        Confirm      = $false
        ErrorAction  = 'Stop'
    } | Out-Null
}
$script:TraceTimeZoneIds = [ordered]@{
    'Eastern'  = 'Eastern Standard Time'
    'Central'  = 'Central Standard Time'
    'Mountain' = 'Mountain Standard Time'
    'Pacific'  = 'Pacific Standard Time'
    'UTC'      = 'UTC'
}
function Get-SelectedTraceTimeZone {
    param([string]$Label)
    $id = $script:TraceTimeZoneIds[$Label]
    if (-not $id) { $id = 'UTC' }
    try {
        return [System.TimeZoneInfo]::FindSystemTimeZoneById($id)
    } catch {
        return [System.TimeZoneInfo]::Utc
    }
}
function ConvertTo-TraceDisplayTime {
    param($UtcValue, [System.TimeZoneInfo]$TimeZone)
    if (-not $UtcValue) { return $UtcValue }
    try {
        $forced = [DateTime]::SpecifyKind([DateTime]$UtcValue, [DateTimeKind]::Utc)
        return [System.TimeZoneInfo]::ConvertTimeFromUtc($forced, $TimeZone)
    } catch {
        return $UtcValue
    }
}
function Format-TraceDetail {
    param($Details, [System.TimeZoneInfo]$TimeZone)
    $sb = New-Object System.Text.StringBuilder
    $rows = @($Details)
    if ($rows.Count -eq 0) {
        [void]$sb.AppendLine("(no detail records returned)")
        return $sb.ToString()
    }
    foreach ($d in $rows) {
        $displayDate = if ($TimeZone) { ConvertTo-TraceDisplayTime -UtcValue $d.Date -TimeZone $TimeZone } else { $d.Date }
        [void]$sb.AppendLine("$displayDate  [$($d.Event)]  $($d.Action)")
        if ($d.Detail) { [void]$sb.AppendLine("    $($d.Detail)") }
    }
    return $sb.ToString()
}
function Show-MessageTraceForm {
    param([string]$Smtp)
    $traceForm = New-Object System.Windows.Forms.Form
    $traceForm.Text = "Message Trace: $Smtp"
    $traceForm.Size = New-Object System.Drawing.Size(1014, 720)
    $traceForm.StartPosition = "CenterParent"
    $traceForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $grpDir = New-Object System.Windows.Forms.GroupBox
    $grpDir.Text = "Direction"
    $grpDir.Location = New-Object System.Drawing.Point(12, 12)
    $grpDir.Size = New-Object System.Drawing.Size(240, 50)
    $traceForm.Controls.Add($grpDir)
    $radSent = New-Object System.Windows.Forms.RadioButton
    $radSent.Text = "Sent (from this mailbox)"
    $radSent.Location = New-Object System.Drawing.Point(10, 16)
    $radSent.AutoSize = $true
    $radSent.Checked = $true
    $grpDir.Controls.Add($radSent)
    $radReceived = New-Object System.Windows.Forms.RadioButton
    $radReceived.Text = "Received (to this mailbox)"
    $radReceived.Location = New-Object System.Drawing.Point(10, 30)
    $radReceived.AutoSize = $true
    $grpDir.Controls.Add($radReceived)
    $lblTraceNote = New-Object System.Windows.Forms.Label
    $lblTraceNote.Text = "Read-only -- no objects are modified. Exchange Online only (Get-MessageTraceV2 / Get-MessageTraceDetailV2), routed through Invoke-EXOCommand."
    $lblTraceNote.Location = New-Object System.Drawing.Point(264, 12)
    $lblTraceNote.Size = New-Object System.Drawing.Size(488, 46)
    $traceForm.Controls.Add($lblTraceNote)
    $lblTZ = New-Object System.Windows.Forms.Label
    $lblTZ.Text = "Show time in:"
    $lblTZ.Location = New-Object System.Drawing.Point(766, 14)
    $lblTZ.AutoSize = $true
    $traceForm.Controls.Add($lblTZ)
    $cboTimeZone = New-Object System.Windows.Forms.ComboBox
    $cboTimeZone.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Eastern', 'Central', 'Mountain', 'Pacific', 'UTC') | ForEach-Object { [void]$cboTimeZone.Items.Add($_) }
    $cboTimeZone.SelectedItem = 'Central'
    $cboTimeZone.Location = New-Object System.Drawing.Point(766, 34)
    $cboTimeZone.Size = New-Object System.Drawing.Size(200, 20)
    $traceForm.Controls.Add($cboTimeZone)
    $lblNarrow = New-Object System.Windows.Forms.Label
    $lblNarrow.Text = "To (optional, narrows recipient):"
    $lblNarrow.Location = New-Object System.Drawing.Point(12, 72)
    $lblNarrow.AutoSize = $true
    $traceForm.Controls.Add($lblNarrow)
    $txtNarrow = New-Object System.Windows.Forms.TextBox
    $txtNarrow.Location = New-Object System.Drawing.Point(270, 69)
    $txtNarrow.Size = New-Object System.Drawing.Size(220, 20)
    $txtNarrow.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtNarrow)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Status:"
    $lblStatus.Location = New-Object System.Drawing.Point(510, 72)
    $lblStatus.AutoSize = $true
    $traceForm.Controls.Add($lblStatus)
    $cboStatus = New-Object System.Windows.Forms.ComboBox
    $cboStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('All Messages', 'Delivered', 'Failed', 'Quarantined', 'FilteredAsSpam', 'Pending', 'Expanded', 'GettingStatus') |
        ForEach-Object { [void]$cboStatus.Items.Add($_) }
    $cboStatus.SelectedIndex = 0
    $cboStatus.Location = New-Object System.Drawing.Point(562, 69)
    $cboStatus.Size = New-Object System.Drawing.Size(180, 20)
    $traceForm.Controls.Add($cboStatus)
    $lblSubject = New-Object System.Windows.Forms.Label
    $lblSubject.Text = "Subject contains (optional):"
    $lblSubject.Location = New-Object System.Drawing.Point(12, 100)
    $lblSubject.AutoSize = $true
    $traceForm.Controls.Add($lblSubject)
    $txtSubject = New-Object System.Windows.Forms.TextBox
    $txtSubject.Location = New-Object System.Drawing.Point(270, 97)
    $txtSubject.Size = New-Object System.Drawing.Size(220, 20)
    $txtSubject.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtSubject)
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = "Search From:"
    $lblStart.Location = New-Object System.Drawing.Point(12, 129)
    $lblStart.AutoSize = $true
    $traceForm.Controls.Add($lblStart)
    $dtpStart = New-Object System.Windows.Forms.DateTimePicker
    $dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtpStart.Location = New-Object System.Drawing.Point(100, 125)
    $dtpStart.Size = New-Object System.Drawing.Size(120, 20)
    $dtpStart.MinDate = (Get-Date).Date.AddDays(-30)
    $dtpStart.MaxDate = (Get-Date).Date
    $dtpStart.Value   = (Get-Date).Date.AddDays(-7)
    $traceForm.Controls.Add($dtpStart)
    $lblEnd = New-Object System.Windows.Forms.Label
    $lblEnd.Text = "Through:"
    $lblEnd.Location = New-Object System.Drawing.Point(240, 129)
    $lblEnd.AutoSize = $true
    $traceForm.Controls.Add($lblEnd)
    $dtpEnd = New-Object System.Windows.Forms.DateTimePicker
    $dtpEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtpEnd.Location = New-Object System.Drawing.Point(300, 125)
    $dtpEnd.Size = New-Object System.Drawing.Size(120, 20)
    $traceForm.Controls.Add($dtpEnd)
    $dtpStart.Add_ValueChanged({
        $maxEnd = $dtpStart.Value.Date.AddDays(9)
        $today  = (Get-Date).Date
        if ($maxEnd -gt $today) { $maxEnd = $today }
        $dtpEnd.MinDate = $dtpStart.Value.Date
        $dtpEnd.MaxDate = $maxEnd
        if ($dtpEnd.Value -lt $dtpStart.Value.Date) { $dtpEnd.Value = $dtpStart.Value.Date }
        if ($dtpEnd.Value -gt $maxEnd) { $dtpEnd.Value = $maxEnd }
    })
    $dtpEnd.MinDate = $dtpStart.Value.Date
    $initialMaxEnd = $dtpStart.Value.Date.AddDays(9)
    if ($initialMaxEnd -gt (Get-Date).Date) { $initialMaxEnd = (Get-Date).Date }
    $dtpEnd.MaxDate = $initialMaxEnd
    $dtpEnd.Value   = $initialMaxEnd
    $btnTraceRun = New-Object System.Windows.Forms.Button
    $btnTraceRun.Text = "Trace"
    $btnTraceRun.Location = New-Object System.Drawing.Point(440, 123)
    $btnTraceRun.Size = New-Object System.Drawing.Size(90, 26)
    $traceForm.Controls.Add($btnTraceRun)
    $lblTraceStatus = New-Object System.Windows.Forms.Label
    $lblTraceStatus.Text = ""
    $lblTraceStatus.Location = New-Object System.Drawing.Point(12, 156)
    $lblTraceStatus.Size = New-Object System.Drawing.Size(974, 40)
    $traceForm.Controls.Add($lblTraceStatus)
    $grdTrace = New-Object System.Windows.Forms.DataGridView
    $grdTrace.Location = New-Object System.Drawing.Point(12, 200)
    $grdTrace.Size = New-Object System.Drawing.Size(974, 260)
    $grdTrace.ReadOnly = $true
    $grdTrace.AllowUserToAddRows = $false
    $grdTrace.AllowUserToDeleteRows = $false
    $grdTrace.AllowUserToResizeRows = $false
    $grdTrace.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdTrace.MultiSelect = $false
    $grdTrace.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdTrace.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdTrace.Columns.Add('Received', 'Received')
    [void]$grdTrace.Columns.Add('Sender', 'Sender Address')
    [void]$grdTrace.Columns.Add('Recipient', 'Recipient Address')
    [void]$grdTrace.Columns.Add('Subject', 'Subject')
    [void]$grdTrace.Columns.Add('Status', 'Status')
    [void]$grdTrace.Columns.Add('Key', 'Key')
    $grdTrace.Columns['Received'].FillWeight  = 15
    $grdTrace.Columns['Sender'].FillWeight    = 20
    $grdTrace.Columns['Recipient'].FillWeight = 20
    $grdTrace.Columns['Subject'].FillWeight   = 30
    $grdTrace.Columns['Status'].FillWeight    = 15
    $grdTrace.Columns['Received'].ValueType   = [DateTime]
    $grdTrace.Columns['Key'].Visible          = $false
    foreach ($col in $grdTrace.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
    $traceForm.Controls.Add($grdTrace)
    $lblDetail = New-Object System.Windows.Forms.Label
    $lblDetail.Text = "Detail (double-click a row above to load it here):"
    $lblDetail.Location = New-Object System.Drawing.Point(12, 468)
    $lblDetail.AutoSize = $true
    $traceForm.Controls.Add($lblDetail)
    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.Multiline = $true
    $txtDetail.ReadOnly = $true
    $txtDetail.ScrollBars = "Vertical"
    $txtDetail.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtDetail.Location = New-Object System.Drawing.Point(12, 488)
    $txtDetail.Size = New-Object System.Drawing.Size(974, 130)
    $txtDetail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtDetail)
    $btnRelease = New-Object System.Windows.Forms.Button
    $btnRelease.Text = "Release from Quarantine"
    $btnRelease.Location = New-Object System.Drawing.Point(12, 628)
    $btnRelease.Size = New-Object System.Drawing.Size(160, 28)
    $btnRelease.Enabled = $false
    $traceForm.Controls.Add($btnRelease)
    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = "Export to CSV"
    $btnExportCsv.Location = New-Object System.Drawing.Point(180, 628)
    $btnExportCsv.Size = New-Object System.Drawing.Size(110, 28)
    $btnExportCsv.Enabled = $false
    $traceForm.Controls.Add($btnExportCsv)
    $btnTraceClose = New-Object System.Windows.Forms.Button
    $btnTraceClose.Text = "Close"
    $btnTraceClose.Location = New-Object System.Drawing.Point(886, 628)
    $btnTraceClose.Size = New-Object System.Drawing.Size(100, 28)
    $traceForm.Controls.Add($btnTraceClose)
    $script:CurrentTraceResults    = @()
    $script:CurrentTraceDetailRaw  = $null
    $script:TraceSearchStart       = $null
    $script:TraceSearchEnd         = $null
    function Update-TraceGridDisplay {
        $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
        $grdTrace.Columns['Received'].HeaderText = "Received ($($cboTimeZone.SelectedItem))"
        $grdTrace.Rows.Clear()
        foreach ($r in $script:CurrentTraceResults) {
            $rowKey = "$($r.MessageTraceId)|$($r.RecipientAddress)"
            $displayReceived = ConvertTo-TraceDisplayTime -UtcValue $r.Received -TimeZone $tz
            [void]$grdTrace.Rows.Add($displayReceived, $r.SenderAddress, $r.RecipientAddress, $r.Subject, $r.Status, $rowKey)
        }
    }
    $cboTimeZone.Add_SelectedIndexChanged({
        Update-TraceGridDisplay
        if ($script:CurrentTraceDetailRaw) {
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $txtDetail.Text = Format-TraceDetail -Details $script:CurrentTraceDetailRaw -TimeZone $tz
        }
    })
    function Update-NarrowLabel {
        if ($radSent.Checked) { $lblNarrow.Text = "To (optional, narrows recipient):" }
        else                  { $lblNarrow.Text = "From (optional, narrows sender):" }
    }
    Update-NarrowLabel
    $radSent.Add_CheckedChanged({ Update-NarrowLabel })
    $radReceived.Add_CheckedChanged({ Update-NarrowLabel })
    $btnTraceRun.Add_Click({
        $btnTraceRun.Enabled = $false
        $btnRelease.Enabled = $false
        $btnExportCsv.Enabled = $false
        $txtDetail.Text = ""
        $grdTrace.Rows.Clear()
        $script:CurrentTraceResults = @()
        $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
        $lblTraceStatus.Text = "Running trace..."
        try {
            $narrow = $txtNarrow.Text.Trim()
            $subject = $txtSubject.Text.Trim()
            $sd = $dtpStart.Value.Date
            $ed = $dtpEnd.Value.Date.AddDays(1).AddSeconds(-1)
            $traceParams = @{ StartDate = $sd; EndDate = $ed }
            if ($radSent.Checked) {
                $traceParams.SenderAddress = $Smtp
                if ($narrow) { $traceParams.RecipientAddress = $narrow }
            } else {
                $traceParams.RecipientAddress = $Smtp
                if ($narrow) { $traceParams.SenderAddress = $narrow }
            }
            if ($subject) { $traceParams.Subject = $subject }
            if ($cboStatus.SelectedItem -and $cboStatus.SelectedItem -ne 'All Messages') {
                $traceParams.Status = $cboStatus.SelectedItem
            }
            $results = @(Invoke-MessageTrace @traceParams)
            $script:CurrentTraceResults = $results
            $script:CurrentTraceDetailRaw = $null
            $script:TraceSearchStart = $sd
            $script:TraceSearchEnd   = $ed
            Update-TraceGridDisplay
            $btnExportCsv.Enabled = ($results.Count -gt 0)
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $statusNote = if ($traceParams.Status) { "  Status: $($traceParams.Status)." } else { "" }
            $lblTraceStatus.Text = "$($results.Count) result(s). $sd through $ed.$statusNote"
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceV2'
        } finally {
            $btnTraceRun.Enabled = $true
        }
    })
    function Find-TraceResultByKey {
        param([string]$RowKey)
        $script:CurrentTraceResults | Where-Object { "$($_.MessageTraceId)|$($_.RecipientAddress)" -eq $RowKey } | Select-Object -First 1
    }
    $grdTrace.Add_SelectionChanged({
        $btnRelease.Enabled = ($grdTrace.SelectedRows.Count -gt 0)
    })
    $grdTrace.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $rowKey = $grdTrace.Rows[$e.RowIndex].Cells['Key'].Value
        $rec = Find-TraceResultByKey -RowKey $rowKey
        if (-not $rec) {
            $txtDetail.Text = "Could not match that row back to a trace result -- run Trace again."
            return
        }
        $txtDetail.Text = "Loading..."
        try {
            $details = Invoke-MessageTraceDetail -MessageTraceId $rec.MessageTraceId -RecipientAddress $rec.RecipientAddress `
                -StartDate $script:TraceSearchStart -EndDate $script:TraceSearchEnd
            $script:CurrentTraceDetailRaw = $details
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $txtDetail.Text = Format-TraceDetail -Details $details -TimeZone $tz
        } catch {
            $script:CurrentTraceDetailRaw = $null
            $txtDetail.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceDetailV2'
        }
    })
    $btnRelease.Add_Click({
        if ($grdTrace.SelectedRows.Count -eq 0) { return }
        $rowKey = $grdTrace.SelectedRows[0].Cells['Key'].Value
        $rec = Find-TraceResultByKey -RowKey $rowKey
        if (-not $rec) {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = "Could not match that row back to a trace result -- run Trace again."
            return
        }
        $confirmSearch = [System.Windows.Forms.MessageBox]::Show(
            "Look up quarantine entries for this message?`nSender: $($rec.SenderAddress)`nRecipient: $($rec.RecipientAddress)`nSubject: $($rec.Subject)",
            "Confirm Quarantine Lookup", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirmSearch -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnRelease.Enabled = $false
        try {
            $matches = @(Find-QuarantineMatch -MessageId $rec.MessageId -RecipientAddress $rec.RecipientAddress `
                -StartReceivedDate $script:TraceSearchStart -EndReceivedDate $script:TraceSearchEnd)
            if ($matches.Count -eq 0) {
                $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
                $lblTraceStatus.Text = "No quarantine entry found for this message (it may not have been quarantined, may already be released, or may have expired -- quarantine data is only retained for a limited time)."
                return
            }
            foreach ($qmsg in $matches) {
                $confirmRelease = [System.Windows.Forms.MessageBox]::Show(
                    "Quarantine match found:`nType: $($qmsg.Type)   Policy: $($qmsg.PolicyName)`nReceived: $($qmsg.ReceivedTime)   Release status: $($qmsg.ReleaseStatus)`n`nRelease this message to ALL original recipients?",
                    "Confirm Release", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($confirmRelease -ne [System.Windows.Forms.DialogResult]::Yes) { continue }
                Invoke-QuarantineRelease -QuarantineIdentity $qmsg.Identity
                Write-CompletedLog -Task "Released QuarantineMessage $($rec.RecipientAddress) ($(Get-SelectedEnvLabel); sender: $($rec.SenderAddress); subject: $($rec.Subject))"
            }
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $lblTraceStatus.Text = "$($matches.Count) quarantine match(es) processed."
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Quarantine release'
        } finally {
            $btnRelease.Enabled = $true
        }
    })
    $btnExportCsv.Add_Click({
        if ($script:CurrentTraceResults.Count -eq 0) { $lblTraceStatus.Text = "No results to export."; return }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $sfd.FileName = "MessageTrace_$($Smtp -replace '[^a-zA-Z0-9]','_')_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($traceForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $script:CurrentTraceResults |
                Select-Object @{Name = "Received ($($cboTimeZone.SelectedItem))"; Expression = { ConvertTo-TraceDisplayTime -UtcValue $_.Received -TimeZone $tz }},
                              SenderAddress, RecipientAddress, Subject, Status, MessageTraceId |
                Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $lblTraceStatus.Text = "Exported $($script:CurrentTraceResults.Count) row(s) to $($sfd.FileName)."
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = "Export failed: $($_.Exception.Message)"
        }
    })
    $btnTraceClose.Add_Click({ $traceForm.Close() })
    [void]$traceForm.ShowDialog($form)
}
#endregion
#region Logging -----------------------------------------------------------------
# Write-CompletedLog now comes from Common.ps1 (was a private copy here).
#endregion
#region WinForms UI ---------------------------------------------------------
$script:FranchiseDCs = @(
    'fran-vm-dc01.franchises.expresspersonnel.com'
    'fran-vm-dc02.franchises.expresspersonnel.com'
    'fran-vm-dc03.franchises.expresspersonnel.com'
    'fran-vm-dc04.franchises.expresspersonnel.com'
    'fran-vm-dc05.franchises.expresspersonnel.com'
    'fran-vm-dc06.franchises.expresspersonnel.com'
    'fran-vm-dc07.franchises.expresspersonnel.com'
    'fran-vm-dc08.franchises.expresspersonnel.com'
)
$script:HQDomain = 'expresspersonnel.com'
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mailbox / Group Inspector (EAC-style Lookup & Edit)"
$form.Size = New-Object System.Drawing.Size(780, 1110)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Connect to Exchange Online before running. On-prem-synced group lookups use the ActiveDirectory (RSAT) module against the selected environment below. Having an on-prem Exchange session ALSO connected in this window is fine -- every EXO call is resolved to the Exchange Online-sourced cmdlet explicitly, regardless of what else is loaded."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(740, 48)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)
$grpEnv = New-Object System.Windows.Forms.GroupBox
$grpEnv.Text = "Environment (governs on-prem AD group lookups only -- EXO data is single-tenant)"
$grpEnv.Location = New-Object System.Drawing.Point(12, 62)
$grpEnv.Size = New-Object System.Drawing.Size(740, 50)
$form.Controls.Add($grpEnv)
$radFranchise = New-Object System.Windows.Forms.RadioButton
$radFranchise.Text = "Franchise"
$radFranchise.Location = New-Object System.Drawing.Point(10, 20)
$radFranchise.AutoSize = $true
$radFranchise.Checked = $true
$grpEnv.Controls.Add($radFranchise)
$radHQ = New-Object System.Windows.Forms.RadioButton
$radHQ.Text = "HQ"
$radHQ.Location = New-Object System.Drawing.Point(120, 20)
$radHQ.AutoSize = $true
$grpEnv.Controls.Add($radHQ)
$lblDC = New-Object System.Windows.Forms.Label
$lblDC.Text = "Domain Controller (Franchise only; HQ is forest-wide):"
$lblDC.Location = New-Object System.Drawing.Point(12, 116)
$lblDC.AutoSize = $true
$form.Controls.Add($lblDC)
$cboDC = New-Object System.Windows.Forms.ComboBox
$cboDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:FranchiseDCs | ForEach-Object { [void]$cboDC.Items.Add($_) }
$cboDC.SelectedIndex = 0
$cboDC.Location = New-Object System.Drawing.Point(320, 113)
$cboDC.Size = New-Object System.Drawing.Size(400, 20)
$cboDC.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($cboDC)
function Update-EnvironmentUI {
    if ($radHQ.Checked) {
        $cboDC.Enabled = $false
        $lblWarn.Text = "Connect to Exchange Online before running. HQ AD lookups use the ActiveDirectory (RSAT) module, forest-wide (no DC pinned) -- allow for replication before re-running Lookup. An on-prem Exchange session may also be connected; EXO calls are resolved explicitly and won't be affected."
    } else {
        $cboDC.Enabled = $true
        $lblWarn.Text = "Connect to Exchange Online before running. Franchise AD lookups use the ActiveDirectory (RSAT) module, pinned to the selected DC below. An on-prem Exchange session may also be connected; EXO calls are resolved explicitly and won't be affected."
    }
}
$radFranchise.Add_CheckedChanged({ Update-EnvironmentUI })
$radHQ.Add_CheckedChanged({ Update-EnvironmentUI })
$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Enter email address (@expresspros.com -- optional):"
$lblInput.Location = New-Object System.Drawing.Point(12, 146)
$lblInput.AutoSize = $true
$form.Controls.Add($lblInput)
$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(12, 166)
$txtInput.Size = New-Object System.Drawing.Size(740, 20)
$txtInput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtInput)
$lblParsed = New-Object System.Windows.Forms.Label
$lblParsed.Text = "Parsed Preview:"
$lblParsed.Location = New-Object System.Drawing.Point(12, 194)
$lblParsed.AutoSize = $true
$form.Controls.Add($lblParsed)
$txtParsed = New-Object System.Windows.Forms.TextBox
$txtParsed.Multiline = $true
$txtParsed.ReadOnly = $true
$txtParsed.ScrollBars = "Vertical"
$txtParsed.BackColor = [System.Drawing.Color]::WhiteSmoke
$txtParsed.Location = New-Object System.Drawing.Point(12, 214)
$txtParsed.Size = New-Object System.Drawing.Size(740, 40)
$txtParsed.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtParsed)
$btnParse = New-Object System.Windows.Forms.Button
$btnParse.Text = "Parse"
$btnParse.Location = New-Object System.Drawing.Point(12, 262)
$btnParse.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnParse)
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run (Lookup)"
$btnRun.Location = New-Object System.Drawing.Point(112, 262)
$btnRun.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnRun)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(232, 262)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)
$lblMbxActions = New-Object System.Windows.Forms.Label
$lblMbxActions.Text = "Mailbox actions (enabled after a Mailbox lookup):"
$lblMbxActions.Location = New-Object System.Drawing.Point(12, 296)
$lblMbxActions.AutoSize = $true
$form.Controls.Add($lblMbxActions)
$btnRules = New-Object System.Windows.Forms.Button
$btnRules.Text = "Inbox Rules"
$btnRules.Location = New-Object System.Drawing.Point(12, 316)
$btnRules.Size = New-Object System.Drawing.Size(120, 28)
$btnRules.Enabled = $false
$form.Controls.Add($btnRules)
$btnOOF = New-Object System.Windows.Forms.Button
$btnOOF.Text = "Out of Office"
$btnOOF.Location = New-Object System.Drawing.Point(140, 316)
$btnOOF.Size = New-Object System.Drawing.Size(130, 28)
$btnOOF.Enabled = $false
$form.Controls.Add($btnOOF)
$btnCalPerm = New-Object System.Windows.Forms.Button
$btnCalPerm.Text = "Calendar Perm"
$btnCalPerm.Location = New-Object System.Drawing.Point(278, 316)
$btnCalPerm.Size = New-Object System.Drawing.Size(130, 28)
$btnCalPerm.Enabled = $false
$form.Controls.Add($btnCalPerm)
$btnHealth = New-Object System.Windows.Forms.Button
$btnHealth.Text = "Health"
$btnHealth.Location = New-Object System.Drawing.Point(416, 316)
$btnHealth.Size = New-Object System.Drawing.Size(90, 28)
$btnHealth.Enabled = $false
$form.Controls.Add($btnHealth)
$btnTrace = New-Object System.Windows.Forms.Button
$btnTrace.Text = "Message Trace"
$btnTrace.Location = New-Object System.Drawing.Point(516, 316)
$btnTrace.Size = New-Object System.Drawing.Size(130, 28)
$btnTrace.Enabled = $false
$form.Controls.Add($btnTrace)
$btnDLMembers = New-Object System.Windows.Forms.Button
$btnDLMembers.Text = "DL Members"
$btnDLMembers.Location = New-Object System.Drawing.Point(650, 316)
$btnDLMembers.Size = New-Object System.Drawing.Size(102, 28)
$btnDLMembers.Enabled = $false
$form.Controls.Add($btnDLMembers)
$btnMailAttribs = New-Object System.Windows.Forms.Button
$btnMailAttribs.Text = "Mail Attribs"
$btnMailAttribs.Location = New-Object System.Drawing.Point(12, 344)
$btnMailAttribs.Size = New-Object System.Drawing.Size(110, 28)
$btnMailAttribs.Enabled = $false
$form.Controls.Add($btnMailAttribs)
$btnAllAttribs = New-Object System.Windows.Forms.Button
$btnAllAttribs.Text = "All Attribs"
$btnAllAttribs.Location = New-Object System.Drawing.Point(132, 344)
$btnAllAttribs.Size = New-Object System.Drawing.Size(110, 28)
$btnAllAttribs.Enabled = $false
$form.Controls.Add($btnAllAttribs)
$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = "Result:"
$lblOut.Location = New-Object System.Drawing.Point(12, 382)
$lblOut.AutoSize = $true
$form.Controls.Add($lblOut)
$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.Location = New-Object System.Drawing.Point(12, 402)
$txtOut.Size = New-Object System.Drawing.Size(740, 180)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOut)
$lblRules = New-Object System.Windows.Forms.Label
$lblRules.Text = "Inbox Rules (double-click a row to open):"
$lblRules.Location = New-Object System.Drawing.Point(12, 588)
$lblRules.AutoSize = $true
$form.Controls.Add($lblRules)
$grdRules = New-Object System.Windows.Forms.DataGridView
$grdRules.Location = New-Object System.Drawing.Point(12, 608)
$grdRules.Size = New-Object System.Drawing.Size(740, 150)
$grdRules.ReadOnly = $true
$grdRules.AllowUserToAddRows = $false
$grdRules.AllowUserToDeleteRows = $false
$grdRules.AllowUserToResizeRows = $false
$grdRules.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grdRules.MultiSelect = $false
$grdRules.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grdRules.Font = New-Object System.Drawing.Font("Consolas", 9)
[void]$grdRules.Columns.Add('Name', 'Rule Name')
[void]$grdRules.Columns.Add('Description', 'Description')
[void]$grdRules.Columns.Add('Enabled', 'Enabled')
[void]$grdRules.Columns.Add('Identity', 'Identity')
$grdRules.Columns['Name'].FillWeight        = 25
$grdRules.Columns['Description'].FillWeight = 55
$grdRules.Columns['Enabled'].FillWeight     = 12
$grdRules.Columns['Identity'].FillWeight    = 8
$grdRules.Columns['Identity'].Visible = $false
foreach ($col in $grdRules.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }
$form.Controls.Add($grdRules)
$grpMbx = New-Object System.Windows.Forms.GroupBox
$grpMbx.Text = "Mailbox Permissions (enabled after a successful Mailbox lookup)"
$grpMbx.Location = New-Object System.Drawing.Point(12, 766)
$grpMbx.Size = New-Object System.Drawing.Size(740, 90)
$grpMbx.Enabled = $false
$form.Controls.Add($grpMbx)
$lblMbxGrantee = New-Object System.Windows.Forms.Label
$lblMbxGrantee.Text = "Grantee (email or logon):"
$lblMbxGrantee.Location = New-Object System.Drawing.Point(10, 25)
$lblMbxGrantee.AutoSize = $true
$grpMbx.Controls.Add($lblMbxGrantee)
$txtMbxGrantee = New-Object System.Windows.Forms.TextBox
$txtMbxGrantee.Location = New-Object System.Drawing.Point(170, 22)
$txtMbxGrantee.Size = New-Object System.Drawing.Size(280, 20)
$txtMbxGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
$grpMbx.Controls.Add($txtMbxGrantee)
$btnFAGrant = New-Object System.Windows.Forms.Button
$btnFAGrant.Text = "Grant Full Access"
$btnFAGrant.Location = New-Object System.Drawing.Point(10, 55)
$btnFAGrant.Size = New-Object System.Drawing.Size(140, 26)
$grpMbx.Controls.Add($btnFAGrant)
$btnFARevoke = New-Object System.Windows.Forms.Button
$btnFARevoke.Text = "Revoke Full Access"
$btnFARevoke.Location = New-Object System.Drawing.Point(155, 55)
$btnFARevoke.Size = New-Object System.Drawing.Size(140, 26)
$grpMbx.Controls.Add($btnFARevoke)
$btnSAGrant = New-Object System.Windows.Forms.Button
$btnSAGrant.Text = "Grant Send As"
$btnSAGrant.Location = New-Object System.Drawing.Point(305, 55)
$btnSAGrant.Size = New-Object System.Drawing.Size(140, 26)
$grpMbx.Controls.Add($btnSAGrant)
$btnSARevoke = New-Object System.Windows.Forms.Button
$btnSARevoke.Text = "Revoke Send As"
$btnSARevoke.Location = New-Object System.Drawing.Point(450, 55)
$btnSARevoke.Size = New-Object System.Drawing.Size(140, 26)
$grpMbx.Controls.Add($btnSARevoke)
$grpGrp = New-Object System.Windows.Forms.GroupBox
$grpGrp.Text = "Group Delegation (enabled after a successful Group lookup)"
$grpGrp.Location = New-Object System.Drawing.Point(12, 864)
$grpGrp.Size = New-Object System.Drawing.Size(740, 120)
$grpGrp.Enabled = $false
$form.Controls.Add($grpGrp)
$lblGrpGrantee = New-Object System.Windows.Forms.Label
$lblGrpGrantee.Text = "Grantee (email or logon):"
$lblGrpGrantee.Location = New-Object System.Drawing.Point(10, 25)
$lblGrpGrantee.AutoSize = $true
$grpGrp.Controls.Add($lblGrpGrantee)
$txtGrpGrantee = New-Object System.Windows.Forms.TextBox
$txtGrpGrantee.Location = New-Object System.Drawing.Point(170, 22)
$txtGrpGrantee.Size = New-Object System.Drawing.Size(280, 20)
$txtGrpGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
$grpGrp.Controls.Add($txtGrpGrantee)
$lblOwnerNote = New-Object System.Windows.Forms.Label
$lblOwnerNote.Text = "On-prem groups: Set Owner REPLACES the single owner (see script header). M365 Groups: Add/Remove Owner requires the grantee to already be a Member."
$lblOwnerNote.Location = New-Object System.Drawing.Point(10, 48)
$lblOwnerNote.Size = New-Object System.Drawing.Size(720, 30)
$grpGrp.Controls.Add($lblOwnerNote)
$btnOwnerSet = New-Object System.Windows.Forms.Button
$btnOwnerSet.Text = "Set/Add Owner"
$btnOwnerSet.Location = New-Object System.Drawing.Point(10, 82)
$btnOwnerSet.Size = New-Object System.Drawing.Size(140, 26)
$grpGrp.Controls.Add($btnOwnerSet)
$btnOwnerRemove = New-Object System.Windows.Forms.Button
$btnOwnerRemove.Text = "Remove Owner"
$btnOwnerRemove.Location = New-Object System.Drawing.Point(155, 82)
$btnOwnerRemove.Size = New-Object System.Drawing.Size(140, 26)
$grpGrp.Controls.Add($btnOwnerRemove)
$btnMemberAdd = New-Object System.Windows.Forms.Button
$btnMemberAdd.Text = "Add Member"
$btnMemberAdd.Location = New-Object System.Drawing.Point(305, 82)
$btnMemberAdd.Size = New-Object System.Drawing.Size(140, 26)
$grpGrp.Controls.Add($btnMemberAdd)
$btnMemberRemove = New-Object System.Windows.Forms.Button
$btnMemberRemove.Text = "Remove Member"
$btnMemberRemove.Location = New-Object System.Drawing.Point(450, 82)
$btnMemberRemove.Size = New-Object System.Drawing.Size(140, 26)
$grpGrp.Controls.Add($btnMemberRemove)
$script:Parsed      = $null
$script:ParsedFrom  = $null
$script:Target      = $null
$script:LastSmtp    = $null
$script:CurrentRules = @()
$script:CalendarIdentityForForm = $null
function Get-SelectedADServer {
    if ($radHQ.Checked) { return $script:HQDomain }
    return $cboDC.SelectedItem
}
function Get-SelectedEnvLabel {
    if ($radHQ.Checked) { return 'HQ' } else { return 'Franchise' }
}
function Update-ParsedPreview {
    $smtp = Get-TargetSmtp -Text $txtInput.Text
    $script:Parsed     = $smtp
    $script:ParsedFrom = $txtInput.Text
    if ($smtp) {
        $txtParsed.Lines = @("Target SMTP : $smtp", "Environment : $(Get-SelectedEnvLabel)")
    } else {
        $txtParsed.Lines = @("No email address entered.")
    }
}
$btnParse.Add_Click({
    $btnParse.Enabled = $false
    $btnRun.Enabled   = $false
    try { Update-ParsedPreview }
    catch { $txtOut.Text = "ERROR: $($_.Exception.Message)" }
    finally { $btnParse.Enabled = $true; $btnRun.Enabled = $true }
})
$btnRun.Add_Click({
    $btnParse.Enabled = $false
    $btnRun.Enabled   = $false
    $grpMbx.Enabled = $false
    $grpGrp.Enabled = $false
    $btnRules.Enabled = $false
    $btnOOF.Enabled = $false
    $btnCalPerm.Enabled = $false
    $btnHealth.Enabled = $false
    $btnTrace.Enabled = $false
    $btnDLMembers.Enabled = $false
    $btnMailAttribs.Enabled = $false
    $btnAllAttribs.Enabled = $false
    $grdRules.Rows.Clear()
    $script:CurrentRules = @()
    try {
        if (-not $script:Parsed -or $txtInput.Text -ne $script:ParsedFrom) { Update-ParsedPreview }
        if (-not $script:Parsed) {
            $txtOut.Text = "Cannot run: enter an email address first."
            return
        }
        $smtp = $script:Parsed
        $adServer = Get-SelectedADServer
        $script:Target   = Resolve-Target -Smtp $smtp -ADServer $adServer
        $script:LastSmtp = $smtp
        $out = New-Object System.Text.StringBuilder
        [void]$out.AppendLine("Environment queried : $(Get-SelectedEnvLabel)  (AD server/domain: $adServer)")
        [void]$out.AppendLine("Target SMTP         : $smtp")
        [void]$out.AppendLine("Detected type        : $($script:Target.Type)")
        [void]$out.AppendLine("")
        switch ($script:Target.Type) {
            'Mailbox' {
                $state = Get-MailboxState -Smtp $smtp
                [void]$out.Append((Format-MailboxState -Heading "CURRENT STATE" -State $state))
                $grpMbx.Enabled = $true
                $btnRules.Enabled = $true
                $btnOOF.Enabled = $true
                $btnCalPerm.Enabled = $true
                $btnHealth.Enabled = $true
                $btnTrace.Enabled = $true
                $btnDLMembers.Enabled = $true
                $btnMailAttribs.Enabled = $true
                $btnAllAttribs.Enabled = $true
            }
            'UnifiedGroup' {
                $state = Get-GroupState -Type 'UnifiedGroup' -Smtp $smtp
                [void]$out.Append((Format-GroupState -Heading "CURRENT STATE" -State $state))
                $grpGrp.Enabled = $true
            }
            'OnpremGroup' {
                $state = Get-GroupState -Type 'OnpremGroup' -Smtp $smtp -ADGroupDN $script:Target.ADGroupDN -ADServer $adServer
                [void]$out.Append((Format-GroupState -Heading "CURRENT STATE" -State $state))
                $grpGrp.Enabled = $true
            }
            default {
                [void]$out.AppendLine("Not found as a Mailbox, Microsoft 365 Group, or on-prem AD group.")
                if ($script:Target.ResolutionErrors.Count -gt 0) {
                    [void]$out.AppendLine("")
                    [void]$out.AppendLine("--- Resolution attempt errors (for troubleshooting) ---")
                    $script:Target.ResolutionErrors | ForEach-Object { [void]$out.AppendLine($_) }
                }
            }
        }
        $txtOut.Text = $out.ToString()
    }
    catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Run (Lookup)'
    }
    finally {
        $btnParse.Enabled = $true
        $btnRun.Enabled   = $true
    }
})
$btnClose.Add_Click({ $form.Close() })
function Invoke-MailboxEdit {
    param([string]$ActionLabel, [string]$Verb, [scriptblock]$Action)
    $grantee = $txtMbxGrantee.Text.Trim()
    if (-not $script:LastSmtp) { $txtOut.Text = "Run a Lookup first."; return }
    if (-not $grantee) { $txtOut.Text = "Enter a grantee (email or logon) first."; return }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$ActionLabel for '$grantee' on mailbox '$($script:LastSmtp)'?",
        "Confirm $ActionLabel",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
    $grpMbx.Enabled = $false
    $btnRun.Enabled = $false
    try {
        $before = Get-MailboxState -Smtp $script:LastSmtp
        $out = New-Object System.Text.StringBuilder
        [void]$out.Append((Format-MailboxState -Heading "BEFORE" -State $before))
        & $Action
        [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
        $after = Get-MailboxState -Smtp $script:LastSmtp
        [void]$out.Append((Format-MailboxState -Heading "AFTER" -State $after))
        $txtOut.Text = $out.ToString()
        Write-CompletedLog -Task "$Verb $($script:LastSmtp) ($(Get-SelectedEnvLabel); grantee: $grantee)"
    }
    catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context $ActionLabel
    }
    finally {
        $grpMbx.Enabled = $true
        $btnRun.Enabled = $true
    }
}
$btnFAGrant.Add_Click({
    Invoke-MailboxEdit -ActionLabel "Grant Full Access" -Verb "Granted FullAccess" -Action {
        Grant-FullAccessPermission -Smtp $script:LastSmtp -Grantee $txtMbxGrantee.Text.Trim()
    }
})
$btnFARevoke.Add_Click({
    Invoke-MailboxEdit -ActionLabel "Revoke Full Access" -Verb "Revoked FullAccess" -Action {
        Revoke-FullAccessPermission -Smtp $script:LastSmtp -Grantee $txtMbxGrantee.Text.Trim()
    }
})
$btnSAGrant.Add_Click({
    Invoke-MailboxEdit -ActionLabel "Grant Send As" -Verb "Granted SendAs" -Action {
        Grant-SendAsPermission -Smtp $script:LastSmtp -Grantee $txtMbxGrantee.Text.Trim()
    }
})
$btnSARevoke.Add_Click({
    Invoke-MailboxEdit -ActionLabel "Revoke Send As" -Verb "Revoked SendAs" -Action {
        Revoke-SendAsPermission -Smtp $script:LastSmtp -Grantee $txtMbxGrantee.Text.Trim()
    }
})
function Invoke-GroupEdit {
    param([string]$ActionLabel, [string]$Verb, [scriptblock]$Action)
    $grantee = $txtGrpGrantee.Text.Trim()
    if (-not $script:LastSmtp -or -not $script:Target) { $txtOut.Text = "Run a Lookup first."; return }
    if (-not $grantee) { $txtOut.Text = "Enter a grantee (email or logon) first."; return }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$ActionLabel for '$grantee' on group '$($script:LastSmtp)'?",
        "Confirm $ActionLabel",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
    $grpGrp.Enabled = $false
    $btnRun.Enabled = $false
    $adServer = Get-SelectedADServer
    try {
        $before = Get-GroupState -Type $script:Target.Type -Smtp $script:LastSmtp -ADGroupDN $script:Target.ADGroupDN -ADServer $adServer
        $out = New-Object System.Text.StringBuilder
        [void]$out.Append((Format-GroupState -Heading "BEFORE" -State $before))
        & $Action
        [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
        $after = Get-GroupState -Type $script:Target.Type -Smtp $script:LastSmtp -ADGroupDN $script:Target.ADGroupDN -ADServer $adServer
        [void]$out.Append((Format-GroupState -Heading "AFTER" -State $after))
        $txtOut.Text = $out.ToString()
        Write-CompletedLog -Task "$Verb $($script:LastSmtp) ($(Get-SelectedEnvLabel); grantee: $grantee)"
    }
    catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context $ActionLabel
    }
    finally {
        $grpGrp.Enabled = $true
        $btnRun.Enabled = $true
    }
}
$btnOwnerSet.Add_Click({
    Invoke-GroupEdit -ActionLabel "Set/Add Owner" -Verb "Set Owner" -Action {
        $grantee = $txtGrpGrantee.Text.Trim()
        if ($script:Target.Type -eq 'UnifiedGroup') { Add-UnifiedOwner -Smtp $script:LastSmtp -Grantee $grantee }
        else { Set-OnpremOwner -ADGroupDN $script:Target.ADGroupDN -Grantee $grantee -ADServer (Get-SelectedADServer) }
    }
})
$btnOwnerRemove.Add_Click({
    Invoke-GroupEdit -ActionLabel "Remove Owner" -Verb "Removed Owner" -Action {
        $grantee = $txtGrpGrantee.Text.Trim()
        if ($script:Target.Type -eq 'UnifiedGroup') { Remove-UnifiedOwner -Smtp $script:LastSmtp -Grantee $grantee }
        else { throw "On-prem groups have a single ManagedBy owner. Use 'Set/Add Owner' with a new value to replace it -- there is no separate remove action for the AD ManagedBy attribute." }
    }
})
$btnMemberAdd.Add_Click({
    Invoke-GroupEdit -ActionLabel "Add Member" -Verb "Added Member" -Action {
        $grantee = $txtGrpGrantee.Text.Trim()
        if ($script:Target.Type -eq 'UnifiedGroup') { Add-UnifiedMember -Smtp $script:LastSmtp -Grantee $grantee }
        else { Add-OnpremMember -ADGroupDN $script:Target.ADGroupDN -Grantee $grantee -ADServer (Get-SelectedADServer) }
    }
})
$btnMemberRemove.Add_Click({
    Invoke-GroupEdit -ActionLabel "Remove Member" -Verb "Removed Member" -Action {
        $grantee = $txtGrpGrantee.Text.Trim()
        if ($script:Target.Type -eq 'UnifiedGroup') { Remove-UnifiedMember -Smtp $script:LastSmtp -Grantee $grantee }
        else { Remove-OnpremMember -ADGroupDN $script:Target.ADGroupDN -Grantee $grantee -ADServer (Get-SelectedADServer) }
    }
})
$btnRules.Add_Click({
    if (-not $script:LastSmtp -or -not $script:Target -or $script:Target.Type -ne 'Mailbox') {
        $txtOut.Text = "Run a Mailbox lookup first."
        return
    }
    $btnRules.Enabled = $false
    try {
        Update-RulesGrid -Smtp $script:LastSmtp
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Inbox Rules lookup'
    } finally {
        $btnRules.Enabled = $true
    }
})
$grdRules.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $identity = $grdRules.Rows[$e.RowIndex].Cells['Identity'].Value
    $rule = $script:CurrentRules | Where-Object { $_.Identity.ToString() -eq $identity } | Select-Object -First 1
    if ($rule) {
        Show-RuleDetailForm -Rule $rule -Smtp $script:LastSmtp
    } else {
        $txtOut.Text = "Could not match that row back to a rule -- click 'Inbox Rules' again to refresh."
    }
})
function Test-MailboxLookupReady {
    if (-not $script:LastSmtp -or -not $script:Target -or $script:Target.Type -ne 'Mailbox') {
        $txtOut.Text = "Run a Mailbox lookup first."
        return $false
    }
    return $true
}
$btnOOF.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-OutOfOfficeForm -Smtp $script:LastSmtp }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Out of Office' }
})
$btnCalPerm.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-CalendarPermissionsForm -Smtp $script:LastSmtp }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Calendar Permissions' }
})
$btnHealth.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-MailboxHealthForm -Smtp $script:LastSmtp }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Health' }
})
$btnTrace.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-MessageTraceForm -Smtp $script:LastSmtp }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Message Trace' }
})
$btnDLMembers.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-DLMembershipsForm -Smtp $script:LastSmtp -ADServer (Get-SelectedADServer) }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'DL Memberships' }
})
$btnMailAttribs.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-MailAttribsForm -Smtp $script:LastSmtp -ADServer (Get-SelectedADServer) }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs' }
})
$btnAllAttribs.Add_Click({
    if (-not (Test-MailboxLookupReady)) { return }
    try { Show-AllAttribsForm -Smtp $script:LastSmtp -ADServer (Get-SelectedADServer) }
    catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'All Attribs' }
})
# Theme if Common.ps1 happens to be loaded; these tools run standalone
# by design, so this must not become a hard dependency.
if (Get-Command Set-ConsoleTheme -ErrorAction SilentlyContinue) {
    Set-ConsoleTheme -Form $form
}
[void]$form.ShowDialog()
#endregion
