#Requires -Version 5.1
<#
    DLManager.ps1
    Distribution List Manager -- scan, browse, and edit on-prem Exchange
    distribution groups, with delegation guardrails.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

param(
    [string]$Environment,
    [string]$DomainController,
    [string]$CacheFilePath = (Join-Path $PSScriptRoot 'DLInventoryCache.json')
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Constants --------------------------------------------------------------
$script:GroupRecipientTypes   = @('MailUniversalDistributionGroup', 'MailUniversalSecurityGroup', 'MailNonUniversalGroup')
$script:ContactRecipientTypes = @('MailContact', 'MailUser')
$script:UserRecipientTypes    = @('UserMailbox', 'RemoteUserMailbox', 'SharedMailbox', 'RemoteSharedMailbox')
#endregion

#region Data layer -- on-prem Exchange (bare cmdlet names -- see header) -------
function Get-DLInventory {
    $groups = @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)
    return @($groups | ForEach-Object {
        [PSCustomObject]@{
            Identity                = $_.Identity.ToString()
            DisplayName             = $_.DisplayName
            PrimarySmtpAddress      = $_.PrimarySmtpAddress.ToString()
            Alias                   = $_.Alias
            GroupType               = if ($_.RecipientTypeDetails -like '*Security*') { 'Security' } else { 'Distribution' }
            ModerationEnabled       = [bool]$_.ModerationEnabled
            HiddenFromAddressLists  = [bool]$_.HiddenFromAddressListsEnabled
            ManagedBy               = @($_.ManagedBy | ForEach-Object { "$_" })
            DistinguishedName       = $_.DistinguishedName
            MemberCount             = $null
            Source                  = $null
        }
    })
}

function Get-DLSourceLabel {
    param([string]$DistinguishedName, [string]$ADServer)
    try {
        $adGroup = Get-ADGroup -Identity $DistinguishedName -Server $ADServer -Properties 'mS-DS-ConsistencyGuid' -ErrorAction Stop
        if ($adGroup.'mS-DS-ConsistencyGuid') { return 'Synced (AAD Connect)' }
        return 'On-Prem Only'
    } catch {
        return "Unknown ($($_.Exception.Message))"
    }
}

function Get-DLMemberCount {
    param([string]$Identity)
    return @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop).Count
}

function Get-DLMemberDetail {
    param([string]$Identity)
    $members = @(Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited -ErrorAction Stop)
    return @($members | ForEach-Object {
        $typeLabel = if ($script:GroupRecipientTypes -contains $_.RecipientType) { 'Group' }
                     elseif ($script:ContactRecipientTypes -contains $_.RecipientType) { 'Contact' }
                     else { 'User' }
        [PSCustomObject]@{
            Identity        = $_.Identity.ToString()
            Name            = $_.DisplayName
            Email           = if ($_.PrimarySmtpAddress) { $_.PrimarySmtpAddress.ToString() } else { '' }
            RecipientType   = "$($_.RecipientType)"
            TypeLabel       = $typeLabel
            IsNestedGroup   = ($script:GroupRecipientTypes -contains $_.RecipientType)
            IsExternalContact = ($_.RecipientType -eq 'MailContact')
            DisabledStatus  = if ($typeLabel -eq 'User') { '?' } else { 'n/a' }
            SamAccountName  = $_.SamAccountName
        }
    })
}
#endregion

#region Inventory cache -- scan-to-file, load-from-file on startup -------------
function Save-DLInventoryCache {
    param([string]$Path, $DLs)
    $payload = [PSCustomObject]@{
        ScanTimestamp = (Get-Date).ToString('o')
        DLs           = @($DLs)
    }
    $payload | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Read-DLInventoryCache {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ("$raw".Trim() -eq '') { return $null }
        $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
        if (-not $parsed.DLs) { return $null }
        $dls = @($parsed.DLs | ForEach-Object {
            $_.ManagedBy = @($_.ManagedBy)
            $_
        })
        return [PSCustomObject]@{
            ScanTimestamp = "$($parsed.ScanTimestamp)"
            DLs           = $dls
        }
    } catch {
        return $null
    }
}
#endregion

#region Guardrail checks --------------------------------------------------------
function Test-RemovingLastOwner {
    param($DL, [string]$OwnerBeingRemoved)
    $owners = @($DL.ManagedBy)
    return ($owners.Count -le 1 -and $owners -contains $OwnerBeingRemoved)
}

function Test-RemovingNestedGroup {
    param($Member)
    return [bool]$Member.IsNestedGroup
}

function Test-AddingDisabledAccount {
    param([string]$SamAccountName, [string]$RecipientType, [string]$ADServer)
    if ($script:UserRecipientTypes -notcontains $RecipientType) { return $false }
    if (-not $SamAccountName) { return $false }
    try {
        $adUser = Get-ADUser -Identity $SamAccountName -Server $ADServer -Properties Enabled -ErrorAction Stop
        return ($adUser.Enabled -eq $false)
    } catch {
        return $null
    }
}

function Test-AddingExternalContact {
    param([string]$RecipientType)
    return ($RecipientType -eq 'MailContact')
}

function Get-SyncedDLWarning {
    param($DL)
    if ($DL.Source -eq 'Synced (AAD Connect)') {
        return "This DL is synced to Azure AD via AAD Connect. Membership changes made here are authoritative, but will not appear in Exchange Online / Azure AD until the next sync cycle runs."
    }
    return $null
}

function Get-ModeratedDLWarning {
    param($DL)
    if ($DL.ModerationEnabled) {
        return "Moderation is enabled on this DL. Messages sent to it (including from members you add) are held for moderator approval until released -- this does not block membership changes themselves, it's a reminder about mail flow."
    }
    return $null
}

function Confirm-GuardrailedAction {
    param(
        [string]$ActionLabel,
        [string[]]$Warnings = @(),
        [string[]]$Guardrails = @(),
        [bool]$OperatorMode
    )
    if ($Guardrails.Count -gt 0 -and -not $OperatorMode) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Blocked by guardrail(s):`n`n$($Guardrails -join "`n")`n`nEnable 'Operator Mode' below to proceed with actions like this.",
            'Blocked -- Operator Mode Required', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    $lines = @()
    if ($Warnings.Count -gt 0) { $lines += $Warnings }
    if ($Guardrails.Count -gt 0) { $lines += ($Guardrails | ForEach-Object { "GUARDRAIL (Operator Mode acknowledged): $_" }) }
    $lines += ''
    $lines += "Proceed with $ActionLabel?"
    $icon = if ($Guardrails.Count -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Question }
    $result = [System.Windows.Forms.MessageBox]::Show(
        ($lines -join "`n"), "Confirm $ActionLabel", [System.Windows.Forms.MessageBoxButtons]::YesNo, $icon)
    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}
#endregion

#region Actions -- writes -------------------------------------------------------
function Add-DLMember {
    param([string]$Identity, [string]$MemberIdentity, [string]$ADServer)
    Add-DistributionGroupMember -Identity $Identity -Member $MemberIdentity -Confirm:$false -DomainController $ADServer -ErrorAction Stop
}
function Remove-DLMember {
    param([string]$Identity, [string]$MemberIdentity, [string]$ADServer)
    Remove-DistributionGroupMember -Identity $Identity -Member $MemberIdentity -Confirm:$false -DomainController $ADServer -ErrorAction Stop
}
function Add-DLMemberBulk {
    param([string]$Identity, [string[]]$MemberIdentities, [string]$ADServer)
    $results = @()
    foreach ($m in $MemberIdentities) {
        $m = $m.Trim()
        if ($m -eq '') { continue }
        try {
            Add-DLMember -Identity $Identity -MemberIdentity $m -ADServer $ADServer
            $results += [PSCustomObject]@{ Member = $m; Success = $true; Error = $null }
        } catch {
            $results += [PSCustomObject]@{ Member = $m; Success = $false; Error = (Format-ErrorDetail -ErrorRecord $_ -Context "Add-DistributionGroupMember ($m)") }
        }
    }
    return $results
}
function Remove-DLMemberBulk {
    param([string]$Identity, [string[]]$MemberIdentities, [string]$ADServer)
    $results = @()
    foreach ($m in $MemberIdentities) {
        $m = $m.Trim()
        if ($m -eq '') { continue }
        try {
            Remove-DLMember -Identity $Identity -MemberIdentity $m -ADServer $ADServer
            $results += [PSCustomObject]@{ Member = $m; Success = $true; Error = $null }
        } catch {
            $results += [PSCustomObject]@{ Member = $m; Success = $false; Error = (Format-ErrorDetail -ErrorRecord $_ -Context "Remove-DistributionGroupMember ($m)") }
        }
    }
    return $results
}
function Export-DLMembersToCsv {
    param($Members, [string]$Path)
    $Members | Select-Object Name, Email, TypeLabel, IsNestedGroup, IsExternalContact, DisabledStatus |
        Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}
#endregion

#region DL Editor window ----------------------------------------------------------
function Show-DLEditorForm {
    param($DL, [string]$ADServer)
    $editForm = New-Object System.Windows.Forms.Form
    $editForm.Text = "DL Editor: $($DL.DisplayName)"
    $editForm.Size = New-Object System.Drawing.Size(900, 760)
    $editForm.StartPosition = "CenterScreen"
    $editForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

    $grpMeta = New-Object System.Windows.Forms.GroupBox
    $grpMeta.Text = "DL Metadata"
    $grpMeta.Location = New-Object System.Drawing.Point(12, 12)
    $grpMeta.Size = New-Object System.Drawing.Size(860, 96)
    $editForm.Controls.Add($grpMeta)
    $lblMeta = New-Object System.Windows.Forms.Label
    $lblMeta.Text = "Display Name : $($DL.DisplayName)`r`nAlias        : $($DL.Alias)`r`nPrimary SMTP : $($DL.PrimarySmtpAddress)`r`nSource       : $($DL.Source)`r`nModeration   : $(if ($DL.ModerationEnabled) {'Enabled'} else {'Disabled'})`r`nHidden From GAL: $($DL.HiddenFromAddressLists)`r`nGroup Type   : $($DL.GroupType)   Owners (ManagedBy): $(($DL.ManagedBy -join ', '))"
    $lblMeta.Location = New-Object System.Drawing.Point(10, 16)
    $lblMeta.Size = New-Object System.Drawing.Size(840, 76)
    $lblMeta.Font = New-Object System.Drawing.Font("Consolas", 8)
    $grpMeta.Controls.Add($lblMeta)

    $lblWarnBanner = New-Object System.Windows.Forms.Label
    $warnings = @((Get-SyncedDLWarning -DL $DL), (Get-ModeratedDLWarning -DL $DL)) | Where-Object { $_ }
    $lblWarnBanner.Text = if ($warnings.Count -gt 0) { ($warnings -join "`r`n`r`n") } else { "No standing warnings for this DL." }
    $lblWarnBanner.Location = New-Object System.Drawing.Point(12, 114)
    $lblWarnBanner.Size = New-Object System.Drawing.Size(860, 50)
    $lblWarnBanner.ForeColor = if ($warnings.Count -gt 0) { [System.Drawing.Color]::DarkOrange } else { [System.Drawing.Color]::DimGray }
    $lblWarnBanner.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
    $editForm.Controls.Add($lblWarnBanner)

    $grdMembers = New-Object System.Windows.Forms.DataGridView
    $grdMembers.Location = New-Object System.Drawing.Point(12, 170)
    $grdMembers.Size = New-Object System.Drawing.Size(860, 300)
    $grdMembers.ReadOnly = $true
    $grdMembers.AllowUserToAddRows = $false
    $grdMembers.AllowUserToDeleteRows = $false
    $grdMembers.AllowUserToResizeRows = $false
    $grdMembers.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdMembers.MultiSelect = $true
    $grdMembers.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdMembers.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdMembers.Columns.Add('Name', 'Name')
    [void]$grdMembers.Columns.Add('Email', 'Email')
    [void]$grdMembers.Columns.Add('TypeLabel', 'Type')
    [void]$grdMembers.Columns.Add('Flags', 'Flags')
    [void]$grdMembers.Columns.Add('Identity', 'Identity')
    $grdMembers.Columns['Name'].FillWeight      = 26
    $grdMembers.Columns['Email'].FillWeight     = 32
    $grdMembers.Columns['TypeLabel'].FillWeight = 14
    $grdMembers.Columns['Flags'].FillWeight     = 28
    $grdMembers.Columns['Identity'].Visible     = $false
    foreach ($col in $grdMembers.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
    $editForm.Controls.Add($grdMembers)

    $script:EditorMembers = @()
    function Update-MembersGrid {
        $grdMembers.Rows.Clear()
        $script:EditorMembers = Get-DLMemberDetail -Identity $DL.Identity
        foreach ($m in $script:EditorMembers) {
            $flags = @()
            if ($m.IsNestedGroup) { $flags += 'NESTED GROUP' }
            if ($m.IsExternalContact) { $flags += 'EXTERNAL CONTACT' }
            if ($m.DisabledStatus -eq $true) { $flags += 'DISABLED ACCOUNT' }
            elseif ($m.DisabledStatus -eq '?') { $flags += '(disabled status unchecked)' }
            [void]$grdMembers.Rows.Add($m.Name, $m.Email, $m.TypeLabel, ($flags -join ', '), $m.Identity)
        }
    }

    $lblOutStatus = New-Object System.Windows.Forms.Label
    $lblOutStatus.Text = ""
    $lblOutStatus.Location = New-Object System.Drawing.Point(12, 476)
    $lblOutStatus.Size = New-Object System.Drawing.Size(860, 40)
    $editForm.Controls.Add($lblOutStatus)

    $grpSingle = New-Object System.Windows.Forms.GroupBox
    $grpSingle.Text = "Add / Remove One Member"
    $grpSingle.Location = New-Object System.Drawing.Point(12, 518)
    $grpSingle.Size = New-Object System.Drawing.Size(860, 66)
    $editForm.Controls.Add($grpSingle)
    $lblGrantee = New-Object System.Windows.Forms.Label
    $lblGrantee.Text = "Member (email or logon):"
    $lblGrantee.Location = New-Object System.Drawing.Point(10, 26)
    $lblGrantee.AutoSize = $true
    $grpSingle.Controls.Add($lblGrantee)
    $txtGrantee = New-Object System.Windows.Forms.TextBox
    $txtGrantee.Location = New-Object System.Drawing.Point(170, 23)
    $txtGrantee.Size = New-Object System.Drawing.Size(260, 22)
    $txtGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grpSingle.Controls.Add($txtGrantee)
    $btnAddOne = New-Object System.Windows.Forms.Button
    $btnAddOne.Text = "Add Member"
    $btnAddOne.Location = New-Object System.Drawing.Point(440, 20)
    $btnAddOne.Size = New-Object System.Drawing.Size(120, 28)
    $grpSingle.Controls.Add($btnAddOne)
    $btnRemoveSelected = New-Object System.Windows.Forms.Button
    $btnRemoveSelected.Text = "Remove Selected"
    $btnRemoveSelected.Location = New-Object System.Drawing.Point(566, 20)
    $btnRemoveSelected.Size = New-Object System.Drawing.Size(140, 28)
    $grpSingle.Controls.Add($btnRemoveSelected)
    $btnRefreshDisabled = New-Object System.Windows.Forms.Button
    $btnRefreshDisabled.Text = "Check Disabled"
    $btnRefreshDisabled.Location = New-Object System.Drawing.Point(712, 20)
    $btnRefreshDisabled.Size = New-Object System.Drawing.Size(130, 28)
    $grpSingle.Controls.Add($btnRefreshDisabled)

    $grpBulk = New-Object System.Windows.Forms.GroupBox
    $grpBulk.Text = "Bulk Add / Bulk Remove (one address per line)"
    $grpBulk.Location = New-Object System.Drawing.Point(12, 590)
    $grpBulk.Size = New-Object System.Drawing.Size(860, 110)
    $editForm.Controls.Add($grpBulk)
    $txtBulk = New-Object System.Windows.Forms.TextBox
    $txtBulk.Multiline = $true
    $txtBulk.ScrollBars = "Vertical"
    $txtBulk.Location = New-Object System.Drawing.Point(10, 20)
    $txtBulk.Size = New-Object System.Drawing.Size(600, 80)
    $txtBulk.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grpBulk.Controls.Add($txtBulk)
    $btnBulkAdd = New-Object System.Windows.Forms.Button
    $btnBulkAdd.Text = "Bulk Add"
    $btnBulkAdd.Location = New-Object System.Drawing.Point(620, 20)
    $btnBulkAdd.Size = New-Object System.Drawing.Size(220, 28)
    $grpBulk.Controls.Add($btnBulkAdd)
    $btnBulkRemove = New-Object System.Windows.Forms.Button
    $btnBulkRemove.Text = "Bulk Remove"
    $btnBulkRemove.Location = New-Object System.Drawing.Point(620, 54)
    $btnBulkRemove.Size = New-Object System.Drawing.Size(220, 28)
    $grpBulk.Controls.Add($btnBulkRemove)

    $chkOperatorMode = New-Object System.Windows.Forms.CheckBox
    $chkOperatorMode.Text = "Operator Mode (required to proceed past a guardrail warning)"
    $chkOperatorMode.Location = New-Object System.Drawing.Point(12, 706)
    $chkOperatorMode.Size = New-Object System.Drawing.Size(500, 22)
    $chkOperatorMode.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $editForm.Controls.Add($chkOperatorMode)

    $btnExportEditor = New-Object System.Windows.Forms.Button
    $btnExportEditor.Text = "Export Members"
    $btnExportEditor.Location = New-Object System.Drawing.Point(560, 704)
    $btnExportEditor.Size = New-Object System.Drawing.Size(140, 26)
    $editForm.Controls.Add($btnExportEditor)
    $btnEditorClose = New-Object System.Windows.Forms.Button
    $btnEditorClose.Text = "Close"
    $btnEditorClose.Location = New-Object System.Drawing.Point(772, 704)
    $btnEditorClose.Size = New-Object System.Drawing.Size(100, 26)
    $editForm.Controls.Add($btnEditorClose)

    Update-MembersGrid

    $btnAddOne.Add_Click({
        $grantee = $txtGrantee.Text.Trim()
        if (-not $grantee) { $lblOutStatus.Text = "Enter a member (email or logon) first."; return }
        $guardrails = @()
        $disabled = Test-AddingDisabledAccount -SamAccountName $grantee -RecipientType 'UserMailbox' -ADServer $ADServer
        if ($disabled -eq $true) { $guardrails += "The account '$grantee' appears to be DISABLED in AD." }
        $warnings = @((Get-SyncedDLWarning -DL $DL), (Get-ModeratedDLWarning -DL $DL)) | Where-Object { $_ }
        if (-not (Confirm-GuardrailedAction -ActionLabel "Add Member '$grantee'" -Warnings $warnings -Guardrails $guardrails -OperatorMode $chkOperatorMode.Checked)) {
            $lblOutStatus.Text = "Cancelled."
            return
        }
        try {
            Add-DLMember -Identity $DL.Identity -MemberIdentity $grantee -ADServer $ADServer
            Write-CompletedLog -Task "Added DL Member $grantee to $($DL.PrimarySmtpAddress) (On-Prem DL Manager; $(Get-SelectedEnvLabel))"
            $lblOutStatus.Text = "Added '$grantee'."
            Update-MembersGrid
        } catch {
            $lblOutStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add-DistributionGroupMember'
        }
    })

    $btnRemoveSelected.Add_Click({
        if ($grdMembers.SelectedRows.Count -eq 0) { $lblOutStatus.Text = "Select one or more members first."; return }
        $selected = @($grdMembers.SelectedRows | ForEach-Object {
            $identity = $_.Cells['Identity'].Value
            $script:EditorMembers | Where-Object { $_.Identity -eq $identity } | Select-Object -First 1
        })
        $guardrails = @()
        $nested = @($selected | Where-Object { $_.IsNestedGroup })
        if ($nested.Count -gt 0) { $guardrails += "Removing nested group member(s): $(($nested | ForEach-Object { $_.Name }) -join ', ')" }
        $names = ($selected | ForEach-Object { $_.Name }) -join ', '
        $warnings = @((Get-SyncedDLWarning -DL $DL), (Get-ModeratedDLWarning -DL $DL)) | Where-Object { $_ }
        if (-not (Confirm-GuardrailedAction -ActionLabel "Remove $($selected.Count) member(s): $names" -Warnings $warnings -Guardrails $guardrails -OperatorMode $chkOperatorMode.Checked)) {
            $lblOutStatus.Text = "Cancelled."
            return
        }
        $results = Remove-DLMemberBulk -Identity $DL.Identity -MemberIdentities ($selected | ForEach-Object { $_.Email }) -ADServer $ADServer
        foreach ($r in ($results | Where-Object { $_.Success })) {
            Write-CompletedLog -Task "Removed DL Member $($r.Member) from $($DL.PrimarySmtpAddress) (On-Prem DL Manager; $(Get-SelectedEnvLabel))"
        }
        $failed = @($results | Where-Object { -not $_.Success })
        $lblOutStatus.Text = "Removed $($results.Count - $failed.Count) of $($results.Count). $(if ($failed.Count -gt 0) { "$($failed.Count) failed -- see log/errors." })"
        Update-MembersGrid
    })

    $btnRefreshDisabled.Add_Click({
        foreach ($m in $script:EditorMembers) {
            if ($m.TypeLabel -eq 'User') {
                $m.DisabledStatus = Test-AddingDisabledAccount -SamAccountName $m.SamAccountName -RecipientType $m.RecipientType -ADServer $ADServer
            }
        }
        Update-MembersGrid
        $lblOutStatus.Text = "Disabled-account status refreshed."
    })

    $btnBulkAdd.Add_Click({
        $lines = @($txtBulk.Text -split "`r`n|`n" | Where-Object { $_.Trim() -ne '' })
        if ($lines.Count -eq 0) { $lblOutStatus.Text = "Enter at least one address to add."; return }
        $warnings = @((Get-SyncedDLWarning -DL $DL), (Get-ModeratedDLWarning -DL $DL)) | Where-Object { $_ }
        $guardrails = @("Bulk adding $($lines.Count) member(s) -- disabled-account and external-contact checks are NOT run per-address for bulk operations (would require one AD/Exchange lookup per line before any writes happen). Review the list above carefully.")
        if (-not (Confirm-GuardrailedAction -ActionLabel "Bulk Add $($lines.Count) member(s)" -Warnings $warnings -Guardrails $guardrails -OperatorMode $chkOperatorMode.Checked)) {
            $lblOutStatus.Text = "Cancelled."
            return
        }
        $results = Add-DLMemberBulk -Identity $DL.Identity -MemberIdentities $lines -ADServer $ADServer
        foreach ($r in ($results | Where-Object { $_.Success })) {
            Write-CompletedLog -Task "Bulk-added DL Member $($r.Member) to $($DL.PrimarySmtpAddress) (On-Prem DL Manager; $(Get-SelectedEnvLabel))"
        }
        $failed = @($results | Where-Object { -not $_.Success })
        $lblOutStatus.Text = "Bulk add: $($results.Count - $failed.Count) succeeded, $($failed.Count) failed."
        if ($failed.Count -gt 0) {
            $lblOutStatus.Text += " First failure: $($failed[0].Error)"
        }
        Update-MembersGrid
    })

    $btnBulkRemove.Add_Click({
        $lines = @($txtBulk.Text -split "`r`n|`n" | Where-Object { $_.Trim() -ne '' })
        if ($lines.Count -eq 0) { $lblOutStatus.Text = "Enter at least one address to remove."; return }
        $warnings = @((Get-SyncedDLWarning -DL $DL), (Get-ModeratedDLWarning -DL $DL)) | Where-Object { $_ }
        $guardrails = @("Bulk removing $($lines.Count) member(s) -- nested-group checks are NOT run per-address for bulk operations. Review the list above carefully.")
        if (-not (Confirm-GuardrailedAction -ActionLabel "Bulk Remove $($lines.Count) member(s)" -Warnings $warnings -Guardrails $guardrails -OperatorMode $chkOperatorMode.Checked)) {
            $lblOutStatus.Text = "Cancelled."
            return
        }
        $results = Remove-DLMemberBulk -Identity $DL.Identity -MemberIdentities $lines -ADServer $ADServer
        foreach ($r in ($results | Where-Object { $_.Success })) {
            Write-CompletedLog -Task "Bulk-removed DL Member $($r.Member) from $($DL.PrimarySmtpAddress) (On-Prem DL Manager; $(Get-SelectedEnvLabel))"
        }
        $failed = @($results | Where-Object { -not $_.Success })
        $lblOutStatus.Text = "Bulk remove: $($results.Count - $failed.Count) succeeded, $($failed.Count) failed."
        if ($failed.Count -gt 0) {
            $lblOutStatus.Text += " First failure: $($failed[0].Error)"
        }
        Update-MembersGrid
    })

    $btnExportEditor.Add_Click({
        try {
            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = "CSV files (*.csv)|*.csv"
            $sfd.FileName = "$($DL.Alias -replace '[^a-zA-Z0-9]','_')_members_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
            if ($sfd.ShowDialog($editForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
            Export-DLMembersToCsv -Members $script:EditorMembers -Path $sfd.FileName
            $lblOutStatus.Text = "Exported $($script:EditorMembers.Count) member(s) to $($sfd.FileName)."
        } catch {
            $lblOutStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export member list'
        }
    })

    $btnEditorClose.Add_Click({ $editForm.Close() })
    Set-ConsoleTheme -Form $editForm
    Set-FormScreenCorner -Form $editForm -Corner 'BottomLeft'
    [void]$editForm.ShowDialog()
}
#endregion

#region Main window --------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Distribution List Manager (On-Prem Exchange)"
$form.Size = New-Object System.Drawing.Size(1040, 720)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "On-prem Exchange only -- requires an already-open Exchange Management Shell session (or Exchange snap-in loaded). Environment/DC below is used only for AD-side lookups (disabled-account and AAD-sync checks), not for Exchange cmdlets."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(1000, 32)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$grpEnv = New-Object System.Windows.Forms.GroupBox
$grpEnv.Text = "Environment (AD lookups only)"
$grpEnv.Location = New-Object System.Drawing.Point(12, 46)
$grpEnv.Size = New-Object System.Drawing.Size(300, 50)
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
$lblDC.Text = "Domain Controller:"
$lblDC.Location = New-Object System.Drawing.Point(324, 60)
$lblDC.AutoSize = $true
$form.Controls.Add($lblDC)
$cboDC = New-Object System.Windows.Forms.ComboBox
$cboDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:FranchiseDCs | ForEach-Object { [void]$cboDC.Items.Add($_) }
$cboDC.SelectedIndex = 0
$cboDC.Location = New-Object System.Drawing.Point(440, 57)
$cboDC.Size = New-Object System.Drawing.Size(360, 20)
$cboDC.Font = New-Object System.Drawing.Font("Consolas", 8)
$lblDC.Visible = $false
$cboDC.Visible = $false
# DC selector hidden 2026-08-03: the toolkit binds to the default franchise
# controller from Common.ps1 and nobody was choosing another one. The control
# is kept rather than deleted so it still supplies that default to the code
# below -- .Visible is the whole change, and reversing it is one word.
$form.Controls.Add($cboDC)

function Get-SelectedEnvLabel { if ($radHQ.Checked) { return 'HQ' } else { return 'Franchise' } }
function Update-EnvironmentUI { $cboDC.Enabled = -not $radHQ.Checked }
$radFranchise.Add_CheckedChanged({ Update-EnvironmentUI })
$radHQ.Add_CheckedChanged({ Update-EnvironmentUI })
if ($Environment -eq 'HQ') { $radHQ.Checked = $true }
elseif ($DomainController -and $script:FranchiseDCs -contains $DomainController) { $cboDC.SelectedItem = $DomainController }
Update-EnvironmentUI

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan All Distribution Groups"
$btnScan.Location = New-Object System.Drawing.Point(812, 46)
$btnScan.Size = New-Object System.Drawing.Size(200, 28)
$form.Controls.Add($btnScan)
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($btnScan, "Runs a live scan against Exchange and saves the result to a cache file ($CacheFilePath). Next time this tool opens, that cache loads automatically instead of re-scanning.")

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search (Name or Email, partial match):"
$lblSearch.Location = New-Object System.Drawing.Point(12, 106)
$lblSearch.AutoSize = $true
$form.Controls.Add($lblSearch)
$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(300, 103)
$txtSearch.Size = New-Object System.Drawing.Size(400, 20)
$txtSearch.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtSearch)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Not scanned yet. Click 'Scan All Distribution Groups' to begin."
$lblStatus.Location = New-Object System.Drawing.Point(12, 132)
$lblStatus.Size = New-Object System.Drawing.Size(900, 20)
$lblStatus.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblStatus)

$btnLoadCounts = New-Object System.Windows.Forms.Button
$btnLoadCounts.Text = "Load Member Counts"
$btnLoadCounts.Location = New-Object System.Drawing.Point(812, 128)
$btnLoadCounts.Size = New-Object System.Drawing.Size(200, 26)
$btnLoadCounts.Enabled = $false
$form.Controls.Add($btnLoadCounts)
$btnCancelCounts = New-Object System.Windows.Forms.Button
$btnCancelCounts.Text = "Cancel"
$btnCancelCounts.Location = New-Object System.Drawing.Point(812, 128)
$btnCancelCounts.Size = New-Object System.Drawing.Size(200, 26)
$btnCancelCounts.Visible = $false
$form.Controls.Add($btnCancelCounts)

$grdDL = New-Object System.Windows.Forms.DataGridView
$grdDL.Location = New-Object System.Drawing.Point(12, 160)
$grdDL.Size = New-Object System.Drawing.Size(1000, 480)
$grdDL.ReadOnly = $true
$grdDL.AllowUserToAddRows = $false
$grdDL.AllowUserToDeleteRows = $false
$grdDL.AllowUserToResizeRows = $false
$grdDL.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grdDL.MultiSelect = $false
$grdDL.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grdDL.Font = New-Object System.Drawing.Font("Consolas", 9)
[void]$grdDL.Columns.Add('Name', 'Name')
[void]$grdDL.Columns.Add('Email', 'Email')
[void]$grdDL.Columns.Add('Source', 'Source')
[void]$grdDL.Columns.Add('MemberCount', 'Members')
[void]$grdDL.Columns.Add('Moderation', 'Moderation')
[void]$grdDL.Columns.Add('HiddenFromGAL', 'Hidden From GAL')
[void]$grdDL.Columns.Add('Identity', 'Identity')
$grdDL.Columns['Name'].FillWeight          = 22
$grdDL.Columns['Email'].FillWeight         = 26
$grdDL.Columns['Source'].FillWeight        = 16
$grdDL.Columns['MemberCount'].FillWeight   = 10
$grdDL.Columns['Moderation'].FillWeight    = 12
$grdDL.Columns['HiddenFromGAL'].FillWeight = 14
$grdDL.Columns['Identity'].Visible         = $false
foreach ($col in $grdDL.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
$form.Controls.Add($grdDL)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(912, 650)
$btnClose.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnClose)

$ctxDL = New-Object System.Windows.Forms.ContextMenuStrip
$miOpen = $ctxDL.Items.Add("Open DL Editor")
$miRefreshCount = $ctxDL.Items.Add("Refresh Member Count")
$miExport = $ctxDL.Items.Add("Export Member List...")
[void]$ctxDL.Items.Add("-")
$miCopyEmail = $ctxDL.Items.Add("Copy Email Address")
$miCopyName = $ctxDL.Items.Add("Copy Name")
$grdDL.ContextMenuStrip = $ctxDL

$script:AllDLs = @()

function Get-SelectedDL {
    if ($grdDL.SelectedRows.Count -eq 0) { return $null }
    $identity = $grdDL.SelectedRows[0].Cells['Identity'].Value
    return $script:AllDLs | Where-Object { $_.Identity -eq $identity } | Select-Object -First 1
}

function Update-DLGridRow {
    param($DL)
    foreach ($row in $grdDL.Rows) {
        if ($row.Cells['Identity'].Value -eq $DL.Identity) {
            $row.Cells['Name'].Value          = $DL.DisplayName
            $row.Cells['Email'].Value         = $DL.PrimarySmtpAddress
            $row.Cells['Source'].Value        = if ($DL.Source) { $DL.Source } else { '(not checked)' }
            $row.Cells['MemberCount'].Value   = if ($null -ne $DL.MemberCount) { "$($DL.MemberCount)" } else { '...' }
            $row.Cells['Moderation'].Value    = if ($DL.ModerationEnabled) { 'Enabled' } else { 'Disabled' }
            $row.Cells['HiddenFromGAL'].Value = "$($DL.HiddenFromAddressLists)"
            return
        }
    }
}

function Update-DLGridFromCurrent {
    $grdDL.Rows.Clear()
    foreach ($dl in $script:AllDLs) {
        [void]$grdDL.Rows.Add(
            $dl.DisplayName, $dl.PrimarySmtpAddress,
            $(if ($dl.Source) { $dl.Source } else { '(not checked)' }),
            $(if ($null -ne $dl.MemberCount) { "$($dl.MemberCount)" } else { '...' }),
            $(if ($dl.ModerationEnabled) { 'Enabled' } else { 'Disabled' }),
            "$($dl.HiddenFromAddressLists)",
            $dl.Identity
        )
    }
    Update-SearchFilter
}

function Update-SearchFilter {
    $needle = $txtSearch.Text.Trim()
    foreach ($row in $grdDL.Rows) {
        if ($needle -eq '') { $row.Visible = $true; continue }
        $name = "$($row.Cells['Name'].Value)"
        $email = "$($row.Cells['Email'].Value)"
        $row.Visible = ($name -like "*$needle*") -or ($email -like "*$needle*")
    }
}
$txtSearch.Add_TextChanged({ Update-SearchFilter })

$btnScan.Add_Click({
    $btnScan.Enabled = $false
    $btnLoadCounts.Enabled = $false
    $lblStatus.Text = "Scanning distribution groups..."
    $form.Refresh()
    try {
        $script:AllDLs = Get-DLInventory
        $adServer = Get-ADServerString -Environment (Get-SelectedEnvLabel) -DomainController $cboDC.SelectedItem
        foreach ($dl in $script:AllDLs) {
            $dl.Source = Get-DLSourceLabel -DistinguishedName $dl.DistinguishedName -ADServer $adServer
        }
        Update-DLGridFromCurrent
        $cacheNote = ''
        try {
            Save-DLInventoryCache -Path $CacheFilePath -DLs $script:AllDLs
            $cacheNote = " Saved to cache ($CacheFilePath)."
        } catch {
            $cacheNote = " (Could not write cache file: $($_.Exception.Message))"
        }
        $lblStatus.Text = "$($script:AllDLs.Count) distribution group(s) found.$cacheNote Member counts not loaded yet -- click 'Load Member Counts'."
        $btnLoadCounts.Enabled = ($script:AllDLs.Count -gt 0)
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup (DL scan)'
    } finally {
        $btnScan.Enabled = $true
    }
})

$script:CountScanCancelRequested = $false
$btnLoadCounts.Add_Click({
    $btnLoadCounts.Visible = $false
    $btnCancelCounts.Visible = $true
    $script:CountScanCancelRequested = $false
    $i = 0
    $total = $script:AllDLs.Count
    foreach ($dl in $script:AllDLs) {
        $i++
        $lblStatus.Text = "Loading member count $i of $total`: $($dl.DisplayName)..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $dl.MemberCount = Get-DLMemberCount -Identity $dl.Identity
        } catch {
            $dl.MemberCount = -1
        }
        Update-DLGridRow -DL $dl
        if ($script:CountScanCancelRequested) { break }
    }
    $cacheNote = ''
    try {
        Save-DLInventoryCache -Path $CacheFilePath -DLs $script:AllDLs
        $cacheNote = ' Cache updated with member counts.'
    } catch {
        $cacheNote = " (Could not write cache file: $($_.Exception.Message))"
    }
    $lblStatus.Text = if ($script:CountScanCancelRequested) { "Member count scan cancelled after $i of $total.$cacheNote" } else { "Member counts loaded for all $total distribution group(s).$cacheNote" }
    $btnLoadCounts.Visible = $true
    $btnCancelCounts.Visible = $false
})
$btnCancelCounts.Add_Click({ $script:CountScanCancelRequested = $true })

$grdDL.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $dl = Get-SelectedDL
    if (-not $dl) { return }
    $adServer = Get-ADServerString -Environment (Get-SelectedEnvLabel) -DomainController $cboDC.SelectedItem
    Show-DLEditorForm -DL $dl -ADServer $adServer
    Update-DLGridRow -DL $dl
})
$miOpen.Add_Click({
    $dl = Get-SelectedDL
    if (-not $dl) { return }
    $adServer = Get-ADServerString -Environment (Get-SelectedEnvLabel) -DomainController $cboDC.SelectedItem
    Show-DLEditorForm -DL $dl -ADServer $adServer
    Update-DLGridRow -DL $dl
})
$miRefreshCount.Add_Click({
    $dl = Get-SelectedDL
    if (-not $dl) { return }
    try {
        $dl.MemberCount = Get-DLMemberCount -Identity $dl.Identity
        Update-DLGridRow -DL $dl
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Refresh member count'
    }
})
$miExport.Add_Click({
    $dl = Get-SelectedDL
    if (-not $dl) { return }
    try {
        $members = Get-DLMemberDetail -Identity $dl.Identity
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $sfd.FileName = "$($dl.Alias -replace '[^a-zA-Z0-9]','_')_members_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Export-DLMembersToCsv -Members $members -Path $sfd.FileName
        $lblStatus.Text = "Exported $($members.Count) member(s) of '$($dl.DisplayName)' to $($sfd.FileName)."
    } catch {
        $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Export member list'
    }
})
$miCopyEmail.Add_Click({
    $dl = Get-SelectedDL
    if ($dl) { [System.Windows.Forms.Clipboard]::SetText($dl.PrimarySmtpAddress) }
})
$miCopyName.Add_Click({
    $dl = Get-SelectedDL
    if ($dl) { [System.Windows.Forms.Clipboard]::SetText($dl.DisplayName) }
})

$btnClose.Add_Click({ $form.Close() })

$script:CacheData = Read-DLInventoryCache -Path $CacheFilePath
if ($script:CacheData -and $script:CacheData.DLs.Count -gt 0) {
    $script:AllDLs = $script:CacheData.DLs
    Update-DLGridFromCurrent
    $btnLoadCounts.Enabled = $true
    $lblStatus.Text = "Loaded $($script:AllDLs.Count) distribution group(s) from cache, scanned $($script:CacheData.ScanTimestamp). Click 'Scan All Distribution Groups' for a fresh live scan."
} elseif (Test-Path -LiteralPath $CacheFilePath) {
    $lblStatus.Text = "Cache file at '$CacheFilePath' exists but could not be read -- click 'Scan All Distribution Groups' to replace it with a fresh one."
}

[void]$form.ShowDialog()
#endregion
