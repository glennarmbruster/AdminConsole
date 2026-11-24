#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-GRAPH-SCOPES: LicenseAssignment.ReadWrite.All, Organization.Read.All, User.ReadWrite.All, Group.ReadWrite.All
<#
    LicenseAssignment.ps1
    What a user is licensed for, where each licence comes from, what the
    tenant has left, and direct add / remove.

    Author   : Glenn Armbruster
    Created  : 2026-09-17
#>
# TOOLKIT-HELP-RUNS: Reads licences for one user and the tenant's SKU counts. Add and Remove change things and each asks first.
# TOOLKIT-HELP-BEGIN
# One user's licences with the fact that matters most and is hardest to see in
# the portal: whether each one was assigned DIRECTLY or INHERITED FROM A GROUP,
# and which group. A licence in an error state (no seats left, conflicting
# plans, no usage location) shows the error.
#
# Below that, every SKU the tenant owns with enabled / consumed / available
# counts, so "do we have a spare E3" is answered on the same screen.
#
# ADD assigns one SKU directly to the user. It needs a usage location on the
# account; if there is none the tool offers to set one first. REMOVE takes a
# DIRECT licence off. A group-inherited licence cannot be removed here -- the
# tool says which group to take the user out of instead, because removing it
# directly would just come back on the next group sync.
#
# Pick a licence to see its service plans and which are provisioned.
#
# The tenant's licence groups (Tenant.json LicenseGroups) are the normal way
# to license people here; this tool is for the exceptions and for finding out
# why a licence is not what somebody expects.
#
# NEEDS: Graph scopes LicenseAssignment.ReadWrite.All, Organization.Read.All,
# User.ReadWrite.All (usage location) and Group.ReadWrite.All (to name the
# group a licence came from), and the License Administrator or User
# Administrator role on YOUR account.
#
# Every add and remove is written to the audit log.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:Scopes = @('LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'User.ReadWrite.All', 'Group.ReadWrite.All')

#region Data ------------------------------------------------------------------

# Friendly names for the SKU part numbers people see most. Anything not here
# shows its part number, which is still unambiguous -- this is a courtesy, not
# a lookup that can fail.
$script:SkuFriendly = @{
    'ENTERPRISEPACK'            = 'Office 365 E3'
    'ENTERPRISEPREMIUM'         = 'Office 365 E5'
    'STANDARDPACK'              = 'Office 365 E1'
    'SPE_E3'                    = 'Microsoft 365 E3'
    'SPE_E5'                    = 'Microsoft 365 E5'
    'SPE_F1'                    = 'Microsoft 365 F1'
    'SPE_F3'                    = 'Microsoft 365 F3'
    'DESKLESSPACK'              = 'Office 365 F3'
    'O365_BUSINESS_PREMIUM'     = 'Microsoft 365 Business Standard'
    'O365_BUSINESS_ESSENTIALS'  = 'Microsoft 365 Business Basic'
    'SPB'                       = 'Microsoft 365 Business Premium'
    'EXCHANGESTANDARD'          = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'        = 'Exchange Online (Plan 2)'
    'EXCHANGEARCHIVE_ADDON'     = 'Exchange Online Archiving'
    'EXCHANGEDESKLESS'          = 'Exchange Online Kiosk'
    'EMS'                       = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM'               = 'Entra ID P1'
    'AAD_PREMIUM_P2'            = 'Entra ID P2'
    'ATP_ENTERPRISE'            = 'Defender for Office 365 (Plan 1)'
    'THREAT_INTELLIGENCE'       = 'Defender for Office 365 (Plan 2)'
    'INTUNE_A'                  = 'Intune'
    'POWER_BI_PRO'              = 'Power BI Pro'
    'POWER_BI_STANDARD'         = 'Power BI (free)'
    'FLOW_FREE'                 = 'Power Automate (free)'
    'PROJECTPROFESSIONAL'       = 'Project Plan 3'
    'PROJECTPREMIUM'            = 'Project Plan 5'
    'VISIOCLIENT'               = 'Visio Plan 2'
    'TEAMS_EXPLORATORY'         = 'Teams Exploratory'
    'MCOEV'                     = 'Teams Phone Standard'
    'MCOMEETADV'                = 'Audio Conferencing'
    'PHONESYSTEM_VIRTUALUSER'   = 'Teams Phone Resource Account'
    'Microsoft_Teams_Rooms_Pro' = 'Teams Rooms Pro'
    'MICROSOFT_365_COPILOT'     = 'Microsoft 365 Copilot'
    'WIN10_PRO_ENT_SUB'         = 'Windows 10/11 Enterprise E3'
    'RIGHTSMANAGEMENT'          = 'Azure Information Protection P1'
}

function Get-SkuName {
    param([string]$PartNumber)
    if ($script:SkuFriendly.ContainsKey($PartNumber)) { return "$($script:SkuFriendly[$PartNumber]) ($PartNumber)" }
    return $PartNumber
}

function Get-TenantSku {
    <#
        Learn 2026-09-17 (List subscribedSkus): GET /subscribedSkus, $select
        only (no $filter). Properties skuId, skuPartNumber, consumedUnits,
        prepaidUnits{enabled,suspended,warning,lockedOut}, capabilityStatus,
        servicePlans.
    #>
    $r = Invoke-ToolkitGraph -Uri 'subscribedSkus'
    if ($r -and $r.PSObject.Properties['value']) { return @($r.value) }
    return @()
}

function Get-UserLicenseState {
    <#
        Learn 2026-09-17 (user resource): licenseAssignmentStates[] carries
        skuId, assignedByGroup (null when direct), state (Active,
        ActiveWithError, Disabled, Error), error, disabledPlans. That is the
        list this tool is built around -- licenseDetails alone cannot say
        where a licence came from.
    #>
    param([Parameter(Mandatory)][string]$UserId)
    return (Invoke-ToolkitGraph -Uri "users/$UserId`?`$select=id,userPrincipalName,displayName,usageLocation,accountEnabled,assignedLicenses,licenseAssignmentStates")
}

function Get-UserLicenseDetail {
    # Learn 2026-09-17 (List licenseDetails): GET /users/{id}/licenseDetails -> skuId, skuPartNumber, servicePlans[].
    param([Parameter(Mandatory)][string]$UserId)
    $r = Invoke-ToolkitGraph -Uri "users/$UserId/licenseDetails"
    if ($r -and $r.PSObject.Properties['value']) { return @($r.value) }
    return @()
}

function Get-GroupName {
    param([string]$GroupId)
    if (-not $GroupId) { return '' }
    if (-not $script:GroupNames) { $script:GroupNames = @{} }
    if ($script:GroupNames.ContainsKey($GroupId)) { return $script:GroupNames[$GroupId] }
    $n = $GroupId
    try { $g = Invoke-ToolkitGraph -Uri "groups/$GroupId`?`$select=displayName"; if ($g.displayName) { $n = "$($g.displayName)" } } catch { }
    $script:GroupNames[$GroupId] = $n
    return $n
}

function Set-UserLicense {
    <#
        Learn 2026-09-17 (user: assignLicense): POST /users/{id}/assignLicense
        with addLicenses[{skuId, disabledPlans}] and removeLicenses[] -- the
        remove list is REQUIRED even when empty. Delegated least-privilege is
        LicenseAssignment.ReadWrite.All.
    #>
    param([Parameter(Mandatory)][string]$UserId, [string[]]$Add = @(), [string[]]$Remove = @())
    $body = @{
        addLicenses    = @($Add | ForEach-Object { @{ skuId = $_; disabledPlans = @() } })
        removeLicenses = @($Remove)
    }
    Invoke-ToolkitGraph -Uri "users/$UserId/assignLicense" -Method POST -Body $body | Out-Null
}

function Set-UserUsageLocation {
    # Learn 2026-09-17 (Update user): PATCH /users/{id} { usageLocation }. Two-letter ISO code.
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][string]$Code)
    Invoke-ToolkitGraph -Uri "users/$UserId" -Method PATCH -Body @{ usageLocation = $Code.ToUpperInvariant() } | Out-Null
}
#endregion

#region UI --------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Licence Assignment'
$form.ClientSize    = New-Object System.Drawing.Size(1120, 700)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$y = 12
$lblUser          = New-Object System.Windows.Forms.Label
$lblUser.Text     = 'User:'
$lblUser.Location = New-Object System.Drawing.Point(12, ($y + 3))
$lblUser.Size     = New-Object System.Drawing.Size(40, 18)
$form.Controls.Add($lblUser)

$txtUser          = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(54, $y)
$txtUser.Size     = New-Object System.Drawing.Size(360, 24)
$txtUser.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtUser.Text     = "$(Get-TargetSmtp -Text $SmtpAddress)"
$form.Controls.Add($txtUser)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load'
$btnLoad.Location = New-Object System.Drawing.Point(426, ($y - 2))
$btnLoad.Size     = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnLoad)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(1008, ($y - 2))
$btnClose.Size     = New-Object System.Drawing.Size(100, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)
$y += 34

$lblWho           = New-Object System.Windows.Forms.Label
$lblWho.Tag       = 'NoTheme'
$lblWho.Location  = New-Object System.Drawing.Point(12, $y)
$lblWho.Size      = New-Object System.Drawing.Size(1096, 18)
$lblWho.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWho.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblWho)
$y += 22

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, $y)
$lblStatus.Size      = New-Object System.Drawing.Size(1096, 18)
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$lblStatus.Text      = 'Enter a user and press Load.'
$form.Controls.Add($lblStatus)
$y += 22

$lblTop          = New-Object System.Windows.Forms.Label
$lblTop.Text     = "This user's licences (pick one for its service plans):"
$lblTop.Location = New-Object System.Drawing.Point(12, $y)
$lblTop.Size     = New-Object System.Drawing.Size(600, 18)
$form.Controls.Add($lblTop)
$y += 20

function New-LicGrid {
    param([object[]]$Columns, [int]$Height, [int]$Top, [int]$Width)
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Location            = New-Object System.Drawing.Point(12, $Top)
    $g.Size                = New-Object System.Drawing.Size($Width, $Height)
    $g.ReadOnly            = $true
    $g.AllowUserToAddRows  = $false
    $g.SelectionMode       = 'FullRowSelect'
    $g.MultiSelect         = $false
    $g.RowHeadersVisible   = $false
    $g.AutoSizeColumnsMode = 'None'
    $g.ScrollBars          = 'Both'
    $g.Font                = New-Object System.Drawing.Font('Consolas', 9)
    foreach ($c in $Columns) {
        $idx = $g.Columns.Add($c[0], $c[0])
        $g.Columns[$idx].Width    = $c[1]
        $g.Columns[$idx].SortMode = 'Automatic'
    }
    return $g
}

$gridUser = New-LicGrid -Top $y -Height 180 -Width 700 -Columns @(
    @('Licence', 300), @('Source', 220), @('State', 110), @('Error', 60))
$form.Controls.Add($gridUser)

$txtPlans            = New-Object System.Windows.Forms.TextBox
$txtPlans.Multiline  = $true
$txtPlans.ReadOnly   = $true
$txtPlans.ScrollBars = 'Both'
$txtPlans.WordWrap   = $false
$txtPlans.Location   = New-Object System.Drawing.Point(720, $y)
$txtPlans.Size       = New-Object System.Drawing.Size(388, 180)
$txtPlans.Font       = New-Object System.Drawing.Font('Consolas', 8)
$txtPlans.BackColor  = [System.Drawing.Color]::WhiteSmoke
$txtPlans.ForeColor  = [System.Drawing.Color]::Black
$txtPlans.Tag        = 'NoTheme'
$txtPlans.Text       = 'Service plans of the selected licence.'
$form.Controls.Add($txtPlans)
$y += 188

$btnRemove          = New-Object System.Windows.Forms.Button
$btnRemove.Text     = 'Remove selected (direct only)'
$btnRemove.Location = New-Object System.Drawing.Point(12, $y)
$btnRemove.Size     = New-Object System.Drawing.Size(210, 28)
$btnRemove.Enabled  = $false
$form.Controls.Add($btnRemove)

$btnLoc          = New-Object System.Windows.Forms.Button
$btnLoc.Text     = 'Set usage location...'
$btnLoc.Location = New-Object System.Drawing.Point(230, $y)
$btnLoc.Size     = New-Object System.Drawing.Size(160, 28)
$btnLoc.Enabled  = $false
$form.Controls.Add($btnLoc)
$y += 40

$lblBottom          = New-Object System.Windows.Forms.Label
$lblBottom.Text     = 'Tenant licences (pick one and press Add to assign it directly):'
$lblBottom.Location = New-Object System.Drawing.Point(12, $y)
$lblBottom.Size     = New-Object System.Drawing.Size(600, 18)
$form.Controls.Add($lblBottom)
$y += 20

$gridSku = New-LicGrid -Top $y -Height (700 - $y - 52) -Width 1096 -Columns @(
    @('Licence', 420), @('Enabled', 80), @('Consumed', 90), @('Available', 90), @('Warning', 80), @('Suspended', 90), @('Status', 110))
$form.Controls.Add($gridSku)
$y = 700 - 40

$btnAdd          = New-Object System.Windows.Forms.Button
$btnAdd.Text     = 'Add selected licence to user'
$btnAdd.Location = New-Object System.Drawing.Point(12, $y)
$btnAdd.Size     = New-Object System.Drawing.Size(210, 28)
$btnAdd.Enabled  = $false
$form.Controls.Add($btnAdd)

$lblNote           = New-Object System.Windows.Forms.Label
$lblNote.Text      = 'Group-inherited licences are managed by group membership, not here. Direct assignments are the exception path.'
$lblNote.Location  = New-Object System.Drawing.Point(232, ($y + 6))
$lblNote.Size      = New-Object System.Drawing.Size(876, 18)
$lblNote.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblNote)

$form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
$script:User    = $null
$script:Skus    = @()
$script:SkuById = @{}
$script:Details = @{}

$loadSkus = {
    $gridSku.Rows.Clear(); $script:SkuById = @{}
    $script:Skus = @(Get-TenantSku | Sort-Object skuPartNumber)
    foreach ($s in $script:Skus) {
        $script:SkuById["$($s.skuId)"] = $s
        $en = [int]$s.prepaidUnits.enabled; $used = [int]$s.consumedUnits
        $i = $gridSku.Rows.Add((Get-SkuName "$($s.skuPartNumber)"), $en, $used, ($en - $used), [int]$s.prepaidUnits.warning, [int]$s.prepaidUnits.suspended, "$($s.capabilityStatus)")
        $gridSku.Rows[$i].Tag = $s
        if (($en - $used) -le 0) { $gridSku.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0) }
    }
}

$loadUser = {
    $gridUser.Rows.Clear(); $txtPlans.Text = ''; $script:Details = @{}
    $btnRemove.Enabled = $false; $btnAdd.Enabled = $false; $btnLoc.Enabled = $false
    $t = "$($txtUser.Text)".Trim()
    if (-not $t) { $lblStatus.Text = 'Enter a user (UPN or email address).'; return }
    $lblStatus.Text = "Resolving $t..."
    $u = Resolve-GraphUser -Text $t
    if (-not $u) { $lblStatus.Text = "No user matches '$t'."; $lblWho.Text = ''; return }
    $lblStatus.Text = 'Reading licences...'
    $st = Get-UserLicenseState -UserId $u.id
    $script:User = $st
    $loc = "$($st.usageLocation)"
    $lblWho.Text = "$($st.displayName)  <$($st.userPrincipalName)>   account $(if ($st.accountEnabled) { 'ENABLED' } else { 'DISABLED' })   usage location: $(if ($loc) { $loc } else { 'NOT SET -- licences cannot be assigned until it is' })"
    $lblWho.ForeColor = $(if ($loc) { [System.Drawing.Color]::FromArgb(0, 110, 60) } else { [System.Drawing.Color]::FromArgb(200, 0, 0) })
    foreach ($d in @(Get-UserLicenseDetail -UserId $u.id)) { $script:Details["$($d.skuId)"] = $d }
    $states = @()
    if ($st.PSObject.Properties['licenseAssignmentStates']) { $states = @($st.licenseAssignmentStates) }
    foreach ($a in $states) {
        $sku  = $script:SkuById["$($a.skuId)"]
        $name = if ($sku) { Get-SkuName "$($sku.skuPartNumber)" } else { "$($a.skuId)" }
        $grp  = "$($a.assignedByGroup)"
        $src  = if ($grp) { "Group: $(Get-GroupName $grp)" } else { 'Direct' }
        $i = $gridUser.Rows.Add($name, $src, "$($a.state)", "$($a.error)")
        $gridUser.Rows[$i].Tag = $a
        if ("$($a.state)" -match 'Error') { $gridUser.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0) }
    }
    $lblStatus.Text = "$($states.Count) licence assignment(s)."
    $btnAdd.Enabled = ($gridSku.Rows.Count -gt 0); $btnLoc.Enabled = $true
    $btnRemove.Enabled = ($states.Count -gt 0)
}

$run = {
    param([scriptblock]$Do, [string]$Context)
    if ($form.Busy) { return }
    $form.Busy = $true
    try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; & $Do }
    catch { $lblStatus.Text = "$Context failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context $Context }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
}

$btnLoad.Add_Click({ & $run { if ($script:Skus.Count -eq 0) { & $loadSkus }; & $loadUser } 'Load licences (Graph)' })
$txtUser.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $btnLoad.PerformClick(); $_.SuppressKeyPress = $true } })

$gridUser.Add_SelectionChanged({
    if ($gridUser.SelectedRows.Count -eq 0) { $txtPlans.Text = ''; return }
    $a = $gridUser.SelectedRows[0].Tag
    if (-not $a) { return }
    $d = $script:Details["$($a.skuId)"]
    $sb = New-Object System.Text.StringBuilder
    $sku = $script:SkuById["$($a.skuId)"]
    [void]$sb.AppendLine("$(if ($sku) { Get-SkuName "$($sku.skuPartNumber)" } else { $a.skuId })")
    [void]$sb.AppendLine("skuId: $($a.skuId)")
    if ($a.disabledPlans -and @($a.disabledPlans).Count) { [void]$sb.AppendLine("disabled plans: $(@($a.disabledPlans).Count)") }
    [void]$sb.AppendLine('')
    if ($d -and $d.servicePlans) {
        foreach ($p in @($d.servicePlans | Sort-Object servicePlanName)) {
            [void]$sb.AppendLine(("  {0,-14} {1}" -f "$($p.provisioningStatus)", "$($p.servicePlanName)"))
        }
    } else {
        [void]$sb.AppendLine('(no service plan detail -- the licence may be in error)')
    }
    $txtPlans.Text = $sb.ToString()
})

$btnRemove.Add_Click({
    & $run {
        if ($gridUser.SelectedRows.Count -eq 0) { $lblStatus.Text = 'Select a licence first.'; return }
        $a = $gridUser.SelectedRows[0].Tag
        $grp = "$($a.assignedByGroup)"
        $sku = $script:SkuById["$($a.skuId)"]
        $name = if ($sku) { Get-SkuName "$($sku.skuPartNumber)" } else { "$($a.skuId)" }
        if ($grp) {
            $null = Show-ToolDialog(
                "$name is inherited from the group '$(Get-GroupName $grp)'.`r`n`r`nRemoving it here would not stick -- group licensing re-applies it on the next evaluation. Remove the user from that group instead (Group Editor).",
                'Group-assigned licence', 'OK', 'Information')
            return
        }
        $ans = Show-ToolDialog("Remove $name from $($script:User.userPrincipalName)?`r`n`r`nAnything that licence provides -- mailbox, apps, Teams -- stops working for them.", 'Confirm', 'OKCancel', 'Warning')
        if ($ans -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            Set-UserLicense -UserId $script:User.id -Remove @("$($a.skuId)")
            Write-CompletedLog -Action 'LicenseRemove' -Target (Get-LogTarget $script:User.userPrincipalName) -Result 'Success' -Detail "$name removed (direct)." -Environment 'Cloud'
            $lblStatus.Text = "Removed $name."
        } catch {
            Write-FailedLog -Action 'LicenseRemove' -Target "$($script:User.userPrincipalName)" -ErrorRecord $_
            throw
        }
        & $loadSkus; & $loadUser
    } 'Remove licence'
})

$btnAdd.Add_Click({
    & $run {
        if (-not $script:User) { $lblStatus.Text = 'Load a user first.'; return }
        if ($gridSku.SelectedRows.Count -eq 0) { $lblStatus.Text = 'Select a tenant licence first.'; return }
        $s = $gridSku.SelectedRows[0].Tag
        $name = Get-SkuName "$($s.skuPartNumber)"
        if (-not "$($script:User.usageLocation)") {
            $null = Show-ToolDialog("$($script:User.userPrincipalName) has no usage location, and Microsoft will not assign a licence without one. Use 'Set usage location...' first.", 'Usage location required', 'OK', 'Warning')
            return
        }
        $avail = [int]$s.prepaidUnits.enabled - [int]$s.consumedUnits
        $warn = if ($avail -le 0) { "`r`n`r`nWARNING: no seats are available ($([int]$s.consumedUnits) of $([int]$s.prepaidUnits.enabled) used). The assignment will land in an error state." } else { "`r`n`r`n$avail seat(s) available." }
        $ans = Show-ToolDialog("Assign $name directly to $($script:User.userPrincipalName)?$warn", 'Confirm', 'OKCancel', 'Question')
        if ($ans -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            Set-UserLicense -UserId $script:User.id -Add @("$($s.skuId)")
            Write-CompletedLog -Action 'LicenseAdd' -Target (Get-LogTarget $script:User.userPrincipalName) -Result 'Success' -Detail "$name assigned (direct)." -Environment 'Cloud'
            $lblStatus.Text = "Assigned $name."
        } catch {
            Write-FailedLog -Action 'LicenseAdd' -Target "$($script:User.userPrincipalName)" -ErrorRecord $_
            throw
        }
        & $loadSkus; & $loadUser
    } 'Add licence'
})

$btnLoc.Add_Click({
    & $run {
        if (-not $script:User) { return }
        $code = Show-InputDialog -Title 'Usage location' -Width 520 -Prompt "Two-letter country code for $($script:User.userPrincipalName) (for example US, GB, AU):" -Default "$($script:User.usageLocation)"
        $code = "$code".Trim().ToUpperInvariant()
        if (-not $code) { return }
        if ($code -notmatch '^[A-Z]{2}$') { $lblStatus.Text = 'Usage location must be exactly two letters.'; return }
        Set-UserUsageLocation -UserId $script:User.id -Code $code
        Write-CompletedLog -Action 'UsageLocationSet' -Target (Get-LogTarget $script:User.userPrincipalName) -Result 'Success' -Detail "Usage location set to $code." -Environment 'Cloud'
        & $loadUser
        $lblStatus.Text = "Usage location set to $code."
    } 'Set usage location'
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Licence Assignment' | Out-Null

$form.Add_Shown({
    $gate = Assert-ToolkitGraphScope -Scopes $script:Scopes
    if (-not $gate.Ok) {
        $lblStatus.Text = $gate.Message
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
        $btnLoad.Enabled = $false
        return
    }
    & $run { & $loadSkus; if ("$($txtUser.Text)".Trim()) { & $loadUser } } 'Load tenant licences (Graph)'
})

[void]$form.ShowDialog()
$form.Dispose()
#endregion
