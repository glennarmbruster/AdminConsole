#Requires -Version 5.1
<#
.SYNOPSIS
    ANZ (cloud-only, expresspros.com.au) termination/disable tool -- hides
    and converts one or more mailboxes to shared, and optionally grants a
    manager Full Access (mailbox) and/or Site Collection Admin (OneDrive).
.DESCRIPTION
    Built the same way as SRGDisable.ps1 -- standalone, multi-user WinForms
    tool, NOT routed through ANZCommon.ps1's Resolve-ToolTarget (that
    function is built for a single resolved target and would force a modal
    prompt / block launching with no target at all). Same shape as
    SendToDL.ps1/SendAsDL.ps1 in the original suite: dot-sources
    ANZCommon.ps1 for its shared helpers (Get-TargetSmtp, Format-ErrorDetail,
    Invoke-EXOCommand, Write-CompletedLog) but owns its own multi-line "one
    or more users" textbox instead of a single-target flow. Accepts
    -SmtpAddress (from ANZAdminConsole's last lookup, if any) purely to
    pre-fill the first line of that textbox -- launching with nothing
    pre-filled is fully supported, unlike every Resolve-ToolTarget-based
    tool in this suite.

    REVISED 2026-07-21 -- license removal and sign-in block, three states
    in one day (matching SRGDisable.ps1). (1) Originally included removing
    the terminated user from an Entra ID license group and blocking
    sign-in, via the AzureAD PowerShell module (Get-AzureADUser /
    Get-AzureADGroup / Remove-AzureADGroupMember / Set-AzureADUser). The
    first real run (of SRGDisable.ps1) hit:
        ERROR [Get-AzureADGroup] -- Access blocked to AAD Graph API for
        this application. https://aka.ms/AzureADGraphMigration.
    Microsoft fully retired the underlying Azure AD Graph API on
    2025-08-31 -- the AzureAD module is permanently dead, not a
    permissions/connection problem. (2) Both steps were dropped, then
    briefly restored using the Microsoft.Entra module (Connect-Entra /
    Get-EntraGroup / Get-EntraUser / Remove-EntraGroupMember /
    Set-EntraUser), a working Graph-backed alternative -- but that required
    one-time tenant-admin consent for the Connect-Entra scopes, and per the
    operator ("graph is not gonna work today") that consent isn't available
    today. (3) Both steps are dropped again -- this script no longer calls
    AzureAD, Microsoft.Entra, or Microsoft Graph in any form.

    In place of the automated steps, a completion popup appears after every
    Disable run, listing the exact users just processed and reminding the
    operator to handle both manually:
        Complete:
        Manually remove from license group and block sign-in for these
        users.
        <one user per line>

    FOUR STEPS PER USER, EACH ATTEMPTED INDEPENDENTLY (a failure in one
    step does not skip the others for that same user -- same "partial data/
    partial success is still useful" philosophy as every other tool in this
    suite, e.g. MailboxHealth's per-field try/catch):
      1+2. Hide from address lists AND convert to a shared mailbox, in one
           Set-Mailbox call (Invoke-EXOCommand-wrapped, per the suite's
           collision-guard convention for every EXO cmdlet).
      3. If a manager was entered AND the "Manager Access" selection
         includes Mailbox: Add-MailboxPermission -AccessRights FullAccess
         only. Per the operator's explicit instruction, this script NEVER
         grants Send-As -- there is no Add-RecipientPermission call
         anywhere in this file.
      4. If a manager was entered AND the "Manager Access" selection
         includes OneDrive: Set-SPOUser -IsSiteCollectionAdmin $true on the
         terminated user's OneDrive personal site, so the manager can
         browse/copy files.

    CMDLET VERIFICATION (per the standards doc -- every cmdlet below was
    checked against Microsoft Learn before this script was delivered, same
    verification pass used for SRGDisable.ps1):
      - Set-Mailbox -Type <Regular|Room|Equipment|Shared> and
        -HiddenFromAddressListsEnabled <$true|$false> -- both confirmed
        Set-Mailbox parameters; combining them in one call changes only the
        mailbox type and GAL visibility, nothing else (existing mail,
        folders, and permissions are retained across the Shared conversion
        per Microsoft's own "Convert a mailbox" documentation).
      - Add-MailboxPermission -AccessRights FullAccess -InheritanceType All
        -- reused verbatim from this suite's own ANZMailboxPermissions.ps1
        (Grant-FullAccessPermission), already a proven call in this
        codebase.
      - Set-SPOUser -Site <url> -LoginName <upn> -IsSiteCollectionAdmin
        $true -- confirmed against Microsoft Learn's Set-SPOUser reference.
        Requires SharePoint Online Management Shell (Connect-SPOService),
        NOT PnP.PowerShell -- same module choice as SRGDisable.ps1
        (operator already has SPO Management Shell set up; it has no
        announced retirement).

    UNVERIFIED / FLAGGED ASSUMPTION -- OneDrive URL construction: the
    personal-site URL for each terminated user is ALGORITHMICALLY BUILT
    from their SMTP address (every '.' and '@' replaced with '_', prefixed
    with "https://$script:OneDriveTenantHost/personal/"), not looked up via
    any API. This matches the operator's ANZ example exactly
    (andy.briggs@expresspros.com.au ->
    .../personal/andy_briggs_expresspros_com_au), but is NOT verified
    against a live tenant beyond that. NOTE the SPO tenant prefix
    ("expressprosau") is NOT simply the mail domain with the TLD stripped
    ("expresspros.com.au" would naively give "expresspros", which is
    wrong) -- it's a separate, hardcoded constant below
    ($script:SpoTenantPrefix) precisely because the two don't match here.
    If a user's OneDrive was provisioned under a different alias than their
    current primary SMTP address (e.g. after a mailbox rename), the
    constructed URL could also be stale or wrong -- Set-SPOUser will simply
    fail with a clear "site not found"-style error in that case (caught and
    reported per user, same as every other step), not silently succeed
    against the wrong site. Confirm the constructed URL shown in the Result
    box matches expectations on first real use.

    Manager Access radio buttons ("Mailbox & OneDrive" / "Mailbox Only" /
    "OneDrive Only") are disabled (greyed out) whenever the Manager textbox
    is blank -- there's nothing for them to control with no manager
    entered. "Mailbox & OneDrive" is the default selection. ONE manager
    field applies to the ENTIRE batch of users being termed in a single run
    (operator's explicit choice, 2026-07-21) -- run the tool again for a
    different manager.

    Connections required (two separate PowerShell sessions/modules, not
    one): Exchange Online (Connect-ExchangeOnline) and SharePoint Online
    Management Shell (Connect-SPOService -Url $script:SpoAdminUrl, shown in
    the warning label). This tool does not attempt to connect either one
    itself -- same "connect before running" convention as every other tool
    in this suite.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21 (built the same way as SRGDisable.ps1)
    Revised : 2026-07-21 (same day) -- dropped license-group-removal and
              sign-in-block steps (AzureAD module permanently dead).
    Revised : 2026-07-21 (same day, later) -- restored license-group-removal
              and sign-in-block using the Microsoft.Entra module instead of
              AzureAD.
    Revised : 2026-07-21 (same day, later still) -- dropped both steps again
              (Entra consent not available today); added a completion popup
              reminding the operator to do both manually. See REVISED note
              above.
    Requires: ExchangeOnlineManagement and Microsoft.Online.SharePoint.PowerShell
              (SharePoint Online Management Shell) modules, connected
              before running.
#>
param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'ANZCommon.ps1')

#region ANZ-specific constants ---------------------------------------------------
# LOCAL to this file (not ANZCommon.ps1) -- only this tool needs them, same
# "don't centralize a constant only one script uses" convention as
# AdminConsole.ps1's own $script:DefaultDC.
$script:SpoTenantPrefix    = 'expressprosau'
$script:OneDriveTenantHost = "$($script:SpoTenantPrefix)-my.sharepoint.com"
$script:SpoAdminUrl        = "https://$($script:SpoTenantPrefix)-admin.sharepoint.com"
# $script:LicenseGroupName ('E5 License ANZ - Full') removed 2026-07-21 --
# license removal and sign-in block are handled manually now, see header
# REVISED note. A completion popup reminds the operator each run.
#endregion

#region Parsing ----------------------------------------------------------------
# Same tolerant newline/comma/semicolon splitting and case-insensitive dedup
# as SendToDL.ps1/SendAsDL.ps1's ConvertTo-NormalizedUserList.
function ConvertTo-NormalizedUserList {
    param([string]$Text)
    if (-not $Text) { return @() }
    $raw = @($Text -split '[\r\n,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $result = @()
    foreach ($u in $raw) {
        if ($seen.Add($u.ToLowerInvariant())) { $result += $u }
    }
    return $result
}

# See header "UNVERIFIED / FLAGGED ASSUMPTION" note -- algorithmic, not
# looked up via any API.
function Get-OneDrivePersonalSiteUrl {
    param([string]$Smtp)
    $slug = ($Smtp -replace '\.', '_') -replace '@', '_'
    return "https://$script:OneDriveTenantHost/personal/$slug"
}
#endregion

#region Actions ------------------------------------------------------------------
function Disable-AnzUser {
    # Runs all four steps for one user, each independently try/caught.
    # Returns the per-user report text (for the Result box) and logs a
    # compact status summary via Write-CompletedLog. License-group-removal
    # and sign-in-block are dropped (again) -- see header REVISED note. The
    # Disable button's click handler shows a completion popup reminding the
    # operator to handle both manually for the whole batch.
    param(
        [string]$Smtp,
        [string]$ManagerSmtp,
        [bool]$GrantMailboxAccess,
        [bool]$GrantOneDriveAccess
    )
    $out = New-Object System.Text.StringBuilder
    [void]$out.AppendLine("=== $Smtp ===")

    # Steps 1+2: hide + convert to shared, in one call.
    $statusHideShared = 'FAILED'
    try {
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $Smtp; HiddenFromAddressListsEnabled = $true; Type = 'Shared'; ErrorAction = 'Stop'
        } | Out-Null
        $statusHideShared = 'OK'
        [void]$out.AppendLine("  Hide + convert to shared mailbox: OK")
    } catch {
        [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context 'Hide + convert to shared'))
    }

    # Step 3: manager Full Access (no Send-As -- see header).
    $statusFullAccess = 'SKIPPED'
    if ($ManagerSmtp -and $GrantMailboxAccess) {
        try {
            Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
                Identity = $Smtp; User = $ManagerSmtp; AccessRights = 'FullAccess'; InheritanceType = 'All'; ErrorAction = 'Stop'
            } | Out-Null
            $statusFullAccess = 'OK'
            [void]$out.AppendLine("  Manager Full Access ($ManagerSmtp): OK")
        } catch {
            $statusFullAccess = 'FAILED'
            [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context 'Manager Full Access'))
        }
    } else {
        $reason = if (-not $ManagerSmtp) { 'no manager entered' } else { 'Manager Access set to OneDrive Only' }
        [void]$out.AppendLine("  Manager Full Access: SKIPPED ($reason)")
    }

    # Step 4: manager OneDrive Site Collection Admin.
    $statusOneDrive = 'SKIPPED'
    if ($ManagerSmtp -and $GrantOneDriveAccess) {
        $odUrl = Get-OneDrivePersonalSiteUrl -Smtp $Smtp
        try {
            Set-SPOUser -Site $odUrl -LoginName $ManagerSmtp -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
            $statusOneDrive = 'OK'
            [void]$out.AppendLine("  Manager OneDrive access ($ManagerSmtp -> $odUrl): OK")
        } catch {
            $statusOneDrive = 'FAILED'
            [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Manager OneDrive access ($odUrl)"))
        }
    } else {
        $reason = if (-not $ManagerSmtp) { 'no manager entered' } else { 'Manager Access set to Mailbox Only' }
        [void]$out.AppendLine("  Manager OneDrive access: SKIPPED ($reason)")
    }

    # License-group-removal and sign-in-block (steps 5 and 6) are dropped
    # again -- see header REVISED note. The Disable button's click handler
    # shows a completion popup reminding the operator to handle both
    # manually for the whole batch, so no per-user reminder is duplicated
    # here.
    [void]$out.AppendLine("  License removal: NOT ATTEMPTED -- handle manually (see completion popup).")
    [void]$out.AppendLine("  Sign-in block: NOT ATTEMPTED -- handle manually (see completion popup).")

    $managerLabel = if ($ManagerSmtp) { $ManagerSmtp } else { '(none)' }
    Write-CompletedLog -Task "Disabled $Smtp (manager: $managerLabel; HideShared=$statusHideShared; FullAccess=$statusFullAccess; OneDrive=$statusOneDrive; LicenseRemoved=NOT ATTEMPTED; SignInBlocked=NOT ATTEMPTED)"

    return $out.ToString()
}
#endregion

#region WinForms UI --------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "ANZ Disable"
$form.Size = New-Object System.Drawing.Size(900, 574)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Connect to Exchange Online and SharePoint Online (Connect-SPOService -Url $script:SpoAdminUrl) before running. License removal and sign-in block are NOT performed by this tool -- handle those manually (a popup will remind you which users when each run finishes)."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(860, 54)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblUsers = New-Object System.Windows.Forms.Label
$lblUsers.Text = "User(s) to disable -- one per line (commas/semicolons also tolerated):"
$lblUsers.Location = New-Object System.Drawing.Point(12, 68)
$lblUsers.AutoSize = $true
$form.Controls.Add($lblUsers)
$txtUsers = New-Object System.Windows.Forms.TextBox
$txtUsers.Multiline = $true
$txtUsers.ScrollBars = "Vertical"
$txtUsers.Location = New-Object System.Drawing.Point(12, 88)
$txtUsers.Size = New-Object System.Drawing.Size(860, 110)
$txtUsers.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtUsers)

$lblManager = New-Object System.Windows.Forms.Label
$lblManager.Text = "Manager (optional):"
$lblManager.Location = New-Object System.Drawing.Point(12, 206)
$lblManager.AutoSize = $true
$form.Controls.Add($lblManager)
$txtManager = New-Object System.Windows.Forms.TextBox
$txtManager.Location = New-Object System.Drawing.Point(12, 226)
$txtManager.Size = New-Object System.Drawing.Size(260, 22)
$txtManager.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtManager)

$grpManagerAccess = New-Object System.Windows.Forms.GroupBox
$grpManagerAccess.Text = "Manager Access (only applies if a manager is entered above)"
$grpManagerAccess.Location = New-Object System.Drawing.Point(286, 214)
$grpManagerAccess.Size = New-Object System.Drawing.Size(586, 50)
$form.Controls.Add($grpManagerAccess)
$radMailboxOneDrive = New-Object System.Windows.Forms.RadioButton
$radMailboxOneDrive.Text = "Mailbox & OneDrive"
$radMailboxOneDrive.Location = New-Object System.Drawing.Point(10, 20)
$radMailboxOneDrive.AutoSize = $true
$radMailboxOneDrive.Checked = $true
$grpManagerAccess.Controls.Add($radMailboxOneDrive)
$radMailboxOnly = New-Object System.Windows.Forms.RadioButton
$radMailboxOnly.Text = "Mailbox Only"
$radMailboxOnly.Location = New-Object System.Drawing.Point(190, 20)
$radMailboxOnly.AutoSize = $true
$grpManagerAccess.Controls.Add($radMailboxOnly)
$radOneDriveOnly = New-Object System.Windows.Forms.RadioButton
$radOneDriveOnly.Text = "OneDrive Only"
$radOneDriveOnly.Location = New-Object System.Drawing.Point(340, 20)
$radOneDriveOnly.AutoSize = $true
$grpManagerAccess.Controls.Add($radOneDriveOnly)

# Radio group is only meaningful once a manager is entered -- greyed out
# otherwise (operator's explicit request). Re-evaluated on every keystroke.
function Update-ManagerAccessEnabled {
    $hasManager = -not [string]::IsNullOrWhiteSpace($txtManager.Text)
    $grpManagerAccess.Enabled = $hasManager
}
$txtManager.Add_TextChanged({ Update-ManagerAccessEnabled })

$btnDisable = New-Object System.Windows.Forms.Button
$btnDisable.Text = "Disable"
$btnDisable.Location = New-Object System.Drawing.Point(12, 286)
$btnDisable.Size = New-Object System.Drawing.Size(110, 30)
$btnDisable.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnDisable)
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(130, 286)
$btnClear.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnClear)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(228, 286)
$btnClose.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnClose)
$form.CancelButton = $btnClose

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Result:"
$lblResult.Location = New-Object System.Drawing.Point(12, 326)
$lblResult.Size = New-Object System.Drawing.Size(860, 20)
$lblResult.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblResult)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.Location = New-Object System.Drawing.Point(12, 350)
$txtOut.Size = New-Object System.Drawing.Size(860, 140)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtOut.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($txtOut)

if ($SmtpAddress) { $txtUsers.Text = $SmtpAddress }
Update-ManagerAccessEnabled

$btnDisable.Add_Click({
    $btnDisable.Enabled = $false
    try {
        $users = ConvertTo-NormalizedUserList -Text $txtUsers.Text
        if ($users.Count -eq 0) { $txtOut.Text = "Enter at least one user to disable first."; return }
        $userSmtps = @($users | ForEach-Object { Get-TargetSmtp -Text $_ })

        $managerSmtp = $null
        $grantMailbox = $false
        $grantOneDrive = $false
        if (-not [string]::IsNullOrWhiteSpace($txtManager.Text)) {
            $managerSmtp = Get-TargetSmtp -Text $txtManager.Text
            $grantMailbox  = $radMailboxOneDrive.Checked -or $radMailboxOnly.Checked
            $grantOneDrive = $radMailboxOneDrive.Checked -or $radOneDriveOnly.Checked
        }

        $managerNote = if ($managerSmtp) {
            $accessBits = @()
            if ($grantMailbox)  { $accessBits += 'mailbox Full Access (no Send-As)' }
            if ($grantOneDrive) { $accessBits += 'OneDrive Site Collection Admin' }
            "Manager '$managerSmtp' will be granted: $($accessBits -join ' and ')."
        } else {
            "No manager entered -- no Full Access or OneDrive access will be granted to anyone."
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Disable $($userSmtps.Count) user(s)?`n`n$($userSmtps -join "`n")`n`nFor each: hide from address lists, convert to shared mailbox.`n`nNOTE: license removal and sign-in block are NOT performed by this tool -- handle those manually.`n`n$managerNote",
            "Confirm Disable ($($userSmtps.Count) user(s))",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }

        $out = New-Object System.Text.StringBuilder
        $n = 0
        foreach ($smtp in $userSmtps) {
            $n++
            $lblResult.Text = "Disabling $n of $($userSmtps.Count): $smtp ..."
            $form.Refresh()
            [void]$out.Append((Disable-AnzUser -Smtp $smtp -ManagerSmtp $managerSmtp -GrantMailboxAccess $grantMailbox -GrantOneDriveAccess $grantOneDrive))
            [void]$out.AppendLine("")
        }
        $lblResult.Text = "Done -- processed $($userSmtps.Count) user(s)."
        $txtOut.Text = $out.ToString()

        # Automated license removal / sign-in block are out (see header
        # REVISED note) -- remind the operator, for this exact batch, right
        # when the run finishes.
        [System.Windows.Forms.MessageBox]::Show(
            "Complete:`n`nManually remove from license group and block sign-in for these users.`n`n$($userSmtps -join "`n")",
            "Manual Steps Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'ANZ Disable'
    } finally {
        $btnDisable.Enabled = $true
    }
})

$btnClear.Add_Click({
    $txtUsers.Text = ''
    $txtManager.Text = ''
    $radMailboxOneDrive.Checked = $true
    Update-ManagerAccessEnabled
    $lblResult.Text = "Result:"
    $txtOut.Text = ''
})

$btnClose.Add_Click({ $form.Close() })

[void]$form.ShowDialog()
#endregion
