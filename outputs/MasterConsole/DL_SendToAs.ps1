#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Add/Remove-ADPermission and the DL cmdlets are
#                                 on-premises Exchange; must not share a process
#                                 with Exchange Online, which exports the same
#                                 distribution-group cmdlet names.
<#
    DL_SendToAs.ps1
    Lookup / Add / Remove Send-To (allowed senders) and/or Send-As
    permission on an on-prem Distribution List -- merged from the operator's
    two separate tools, SendToDL.ps1 and SendAsDL.ps1, into a single
    console.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# Controls who may send TO a distribution list, and who may send AS it.
#
# Two separate permissions that get confused constantly:
#
#   * Send To -- the list is restricted, and only approved senders get through.
#     Everyone else gets a bounce.
#   * Send As -- mail appears to come FROM the list rather than from the person.
#
# "Nobody outside can email this list" is Send To. "I need my replies to look
# like they came from the team address" is Send As.
#
# On-premises lists only, and it opens in its own window so it cannot be
# confused with the cloud.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Parsing ----------------------------------------------------------------
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
#endregion

#region Send-To actions -- from SendToDL.ps1, unchanged ------------------------
function Get-SendToState {
    param([string]$DLIdentity, [hashtable]$DCParam)
    $grp = Get-DistributionGroup -Identity $DLIdentity @DCParam -ErrorAction Stop
    $rawSenders = @($grp.AcceptMessagesOnlyFromSendersOrMembers)
    $resolved = @($rawSenders | ForEach-Object {
        try {
            $r = Get-Recipient -Identity "$_" @DCParam -ErrorAction Stop
            [PSCustomObject]@{ Display = "$($r.Name) ($($r.SamAccountName))"; DN = $r.DistinguishedName }
        } catch {
            [PSCustomObject]@{ Display = "$_ (could not re-resolve -- may be stale/deleted)"; DN = $null }
        }
    })
    return [PSCustomObject]@{
        DisplayName                        = $grp.DisplayName
        Identity                           = $grp.Identity.ToString()
        RequireSenderAuthenticationEnabled = [bool]$grp.RequireSenderAuthenticationEnabled
        Senders                             = $resolved
    }
}

function Format-SendToState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading ---")
    [void]$sb.AppendLine("  DL              : $($State.DisplayName)  ($($State.Identity))")
    if ($State.RequireSenderAuthenticationEnabled) {
        [void]$sb.AppendLine("  Sender auth req : ON  (this allowed-senders list IS enforced)")
    } else {
        [void]$sb.AppendLine("  Sender auth req : OFF -- this allowed-senders list currently has NO EFFECT until this is turned on.")
    }
    [void]$sb.AppendLine("  Allowed senders ($($State.Senders.Count)):")
    if ($State.Senders.Count -eq 0) {
        [void]$sb.AppendLine("    (none -- DL is not restricted to specific senders)")
    } else {
        $State.Senders | ForEach-Object { [void]$sb.AppendLine("    $($_.Display)") }
    }
    return $sb.ToString()
}

function Invoke-SendToChange {
    param(
        [string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb,
        [hashtable]$DCParam,
        [bool]$AlsoEnableAuth
    )
    $out = New-Object System.Text.StringBuilder
    $before = Get-SendToState -DLIdentity $DLIdentity -DCParam $DCParam
    [void]$out.Append((Format-SendToState -Heading "BEFORE" -State $before))
    [void]$out.AppendLine("")

    $resolvedUsers = @()
    foreach ($u in $Users) {
        try {
            $r = Get-Recipient -Identity $u @DCParam -ErrorAction Stop
            $resolvedUsers += [PSCustomObject]@{ Display = "$($r.Name) ($($r.SamAccountName))"; DN = $r.DistinguishedName }
        } catch {
            [void]$out.AppendLine("  ERROR resolving '$u' -- skipped, not included in this change: $($_.Exception.Message)")
        }
    }

    $currentDNs = @($before.Senders | Where-Object { $_.DN } | ForEach-Object { $_.DN })
    $currentSet = New-Object System.Collections.Generic.HashSet[string]([string[]]$currentDNs, [System.StringComparer]::OrdinalIgnoreCase)
    $newDNs = New-Object System.Collections.Generic.List[string]
    $currentDNs | ForEach-Object { [void]$newDNs.Add($_) }

    $changed = $false
    $actedOn = @()
    foreach ($ru in $resolvedUsers) {
        $already = $currentSet.Contains($ru.DN)
        if ($Verb -eq 'Add') {
            if ($already) {
                [void]$out.AppendLine("  SKIP (already allowed to send to this DL): $($ru.Display)")
            } else {
                [void]$newDNs.Add($ru.DN)
                [void]$currentSet.Add($ru.DN)
                [void]$out.AppendLine("  ADDED: $($ru.Display)")
                $actedOn += $ru.Display
                $changed = $true
            }
        } else {
            if (-not $already) {
                [void]$out.AppendLine("  SKIP (not currently on the allowed-senders list): $($ru.Display)")
            } else {
                [void]$newDNs.Remove($ru.DN) | Out-Null
                [void]$currentSet.Remove($ru.DN)
                [void]$out.AppendLine("  REMOVED: $($ru.Display)")
                $actedOn += $ru.Display
                $changed = $true
            }
        }
    }

    if (-not $changed) {
        [void]$out.AppendLine("")
        [void]$out.AppendLine("No change -- every requested user was already in the desired state. Nothing written.")
        return [PSCustomObject]@{ Output = $out.ToString(); Changed = $false; ActedOn = @() }
    }

    Set-DistributionGroup -Identity $DLIdentity -AcceptMessagesOnlyFromSendersOrMembers $newDNs.ToArray() @DCParam -Confirm:$false -ErrorAction Stop

    if ($Verb -eq 'Add' -and $AlsoEnableAuth -and -not $before.RequireSenderAuthenticationEnabled) {
        Set-DistributionGroup -Identity $DLIdentity -RequireSenderAuthenticationEnabled $true @DCParam -Confirm:$false -ErrorAction Stop
        [void]$out.AppendLine("")
        [void]$out.AppendLine("Also enabled 'require sender authentication' (was OFF) -- this list is now actually enforced.")
    }

    $after = Get-SendToState -DLIdentity $DLIdentity -DCParam $DCParam
    [void]$out.AppendLine("")
    [void]$out.Append((Format-SendToState -Heading "AFTER" -State $after))

    return [PSCustomObject]@{ Output = $out.ToString(); Changed = $true; ActedOn = $actedOn }
}
#endregion

#region Send-As actions -- from SendAsDL.ps1, unchanged -------------------------
function Get-SamAccountNameFromUserString {
    param([string]$UserString)
    if ($UserString -match '\\([^\\]+)$') { return $Matches[1] }
    return $UserString
}

function Resolve-SamAccountName {
    param([string]$Identity, [string]$ADServer)
    $u = Get-ADUser -Identity $Identity -Server $ADServer -ErrorAction Stop
    return $u.SamAccountName
}

function Get-SendAsState {
    param([string]$DLIdentity, [hashtable]$DCParam)
    $perms = @(Get-ADPermission -Identity $DLIdentity @DCParam -ErrorAction Stop)
    $grants = @($perms | Where-Object {
        ($_.ExtendedRights -like '*Send-As*') -and
        (-not $_.IsInherited) -and
        ("$($_.User)" -notlike '*\SELF')
    })
    return [PSCustomObject]@{
        Identity = $DLIdentity
        Grants   = @($grants | ForEach-Object {
            [PSCustomObject]@{ User = "$($_.User)"; Deny = [bool]$_.Deny; Sam = (Get-SamAccountNameFromUserString -UserString "$($_.User)") }
        })
    }
}

function Format-SendAsState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading ---")
    [void]$sb.AppendLine("  DL              : $($State.Identity)")
    [void]$sb.AppendLine("  Send-As grants ($($State.Grants.Count)), inherited/SELF excluded:")
    if ($State.Grants.Count -eq 0) {
        [void]$sb.AppendLine("    (none)")
    } else {
        $State.Grants | ForEach-Object {
            $tag = if ($_.Deny) { ' [DENY]' } else { '' }
            [void]$sb.AppendLine("    $($_.User)$tag")
        }
    }
    return $sb.ToString()
}

function Add-SendAsGrant {
    param([string]$DLIdentity, [string]$UserIdentity, [hashtable]$DCParam)
    Add-ADPermission -Identity $DLIdentity -User $UserIdentity -ExtendedRights "Send-As" -Confirm:$false @DCParam -ErrorAction Stop
}

function Remove-SendAsGrant {
    param([string]$DLIdentity, [string]$UserIdentity, [hashtable]$DCParam)
    Remove-ADPermission -Identity $DLIdentity -User $UserIdentity -ExtendedRights "Send-As" -Confirm:$false @DCParam -ErrorAction Stop
}

function Invoke-SendAsChange {
    param(
        [string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb,
        [string]$ADServer,
        [hashtable]$DCParam
    )
    $out = New-Object System.Text.StringBuilder
    $before = Get-SendAsState -DLIdentity $DLIdentity -DCParam $DCParam
    [void]$out.Append((Format-SendAsState -Heading "BEFORE" -State $before))
    [void]$out.AppendLine("")

    $beforeSams = New-Object System.Collections.Generic.HashSet[string]([string[]]@($before.Grants | ForEach-Object { $_.Sam }), [System.StringComparer]::OrdinalIgnoreCase)

    $changed = $false
    $actedOn = @()
    foreach ($u in $Users) {
        try {
            $sam = Resolve-SamAccountName -Identity $u -ADServer $ADServer
        } catch {
            [void]$out.AppendLine("  ERROR resolving '$u' via Get-ADUser -- skipped, not included in this change: $($_.Exception.Message)")
            continue
        }
        $already = $beforeSams.Contains($sam)
        if ($Verb -eq 'Add') {
            if ($already) {
                [void]$out.AppendLine("  SKIP (already has Send-As on this DL): $sam")
                continue
            }
            try {
                Add-SendAsGrant -DLIdentity $DLIdentity -UserIdentity $sam -DCParam $DCParam
                [void]$out.AppendLine("  ADDED: $sam")
                $actedOn += $sam
                $changed = $true
            } catch {
                [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Add-ADPermission ($sam -> $DLIdentity)"))
            }
        } else {
            if (-not $already) {
                [void]$out.AppendLine("  SKIP (does not currently have Send-As on this DL): $sam")
                continue
            }
            try {
                Remove-SendAsGrant -DLIdentity $DLIdentity -UserIdentity $sam -DCParam $DCParam
                [void]$out.AppendLine("  REMOVED: $sam")
                $actedOn += $sam
                $changed = $true
            } catch {
                [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Remove-ADPermission ($sam -> $DLIdentity)"))
            }
        }
    }

    if (-not $changed) {
        [void]$out.AppendLine("")
        [void]$out.AppendLine("No change -- every requested user was already in the desired state (or failed to resolve). Nothing written.")
        return [PSCustomObject]@{ Output = $out.ToString(); Changed = $false; ActedOn = @() }
    }

    $after = Get-SendAsState -DLIdentity $DLIdentity -DCParam $DCParam
    [void]$out.AppendLine("")
    [void]$out.Append((Format-SendAsState -Heading "AFTER" -State $after))

    return [PSCustomObject]@{ Output = $out.ToString(); Changed = $true; ActedOn = $actedOn }
}
#endregion

#region Combined orchestration ---------------------------------------------------

function Invoke-DLAccessChange {
    param(
        [string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb,
        [bool]$DoSendTo,
        [bool]$DoSendAs,
        [bool]$AlsoEnableAuth,
        [hashtable]$DCParam,
        [string]$ADServer
    )
    $out = New-Object System.Text.StringBuilder
    $sendToChanged = $false; $sendToActedOn = @()
    $sendAsChanged = $false; $sendAsActedOn = @()

    if ($DoSendTo) {
        [void]$out.AppendLine("=== Send-To (allowed senders) ===")
        $r = Invoke-SendToChange -DLIdentity $DLIdentity -Users $Users -Verb $Verb -DCParam $DCParam -AlsoEnableAuth ($Verb -eq 'Add' -and $AlsoEnableAuth)
        [void]$out.Append($r.Output)
        [void]$out.AppendLine("")
        $sendToChanged = $r.Changed
        $sendToActedOn = $r.ActedOn
    }
    if ($DoSendAs) {
        [void]$out.AppendLine("=== Send-As (permission delegation) ===")
        $r2 = Invoke-SendAsChange -DLIdentity $DLIdentity -Users $Users -Verb $Verb -ADServer $ADServer -DCParam $DCParam
        [void]$out.Append($r2.Output)
        [void]$out.AppendLine("")
        $sendAsChanged = $r2.Changed
        $sendAsActedOn = $r2.ActedOn
    }

    return [PSCustomObject]@{
        Output        = $out.ToString()
        SendToChanged = $sendToChanged
        SendToActedOn = $sendToActedOn
        SendAsChanged = $sendAsChanged
        SendAsActedOn = $sendAsActedOn
    }
}
#endregion

#region WinForms UI -------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "DL Access Manager -- Send-To / Send-As (on-prem)"
$form.Size = New-Object System.Drawing.Size(780, 745)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Connect to on-prem Exchange ONLY (do NOT connect to Exchange Online -- Get-/Set-DistributionGroup and Get-Recipient would resolve against the wrong environment). HQ and Franchise are checked automatically per DL name."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(740, 40)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$script:DefaultDC = 'fran-vm-dc01.franchises.expresspersonnel.com'

function Resolve-DLEnvironment {
    param([string]$DLIdentity)
    $errors = @()
    try {
        $null = Get-DistributionGroup -Identity $DLIdentity -ErrorAction Stop
        return [PSCustomObject]@{ Environment = 'HQ'; DCParam = @{}; ADServer = $script:HQDomain; Errors = $errors }
    } catch {
        $errors += "HQ: $($_.Exception.Message)"
    }
    try {
        $null = Get-DistributionGroup -Identity $DLIdentity -DomainController $script:DefaultDC -ErrorAction Stop
        return [PSCustomObject]@{ Environment = 'Franchise'; DCParam = @{ DomainController = $script:DefaultDC }; ADServer = $script:DefaultDC; Errors = $errors }
    } catch {
        $errors += "Franchise: $($_.Exception.Message)"
    }
    return [PSCustomObject]@{ Environment = $null; DCParam = $null; ADServer = $null; Errors = $errors }
}

$lblDL = New-Object System.Windows.Forms.Label
$lblDL.Text = "Distribution List (name, not SMTP):"
$lblDL.Location = New-Object System.Drawing.Point(12, 58)
$lblDL.AutoSize = $true
$form.Controls.Add($lblDL)
$txtDL = New-Object System.Windows.Forms.TextBox
$txtDL.Location = New-Object System.Drawing.Point(12, 78)
$txtDL.Size = New-Object System.Drawing.Size(740, 22)
$txtDL.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtDL)

$lblUsers = New-Object System.Windows.Forms.Label
$lblUsers.Text = "User(s) -- optional for Lookup, required for Add/Remove (firstname.lastname, one per line):"
$lblUsers.Location = New-Object System.Drawing.Point(12, 108)
$lblUsers.AutoSize = $true
$form.Controls.Add($lblUsers)
$txtUsers = New-Object System.Windows.Forms.TextBox
$txtUsers.Multiline = $true
$txtUsers.ScrollBars = "Vertical"
$txtUsers.Location = New-Object System.Drawing.Point(12, 128)
$txtUsers.Size = New-Object System.Drawing.Size(740, 110)
$txtUsers.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtUsers)

$grpAccessType = New-Object System.Windows.Forms.GroupBox
$grpAccessType.Text = "Access Type -- select one or both"
$grpAccessType.Size = New-Object System.Drawing.Size(740, 45)
$grpAccessType.Location = New-Object System.Drawing.Point(12, 246)
$form.Controls.Add($grpAccessType)

$chkSendTo = New-Object System.Windows.Forms.CheckBox
$chkSendTo.Text = "Send-To (allowed senders)"
$chkSendTo.AutoSize = $true
$chkSendTo.Location = New-Object System.Drawing.Point(15, 18)
$chkSendTo.Checked = $true
$grpAccessType.Controls.Add($chkSendTo)

$chkSendAs = New-Object System.Windows.Forms.CheckBox
$chkSendAs.Text = "Send-As (permission delegation)"
$chkSendAs.AutoSize = $true
$chkSendAs.Location = New-Object System.Drawing.Point(220, 18)
$grpAccessType.Controls.Add($chkSendAs)

$chkAlsoEnable = New-Object System.Windows.Forms.CheckBox
$chkAlsoEnable.Text = "On Add: also enable 'require sender authentication' if currently OFF (Send-To only -- needed for that list to have any effect; see warning)"
$chkAlsoEnable.Location = New-Object System.Drawing.Point(12, 299)
$chkAlsoEnable.Size = New-Object System.Drawing.Size(740, 22)
$form.Controls.Add($chkAlsoEnable)

$chkSendTo.Add_CheckedChanged({
    $chkAlsoEnable.Enabled = $chkSendTo.Checked
    if (-not $chkSendTo.Checked) { $chkAlsoEnable.Checked = $false }
})

$btnLookup = New-Object System.Windows.Forms.Button
$btnLookup.Text = "Lookup"
$btnLookup.Location = New-Object System.Drawing.Point(12, 331)
$btnLookup.Size = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnLookup)
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add Access"
$btnAdd.Location = New-Object System.Drawing.Point(130, 331)
$btnAdd.Size = New-Object System.Drawing.Size(140, 30)
$btnAdd.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnAdd)
$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remove Access"
$btnRemove.Location = New-Object System.Drawing.Point(278, 331)
$btnRemove.Size = New-Object System.Drawing.Size(150, 30)
$btnRemove.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRemove)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(436, 331)
$btnClose.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnClose)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Result:"
$lblResult.Location = New-Object System.Drawing.Point(12, 371)
$lblResult.AutoSize = $true
$form.Controls.Add($lblResult)
$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Both"
$txtOut.WordWrap = $false
$txtOut.Location = New-Object System.Drawing.Point(12, 391)
$txtOut.Size = New-Object System.Drawing.Size(740, 300)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOut)

function Set-ButtonsEnabled {
    param([bool]$Enabled)
    $btnLookup.Enabled = $Enabled
    $btnAdd.Enabled = $Enabled
    $btnRemove.Enabled = $Enabled
}

function Get-AccessTypeSummary {
    $parts = @()
    if ($chkSendTo.Checked) { $parts += 'Send-To' }
    if ($chkSendAs.Checked) { $parts += 'Send-As' }
    return ($parts -join ' and ')
}

$btnLookup.Add_Click({
    Set-ButtonsEnabled -Enabled $false
    try {
        if (-not $chkSendTo.Checked -and -not $chkSendAs.Checked) {
            $txtOut.Text = "Check Send-To and/or Send-As above first."
            return
        }
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }
        $out = New-Object System.Text.StringBuilder
        [void]$out.AppendLine("Environment: $($envResult.Environment)")
        [void]$out.AppendLine("")
        if ($chkSendTo.Checked) {
            $stateTo = Get-SendToState -DLIdentity $dl -DCParam $envResult.DCParam
            [void]$out.AppendLine("=== Send-To (allowed senders) ===")
            [void]$out.Append((Format-SendToState -Heading "CURRENT STATE" -State $stateTo))
            [void]$out.AppendLine("")
        }
        if ($chkSendAs.Checked) {
            $stateAs = Get-SendAsState -DLIdentity $dl -DCParam $envResult.DCParam
            [void]$out.AppendLine("=== Send-As (permission delegation) ===")
            [void]$out.Append((Format-SendAsState -Heading "CURRENT STATE" -State $stateAs))
        }
        $txtOut.Text = $out.ToString()
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Lookup DL Access'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnAdd.Add_Click({
    Set-ButtonsEnabled -Enabled $false
    try {
        if (-not $chkSendTo.Checked -and -not $chkSendAs.Checked) {
            $txtOut.Text = "Check Send-To and/or Send-As above first."
            return
        }
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $users = ConvertTo-NormalizedUserList -Text $txtUsers.Text
        if ($users.Count -eq 0) { $txtOut.Text = "Enter at least one user in the second box first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }

        $authNote = ''
        if ($chkSendTo.Checked -and -not $chkAlsoEnable.Checked) {
            $pre = Get-SendToState -DLIdentity $dl -DCParam $envResult.DCParam
            if (-not $pre.RequireSenderAuthenticationEnabled) {
                $authNote = "`n`nNOTE: 'require sender authentication' is currently OFF for this DL -- the Send-To list will have NO EFFECT until that's turned on (check the box above, or enable it separately)."
            }
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Add $($users.Count) user(s) to $(Get-AccessTypeSummary) on '$dl' ($($envResult.Environment))?$authNote",
            "Confirm Add Access",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }

        $result = Invoke-DLAccessChange -DLIdentity $dl -Users $users -Verb 'Add' `
            -DoSendTo $chkSendTo.Checked -DoSendAs $chkSendAs.Checked -AlsoEnableAuth $chkAlsoEnable.Checked `
            -DCParam $envResult.DCParam -ADServer $envResult.ADServer
        $txtOut.Text = $result.Output

        if ($result.SendToChanged -and $result.SendToActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-To $dl ($($envResult.Environment); users: $($result.SendToActedOn -join '+'))"
        }
        if ($result.SendAsChanged -and $result.SendAsActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-As $dl ($($envResult.Environment); users: $($result.SendAsActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add DL Access'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnRemove.Add_Click({
    Set-ButtonsEnabled -Enabled $false
    try {
        if (-not $chkSendTo.Checked -and -not $chkSendAs.Checked) {
            $txtOut.Text = "Check Send-To and/or Send-As above first."
            return
        }
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $users = ConvertTo-NormalizedUserList -Text $txtUsers.Text
        if ($users.Count -eq 0) { $txtOut.Text = "Enter at least one user in the second box first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($users.Count) user(s) from $(Get-AccessTypeSummary) on '$dl' ($($envResult.Environment))?",
            "Confirm Remove Access",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }

        $result = Invoke-DLAccessChange -DLIdentity $dl -Users $users -Verb 'Remove' `
            -DoSendTo $chkSendTo.Checked -DoSendAs $chkSendAs.Checked -AlsoEnableAuth $false `
            -DCParam $envResult.DCParam -ADServer $envResult.ADServer
        $txtOut.Text = $result.Output

        if ($result.SendToChanged -and $result.SendToActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-To $dl ($($envResult.Environment); users: $($result.SendToActedOn -join '+'))"
        }
        if ($result.SendAsChanged -and $result.SendAsActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-As $dl ($($envResult.Environment); users: $($result.SendAsActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove DL Access'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Permissions' | Out-Null

[void]$form.ShowDialog()
#endregion
