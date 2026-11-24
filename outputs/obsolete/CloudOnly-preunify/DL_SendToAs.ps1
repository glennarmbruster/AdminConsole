#Requires -Version 5.1
<#
    DL_SendToAs.ps1
    Lookup / Add / Remove Send-To (allowed senders) and/or Send-As
    permission on a cloud Distribution List.

    Cloud-only build. Every call goes through Invoke-EXOCommand, so there is
    no domain controller to choose and no sAMAccountName anywhere -- addresses
    are the identity throughout.

    Author   : Glenn Armbruster
    Revised  : 2026-08-03
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
    [string]$SmtpAddress
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

#region Send-To actions --------------------------------------------------------
function Get-SendToState {
    <#
        Reads the allowed-senders list.

        Cloud DLs return AcceptMessagesOnlyFromSendersOrMembers as names, not
        distinguished names, so every entry is re-resolved to a primary SMTP
        address. That address is what the list is rewritten with -- keying on
        anything else means the set comparison below can miss a match that is
        really there, and silently add a duplicate.

        An entry that will not re-resolve is reported rather than dropped: it
        is usually a deleted recipient still sitting on the list, and quietly
        removing it on the next write would be a change nobody asked for.
    #>
    param([Parameter(Mandatory)][string]$DLIdentity)
    $grp = Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
        Identity = $DLIdentity; ErrorAction = 'Stop' }
    $rawSenders = @($grp.AcceptMessagesOnlyFromSendersOrMembers)
    $resolved = @($rawSenders | ForEach-Object {
        $entry = "$_"
        try {
            $r = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
                Identity = $entry; ErrorAction = 'Stop' }
            [PSCustomObject]@{
                Display = "$($r.DisplayName) <$($r.PrimarySmtpAddress)>"
                Key     = "$($r.PrimarySmtpAddress)"
            }
        } catch {
            [PSCustomObject]@{
                Display = "$entry (could not re-resolve -- may be stale/deleted)"
                Key     = $null
            }
        }
    })
    return [PSCustomObject]@{
        DisplayName                        = $grp.DisplayName
        Identity                           = "$($grp.Identity)"
        RequireSenderAuthenticationEnabled = [bool]$grp.RequireSenderAuthenticationEnabled
        Senders                            = $resolved
        RawSenders                         = $rawSenders
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
    <#
        Set-DistributionGroup replaces the whole list, so the current list is
        read, edited in memory, and written back. Unresolvable entries are
        carried through untouched -- see Get-SendToState.
    #>
    param(
        [Parameter(Mandatory)][string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb,
        [bool]$AlsoEnableAuth
    )
    $out = New-Object System.Text.StringBuilder
    $before = Get-SendToState -DLIdentity $DLIdentity
    [void]$out.Append((Format-SendToState -Heading "BEFORE" -State $before))
    [void]$out.AppendLine("")

    $resolvedUsers = @()
    foreach ($u in $Users) {
        try {
            $r = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
                Identity = $u; ErrorAction = 'Stop' }
            $resolvedUsers += [PSCustomObject]@{
                Display = "$($r.DisplayName) <$($r.PrimarySmtpAddress)>"
                Key     = "$($r.PrimarySmtpAddress)"
            }
        } catch {
            [void]$out.AppendLine("  ERROR resolving '$u' -- skipped, not included in this change: $($_.Exception.Message)")
        }
    }

    $currentKeys = @($before.Senders | Where-Object { $_.Key } | ForEach-Object { $_.Key })
    $currentSet  = New-Object System.Collections.Generic.HashSet[string]([string[]]$currentKeys, [System.StringComparer]::OrdinalIgnoreCase)
    $newList     = New-Object System.Collections.ArrayList
    foreach ($k in $currentKeys) { [void]$newList.Add($k) }
    foreach ($s in $before.Senders) {
        if (-not $s.Key) { [void]$newList.Add(($s.Display -replace ' \(could not re-resolve.*$', '')) }
    }

    $changed = $false
    $actedOn = @()
    foreach ($ru in $resolvedUsers) {
        $already = $currentSet.Contains($ru.Key)
        if ($Verb -eq 'Add') {
            if ($already) {
                [void]$out.AppendLine("  SKIP (already allowed to send to this DL): $($ru.Display)")
            } else {
                [void]$newList.Add($ru.Key)
                [void]$currentSet.Add($ru.Key)
                [void]$out.AppendLine("  ADDED: $($ru.Display)")
                $actedOn += $ru.Display
                $changed = $true
            }
        } else {
            if (-not $already) {
                [void]$out.AppendLine("  SKIP (not currently on the allowed-senders list): $($ru.Display)")
            } else {
                $newList.Remove($ru.Key)
                [void]$currentSet.Remove($ru.Key)
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

    Invoke-EXOCommand -Name 'Set-DistributionGroup' -Parameters @{
        Identity = $DLIdentity
        AcceptMessagesOnlyFromSendersOrMembers = $newList.ToArray()
        Confirm = $false; ErrorAction = 'Stop' } | Out-Null

    if ($Verb -eq 'Add' -and $AlsoEnableAuth -and -not $before.RequireSenderAuthenticationEnabled) {
        Invoke-EXOCommand -Name 'Set-DistributionGroup' -Parameters @{
            Identity = $DLIdentity; RequireSenderAuthenticationEnabled = $true
            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
        [void]$out.AppendLine("")
        [void]$out.AppendLine("Also enabled 'require sender authentication' (was OFF) -- this list is now actually enforced.")
    }

    $after = Get-SendToState -DLIdentity $DLIdentity
    [void]$out.AppendLine("")
    [void]$out.Append((Format-SendToState -Heading "AFTER" -State $after))

    return [PSCustomObject]@{ Output = $out.ToString(); Changed = $true; ActedOn = $actedOn }
}
#endregion

#region Send-As actions --------------------------------------------------------
function Resolve-CloudRecipient {
    <#
        Confirms a user exists and returns the handle to grant against.

        Cloud-only, so there is no sAMAccountName to find. Send As is granted
        to a recipient and the primary SMTP address is the stable identifier,
        which is also what Get-RecipientPermission reports back as Trustee --
        so the same value compares cleanly on both sides.
    #>
    param([Parameter(Mandatory)][string]$Identity)
    $r = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
        Identity = $Identity; ErrorAction = 'Stop' }
    return "$($r.PrimarySmtpAddress)"
}

function Get-SendAsState {
    <#
        Who currently holds Send As on this group.

        Get-RecipientPermission is the cloud equivalent of the on-premises
        Get-ADPermission -ExtendedRights Send-As. NT AUTHORITY\SELF is on every
        object and is the platform's own grant, so it is filtered rather than
        reported as somebody's delegation.

        Deny entries are kept and flagged. A Deny outranks an Allow, so hiding
        one would mean reporting access that does not actually work.
    #>
    param([Parameter(Mandatory)][string]$DLIdentity)
    $grants = @()
    $perms = @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{
        Identity = $DLIdentity; ErrorAction = 'Stop' })
    foreach ($p in $perms) {
        $trustee = "$($p.Trustee)"
        if ($trustee -match 'NT AUTHORITY\\SELF') { continue }
        if ("$($p.AccessRights)" -notmatch 'SendAs') { continue }
        $grants += [PSCustomObject]@{
            User = $trustee
            Key  = $trustee
            Deny = ("$($p.AccessControlType)" -eq 'Deny')
        }
    }
    return [PSCustomObject]@{
        Identity = $DLIdentity
        Grants   = @($grants | Sort-Object User)
    }
}

function Format-SendAsState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading ---")
    [void]$sb.AppendLine("  DL              : $($State.Identity)")
    [void]$sb.AppendLine("  Send-As grants ($($State.Grants.Count)), SELF excluded:")
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
    param([Parameter(Mandatory)][string]$DLIdentity, [Parameter(Mandatory)][string]$UserIdentity)
    Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
        Identity = $DLIdentity; Trustee = $UserIdentity; AccessRights = 'SendAs'
        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
}

function Remove-SendAsGrant {
    param([Parameter(Mandatory)][string]$DLIdentity, [Parameter(Mandatory)][string]$UserIdentity)
    Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
        Identity = $DLIdentity; Trustee = $UserIdentity; AccessRights = 'SendAs'
        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
}

function Invoke-SendAsChange {
    param(
        [Parameter(Mandatory)][string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb
    )
    $out = New-Object System.Text.StringBuilder
    $before = Get-SendAsState -DLIdentity $DLIdentity
    [void]$out.Append((Format-SendAsState -Heading "BEFORE" -State $before))
    [void]$out.AppendLine("")

    $beforeKeys = New-Object System.Collections.Generic.HashSet[string](
        [string[]]@($before.Grants | ForEach-Object { $_.Key }),
        [System.StringComparer]::OrdinalIgnoreCase)

    $changed = $false
    $actedOn = @()
    foreach ($u in $Users) {
        try {
            $key = Resolve-CloudRecipient -Identity $u
        } catch {
            [void]$out.AppendLine("  ERROR resolving '$u' via Get-Recipient -- skipped, not included in this change: $($_.Exception.Message)")
            continue
        }
        $already = $beforeKeys.Contains($key)
        if ($Verb -eq 'Add') {
            if ($already) {
                [void]$out.AppendLine("  SKIP (already has Send-As on this DL): $key")
                continue
            }
            try {
                Add-SendAsGrant -DLIdentity $DLIdentity -UserIdentity $key
                [void]$out.AppendLine("  ADDED: $key")
                $actedOn += $key
                $changed = $true
            } catch {
                [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Add-RecipientPermission ($key -> $DLIdentity)"))
            }
        } else {
            if (-not $already) {
                [void]$out.AppendLine("  SKIP (does not currently have Send-As on this DL): $key")
                continue
            }
            try {
                Remove-SendAsGrant -DLIdentity $DLIdentity -UserIdentity $key
                [void]$out.AppendLine("  REMOVED: $key")
                $actedOn += $key
                $changed = $true
            } catch {
                [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Remove-RecipientPermission ($key -> $DLIdentity)"))
            }
        }
    }

    if (-not $changed) {
        [void]$out.AppendLine("")
        [void]$out.AppendLine("No change -- every requested user was already in the desired state (or failed to resolve). Nothing written.")
        return [PSCustomObject]@{ Output = $out.ToString(); Changed = $false; ActedOn = @() }
    }

    $after = Get-SendAsState -DLIdentity $DLIdentity
    [void]$out.AppendLine("")
    [void]$out.Append((Format-SendAsState -Heading "AFTER" -State $after))

    return [PSCustomObject]@{ Output = $out.ToString(); Changed = $true; ActedOn = $actedOn }
}
#endregion

#region Combined orchestration ---------------------------------------------------

function Invoke-DLAccessChange {
    param(
        [Parameter(Mandatory)][string]$DLIdentity,
        [string[]]$Users,
        [ValidateSet('Add', 'Remove')][string]$Verb,
        [bool]$DoSendTo,
        [bool]$DoSendAs,
        [bool]$AlsoEnableAuth
    )
    $out = New-Object System.Text.StringBuilder
    $sendToChanged = $false; $sendToActedOn = @()
    $sendAsChanged = $false; $sendAsActedOn = @()

    if ($DoSendTo) {
        [void]$out.AppendLine("=== Send-To (allowed senders) ===")
        $r = Invoke-SendToChange -DLIdentity $DLIdentity -Users $Users -Verb $Verb -AlsoEnableAuth ($Verb -eq 'Add' -and $AlsoEnableAuth)
        [void]$out.Append($r.Output)
        [void]$out.AppendLine("")
        $sendToChanged = $r.Changed
        $sendToActedOn = $r.ActedOn
    }
    if ($DoSendAs) {
        [void]$out.AppendLine("=== Send-As (permission delegation) ===")
        $r2 = Invoke-SendAsChange -DLIdentity $DLIdentity -Users $Users -Verb $Verb
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
$form.Text = "DL Access Manager -- Send-To / Send-As (cloud)"
$form.Size = New-Object System.Drawing.Size(780, 720)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Exchange Online only. Every call is routed through the cloud session explicitly, so a stray on-premises session in this window cannot capture Get-/Set-DistributionGroup."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(740, 32)
$form.Controls.Add($lblWarn)

function Resolve-DLEnvironment {
    <#
        Confirms the distribution list exists before anything tries to change
        it, so a typo produces one clear message instead of two half-finished
        Send-To / Send-As passes.
    #>
    param([Parameter(Mandatory)][string]$DLIdentity)

    $errors = @()
    try {
        $g = Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
            Identity = $DLIdentity; ErrorAction = 'Stop' }
        return [PSCustomObject]@{ Found = $true; Label = "$($g.PrimarySmtpAddress)"; Errors = $errors }
    } catch {
        $errors += $_.Exception.Message
    }
    return [PSCustomObject]@{ Found = $false; Label = $null; Errors = $errors }
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
        if (-not $envResult.Found) {
            $txtOut.Text = "'$dl' was not found. Confirm the name, and confirm this window is connected to Exchange Online.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }
        $out = New-Object System.Text.StringBuilder
        [void]$out.AppendLine("DL: $($envResult.Label)")
        [void]$out.AppendLine("")
        if ($chkSendTo.Checked) {
            $stateTo = Get-SendToState -DLIdentity $dl
            [void]$out.AppendLine("=== Send-To (allowed senders) ===")
            [void]$out.Append((Format-SendToState -Heading "CURRENT STATE" -State $stateTo))
            [void]$out.AppendLine("")
        }
        if ($chkSendAs.Checked) {
            $stateAs = Get-SendAsState -DLIdentity $dl
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
        if (-not $envResult.Found) {
            $txtOut.Text = "'$dl' was not found. Confirm the name, and confirm this window is connected to Exchange Online.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }

        $authNote = ''
        if ($chkSendTo.Checked -and -not $chkAlsoEnable.Checked) {
            $pre = Get-SendToState -DLIdentity $dl
            if (-not $pre.RequireSenderAuthenticationEnabled) {
                $authNote = "`n`nNOTE: 'require sender authentication' is currently OFF for this DL -- the Send-To list will have NO EFFECT until that's turned on (check the box above, or enable it separately)."
            }
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Add $($users.Count) user(s) to $(Get-AccessTypeSummary) on '$dl' ($($envResult.Label))?$authNote",
            "Confirm Add Access",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }

        $result = Invoke-DLAccessChange -DLIdentity $dl -Users $users -Verb 'Add' `
            -DoSendTo $chkSendTo.Checked -DoSendAs $chkSendAs.Checked -AlsoEnableAuth $chkAlsoEnable.Checked
        $txtOut.Text = $result.Output

        if ($result.SendToChanged -and $result.SendToActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-To $dl ($($envResult.Found); users: $($result.SendToActedOn -join '+'))"
        }
        if ($result.SendAsChanged -and $result.SendAsActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-As $dl ($($envResult.Found); users: $($result.SendAsActedOn -join '+'))"
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
        if (-not $envResult.Found) {
            $txtOut.Text = "'$dl' was not found. Confirm the name, and confirm this window is connected to Exchange Online.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($users.Count) user(s) from $(Get-AccessTypeSummary) on '$dl' ($($envResult.Label))?",
            "Confirm Remove Access",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }

        $result = Invoke-DLAccessChange -DLIdentity $dl -Users $users -Verb 'Remove' `
            -DoSendTo $chkSendTo.Checked -DoSendAs $chkSendAs.Checked -AlsoEnableAuth $false
        $txtOut.Text = $result.Output

        if ($result.SendToChanged -and $result.SendToActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-To $dl ($($envResult.Found); users: $($result.SendToActedOn -join '+'))"
        }
        if ($result.SendAsChanged -and $result.SendAsActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-As $dl ($($envResult.Found); users: $($result.SendAsActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove DL Access'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Permissions' | Out-Null
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
#endregion
