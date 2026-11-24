#Requires -Version 5.1
<#
    SendToDL.ps1
    Lookup / Add / Remove who is allowed to send messages TO an on-prem
    Distribution List -- standalone, no target lookup.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

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

#region Actions -- on-prem Exchange, bare cmdlet names (see header collision warning)
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

#region WinForms UI -------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Send-To DL -- Allowed Senders (on-prem)"
$form.Size = New-Object System.Drawing.Size(780, 706)
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
        return [PSCustomObject]@{ Environment = 'HQ'; DCParam = @{}; Errors = $errors }
    } catch {
        $errors += "HQ: $($_.Exception.Message)"
    }
    try {
        $null = Get-DistributionGroup -Identity $DLIdentity -DomainController $script:DefaultDC -ErrorAction Stop
        return [PSCustomObject]@{ Environment = 'Franchise'; DCParam = @{ DomainController = $script:DefaultDC }; Errors = $errors }
    } catch {
        $errors += "Franchise: $($_.Exception.Message)"
    }
    return [PSCustomObject]@{ Environment = $null; DCParam = $null; Errors = $errors }
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

$chkAlsoEnable = New-Object System.Windows.Forms.CheckBox
$chkAlsoEnable.Text = "On Add: also enable 'require sender authentication' if currently OFF (needed for this list to have any effect -- see warning)"
$chkAlsoEnable.Location = New-Object System.Drawing.Point(12, 246)
$chkAlsoEnable.Size = New-Object System.Drawing.Size(740, 22)
$form.Controls.Add($chkAlsoEnable)

$btnLookup = New-Object System.Windows.Forms.Button
$btnLookup.Text = "Lookup"
$btnLookup.Location = New-Object System.Drawing.Point(12, 274)
$btnLookup.Size = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnLookup)
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add Send-To"
$btnAdd.Location = New-Object System.Drawing.Point(130, 274)
$btnAdd.Size = New-Object System.Drawing.Size(140, 30)
$btnAdd.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnAdd)
$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Remove Send-To"
$btnRemove.Location = New-Object System.Drawing.Point(278, 274)
$btnRemove.Size = New-Object System.Drawing.Size(150, 30)
$btnRemove.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRemove)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(436, 274)
$btnClose.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnClose)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Result:"
$lblResult.Location = New-Object System.Drawing.Point(12, 314)
$lblResult.AutoSize = $true
$form.Controls.Add($lblResult)
$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Both"
$txtOut.WordWrap = $false
$txtOut.Location = New-Object System.Drawing.Point(12, 334)
$txtOut.Size = New-Object System.Drawing.Size(740, 344)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOut)

function Set-ButtonsEnabled {
    param([bool]$Enabled)
    $btnLookup.Enabled = $Enabled
    $btnAdd.Enabled = $Enabled
    $btnRemove.Enabled = $Enabled
}

$btnLookup.Add_Click({
    Set-ButtonsEnabled -Enabled $false
    try {
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }
        $state = Get-SendToState -DLIdentity $dl -DCParam $envResult.DCParam
        $txtOut.Text = "Environment: $($envResult.Environment)`r`n`r`n" + (Format-SendToState -Heading "CURRENT STATE" -State $state)
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Lookup Send-To'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnAdd.Add_Click({
    Set-ButtonsEnabled -Enabled $false
    try {
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $users = ConvertTo-NormalizedUserList -Text $txtUsers.Text
        if ($users.Count -eq 0) { $txtOut.Text = "Enter at least one user in the second box first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }
        $dcParam = $envResult.DCParam
        $pre = Get-SendToState -DLIdentity $dl -DCParam $dcParam
        $authNote = ''
        if (-not $pre.RequireSenderAuthenticationEnabled -and -not $chkAlsoEnable.Checked) {
            $authNote = "`n`nNOTE: 'require sender authentication' is currently OFF for this DL -- this allowed-senders list will have NO EFFECT until that's turned on (check the box above, or enable it separately)."
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Add $($users.Count) user(s) to the Send-To (allowed senders) list for '$dl' ($($envResult.Environment))?$authNote",
            "Confirm Add Send-To",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $result = Invoke-SendToChange -DLIdentity $dl -Users $users -Verb 'Add' -DCParam $dcParam -AlsoEnableAuth $chkAlsoEnable.Checked
        $txtOut.Text = $result.Output
        if ($result.Changed -and $result.ActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Added Send-To $dl ($($envResult.Environment); users: $($result.ActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add Send-To'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnRemove.Add_Click({
    Set-ButtonsEnabled -Enabled $false
    try {
        $dl = $txtDL.Text.Trim()
        if (-not $dl) { $txtOut.Text = "Enter a Distribution List name first."; return }
        $users = ConvertTo-NormalizedUserList -Text $txtUsers.Text
        if ($users.Count -eq 0) { $txtOut.Text = "Enter at least one user in the second box first."; return }
        $envResult = Resolve-DLEnvironment -DLIdentity $dl
        if (-not $envResult.Environment) {
            $txtOut.Text = "'$dl' was not found in HQ or Franchise.`r`n`r`n" + ($envResult.Errors -join "`r`n")
            return
        }
        $dcParam = $envResult.DCParam
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($users.Count) user(s) from the Send-To (allowed senders) list for '$dl' ($($envResult.Environment))?",
            "Confirm Remove Send-To",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $result = Invoke-SendToChange -DLIdentity $dl -Users $users -Verb 'Remove' -DCParam $dcParam -AlsoEnableAuth $false
        $txtOut.Text = $result.Output
        if ($result.Changed -and $result.ActedOn.Count -gt 0) {
            Write-CompletedLog -Task "Removed Send-To $dl ($($envResult.Environment); users: $($result.ActedOn -join '+'))"
        }
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove Send-To'
    } finally {
        Set-ButtonsEnabled -Enabled $true
    }
})

$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form
[void]$form.ShowDialog()
#endregion
