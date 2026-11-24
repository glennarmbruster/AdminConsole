#Requires -Version 5.1
<#
    NotDeadYet.ps1
    Re-enable a returning user.

    Author   : Glenn Armbruster
    Revised  : 2026-07-28
#>
# TOOLKIT-HELP-BEGIN
# Reverses a termination for someone who is coming back, or who was disabled
# by mistake.
#
# It undoes what the disable pipeline did:
#
#   * un-hides the person from the address lists
#   * turns the mailbox back into a regular user mailbox (it is left as a
#     shared mailbox by the disable)
#   * clears any forwarding that was set up to redirect their mail
#
# What it does NOT do: re-enable the AD account, re-assign a licence, or
# restore group memberships. Those are separate steps and this tool
# deliberately does not guess at them.
#
# It touches both directories -- the on-premises object and the cloud
# mailbox -- so both connections need to be live.
# TOOLKIT-HELP-END

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

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }


#region Parsing --------------------------------------------------------------

function Get-LogonList {
    param([string]$Text)

    if (-not $Text) { return @() }

    $items = $Text -split '[\r\n,;]+' |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -ne '' } |
        Select-Object -Unique

    return @($items)
}

function Parse-Request {
    param([string]$Text)

    $logons = Get-LogonList -Text $Text

    [PSCustomObject]@{
        Logons = $logons
        Count  = $logons.Count
        Valid  = ($logons.Count -gt 0 -and $logons.Count -le 15)
    }
}

#endregion

#region Actions ---------------------------------------------------------------


# The licence group comes from Tenant.json, not from a literal.
#
# It was 'E5 License Fran - Full'. $script:DC has always resolved from config
# while the group it pairs with was hardcoded, so in any tenant without that
# exact group name this script pointed a correctly-resolved domain controller
# at a group that does not exist there.
# THIS TOOL IS FOR ONE DIRECTORY, AND IT IS NAMED HERE.
#
# It used to take whichever directory was marked default -- which is HQ. So a
# franchise restore looked up an HQ licence group, and the window title said
# "Restore HQ Users" while sitting in front of someone restoring a franchise
# account. The title was not lying; it was reporting the wrong value faithfully.
#
# Named rather than inferred. "The one that is not the default" works with two
# directories and breaks with three, and a tool that writes group membership
# should not be guessing which directory it is aimed at.
$script:RestoreDirectoryId = 'Franchise'

# Which licence a returning user gets back. Matched against the Sku column in
# Tenant Manager, NOT against the group's position in the list -- the list is
# ordered by whatever was typed first, and reading $groups[0] is what removed
# an E1 licence during a termination.
$script:RestoreSku = 'E5'

$script:ADScopeWarning = ''
$script:DirectoryId = ''
foreach ($d in @(Get-TenantDirectory)) {
    if ("$($d.Id)" -ieq $script:RestoreDirectoryId) { $script:DirectoryId = "$($d.Id)" }
}
if (-not $script:DirectoryId) {
    $null = Show-ToolDialog(
        ("This tool restores accounts in the '$($script:RestoreDirectoryId)' directory, and " +
         "Tenant.json does not have one." + "`r`n`r`n" +
         "Add it in Tenant Manager under Directories, or change " +
         "`$script:RestoreDirectoryId at the top of this script."),
        'Directory not configured', 'OK', 'Warning')
    return
}

# AFTER the directory is chosen, and scoped to it. This used to run higher up
# with no argument, so it returned the DEFAULT directory's server -- which is
# how the banner came to say "pinned to expresspersonnel.com" on a tool that
# only ever touches franchise accounts.
# -Id, not -DirectoryId. The parameter is named -Id on this function and
# -DirectoryId on its neighbours, and PowerShell does not error loudly enough
# about it: the binding simply fails, $script:DC stays $null, and every
# -Server $null call quietly falls back to the CURRENT domain. That is how a
# franchise lookup ended up searching DC=expresspersonnel,DC=com.
$script:DC = (Get-TenantDirectoryServer -Id $script:DirectoryId)
if (-not $script:DC) {
    $null = Show-ToolDialog(
        ("No server could be resolved for directory '$($script:DirectoryId)'." + "`r`n`r`n" +
         "Check its Domain and DCs columns in Tenant Manager."),
        'Directory server not resolved', 'OK', 'Warning')
    return
}

$script:Group = Resolve-TenantLicenseGroupName -DirectoryId $script:DirectoryId -Sku $script:RestoreSku

function Test-LicenceGroupConfigured {
    <#
        Guard for the two places that name the group to AD.

        An unconfigured group does not fail politely anywhere downstream: the
        membership test throws about a value the operator cannot act on, and an
        empty group name would otherwise be passed to a write. Saying WHICH
        setting is missing is the entire value here.
    #>
    if ($script:Group) { return $true }
    $null = Show-ToolDialog(
        ("No '$($script:RestoreSku)' licence group is configured for directory " +
         "'$($script:DirectoryId)'." + "`r`n`r`n" +
         "Add one in Tenant Manager under Licence groups -- the Sku column must say " +
         "'$($script:RestoreSku)' and the Directory column '$($script:DirectoryId)' -- " +
         "then reopen this tool."),
        'Licence group not configured', 'OK', 'Warning')
    return $false
}

function Initialize-ADSession {
    <#
        $script:DC is whatever the directory resolves to, and that is a DOMAIN
        when the directory names several controllers. Correct for -Server on the
        AD calls below; rejected by -PreferredServer with error 0x52.

        Scoped to THIS directory. Without that, a controller belonging to another
        directory validates and gets pinned -- which is how an HQ write ends up
        aimed at a franchise controller and fails as a permissions problem.

        A refused pin no longer stops the tool: Set-ADSessionScope falls back to
        forest-wide and reports it.
    #>
    $scope = Set-ADSessionScope -Candidate $script:DC -DirectoryId $script:DirectoryId
    $script:PinDC    = $scope.Pinned
    $script:ADScope  = $scope.Detail
    if ($scope.Fallback) { $script:ADScopeWarning = $scope.Detail } else { $script:ADScopeWarning = '' }
}

function Get-OnPremState {
    param([string]$LogonName)

    Get-RemoteMailbox -Identity $LogonName -ErrorAction Stop |
        Select-Object PrimarySmtpAddress,
                      HiddenFromAddressListsEnabled,
                      RecipientTypeDetails
}

function Get-CloudState {
    param([string]$PrimarySmtpAddress)

    # Cloud, explicitly. This script calls Set-RemoteMailbox a few lines down,
    # so BOTH Exchanges are loaded and a bare Get-Mailbox goes to whichever was
    # imported last. See Get-ExchangeModuleMap.
    Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $PrimarySmtpAddress; ErrorAction = 'Stop' } |
        Select-Object HiddenFromAddressListsEnabled,
                      RecipientTypeDetails,
                      ForwardingAddress,
                      ForwardingSmtpAddress
}

function Test-GroupMembership {
    param([string]$LogonName)

    # Returning $false on an unconfigured group would read as "not a member",
    # and the caller would then try to add them to a group named by an empty
    # string. Throw instead, so the reason reaches the operator.
    if (-not $script:Group) {
        throw "No licence group is configured for directory '$($script:DirectoryId)'. Add one in Tenant Manager."
    }
    # NOT Get-ADGroupMember. IT CAPS AT 5,000 MEMBERS AND THROWS.
    #
    # "The size limit for this request was exceeded" is what a licence group
    # larger than the AD web service's MaxGroupOrMemberEntries returns, and a
    # company-wide E5 group passes that long before anybody notices. The tool
    # then fails for EVERY user rather than for some of them.
    #
    # It was the wrong shape at any size, too: enumerating a whole group to
    # answer a yes/no question about ONE person. Get-UserLicenseState asks the
    # opposite question -- it binds to the user and tests membership from that
    # end -- so it costs the same whether the group holds ten people or forty
    # thousand.
    #
    # NESTED MEMBERSHIP COUNTS, AND THAT MATTERS HERE. Licences are granted
    # through nested groups, so a direct-membership test would report "not a
    # member" for somebody who IS licensed, and this tool would then add them
    # to a group they are already effectively in. Get-UserLicenseState resolves
    # the chain with LDAP_MATCHING_RULE_IN_CHAIN -- the same code Mailbox Health
    # and New User call, rather than a third copy of the rule (FUN001).
    $lic = Get-UserLicenseState -SamAccountName $LogonName `
               -DirectoryId $script:DirectoryId -Server $script:DC -ErrorAction Stop

    # UNKNOWN IS NOT "NO". An unreachable controller returning $false here would
    # read as "not a member", and the caller would add somebody to a group it
    # could not actually see -- the same failure the guard above exists to stop.
    if (-not $lic.Known) {
        throw "Could not read licence group membership for '$LogonName': $($lic.Reason)"
    }
    return [bool](@($lic.Groups) -contains $script:Group)
}

function Format-State {
    param(
        [string]$Heading,
        $OnPrem,
        $Cloud,
        [bool]$InGroup
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("  --- $Heading ---")

    if ($OnPrem) {
        [void]$sb.AppendLine("  On-Prem PrimarySmtp   : $($OnPrem.PrimarySmtpAddress)")
        [void]$sb.AppendLine("  On-Prem Hidden        : $($OnPrem.HiddenFromAddressListsEnabled)")
        [void]$sb.AppendLine("  On-Prem RecipientType : $($OnPrem.RecipientTypeDetails)")
    } else {
        [void]$sb.AppendLine("  On-Prem mailbox not found.")
    }

    if ($Cloud) {
        [void]$sb.AppendLine("  EXO Hidden            : $($Cloud.HiddenFromAddressListsEnabled)")
        [void]$sb.AppendLine("  EXO RecipientType     : $($Cloud.RecipientTypeDetails)")
        [void]$sb.AppendLine("  EXO FwdAddress        : $($Cloud.ForwardingAddress)")
        [void]$sb.AppendLine("  EXO FwdSmtp           : $($Cloud.ForwardingSmtpAddress)")
    } else {
        [void]$sb.AppendLine("  EXO mailbox not found / not yet resolvable.")
    }

    [void]$sb.AppendLine("  In '$script:Group'     : $InGroup")

    $sb.ToString()
}

# Write-CompletedLog now comes from Common.ps1 (was a private copy here).
function Restore-User {
    param([string]$LogonName)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("===== $LogonName =====")

    Initialize-ADSession

    $beforeOnPrem  = Get-OnPremState -LogonName $LogonName
    $beforeInGroup = Test-GroupMembership -LogonName $LogonName
    $primarySmtp   = $beforeOnPrem.PrimarySmtpAddress

    $beforeCloud = $null
    if ($primarySmtp) {
        try { $beforeCloud = Get-CloudState -PrimarySmtpAddress $primarySmtp } catch { $beforeCloud = $null }
    }
    [void]$sb.Append((Format-State -Heading "BEFORE" -OnPrem $beforeOnPrem -Cloud $beforeCloud -InGroup $beforeInGroup))

    Invoke-OnPremCommand -Name 'Set-RemoteMailbox' -Parameters @{ Identity = $LogonName; HiddenFromAddressListsEnabled = $false; ErrorAction = 'Stop' }

    Invoke-OnPremCommand -Name 'Set-RemoteMailbox' -Parameters @{ Identity = $LogonName; Type = 'Regular'; ErrorAction = 'Stop' }

    if ($beforeInGroup) {
        [void]$sb.AppendLine("  SKIP: already a member of '$script:Group'.")
    }
    else {
        # -Filter returns NOTHING when there is no match -- that is not an
        # error, so -ErrorAction Stop never fires and $null travels on to
        # Add-ADGroupMember, which complains about a parameter instead of
        # naming the account. Checked here so the message says who and where.
        #
        # Worth knowing when this fires: the Set-RemoteMailbox calls above
        # will already have SUCCEEDED, because Exchange searches the whole
        # forest while Get-ADUser -Server searches one domain. An account in
        # another directory is found by one and not the other.
        $adUser = Get-ADUser -Filter "SamAccountName -eq '$LogonName'" -Server $script:DC -ErrorAction Stop
        if (-not $adUser) {
            [void]$sb.AppendLine("  FAILED: '$LogonName' was not found in Active Directory on '$script:DC'.")
            [void]$sb.AppendLine("          The mailbox was un-hidden and set back to Regular, but the")
            [void]$sb.AppendLine("          licence group was NOT changed -- this restore is incomplete.")
            [void]$sb.AppendLine("          If the account lives in another directory, run this from that one.")
            Write-FailedLog -Action 'Restore' -Target $LogonName `
                -Detail "Not found in AD on '$script:DC'; licence group '$script:Group' not updated"
        }
        else {
            Get-ADGroup -Identity $script:Group -Server $script:DC -ErrorAction Stop |
                Add-ADGroupMember -Members $adUser -ErrorAction Stop
            [void]$sb.AppendLine("  Added to '$script:Group'.")
        }
    }

    if ($primarySmtp) {
        # Cloud, explicitly -- the mailbox lives in Exchange Online; the
        # on-premises object was handled above with Set-RemoteMailbox. Bare
        # Set-Mailbox here would clear forwarding on whichever directory won
        # the name, and on the wrong one it silently does nothing useful.
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity              = $primarySmtp
            Type                  = 'Regular'
            ForwardingAddress     = $null
            ForwardingSmtpAddress = $null
            ErrorAction           = 'Stop'
        }
    }
    else {
        [void]$sb.AppendLine("  WARNING: no PrimarySmtpAddress resolved on-prem; skipped EXO mailbox update.")
    }

    Start-Sleep -Seconds 10

    $afterOnPrem  = Get-OnPremState -LogonName $LogonName
    $afterInGroup = Test-GroupMembership -LogonName $LogonName
    $afterCloud   = $null
    if ($primarySmtp) {
        try { $afterCloud = Get-CloudState -PrimarySmtpAddress $primarySmtp } catch { $afterCloud = $null }
    }
    [void]$sb.Append((Format-State -Heading "AFTER" -OnPrem $afterOnPrem -Cloud $afterCloud -InGroup $afterInGroup))

    Write-CompletedLog -Action 'Restore' -Target $LogonName -Result 'Success' `
        -Environment "$script:DirectoryId" `
        -Detail "Restored returning user $LogonName in $script:DirectoryId."

    $sb.ToString()
}

function Verify-User {
    param([string]$LogonName)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("===== $LogonName (Verify) =====")
    [void]$sb.AppendLine("  DC queried: $script:DC")

    Initialize-ADSession

    $onPrem  = Get-OnPremState -LogonName $LogonName
    $inGroup = Test-GroupMembership -LogonName $LogonName
    $cloud   = $null
    if ($onPrem.PrimarySmtpAddress) {
        try { $cloud = Get-CloudState -PrimarySmtpAddress $onPrem.PrimarySmtpAddress } catch { $cloud = $null }
    }
    [void]$sb.Append((Format-State -Heading "CURRENT STATE" -OnPrem $onPrem -Cloud $cloud -InGroup $inGroup))

    $sb.ToString()
}

#endregion

#region WinForms UI -------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Restore $script:DirectoryId Users (Returning from Term)"
$form.ClientSize = New-Object System.Drawing.Size(744, 686)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Make sure you are connected to BOTH on-prem Exchange ($script:DirectoryId) AND Exchange Online before running.`nThis tool is pinned to $script:DC only -- do not run it against another DC."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(720, 34)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblDC = New-Object System.Windows.Forms.Label
$lblDC.Text = "Domain Controller (fixed):"
$lblDC.Location = New-Object System.Drawing.Point(12, 50)
$lblDC.AutoSize = $true
$form.Controls.Add($lblDC)

$txtDC = New-Object System.Windows.Forms.TextBox
$txtDC.Text = $script:DC
$txtDC.Location = New-Object System.Drawing.Point(150, 47)
$txtDC.Size = New-Object System.Drawing.Size(400, 20)
$txtDC.ReadOnly = $true
$txtDC.Enabled = $false
$txtDC.Font = New-Object System.Drawing.Font("Consolas", 9)
$lblDC.Visible = $false
$txtDC.Visible = $false
# DC selector hidden 2026-08-03: the toolkit binds to the default franchise
# controller from Common.ps1 and nobody was choosing another one. The control
# is kept rather than deleted so it still supplies that default to the code
# below -- .Visible is the whole change, and reversing it is one word.
$form.Controls.Add($txtDC)

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Logon IDs (one per line -- max 15):"
$lblInput.Location = New-Object System.Drawing.Point(12, 78)
$lblInput.AutoSize = $true
$form.Controls.Add($lblInput)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Multiline = $true
$txtInput.ScrollBars = "Vertical"
$txtInput.Location = New-Object System.Drawing.Point(12, 98)
$txtInput.Size = New-Object System.Drawing.Size(720, 140)
$txtInput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtInput)

$lblParsed = New-Object System.Windows.Forms.Label
$lblParsed.Text = "Parsed Preview:"
$lblParsed.Location = New-Object System.Drawing.Point(12, 244)
$lblParsed.AutoSize = $true
$form.Controls.Add($lblParsed)

$txtParsed = New-Object System.Windows.Forms.TextBox
$txtParsed.Multiline = $true
$txtParsed.ReadOnly = $true
$txtParsed.ScrollBars = "Vertical"
$txtParsed.BackColor = [System.Drawing.Color]::WhiteSmoke
$txtParsed.Location = New-Object System.Drawing.Point(12, 264)
$txtParsed.Size = New-Object System.Drawing.Size(720, 110)
$txtParsed.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtParsed)

$grpMode = New-Object System.Windows.Forms.GroupBox
$grpMode.Text = "Mode"
$grpMode.Location = New-Object System.Drawing.Point(12, 382)
$grpMode.Size = New-Object System.Drawing.Size(160, 60)
$form.Controls.Add($grpMode)

$radRestore = New-Object System.Windows.Forms.RadioButton
$radRestore.Text = "Restore"
$radRestore.Location = New-Object System.Drawing.Point(10, 18)
$radRestore.AutoSize = $true
$radRestore.Checked = $true
$grpMode.Controls.Add($radRestore)

$radVerify = New-Object System.Windows.Forms.RadioButton
$radVerify.Text = "Verify"
$radVerify.Location = New-Object System.Drawing.Point(10, 36)
$radVerify.AutoSize = $true
$grpMode.Controls.Add($radVerify)

$btnParse = New-Object System.Windows.Forms.Button
$btnParse.Text = "Parse"
$btnParse.Location = New-Object System.Drawing.Point(190, 400)
$btnParse.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnParse)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run"
$btnRun.Location = New-Object System.Drawing.Point(290, 400)
$btnRun.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnRun)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(390, 400)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)

$lblOut = New-Object System.Windows.Forms.Label
$lblOut.Text = "Result:"
$lblOut.Location = New-Object System.Drawing.Point(12, 454)
$lblOut.AutoSize = $true
$form.Controls.Add($lblOut)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.Location = New-Object System.Drawing.Point(12, 474)
$txtOut.Size = New-Object System.Drawing.Size(720, 200)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOut)

$script:Parsed     = $null
$script:ParsedFrom = $null

function Update-ParsedPreview {
    $script:Parsed     = Parse-Request -Text $txtInput.Text
    $script:ParsedFrom = $txtInput.Text

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Logon count: $($script:Parsed.Count) (max 15)")
    if ($script:Parsed.Count -eq 0) {
        $lines.Add("No logon IDs found.")
    }
    elseif ($script:Parsed.Count -gt 15) {
        $lines.Add("ERROR: more than 15 logon IDs pasted. Split into batches of 15 or fewer.")
    }
    foreach ($l in $script:Parsed.Logons) { $lines.Add("  $l") }

    $txtParsed.Lines = $lines.ToArray()
}

$btnParse.Add_Click({
    $btnParse.Enabled = $false
    $btnRun.Enabled   = $false
    try {
        Update-ParsedPreview
    }
    catch {
        $txtOut.Text = "ERROR: $($_.Exception.Message)"
    }
    finally {
        $btnParse.Enabled = $true
        $btnRun.Enabled   = $true
    }
})

$btnRun.Add_Click({
    $btnParse.Enabled = $false
    $btnRun.Enabled   = $false
    try {
        if (-not $script:Parsed -or $txtInput.Text -ne $script:ParsedFrom) {
            Update-ParsedPreview
        }

        if (-not $script:Parsed.Valid) {
            $txtOut.Text = "Cannot run: paste between 1 and 15 logon IDs (see Parsed Preview)."
            return
        }

        $logons = $script:Parsed.Logons
        $mode   = if ($radRestore.Checked) { "Restore" } else { "Verify" }

        if ($mode -eq "Restore") {
            $list = ($logons -join "`n  ")
            $confirm = Show-ToolDialog(
                "Restore $($logons.Count) user(s) in $script:DirectoryId (DC: $script:DC):`n`n  $list`n`nThis will unhide each mailbox, clear forwarding, set mailbox type to Regular, and add to '$script:Group' if missing.`n`nContinue?",
                "Confirm Restore",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtOut.Text = "Cancelled."
                return
            }
        }

        $all = New-Object System.Text.StringBuilder
        foreach ($logon in $logons) {
            try {
                if ($mode -eq "Restore") {
                    [void]$all.Append((Restore-User -LogonName $logon))
                }
                else {
                    [void]$all.Append((Verify-User -LogonName $logon))
                }
            }
            catch {
                [void]$all.AppendLine("===== $logon =====")
                [void]$all.AppendLine("  ERROR: $($_.Exception.Message)")
            }
            [void]$all.AppendLine("")
        }

        # A fallback that nobody sees is the same as a silent one. If the
        # configured controller was refused, that goes at the TOP of the result
        # -- above the per-user output -- because it changes which directory
        # server answered everything below it.
        $txtOut.Text = $(if ($script:ADScopeWarning) {
            "NOTE: $($script:ADScopeWarning)" + [Environment]::NewLine + [Environment]::NewLine + $all.ToString()
        } else { $all.ToString() })
    }
    catch {
        $txtOut.Text = "ERROR: $($_.Exception.Message)"
    }
    finally {
        $btnParse.Enabled = $true
        $btnRun.Enabled   = $true
    }
})

$btnClose.Add_Click({ $form.Close() })

# Set-FormScreenCorner, not a fourth hand-rolled copy.
#
# This computed the working area and bottom-left corner inline -- the same
# arithmetic as Common.ps1's helper, which two other tools also duplicated.
# Four copies of one calculation is how one of them ends up subtly different.
#
# The old comment nearby said these tools must not take a hard dependency on
# Common.ps1. That was already untrue: this file dot-sources it at the top and
# calls several of its functions, so the helper is guaranteed to be there.
#
# This copy also had NO try/catch, unlike the other three. A host with no
# PrimaryScreen would have thrown here rather than falling back to centre.
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Theme if Common.ps1 happens to be loaded; these tools run standalone
# by design, so this must not become a hard dependency.
if (Get-Command Set-ConsoleTheme -ErrorAction SilentlyContinue) {
    Set-ConsoleTheme -Form $form
}
# Consistent '?' in the top-right of every tool. Added after the theme so
# it is not restyled into a normal button, and BringToFront'd inside the
# helper because WinForms draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Returning User' | Out-Null

[void]$form.ShowDialog()

#endregion
