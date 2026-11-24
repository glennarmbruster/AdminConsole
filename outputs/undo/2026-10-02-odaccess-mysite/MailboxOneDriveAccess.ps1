#Requires -Version 5.1
<#
    MailboxOneDriveAccess.ps1
    WinForms tool to grant or remove Exchange mailbox access (Full Access,
    Send-As) and/or OneDrive site collection admin access -- merged from the
    operator's two separate tools, Set-MailboxAccess.ps1 and
    onedriveaccess.ps1, into a single console.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# THE TOOL OPENS EVERYWHERE; THE BUTTONS DECIDE WHAT IS POSSIBLE.
#
# This script calls Set-SPOUser, so the capability scanner read it as needing
# SharePoint Online and hid the entire tool in a tenant with no SPO -- taking
# the mailbox half, which is what people actually open it for, with it. The
# directive below tells the scanner the SPO part is optional. The contract
# that comes with saying so is honoured in Update-ConnectionGating: the window
# opens regardless and disables only what the live connections cannot do.
# TOOLKIT-CAPABILITY-OPTIONAL: ExO, SPO
#
# BOTH, not just SPO. Exchange Online is optional here for the same reason:
# with nothing connected the console greyed the button, so the one window that
# explains which service is missing could not be opened. The tool is now
# responsible for its own gating -- see Update-ConnectionGating -- and with
# neither service connected it opens and offers nothing, which is the honest
# answer and a readable one.
# TOOLKIT-HELP-BEGIN
# Grants or removes access to somebody's mailbox and their OneDrive, in one
# place.
#
# The usual case is someone leaving or being off unexpectedly, when their
# manager needs both. Doing them separately is how one of the two gets
# forgotten.
#
#   * Mailbox -- Full Access and Send As
#   * OneDrive -- site collection administrator on their personal site
#
# OneDrive access is visible to the person if they return. It is appropriate for
# a departure, less so for a temporary absence, and either way it should be
# removed when it is no longer needed.
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


#region Config  ---------------------------------------------------------------

$Domain              = "@$(Get-MailDomain)"
$MySiteBase          = "$((Get-TenantConfig).SharePoint.MySiteHost)/personal"
# The log path now comes from Common.ps1 (Get-CompletedLogPath), which is
# also what decides the schema. Leaving these here would be a second
# definition of where the audit log lives.

#endregion

#region Helpers  --------------------------------------------------------------

function Get-AddressList {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }
    $tokens = $Raw -split "[\r\n,;]+"
    $clean = foreach ($t in $tokens) {
        $name = $t.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($name -notmatch '@') { $name = $name + $Domain }
        $name
    }
    return @($clean | Sort-Object -Unique)
}

function Get-OneDriveUrl {
    <#
        A personal site URL: <mysite>/personal/<local>_<domain with dots as
        underscores>.

        THE DOMAIN IS DERIVED, NOT LITERAL. This used to end in a hardcoded
        '_contoso_com'-style suffix: the tenant domain written the way OneDrive
        writes it -- dots replaced by underscores.

        That form is invisible to every leakage rule the toolkit has. TEN002
        matches the configured domain 'contoso.com', and 'contoso_com'
        is not that string; the word-boundary check cannot see it either
        because the underscores make it one long token. So a hardcoded tenant
        domain sat in shipping code, in the URL-building function of a tool
        that ships to every tenant, and every check passed.

        On any other tenant this produced a URL for the wrong organisation --
        the lookup simply fails, which reads as "that user has no OneDrive"
        rather than as a bug in the toolkit.
    #>
    param([string]$Email)
    $localPart = ($Email -split '@')[0]
    $urlSafe   = $localPart -replace '\.', '_'

    # Prefer the address's OWN domain. A tenant with several accepted domains
    # has personal sites under each, so assuming the primary would break every
    # secondary-domain user.
    $domain = ($Email -split '@')[1]
    if (-not $domain) { $domain = Get-MailDomain }
    $domainSafe = "$domain" -replace '\.', '_'

    return "$MySiteBase/${urlSafe}_$domainSafe"
}

# Write-CompletedLog now comes from Common.ps1 (was a private copy here).
#endregion

#region Access Type mapping  ---------------------------------------------------

function Get-AccessFlags {
    param([string]$AccessType)
    switch ($AccessType) {
        'Full'         { return [PSCustomObject]@{ Full = $true;  SendAs = $false; OneDrive = $false } }
        'SendAs'       { return [PSCustomObject]@{ Full = $false; SendAs = $true;  OneDrive = $false } }
        'OneDrive'     { return [PSCustomObject]@{ Full = $false; SendAs = $false; OneDrive = $true  } }
        'FullSendAs'   { return [PSCustomObject]@{ Full = $true;  SendAs = $true;  OneDrive = $false } }
        'FullOneDrive' { return [PSCustomObject]@{ Full = $true;  SendAs = $false; OneDrive = $true  } }
    }
}

function Get-AccessTypeText {
    param([string]$AccessType)
    switch ($AccessType) {
        'Full'         { return 'Full Access' }
        'SendAs'       { return 'Send-As' }
        'OneDrive'     { return 'OneDrive' }
        'FullSendAs'   { return 'Full Access & Send-As' }
        'FullOneDrive' { return 'Full Access & OneDrive' }
    }
}

#endregion

#region Parsing  --------------------------------------------------------------

function Get-ParsedInputs {
    param(
        [string]$MailboxRaw,
        [string]$UserRaw,
        [string]$AccessType,
        [string]$Mode
    )
    $mailboxes = Get-AddressList -Raw $MailboxRaw
    $users     = Get-AddressList -Raw $UserRaw

    $users = @($users | Where-Object { $mailboxes -notcontains $_ })

    return [PSCustomObject]@{
        Mailboxes  = $mailboxes
        Users      = $users
        AccessType = $AccessType
        Mode       = $Mode
        OpCount    = ($mailboxes.Count * $users.Count)
    }
}

#endregion

#region Actions  --------------------------------------------------------------

function Invoke-MailboxOneDriveAccess {
    param(
        [string[]]$Mailboxes,
        [string[]]$Users,
        [string]$AccessType,
        [string]$Mode
    )

    $flags      = Get-AccessFlags -AccessType $AccessType
    $isGrant    = ($Mode -eq 'Grant')
    $verb       = if ($isGrant) { 'Granted' } else { 'Removed' }
    $actionText = "$(if ($isGrant) { 'Grant' } else { 'Remove' }) $(Get-AccessTypeText -AccessType $AccessType)"

    $log = New-Object System.Text.StringBuilder

    [void]$log.AppendLine("=== Mailbox & OneDrive Access Manager ===")
    [void]$log.AppendLine("Action        : $actionText")
    [void]$log.AppendLine("Mailbox(es)   : $($Mailboxes -join ', ')")
    [void]$log.AppendLine("User(s)       : $($Users -join ', ')")
    [void]$log.AppendLine("")

    # MICROSOFT 365 GROUPS ARE NOT MAILBOXES, AND Get-Mailbox WILL NOT RETURN ONE.
    #
    # The pre-flight used Get-Mailbox for everything, so three perfectly real
    # Microsoft 365 Groups came back as "could not be resolved" -- which reads as
    # a typo and sends you checking the addresses. They were fine; the question
    # was wrong.
    #
    # Get-Recipient answers for every recipient type AND reports which one it
    # found, so the check now identifies rather than merely confirms. That
    # matters because what you can do to a Microsoft 365 Group is not what you
    # can do to a mailbox -- see the branch below.
    #
    # NOT Get-EXOMailbox: its -GroupMailbox switch is documented as "reserved for
    # internal Microsoft use" and is listed among the Get-Mailbox parameters that
    # "aren't available or functional" there.
    $script:GroupTargets = @{}

    if ($flags.Full -or $flags.SendAs) {
        $notFound = @()

        foreach ($m in $Mailboxes) {
            try {
                # -ResultSize 1 AND an explicit empty-check: Get-Recipient with a
                # null Identity returns EVERY recipient, so "did it return
                # something" is not the same question as "did it find this one".
                $r = @(Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
                    Identity = $m; ResultSize = 1; WarningAction = 'SilentlyContinue'; ErrorAction = 'Stop' })
                if ($r.Count -eq 0) { $notFound += $m; continue }
                if ("$($r[0].RecipientTypeDetails)" -match '(?i)groupmailbox') {
                    # Name, not the SMTP address -- Microsoft's own documented
                    # sample for granting Send As on a group passes .Name.
                    $script:GroupTargets[$m] = "$($r[0].Name)"
                }
            } catch { $notFound += $m }
        }

        foreach ($u in $Users) {
            try {
                $r = @(Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
                    Identity = $u; ResultSize = 1; WarningAction = 'SilentlyContinue'; ErrorAction = 'Stop' })
                if ($r.Count -eq 0) { $notFound += $u }
            } catch { $notFound += $u }
        }

        if ($notFound.Count -gt 0) {
            $list = ($notFound | Sort-Object -Unique) -join "`r`n    "
            [void]$log.AppendLine("PRE-FLIGHT FAILED -- the following could not be resolved:")
            [void]$log.AppendLine("    $list")
            [void]$log.AppendLine("")
            [void]$log.AppendLine("No changes were made. Correct the above entries and try again.")
            return $log.ToString()
        }

        # A GROUP IS REPORTED BEFORE ANYTHING RUNS, not discovered halfway through.
        if ($script:GroupTargets.Count -gt 0) {
            [void]$log.AppendLine("Pre-flight : all identities resolved OK.")
            [void]$log.AppendLine("")
            [void]$log.AppendLine("MICROSOFT 365 GROUPS DETECTED ($($script:GroupTargets.Count)):")
            foreach ($k in ($script:GroupTargets.Keys | Sort-Object)) {
                [void]$log.AppendLine("    $k")
            }
            [void]$log.AppendLine("")

            if ($flags.Full) {
                # FULL ACCESS ON A GROUP IS REFUSED, AND THIS IS A JUDGEMENT CALL
                # WORTH STATING.
                #
                # Add-MailboxPermission DOES have a -GroupMailbox switch, and
                # Microsoft says it is "required to add permissions to a Microsoft
                # 365 Group mailbox" -- so the command exists and would probably
                # succeed. What Microsoft documents NOWHERE is what the grantee can
                # then actually do: there is no example, and no description of the
                # resulting client behaviour, anywhere in their docs.
                #
                # Meanwhile the entire documented access story for these groups runs
                # through MEMBERSHIP (Add-UnifiedGroupLinks -LinkType Members).
                #
                # Issuing a command that reports success and may grant nothing usable
                # is worse than declining: the ticket gets closed and the person
                # still cannot open the group.
                [void]$log.AppendLine("FULL ACCESS IS NOT APPLIED TO MICROSOFT 365 GROUPS.")
                [void]$log.AppendLine("Access to a group mailbox comes from MEMBERSHIP, not mailbox permissions.")
                [void]$log.AppendLine("Add the user to the group (Group Editor, or the Microsoft 365 admin centre).")
                [void]$log.AppendLine("")
                [void]$log.AppendLine("Microsoft does document an Add-MailboxPermission -GroupMailbox switch, but")
                [void]$log.AppendLine("never describes what the grantee can do with it -- so this tool will not")
                [void]$log.AppendLine("report success for a change whose effect is undefined.")
                [void]$log.AppendLine("")
                [void]$log.AppendLine("No changes were made.")
                return $log.ToString()
            }

            [void]$log.AppendLine("Send-As on a Microsoft 365 Group is supported and will be applied.")
            [void]$log.AppendLine("NOTE: Send As and Send on Behalf are mutually exclusive per user --")
            [void]$log.AppendLine("      if both are set, Send As wins. Guests cannot be granted either.")
            [void]$log.AppendLine("")
        } else {
            [void]$log.AppendLine("Pre-flight : all identities resolved OK.")
            [void]$log.AppendLine("")
        }
    }

    $succeeded = [System.Collections.Generic.List[PSCustomObject]]::new()
    $failed    = [System.Collections.Generic.List[string]]::new()
    $skipped   = [System.Collections.Generic.List[string]]::new()
    $oneDriveOkMailboxes = New-Object System.Collections.Generic.List[string]

    foreach ($m in $Mailboxes) {

        if ($flags.SendAs -and $isGrant) {
            # SKIPPED FOR A MICROSOFT 365 GROUP. Set-Mailbox does not act on a
            # GroupMailbox -- it answers "The current operation is not supported
            # on GroupMailbox" -- so this would log a warning on every group for
            # a setting that does not apply to one.
            if ($script:GroupTargets.ContainsKey($m)) {
                [void]$log.AppendLine("  [SKIP] MessageCopyForSentAs does not apply to a Microsoft 365 Group ($m)")
            }
            else {
                try {
                    # MessageCopyForSentAsEnabled -- THE FULL NAME, deliberately.
                    #
                    # The previous 'MessageCopyForSentAs' worked, and that is
                    # exactly why it was worth changing. PowerShell resolves
                    # partial parameter names, including in a splat, so it bound
                    # to the real parameter as long as nothing else shared the
                    # prefix. The day Microsoft adds a second parameter starting
                    # with those characters, the binding becomes ambiguous and
                    # this fails at run time in a tool nobody has touched.
                    Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{ Identity = $m; MessageCopyForSentAsEnabled = $true; ErrorAction = 'Stop' }
                    [void]$log.AppendLine("  [OK]   MessageCopyForSentAs = true on $m")
                } catch {
                    Write-FailedLog -Action 'SetMessageCopyForSentAs' -Target $m -ErrorRecord $_ -Detail 'Send-As sent-items copy'
                    [void]$log.AppendLine("  [WARN] Could not set MessageCopyForSentAs on $m : $($_.Exception.Message)")
                }
            }
        }

        $siteUrl = $null
        if ($flags.OneDrive) {
            $siteUrl = Get-OneDriveUrl -Email $m
            [void]$log.AppendLine("  OneDrive site for $m : $siteUrl")
        }

        foreach ($u in $Users) {

            if ($flags.Full) {
                if ($isGrant) {
                    try {
                        Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{ Identity = $m; User = $u; AccessRights = 'FullAccess'; AutoMapping = $true; ErrorAction = 'Stop' } | Out-Null
                        [void]$log.AppendLine("  [OK]   Granted Full Access : $u -> $m")
                        $succeeded.Add([PSCustomObject]@{ Mailbox = $m; User = $u; Op = 'Full' })
                    } catch {
                        Write-FailedLog -Action 'GrantFullAccess' -Target $m -ErrorRecord $_ -Detail "user: $u"
                        [void]$log.AppendLine("  [FAIL] Full Access : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("Full: $u -> $m")
                    }
                } else {
                    try {
                        $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $m; User = $u; ErrorAction = 'SilentlyContinue' } |
                              Where-Object { $_.AccessRights -contains 'FullAccess' }
                        if ($fa) {
                            Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{ Identity = $m; User = $u; AccessRights = 'FullAccess'; InheritanceType = 'All'; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                            [void]$log.AppendLine("  [OK]   Removed Full Access : $u -> $m")
                            $succeeded.Add([PSCustomObject]@{ Mailbox = $m; User = $u; Op = 'Full' })
                        } else {
                            [void]$log.AppendLine("  [SKIP] No Full Access found : $u -> $m")
                            $skipped.Add("Full: $u -> $m")
                        }
                    } catch {
                        Write-FailedLog -Action 'RemoveFullAccess' -Target $m -ErrorRecord $_ -Detail "user: $u"
                        [void]$log.AppendLine("  [FAIL] Remove Full Access : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("RemoveFull: $u -> $m")
                    }
                }
            }

            if ($flags.SendAs) {
                if ($isGrant) {
                    try {
                        # A MICROSOFT 365 GROUP IS ADDRESSED BY NAME, NOT ADDRESS.
                        # Microsoft's own documented sample for granting Send As on
                        # a group resolves it with Get-Recipient and then passes
                        # $groupsRecipientDetails.Name -- captured in pre-flight.
                        $sendAsId = if ($script:GroupTargets.ContainsKey($m)) { $script:GroupTargets[$m] } else { $m }
                        Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{ Identity = $sendAsId; Trustee = $u; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        [void]$log.AppendLine("  [OK]   Granted Send-As : $u -> $m")
                        $succeeded.Add([PSCustomObject]@{ Mailbox = $m; User = $u; Op = 'SendAs' })
                    } catch {
                        Write-FailedLog -Action 'GrantSendAs' -Target $m -ErrorRecord $_ -Detail "user: $u"
                        [void]$log.AppendLine("  [FAIL] Send-As : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("SendAs: $u -> $m")
                    }
                } else {
                    try {
                        # SAME IDENTITY THE REMOVE WILL USE. Reading with the
                        # address and writing with the group Name would make the
                        # "is it already granted" check answer for a different
                        # object than the one being changed.
                        $sendAsId = if ($script:GroupTargets.ContainsKey($m)) { $script:GroupTargets[$m] } else { $m }
                        $sa = Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{
                                  Identity = $sendAsId; Trustee = $u; ErrorAction = 'SilentlyContinue' } |
                              Where-Object { $_.AccessRights -contains 'SendAs' }
                        if ($sa) {
                            Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{ Identity = $sendAsId; Trustee = $u; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                            [void]$log.AppendLine("  [OK]   Removed Send-As : $u -> $m")
                            $succeeded.Add([PSCustomObject]@{ Mailbox = $m; User = $u; Op = 'SendAs' })
                        } else {
                            [void]$log.AppendLine("  [SKIP] No Send-As found : $u -> $m")
                            $skipped.Add("SendAs: $u -> $m")
                        }
                    } catch {
                        Write-FailedLog -Action 'RemoveSendAs' -Target $m -ErrorRecord $_ -Detail "user: $u"
                        [void]$log.AppendLine("  [FAIL] Remove Send-As : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("RemoveSendAs: $u -> $m")
                    }
                }
            }

            if ($flags.OneDrive) {
                try {
                    Set-SPOUser -Site $siteUrl -LoginName $u -IsSiteCollectionAdmin $isGrant -ErrorAction Stop | Out-Null
                    [void]$log.AppendLine("  [OK]   $verb OneDrive admin : $u -> $m")
                    $succeeded.Add([PSCustomObject]@{ Mailbox = $m; User = $u; Op = 'OneDrive' })
                    if ($isGrant -and -not $oneDriveOkMailboxes.Contains($m)) {
                        $oneDriveOkMailboxes.Add($m)
                    }
                } catch {
                    Write-FailedLog -Action 'SetOneDriveAdmin' -Target $m -ErrorRecord $_ -Detail "user: $u; site: $siteUrl"
                    [void]$log.AppendLine("  [FAIL] OneDrive admin : $u -> $m : $($_.Exception.Message)")
                    $failed.Add("OneDrive: $u -> $m")
                }
            }
        }
    }

    [void]$log.AppendLine("")
    [void]$log.AppendLine(("Summary : {0} succeeded, {1} failed, {2} with nothing to do." -f `
        $succeeded.Count, $failed.Count, $skipped.Count))

    [void]$log.AppendLine("")
    [void]$log.AppendLine("=== Verification ===")
    foreach ($m in $Mailboxes) {
        [void]$log.AppendLine("")
        if ($flags.Full) {
            [void]$log.AppendLine("Full Access on $m :")
            try {
                $faList = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $m; ErrorAction = 'Stop' } |
                          Where-Object { $_.User -notlike 'NT AUTHORITY\*' -and $_.User -notlike 'S-1-5-*' }
                if ($faList) {
                    foreach ($p in $faList) { [void]$log.AppendLine("    $($p.User)  :  $($p.AccessRights -join ', ')") }
                } else {
                    [void]$log.AppendLine("    (none)")
                }
            } catch {
                [void]$log.AppendLine("    (error: $($_.Exception.Message))")
            }
        }
        if ($flags.SendAs) {
            [void]$log.AppendLine("Send-As on $m :")
            try {
                $listId = if ($script:GroupTargets.ContainsKey($m)) { $script:GroupTargets[$m] } else { $m }
                $saList = Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{
                              Identity = $listId; ErrorAction = 'Stop' } |
                          Where-Object { $_.Trustee -notlike 'NT AUTHORITY\*' }
                if ($saList) {
                    foreach ($p in $saList) { [void]$log.AppendLine("    $($p.Trustee)  :  $($p.AccessRights -join ', ')") }
                } else {
                    [void]$log.AppendLine("    (none)")
                }
            } catch {
                [void]$log.AppendLine("    (error: $($_.Exception.Message))")
            }
        }
        if ($flags.OneDrive) {
            $verifySiteUrl = Get-OneDriveUrl -Email $m
            [void]$log.AppendLine("OneDrive site collection admins on $m ($verifySiteUrl) :")
            try {
                $admins = Get-SPOUser -Site $verifySiteUrl -ErrorAction Stop | Where-Object { $_.IsSiteCollectionAdmin -eq $true }
                if ($admins) {
                    foreach ($a in $admins) { [void]$log.AppendLine("    $($a.LoginName)") }
                } else {
                    [void]$log.AppendLine("    (none)")
                }
            } catch {
                [void]$log.AppendLine("    (error reading admins: $($_.Exception.Message))")
            }
        }
    }

    if ($succeeded.Count -gt 0) {
        $userList = $Users -join '+'
        foreach ($m in $Mailboxes) {
            $mSuccess = @($succeeded | Where-Object { $_.Mailbox -eq $m })
            if ($mSuccess.Count -gt 0) {
                Write-CompletedLog -Task "$verb MailboxOneDriveAccess [$actionText] $m (users: $userList)"
            }
        }
    }

    if ($oneDriveOkMailboxes.Count -gt 0) {
        $linkLines = New-Object System.Text.StringBuilder
        foreach ($m in $oneDriveOkMailboxes) {
            [void]$linkLines.AppendLine("OneDrive can be accessed via the following link: $(Get-OneDriveUrl -Email $m)")
        }
        [void]$linkLines.AppendLine("")
        return $linkLines.ToString() + $log.ToString()
    }

    return $log.ToString()
}

#endregion

#region WinForms UI  ------------------------------------------------------------

$script:Parsed     = $null
$script:ParsedFrom = $null

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Mailbox & OneDrive Access Manager"
$form.ClientSize      = New-Object System.Drawing.Size(764, 869)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Location  = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size      = New-Object System.Drawing.Size(740, 32)
$lblWarn.Font      = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblMailbox          = New-Object System.Windows.Forms.Label
$lblMailbox.Text     = "Mailbox(es) -- one per line (enter email address, @$(Get-MailDomain) optional):"
$lblMailbox.AutoSize = $true
$lblMailbox.Location = New-Object System.Drawing.Point(12, 50)
$form.Controls.Add($lblMailbox)

$txtMailbox               = New-Object System.Windows.Forms.TextBox
$txtMailbox.Multiline     = $true
$txtMailbox.ScrollBars    = "Vertical"
$txtMailbox.Size          = New-Object System.Drawing.Size(740, 100)
$txtMailbox.Location      = New-Object System.Drawing.Point(12, 70)
$txtMailbox.Font          = New-Object System.Drawing.Font("Consolas", 9)
$txtMailbox.AcceptsReturn = $true
$form.Controls.Add($txtMailbox)

$lblUser          = New-Object System.Windows.Forms.Label
$lblUser.Text     = "User(s)/Manager(s) to grant/remove access -- one per line (enter email address, @$(Get-MailDomain) optional):"
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(12, 178)
$form.Controls.Add($lblUser)

$txtUser               = New-Object System.Windows.Forms.TextBox
$txtUser.Multiline     = $true
$txtUser.ScrollBars    = "Vertical"
$txtUser.Size          = New-Object System.Drawing.Size(740, 100)
$txtUser.Location      = New-Object System.Drawing.Point(12, 198)
$txtUser.Font          = New-Object System.Drawing.Font("Consolas", 9)
$txtUser.AcceptsReturn = $true
$form.Controls.Add($txtUser)

$lblParsed          = New-Object System.Windows.Forms.Label
$lblParsed.Text     = "Parsed preview -- review before clicking Run:"
$lblParsed.AutoSize = $true
$lblParsed.Location = New-Object System.Drawing.Point(12, 306)
$form.Controls.Add($lblParsed)

$txtParsed              = New-Object System.Windows.Forms.TextBox
$txtParsed.Multiline    = $true
$txtParsed.ScrollBars   = "Vertical"
$txtParsed.ReadOnly     = $true
$txtParsed.BackColor    = [System.Drawing.Color]::WhiteSmoke
$txtParsed.Size         = New-Object System.Drawing.Size(740, 75)
$txtParsed.Location     = New-Object System.Drawing.Point(12, 326)
$txtParsed.Font         = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtParsed)

$grpMode          = New-Object System.Windows.Forms.GroupBox
$grpMode.Text     = "Mode"
$grpMode.Size     = New-Object System.Drawing.Size(740, 50)
$grpMode.Location = New-Object System.Drawing.Point(12, 409)
$form.Controls.Add($grpMode)

$rbGrant          = New-Object System.Windows.Forms.RadioButton
$rbGrant.Text     = "Grant"
$rbGrant.AutoSize = $true
$rbGrant.Location = New-Object System.Drawing.Point(15, 20)
$rbGrant.Checked  = $true
$grpMode.Controls.Add($rbGrant)

$rbRemoveMode          = New-Object System.Windows.Forms.RadioButton
$rbRemoveMode.Text     = "Remove"
$rbRemoveMode.AutoSize = $true
$rbRemoveMode.Location = New-Object System.Drawing.Point(150, 20)
$grpMode.Controls.Add($rbRemoveMode)

$grpAccessType          = New-Object System.Windows.Forms.GroupBox
$grpAccessType.Text     = "Access Type"
$grpAccessType.Size     = New-Object System.Drawing.Size(740, 170)
$grpAccessType.Location = New-Object System.Drawing.Point(12, 467)
$form.Controls.Add($grpAccessType)

$rbFull          = New-Object System.Windows.Forms.RadioButton
$rbFull.Text     = "Full Access"
$rbFull.AutoSize = $true
$rbFull.Location = New-Object System.Drawing.Point(15, 25)
$rbFull.Checked  = $true
$grpAccessType.Controls.Add($rbFull)

$rbSendAs          = New-Object System.Windows.Forms.RadioButton
$rbSendAs.Text     = "Send-As"
$rbSendAs.AutoSize = $true
$rbSendAs.Location = New-Object System.Drawing.Point(15, 52)
$grpAccessType.Controls.Add($rbSendAs)

$rbOneDrive          = New-Object System.Windows.Forms.RadioButton
$rbOneDrive.Text     = "OneDrive"
$rbOneDrive.AutoSize = $true
$rbOneDrive.Location = New-Object System.Drawing.Point(15, 79)
$grpAccessType.Controls.Add($rbOneDrive)

$rbFullSendAs          = New-Object System.Windows.Forms.RadioButton
$rbFullSendAs.Text     = "Full Access && Send-As"
$rbFullSendAs.AutoSize = $true
$rbFullSendAs.Location = New-Object System.Drawing.Point(15, 106)
$grpAccessType.Controls.Add($rbFullSendAs)

$rbFullOneDrive          = New-Object System.Windows.Forms.RadioButton
$rbFullOneDrive.Text     = "Full Access && OneDrive"
$rbFullOneDrive.AutoSize = $true
$rbFullOneDrive.Location = New-Object System.Drawing.Point(15, 133)
$grpAccessType.Controls.Add($rbFullOneDrive)

# WHY THE STATE IS ON THE FORM AND NOT IN A POPUP. A greyed radio with no
# explanation reads as a broken tool. This says which service is missing, so
# the answer ("run Connect-SPO") is visible at the moment the button is grey.
$lblConn          = New-Object System.Windows.Forms.Label
$lblConn.Location = New-Object System.Drawing.Point(300, 22)
$lblConn.Size     = New-Object System.Drawing.Size(424, 92)
$grpAccessType.Controls.Add($lblConn)

# Connections are per-process and can be made while this window is open, so
# the operator needs a way to re-probe without losing what they have typed.
$btnRecheck          = New-Object System.Windows.Forms.Button
$btnRecheck.Text     = "Recheck connections"
$btnRecheck.Size     = New-Object System.Drawing.Size(150, 26)
$btnRecheck.Location = New-Object System.Drawing.Point(300, 122)
$grpAccessType.Controls.Add($btnRecheck)

$btnParse          = New-Object System.Windows.Forms.Button
$btnParse.Text     = "Parse"
$btnParse.Size     = New-Object System.Drawing.Size(80, 28)
$btnParse.Location = New-Object System.Drawing.Point(482, 649)
$form.Controls.Add($btnParse)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = "Run"
$btnRun.Size     = New-Object System.Drawing.Size(80, 28)
$btnRun.Location = New-Object System.Drawing.Point(570, 649)
$form.Controls.Add($btnRun)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = "Close"
$btnClose.Size     = New-Object System.Drawing.Size(84, 28)
$btnClose.Location = New-Object System.Drawing.Point(658, 649)
$form.Controls.Add($btnClose)

$lblResult          = New-Object System.Windows.Forms.Label
$lblResult.Text     = "Result:"
$lblResult.AutoSize = $true
$lblResult.Location = New-Object System.Drawing.Point(12, 687)
$form.Controls.Add($lblResult)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.ReadOnly   = $true
$txtOut.Size       = New-Object System.Drawing.Size(740, 150)
$txtOut.Location   = New-Object System.Drawing.Point(12, 707)
$txtOut.Font       = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtOut)

function Update-ConnectionGating {
    <#
        Enable only the access types the LIVE connections can actually deliver.

        EXO  -> Full Access, Send-As, Full Access & Send-As
        SPO  -> OneDrive
        both -> everything, including Full Access & OneDrive
        none -> nothing runnable, and the banner says why

        PROBED, NOT ASSUMED FROM CONFIG. Tenant.json says what a tenant HAS;
        these say what this PowerShell process is connected to right now, and
        those differ every time somebody opens the console before connecting.
        Connect-ExchangeOnline is per-process, so a session in another window
        does not count -- which is exactly the confusion this banner ends.
    #>
    $exo = $false; $spo = $false
    $exoWhy = 'Exchange Online: not checked.'
    $spoWhy = 'SharePoint Online: not checked.'
    try { $r = Test-ExoConnectionState; $exo = [bool]$r.Connected; $exoWhy = "Exchange Online: $(if ($exo) { 'connected' } else { 'NOT connected' })" } catch { $exoWhy = "Exchange Online: probe failed -- $($_.Exception.Message)" }
    try { $r = Test-SpoConnectionState; $spo = [bool]$r.Connected; $spoWhy = "SharePoint Online: $(if ($spo) { 'connected' } else { 'NOT connected' })" } catch { $spoWhy = "SharePoint Online: probe failed -- $($_.Exception.Message)" }

    $script:ExoOk = $exo
    $script:SpoOk = $spo

    $rbFull.Enabled         = $exo
    $rbSendAs.Enabled       = $exo
    $rbFullSendAs.Enabled   = $exo
    $rbOneDrive.Enabled     = $spo
    $rbFullOneDrive.Enabled = ($exo -and $spo)

    # A DISABLED RADIO CAN STILL BE THE CHECKED ONE. WinForms does not move the
    # selection when you disable it, so without this the form would sit on a
    # greyed choice and Run would attempt exactly the thing that is impossible.
    $current = @($rbFull, $rbSendAs, $rbOneDrive, $rbFullSendAs, $rbFullOneDrive) |
               Where-Object { $_.Checked } | Select-Object -First 1
    if ($current -and -not $current.Enabled) {
        $first = @($rbFull, $rbSendAs, $rbOneDrive, $rbFullSendAs, $rbFullOneDrive) |
                 Where-Object { $_.Enabled } | Select-Object -First 1
        if ($first) { $first.Checked = $true } else { $current.Checked = $false }
    }

    $any = ($exo -or $spo)
    $btnRun.Enabled   = $any
    $btnParse.Enabled = $any

    if ($any) {
        $lblConn.Text = "$exoWhy`r`n$spoWhy`r`n`r`nGreyed options need a service this process is not connected to."
        $lblConn.ForeColor = $(if ($exo -and $spo) { (Get-ToolStatusColour -Kind Muted) } else { (Get-ToolStatusColour -Kind Warn) })
    } else {
        $lblConn.Text = "$exoWhy`r`n$spoWhy`r`n`r`nNothing can be granted or removed. Connect to Exchange Online and/or SharePoint Online, then press Recheck."
        $lblConn.ForeColor = (Get-ToolStatusColour -Kind Warn)
    }
    return $any
}

function Test-AccessTypeAvailable {
    # Re-checked at RUN time, not just at open: a session can expire, or the
    # operator can connect, while the window sits there.
    param([string]$AccessType)
    switch ($AccessType) {
        'Full'         { return $script:ExoOk }
        'SendAs'       { return $script:ExoOk }
        'FullSendAs'   { return $script:ExoOk }
        'OneDrive'     { return $script:SpoOk }
        'FullOneDrive' { return ($script:ExoOk -and $script:SpoOk) }
    }
    return $false
}

function Get-SelectedAccessType {
    if ($rbFull.Checked)         { return 'Full' }
    if ($rbSendAs.Checked)       { return 'SendAs' }
    if ($rbOneDrive.Checked)     { return 'OneDrive' }
    if ($rbFullSendAs.Checked)   { return 'FullSendAs' }
    if ($rbFullOneDrive.Checked) { return 'FullOneDrive' }
}
function Get-SelectedMode {
    if ($rbRemoveMode.Checked) { return 'Remove' }
    return 'Grant'
}

function Update-ConnectionWarning {
    $accessType = Get-SelectedAccessType
    $flags = Get-AccessFlags -AccessType $accessType
    $parts = @()
    if ($flags.Full -or $flags.SendAs) { $parts += 'Exchange Online' }
    if ($flags.OneDrive)               { $parts += 'SharePoint Online' }
    $lblWarn.Text = "Make sure you are connected to $($parts -join ' and ') before running."
}
$rbFull.Add_CheckedChanged({ Update-ConnectionWarning })
$rbSendAs.Add_CheckedChanged({ Update-ConnectionWarning })
$rbOneDrive.Add_CheckedChanged({ Update-ConnectionWarning })
$rbFullSendAs.Add_CheckedChanged({ Update-ConnectionWarning })
$rbFullOneDrive.Add_CheckedChanged({ Update-ConnectionWarning })
Update-ConnectionWarning

function Update-ParsedPreview {
    param($Parsed)
    $modeText   = if ($Parsed.Mode -eq 'Remove') { 'Remove' } else { 'Grant' }
    $accessText = Get-AccessTypeText -AccessType $Parsed.AccessType

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("Mode   : $modeText")
    $lines.Add("Access : $accessText")
    $lines.Add("Mailboxes ($($Parsed.Mailboxes.Count)) : $($Parsed.Mailboxes -join ', ')")
    $lines.Add("Users/Managers ($($Parsed.Users.Count)) : $($Parsed.Users -join ', ')")
    $lines.Add("Total operations : $($Parsed.OpCount)")

    $txtParsed.Lines = $lines.ToArray()
}

$btnParse.Add_Click({
    $btnParse.Enabled = $false
    $btnRun.Enabled   = $false
    try {
        $accessType = Get-SelectedAccessType
        $mode       = Get-SelectedMode
        $p          = Get-ParsedInputs -MailboxRaw $txtMailbox.Text -UserRaw $txtUser.Text -AccessType $accessType -Mode $mode

        if ($p.Mailboxes.Count -lt 1) {
            $txtOut.Text = "Please enter at least one mailbox."
            return
        }
        if ($p.Users.Count -lt 1) {
            $txtOut.Text = "Please enter at least one user/manager. (If entries were removed due to mailbox overlap, re-check your entries.)"
            return
        }

        $script:Parsed     = $p
        $script:ParsedFrom = $txtMailbox.Text + "|" + $txtUser.Text + "|" + $accessType + "|" + $mode

        Update-ParsedPreview -Parsed $p
        $txtOut.Text = "Parsed OK. Review the preview above, then click Run to proceed."
    } catch {
        $txtOut.Text = "Parse error: $($_.Exception.Message)"
    } finally {
        $btnParse.Enabled = $true
        $btnRun.Enabled   = $true
    }
})

$btnRun.Add_Click({
    $btnRun.Enabled   = $false
    $btnParse.Enabled = $false
    try {
        $accessType = Get-SelectedAccessType

        # THE STATE IS RE-READ HERE, not trusted from when the window opened.
        [void](Update-ConnectionGating)
        if (-not (Test-AccessTypeAvailable -AccessType $accessType)) {
            $txtOut.Text = "'$accessType' needs a service this process is not connected to. The options have been re-checked -- see the panel above."
            return
        }
        $mode       = Get-SelectedMode
        $stateKey   = $txtMailbox.Text + "|" + $txtUser.Text + "|" + $accessType + "|" + $mode

        if (-not $script:Parsed -or $stateKey -ne $script:ParsedFrom) {
            $script:Parsed     = Get-ParsedInputs -MailboxRaw $txtMailbox.Text -UserRaw $txtUser.Text -AccessType $accessType -Mode $mode
            $script:ParsedFrom = $stateKey
            Update-ParsedPreview -Parsed $script:Parsed
        }

        $p = $script:Parsed

        if ($p.Mailboxes.Count -lt 1) {
            $txtOut.Text = "Please enter at least one mailbox."
            return
        }
        if ($p.Users.Count -lt 1) {
            $txtOut.Text = "Please enter at least one user/manager. (If entries were removed due to mailbox overlap, re-check your entries.)"
            return
        }

        $modeText   = if ($mode -eq 'Remove') { 'Remove' } else { 'Grant' }
        $accessText = Get-AccessTypeText -AccessType $accessType
        $mbList     = ($p.Mailboxes | ForEach-Object { "  $_" }) -join "`r`n"
        $usrList    = ($p.Users     | ForEach-Object { "  $_" }) -join "`r`n"

        $confirmMsg = ("Mode: {0}`r`nAccess: {1}`r`n`r`nMailbox(es) ({2}):`r`n{3}`r`n`r`nUser(s)/Manager(s) ({4}):`r`n{5}`r`n`r`nTotal operations: {6}`r`n`r`nProceed?" -f `
            $modeText, $accessText, $p.Mailboxes.Count, $mbList, $p.Users.Count, $usrList, $p.OpCount)

        $confirm = Show-ToolDialog(
            $confirmMsg,
            "Confirm Action",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $txtOut.Text = "Cancelled."
            return
        }

        $txtOut.Text = "Running -- please wait..."
        $form.Refresh()

        $result      = Invoke-MailboxOneDriveAccess -Mailboxes $p.Mailboxes -Users $p.Users -AccessType $p.AccessType -Mode $p.Mode
        $txtOut.Text = $result

    } catch {
        $txtOut.Text = "ERROR: $($_.Exception.Message)"
    } finally {
        $btnRun.Enabled   = $true
        $btnParse.Enabled = $true
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
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

# Theme if Common.ps1 happens to be loaded; these tools run standalone
# by design, so this must not become a hard dependency.
if (Get-Command Set-ConsoleTheme -ErrorAction SilentlyContinue) {
    Set-ConsoleTheme -Form $form
}
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
$btnRecheck.Add_Click({ [void](Update-ConnectionGating) })

# LAST, AFTER EVERY CONTROL EXISTS. It reads and writes the radios, the two
# buttons and the banner, so running it earlier would be a method call on a
# control that has not been created yet -- the failure that stopped the Person
# Editor launching on 2026-09-20.
[void](Update-ConnectionGating)

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mailbox / Onedrive Perms' | Out-Null

[void]$form.ShowDialog()

#endregion
