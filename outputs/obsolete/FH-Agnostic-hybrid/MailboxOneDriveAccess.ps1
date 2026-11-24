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

#region Config  ---------------------------------------------------------------

<#
    Nothing here is hardcoded to an organisation.

    The mail domain and the OneDrive host are both derived at run time. The
    OneDrive host is the SharePoint tenant name with '-my' appended -- read from
    the connected SPO session rather than assumed, because getting it wrong
    silently produces URLs for a tenant that is not yours.

    Logs go beside this script, not to an absolute path, so the toolkit runs
    from wherever it is dropped.
#>
$Domain              = ''    # resolved by Get-MailSuffix below
$MySiteBase          = ''    # resolved by Get-MySiteBase below
# Get-LogPath, not a second copy of the path: this script had its own private
# definition, and two definitions of one location is how they end up disagreeing.
$CompletedTaskFolder = Get-LogPath
$CompletedTaskFile   = Join-Path $CompletedTaskFolder 'CompletedTasks.csv'

function Get-MailSuffix {
    <#
        '@domain' to append to a bare alias, or '' if no domain can be worked
        out. Empty means a bare alias is left alone rather than turned into a
        broken address ending in '@'.
    #>
    if ($script:ResolvedMailSuffix) { return $script:ResolvedMailSuffix }
    $d = Get-MailDomain
    $script:ResolvedMailSuffix = if ($d) { "@$d" } else { '' }
    return $script:ResolvedMailSuffix
}

function Get-MySiteBase {
    <#
        The personal-site root for this tenant, e.g.
        the SharePoint host with '-my' inserted, plus /personal

        Read from Get-SPOSite rather than built from a name: the root SharePoint
        URL is authoritative, and deriving the '-my' host from it is the same
        transformation Microsoft applies. Throws rather than guessing, because a
        guessed host produces URLs that look right and point nowhere.
    #>
    if ($script:ResolvedMySiteBase) { return $script:ResolvedMySiteBase }

    if (-not (Get-Command -Name 'Get-SPOSite' -ErrorAction SilentlyContinue)) {
        throw 'SharePoint Online is not connected, so the OneDrive host cannot be determined. Run Connect-SPOService first.'
    }
    $root = $null
    try {
        $root = @(Get-SPOSite -Limit 1 -ErrorAction Stop)[0].Url
    } catch {
        throw "Could not read the SharePoint tenant URL: $($_.Exception.Message)"
    }
    if ("$root" -notmatch '^(https://[^./]+)\.sharepoint\.com') {
        throw "Unexpected SharePoint URL shape '$root' -- cannot derive the OneDrive host from it."
    }
    $script:ResolvedMySiteBase = "$($Matches[1])-my.sharepoint.com/personal"
    return $script:ResolvedMySiteBase
}

#endregion

#region Helpers  --------------------------------------------------------------

function Get-AddressList {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }
    $tokens = $Raw -split "[\r\n,;]+"
    $clean = foreach ($t in $tokens) {
        $name = $t.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($name -notmatch '@') {
            $suffix = Get-MailSuffix
            if ($suffix) { $name = $name + $suffix }
        }
        $name
    }
    return @($clean | Sort-Object -Unique)
}

function Get-OneDriveUrl {
    <#
        A personal site URL is the whole UPN with every '.' and '@' replaced by
        '_'. Built from the full address rather than from the local part plus a
        fixed domain suffix, so a user in any accepted domain resolves correctly
        instead of only those in the primary one.
    #>
    param([string]$Email)
    $urlSafe = ($Email -replace '[.@]', '_')
    return "$(Get-MySiteBase)/$urlSafe"
}

function Write-CompletedLog {
    param([string]$Task)
    if (-not (Test-Path -LiteralPath $CompletedTaskFolder)) {
        New-Item -Path $CompletedTaskFolder -ItemType Directory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $CompletedTaskFile)) {
        "Task,DateTime" | Out-File -FilePath $CompletedTaskFile -Encoding utf8
    }
    $dt = Get-Date -Format "yyyy-MM-dd HHmm"
    if ($Task -and $Task.Trim() -ne "") {
        Add-Content -Path $CompletedTaskFile -Value ('"{0}","{1}"' -f $Task.Trim(), $dt)
    }
}

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

    if ($flags.Full -or $flags.SendAs) {
        $notFound = @()
        foreach ($m in $Mailboxes) {
            try   { Get-Mailbox -Identity $m -ErrorAction Stop | Out-Null }
            catch { $notFound += $m }
        }
        foreach ($u in $Users) {
            try   { Get-Mailbox -Identity $u -ErrorAction Stop | Out-Null }
            catch { $notFound += $u }
        }
        if ($notFound.Count -gt 0) {
            $list = ($notFound | Sort-Object -Unique) -join "`r`n    "
            [void]$log.AppendLine("PRE-FLIGHT FAILED -- the following could not be resolved:")
            [void]$log.AppendLine("    $list")
            [void]$log.AppendLine("")
            [void]$log.AppendLine("No changes were made. Correct the above entries and try again.")
            return $log.ToString()
        }
        [void]$log.AppendLine("Pre-flight : all identities resolved OK.")
        [void]$log.AppendLine("")
    }

    $succeeded = [System.Collections.Generic.List[PSCustomObject]]::new()
    $failed    = [System.Collections.Generic.List[string]]::new()
    $skipped   = [System.Collections.Generic.List[string]]::new()
    $oneDriveOkMailboxes = New-Object System.Collections.Generic.List[string]

    foreach ($m in $Mailboxes) {

        if ($flags.SendAs -and $isGrant) {
            try {
                Set-Mailbox -Identity $m -MessageCopyForSentAs $true -ErrorAction Stop
                [void]$log.AppendLine("  [OK]   MessageCopyForSentAs = true on $m")
            } catch {
                [void]$log.AppendLine("  [WARN] Could not set MessageCopyForSentAs on $m : $($_.Exception.Message)")
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
                        Add-MailboxPermission -Identity $m -User $u -AccessRights FullAccess -AutoMapping:$true -ErrorAction Stop | Out-Null
                        [void]$log.AppendLine("  [OK]   Granted Full Access : $u -> $m")
                        $succeeded.Add([PSCustomObject]@{ Mailbox = $m; User = $u; Op = 'Full' })
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Full Access : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("Full: $u -> $m")
                    }
                } else {
                    try {
                        $fa = Get-MailboxPermission -Identity $m -User $u -ErrorAction SilentlyContinue |
                              Where-Object { $_.AccessRights -contains 'FullAccess' }
                        if ($fa) {
                            Remove-MailboxPermission -Identity $m -User $u -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                            [void]$log.AppendLine("  [OK]   Removed Full Access : $u -> $m")
                            $succeeded.Add([PSCustomObject]@{ Mailbox = $m; User = $u; Op = 'Full' })
                        } else {
                            [void]$log.AppendLine("  [SKIP] No Full Access found : $u -> $m")
                            $skipped.Add("Full: $u -> $m")
                        }
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Remove Full Access : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("RemoveFull: $u -> $m")
                    }
                }
            }

            if ($flags.SendAs) {
                if ($isGrant) {
                    try {
                        Add-RecipientPermission -Identity $m -Trustee $u -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                        [void]$log.AppendLine("  [OK]   Granted Send-As : $u -> $m")
                        $succeeded.Add([PSCustomObject]@{ Mailbox = $m; User = $u; Op = 'SendAs' })
                    } catch {
                        [void]$log.AppendLine("  [FAIL] Send-As : $u -> $m : $($_.Exception.Message)")
                        $failed.Add("SendAs: $u -> $m")
                    }
                } else {
                    try {
                        $sa = Get-RecipientPermission -Identity $m -Trustee $u -ErrorAction SilentlyContinue |
                              Where-Object { $_.AccessRights -contains 'SendAs' }
                        if ($sa) {
                            Remove-RecipientPermission -Identity $m -Trustee $u -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            [void]$log.AppendLine("  [OK]   Removed Send-As : $u -> $m")
                            $succeeded.Add([PSCustomObject]@{ Mailbox = $m; User = $u; Op = 'SendAs' })
                        } else {
                            [void]$log.AppendLine("  [SKIP] No Send-As found : $u -> $m")
                            $skipped.Add("SendAs: $u -> $m")
                        }
                    } catch {
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
                $faList = Get-MailboxPermission -Identity $m -ErrorAction Stop |
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
                $saList = Get-RecipientPermission -Identity $m -ErrorAction Stop |
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
$form.Size            = New-Object System.Drawing.Size(780, 890)
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
$lblMailbox.Text     = "Mailbox(es) -- one per line (enter email address,  optional):"
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
$lblUser.Text     = "User(s)/Manager(s) to grant/remove access -- one per line (enter email address,  optional):"
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

        $confirm = [System.Windows.Forms.MessageBox]::Show(
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

try {
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $y  = $wa.Bottom - $form.Height
    if ($y -lt $wa.Top) { $y = $wa.Top }
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $form.Location = New-Object System.Drawing.Point($wa.Left, $y)
} catch {
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
}

# Theme if Common.ps1 happens to be loaded; these tools run standalone
# by design, so this must not become a hard dependency.
if (Get-Command Set-ConsoleTheme -ErrorAction SilentlyContinue) {
    Set-ConsoleTheme -Form $form
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mailbox / Onedrive Perms' | Out-Null
}
[void]$form.ShowDialog()

#endregion
