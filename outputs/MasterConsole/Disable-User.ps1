#Requires -Version 5.1
<#
    Disable-User.ps1
    Termination / disable processing for HQ and Franchise, in one tool.

    Replaces Disable-Intake.ps1, Disable-New.ps1 and Disable-New-HQ.ps1.

    The reasoning -- why one script, why per-user rather than per-step, why
    check-then-act, why notifications are the one exception -- is in
    ScriptNotes/Disable-User.ps1.notes.md and ARCHITECTURE.md 21.

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# The termination pipeline. Runs the full set of steps for someone leaving,
# in order, with a preview first.
#
# Typically: hides them from the address lists, converts the mailbox to a
# shared mailbox, grants the manager or replacement access to it, sets up
# forwarding, removes them from distribution lists, and clears inbox rules
# that would redirect mail away.
#
#   * It shows you every step it intends to take before taking any of them.
#   * Each step is logged, and a step that fails does not stop the rest.
#   * The audit log is the record of what was done, and is what you go back
#     to if something needs reversing.
#
# THE RETRY QUEUE (second tab)
#
# Anything that does not finish cleanly lands in the retry queue rather than
# being lost or silently skipped. It is a CSV -- disable-retry.csv, in the
# request folder -- so it survives closing the tool.
#
# Each row records which step failed, which steps had ALREADY succeeded, the
# error text, how many attempts have been made, and when it first failed.
#
#   * Retrying is safe. Steps already completed are skipped, so a retry
#     resumes rather than starting over.
#   * A row disappears from the queue as soon as it fully succeeds.
#   * Most failures are timing, not faults. The two usual ones are the person
#     having no mailbox yet (onboarding never enabled it) and directory sync
#     not having caught up -- both fix themselves, and the queue tells you
#     which one you are looking at.
#
# A row that has been sitting there through several attempts is telling you
# something the tool cannot fix on its own -- usually a bad identity in the
# ticket.
#
# To reverse a termination, use Returning User -- do not undo these steps by
# hand.
#
# This is a destructive operation on a live account. Read the preview.
# TOOLKIT-HELP-END

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------

$script:LicenseGroupName  = 'E5 License Fran - Full'
$script:OneDriveBaseUrl   = 'https://expresspros-my.sharepoint.com/personal'
$script:SmtpServer        = 'smtp.expresspros.com'
$script:DoNotReplyAddress = 'DO_NOT_REPLY@ExpressPros.com'

$script:NotificationCopyTo = Get-OperatorEmail

$script:RequestFolder = Get-DisableRequestPath
$script:RetryCsvPath  = Join-Path $script:RequestFolder 'disable-retry.csv'
$script:ImportCsvPath = Join-Path $script:RequestFolder 'disablelist.csv'

$script:StepPlan = @{
    Franchise = @(
        'ConvertShared', 'HideFromGAL', 'GrantDelegate', 'RemoveLicence',
        'GrantOneDrive', 'NotifyOneDrive', 'DisableForwarding', 'ValidateLicence'
    )
    HQ = @(
        'ConvertShared', 'HideFromGAL', 'GrantDelegate',
        'GrantOneDrive', 'NotifyOneDrive', 'DisableForwarding'
    )
}

$script:StateTrackedSteps = @('NotifyOneDrive')

$script:StepLabels = @{
    ConvertShared     = 'Convert to shared mailbox'
    HideFromGAL       = 'Hide from address lists'
    GrantDelegate     = 'Grant delegate full access'
    RemoveLicence     = 'Remove licence group membership'
    GrantOneDrive     = 'Grant OneDrive site admin'
    NotifyOneDrive    = 'Notify new OneDrive owner'
    DisableForwarding = 'Disable forwarding inbox rules'
    ValidateLicence   = 'Verify licence removal'
}

#endregion

#region Input parsing --------------------------------------------------------

function ConvertTo-NormalValue {
    param([string]$Value)
    if (-not $Value) { return '' }
    $trimmed = $Value.Trim()
    $emptyTokens = @('NA', 'N/A', 'NONE', 'NULL', '-', 'NO')
    if ($emptyTokens -contains $trimmed.ToUpper()) { return '' }
    if ($trimmed -match '@') { return $trimmed.Split('@')[0] }
    return $trimmed
}

function ConvertTo-NormalLine {
    param([string]$Text)
    if (-not $Text) { return '' }
    $clean = $Text
    $clean = $clean.Replace([char]0xA0, ' ')
    $clean = $clean.Replace([char]0x2007, ' ')
    $clean = $clean.Replace([char]0x202F, ' ')
    $clean = $clean.Replace([string][char]0x00AD, '')
    return $clean.Trim()
}

function ConvertTo-NormalManager {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $v = $Value.Trim()
    if ($v -match '@') { $v = $v.Split('@')[0] }
    $v = $v -replace '\s+', '.'
    if ($v -notmatch '^[A-Za-z0-9._\-]+$') { return $null }
    return $v
}

function Test-LogonName {
    param([string]$Value)
    return ($Value -match '^[A-Za-z0-9]+$')
}

function Test-DisableIdentity {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $v = $Value.Trim()
    if ($v -match '^[A-Za-z0-9]+$') { return $true }
    if ($v -match '^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$') { return $true }
    return $false
}

function Show-ParseError {
    param([string]$Message)
    [void][System.Windows.Forms.MessageBox]::Show($Message, 'Input Error', 'OK', 'Error')
}

function Get-NormalizedLines {
    param([string]$Text)
    $raw = $Text -split "\r?\n"
    return @($raw | ForEach-Object { ConvertTo-NormalLine $_ } | Where-Object { $_ -ne '' })
}

function New-DisableRecord {
    param([string]$Type)
    return [pscustomobject]@{
        Type        = $Type
        MB          = ''
        Delegate    = ''
        OD          = ''
        Employee    = ''
        LWD         = ''
        Attempts    = 0
        FirstFailed = ''
        StepsDone   = @()
    }
}

function ConvertFrom-HQText {
    param([string]$Text)
    $lines = Get-NormalizedLines -Text $Text
    if ($lines.Count -eq 0) { return [pscustomobject]@{ Ok = $true; Records = @() } }

    $records = @()
    $cur     = $null
    foreach ($line in $lines) {
        if ($line -match '^\s*([^:]+?)\s*:\s*(.*)$') {
            $label = $Matches[1].Trim().ToLower()
            $val   = $Matches[2].Trim()
            switch ($label) {
                'employee name' {
                    if ($cur) { $records += $cur }
                    $cur = New-DisableRecord 'HQ'
                    $cur.Employee = $val
                }
                'upn' {
                    if (-not $cur -or $cur.MB -ne '') {
                        if ($cur) { $records += $cur }
                        $cur = New-DisableRecord 'HQ'
                    }
                    $cur.MB = $val
                }
                'last working day' { if ($cur) { $cur.LWD = $val } }
                default { }
            }
        }
    }
    if ($cur) { $records += $cur }

    if ($records.Count -eq 0) {
        Show-ParseError "The HQ box has text but no 'Employee Name:' or 'UPN:' entries were found."
        return [pscustomobject]@{ Ok = $false; Records = @() }
    }
    foreach ($r in $records) {
        if ([string]::IsNullOrWhiteSpace($r.MB)) {
            $who = if ($r.Employee) { $r.Employee } else { '(unnamed HQ block)' }
            Show-ParseError "An HQ block is missing its UPN:`r`n`r`n$who"
            return [pscustomobject]@{ Ok = $false; Records = @() }
        }
        if (-not (Test-DisableIdentity $r.MB)) {
            Show-ParseError "Invalid HQ UPN / identity:`r`n`r`n$($r.MB)"
            return [pscustomobject]@{ Ok = $false; Records = @() }
        }
    }
    return [pscustomobject]@{ Ok = $true; Records = $records }
}

function ConvertFrom-FranchiseText {
    param([string]$Text)
    $lines = Get-NormalizedLines -Text $Text
    if ($lines.Count -eq 0) { return [pscustomobject]@{ Ok = $true; Records = @() } }

    $records = @()
    $cur     = $null
    foreach ($line in $lines) {
        if ($line -match '^\s*([^:]+?)\s*:\s*(.*)$') {
            $label = $Matches[1].Trim().ToLower()
            $val   = $Matches[2].Trim()
            switch ($label) {
                'logon name' {
                    if ($cur) { $records += $cur }
                    $cur = New-DisableRecord 'Franchise'
                    $cur.MB = $val
                }
                'delegate' { if ($cur) { $cur.Delegate = ConvertTo-NormalValue $val } }
                { $_ -in @('onedrive', 'one drive') } { if ($cur) { $cur.OD = ConvertTo-NormalValue $val } }
                default { }
            }
        }
    }
    if ($cur) { $records += $cur }

    if ($records.Count -eq 0) {
        Show-ParseError "The Franchise box has text but no 'Logon Name:' entries were found."
        return [pscustomobject]@{ Ok = $false; Records = @() }
    }

    $last = $records[$records.Count - 1]
    if ($records.Count -gt 1) {
        foreach ($r in $records) {
            if (-not $r.Delegate -and $last.Delegate) { $r.Delegate = $last.Delegate }
            if (-not $r.OD       -and $last.OD)       { $r.OD       = $last.OD }
        }
    }
    foreach ($r in $records) {
        if (-not (Test-LogonName $r.MB)) {
            Show-ParseError "Franchise Logon Name must be alphanumeric (a sAMAccountName, not an email):`r`n`r`n$($r.MB)"
            return [pscustomobject]@{ Ok = $false; Records = @() }
        }
        if ($r.Delegate -ne '' -and $r.Delegate -notmatch '^[A-Za-z0-9._\-]+$') {
            Show-ParseError "A delegate value does not look valid:`r`n`r`n$($r.Delegate)"
            return [pscustomobject]@{ Ok = $false; Records = @() }
        }
    }
    return [pscustomobject]@{ Ok = $true; Records = $records }
}

function Import-DisableListCsv {
    param([string]$Path)
    $records = @()
    try { $rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop) }
    catch {
        Show-ParseError "Could not read $Path`r`n`r`n$($_.Exception.Message)"
        return @()
    }
    foreach ($row in $rows) {
        $mb = "$($row.MB)".Trim()
        if (-not $mb) { continue }
        # No Type column means the old three-column format, which was Franchise.
        $type = "$($row.Type)".Trim()
        if ($type -ne 'HQ') { $type = 'Franchise' }
        $r = New-DisableRecord $type
        $r.MB       = $mb
        $r.Delegate = ConvertTo-NormalValue "$($row.Delegate)"
        $r.OD       = ConvertTo-NormalValue "$($row.OD)"
        $records += $r
    }
    return $records
}

#endregion

#region Retry queue ----------------------------------------------------------

$script:RetryColumns = @(
    'Type', 'MB', 'Delegate', 'OD', 'Employee',
    'FirstFailed', 'LastAttempt', 'Attempts', 'FailedStep', 'StepsDone', 'LastError'
)

function Import-RetryQueue {
    if (-not (Test-Path -LiteralPath $script:RetryCsvPath)) { return @() }
    try { $rows = @(Import-Csv -LiteralPath $script:RetryCsvPath -ErrorAction Stop) }
    catch {
        Write-Warning "Could not read the retry queue: $($_.Exception.Message)"
        return @()
    }
    # ArrayList, not List[object]: a generic List returned through @() throws
    # "Argument types do not match" in PS 5.1 (ARCHITECTURE 14a). Mutates in
    # place, which is required -- '+=' inside a scriptblock only updates a local.
    $out = New-Object System.Collections.ArrayList
    foreach ($row in $rows) {
        if (-not "$($row.MB)".Trim()) { continue }
        $r = New-DisableRecord "$($row.Type)"
        $r.MB          = "$($row.MB)".Trim()
        $r.Delegate    = "$($row.Delegate)".Trim()
        $r.OD          = "$($row.OD)".Trim()
        $r.Employee    = "$($row.Employee)".Trim()
        $r.FirstFailed = "$($row.FirstFailed)".Trim()
        $n = 0; [void][int]::TryParse("$($row.Attempts)", [ref]$n)
        $r.Attempts    = $n
        $r.StepsDone   = @("$($row.StepsDone)".Split(';') | Where-Object { $_ })
        Add-Member -InputObject $r -NotePropertyName 'FailedStep'  -NotePropertyValue "$($row.FailedStep)"  -Force
        Add-Member -InputObject $r -NotePropertyName 'LastError'   -NotePropertyValue "$($row.LastError)"   -Force
        Add-Member -InputObject $r -NotePropertyName 'LastAttempt' -NotePropertyValue "$($row.LastAttempt)" -Force
        [void]$out.Add($r)
    }
    return @($out)
}

function Export-RetryQueue {
    param($Records)
    $dir = Split-Path -Parent $script:RetryCsvPath
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }

    if (@($Records).Count -eq 0) {
        ($script:RetryColumns -join ',') | Set-Content -LiteralPath $script:RetryCsvPath -Encoding UTF8
        return
    }
    $rows = foreach ($r in $Records) {
        [PSCustomObject]@{
            Type        = $r.Type
            MB          = $r.MB
            Delegate    = $r.Delegate
            OD          = $r.OD
            Employee    = $r.Employee
            FirstFailed = $r.FirstFailed
            LastAttempt = $r.LastAttempt
            Attempts    = $r.Attempts
            FailedStep  = $r.FailedStep
            StepsDone   = (@($r.StepsDone) -join ';')
            LastError   = $r.LastError
        }
    }
    $rows | Export-Csv -LiteralPath $script:RetryCsvPath -NoTypeInformation -Encoding UTF8
}

function Update-RetryQueue {
    param(
        [Parameter(Mandatory)]$Record,
        [Parameter(Mandatory)][ValidateSet('Success', 'Partial', 'Failed')][string]$Outcome,
        [string]$FailedStep,
        [string]$LastError,
        [string[]]$StepsDone
    )
    $queue = New-Object System.Collections.ArrayList
    foreach ($q in (Import-RetryQueue)) {
        if ($q.MB -eq $Record.MB -and $q.Type -eq $Record.Type) { continue }
        [void]$queue.Add($q)
    }

    if ($Outcome -ne 'Success') {
        $now = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        $r = New-DisableRecord $Record.Type
        $r.MB          = $Record.MB
        $r.Delegate    = $Record.Delegate
        $r.OD          = $Record.OD
        $r.Employee    = $Record.Employee
        $r.Attempts    = [int]$Record.Attempts + 1
        $r.FirstFailed = if ($Record.FirstFailed) { $Record.FirstFailed } else { $now }
        $r.StepsDone   = @($StepsDone)
        Add-Member -InputObject $r -NotePropertyName 'FailedStep'  -NotePropertyValue $FailedStep  -Force
        Add-Member -InputObject $r -NotePropertyName 'LastError'   -NotePropertyValue $LastError   -Force
        Add-Member -InputObject $r -NotePropertyName 'LastAttempt' -NotePropertyValue $now         -Force
        [void]$queue.Add($r)
    }
    Export-RetryQueue -Records @($queue)
}

#endregion

#region Target resolution ----------------------------------------------------

function Get-DisableADServer {
    param([string]$Type)
    if ($Type -eq 'HQ') { return $script:HQDomain }
    return $script:FranchiseAutoDetectDC
}

function Resolve-DisableTarget {
    param($Record)

    $server = Get-DisableADServer -Type $Record.Type
    $out = [PSCustomObject]@{
        Ok       = $false
        Smtp     = ''
        AdUser   = $null
        Reason   = ''
        Fix      = ''
    }

    # --- AD account -------------------------------------------------------
    $ad = $null
    try {
        if ($Record.MB -match '@') {
            $upn = ConvertTo-LdapSafe $Record.MB
            $ad = Get-ADUser -LDAPFilter "(userPrincipalName=$upn)" -Server $server -Properties MemberOf, mail -ErrorAction Stop |
                  Select-Object -First 1
        } else {
            $ad = Get-ADUser -Identity $Record.MB -Server $server -Properties MemberOf, mail -ErrorAction Stop
        }
    } catch { $ad = $null }

    if (-not $ad) {
        $out.Reason = "No AD account found for '$($Record.MB)' in $server."
        $out.Fix    = 'Most likely a typo in the ticket. Verify the identity and correct it, or remove the row.'
        return $out
    }
    $out.AdUser = $ad

    # --- on-premises remote mailbox --------------------------------------
    $remote = $null
    try { $remote = Get-RemoteMailbox -Identity $ad.SamAccountName -ErrorAction Stop } catch { $remote = $null }

    if (-not $remote) {
        $out.Reason = "The AD account '$($ad.SamAccountName)' exists but has no remote mailbox."
        $out.Fix    = 'Onboarding almost certainly never raised the ticket to enable it. Enable the remote mailbox, wait for directory sync, then retry this row.'
        return $out
    }

    # --- cloud mailbox ----------------------------------------------------
    $cloud = $null
    try { $cloud = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $ad.SamAccountName; ErrorAction = 'Stop' } }
    catch { $cloud = $null }

    if (-not $cloud) {
        $out.Reason = "'$($ad.SamAccountName)' has a remote mailbox on-premises but no Exchange Online mailbox yet."
        $out.Fix    = 'Directory sync has not caught up. Wait for the next sync cycle and retry.'
        return $out
    }

    $out.Smtp = "$($cloud.PrimarySmtpAddress)"
    $out.Ok   = $true
    return $out
}

function ConvertTo-LdapSafe {
    # Escape the five characters an LDAP filter treats as syntax.
    param([string]$Value)
    $v = "$Value"
    $v = $v -replace '\\', '\5c'
    $v = $v -replace '\(', '\28'
    $v = $v -replace '\)', '\29'
    $v = $v -replace '\*', '\2a'
    $v = $v -replace "`0", '\00'
    return $v
}

#endregion

#region Step implementations -------------------------------------------------

function New-StepResult {
    param(
        [ValidateSet('Done', 'Skipped', 'Failed')][string]$Status,
        [string]$Message
    )
    return [PSCustomObject]@{ Status = $Status; Message = $Message }
}

function Get-OneDriveUrl {
    param([string]$Smtp)
    return "$($script:OneDriveBaseUrl)/$($Smtp.Replace('@','_').Replace('.','_'))"
}

function Get-OneDriveNotificationBody {
    param([string]$Type, [string]$Smtp, [string]$Url)

    if ($Type -eq 'HQ') {
        return @"
<html>
<body>
<p>To access the user's OneDrive for Business, you can use the following link: <a href='$Url'>OneDrive for Business</a></p>
</body>
</html>
"@
    }

    return @"
<html>
<body>
<p>Greetings,</p>
<p>The user account $Smtp has been disabled in the Active Directory. However, please note that their OneDrive for Business data will be preserved for 30 days. During this period, you are the temporary owner of all the documents saved to their OneDrive for Business.</p>
<p>If you wish to save any important documents beyond the 30-day retention period, you have the option to copy them to another location. Additionally, you can contact the Asssistance Center to reassign ownership to another OneDrive for Business owner. Please be aware that after the 30-day period, the user's OneDrive for Business data will be permanently deleted.</p>
<p>To access the user's OneDrive for Business, you can use the following link: <a href='$Url'>OneDrive for Business</a></p>
<p>If you have any questions or need further assistance, please contact the Assistance Center.</p>
</body>
</html>
"@
}

function Invoke-DisableStep {
    param([string]$Step, $Ctx)

    switch ($Step) {

        'ConvertShared' {
            $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Ctx.Sam; ErrorAction = 'Stop' }
            if ("$($mb.RecipientTypeDetails)" -eq 'SharedMailbox') {
                return New-StepResult 'Skipped' 'Already a shared mailbox.'
            }
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{ Identity = $Ctx.Sam; Type = 'Shared'; ErrorAction = 'Stop' }
            Set-RemoteMailbox -Identity $Ctx.Sam -Type Shared -ErrorAction Stop
            return New-StepResult 'Done' 'Converted to shared.'
        }

        'HideFromGAL' {
            $rm = Get-RemoteMailbox -Identity $Ctx.Sam -ErrorAction Stop
            if ($rm.HiddenFromAddressListsEnabled) {
                return New-StepResult 'Skipped' 'Already hidden from address lists.'
            }
            Set-RemoteMailbox -Identity $Ctx.Sam -HiddenFromAddressListsEnabled $true -ErrorAction Stop
            return New-StepResult 'Done' 'Hidden from address lists.'
        }

        'GrantDelegate' {
            if (-not $Ctx.DelegateSmtp) { return New-StepResult 'Skipped' 'No delegate specified.' }
            $existing = @()
            try {
                $existing = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{
                    Identity = $Ctx.Sam; User = $Ctx.DelegateSmtp; ErrorAction = 'SilentlyContinue' })
            } catch { $existing = @() }
            if ($existing | Where-Object { "$($_.AccessRights)" -match 'FullAccess' -and -not $_.Deny }) {
                return New-StepResult 'Skipped' "$($Ctx.DelegateSmtp) already has FullAccess."
            }
            Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
                Identity = $Ctx.Sam; User = $Ctx.DelegateSmtp; AccessRights = 'FullAccess'; ErrorAction = 'Stop' } | Out-Null
            return New-StepResult 'Done' "FullAccess granted to $($Ctx.DelegateSmtp)."
        }

        'RemoveLicence' {
            $dnPart = "CN=$($script:LicenseGroupName)"
            $groups = @($Ctx.AdUser.MemberOf)
            if (-not ($groups | Where-Object { $_ -like "$dnPart,*" })) {
                return New-StepResult 'Skipped' "Not a member of '$($script:LicenseGroupName)'."
            }
            $grp = Get-ADGroup -Identity $script:LicenseGroupName -Server $Ctx.AdServer -ErrorAction Stop
            Remove-ADGroupMember -Identity $grp -Members $Ctx.Sam -Server $Ctx.AdServer -Confirm:$false -ErrorAction Stop
            return New-StepResult 'Done' "Removed from '$($script:LicenseGroupName)'."
        }

        'GrantOneDrive' {
            if (-not $Ctx.OdSmtp) { return New-StepResult 'Skipped' 'No OneDrive owner specified.' }
            $url = Get-OneDriveUrl -Smtp $Ctx.Smtp
            $already = $false
            try {
                $u = Get-SPOUser -Site $url -LoginName $Ctx.OdSmtp -ErrorAction Stop
                if ($u -and $u.IsSiteCollectionAdmin) { $already = $true }
            } catch { $already = $false }
            if ($already) { return New-StepResult 'Skipped' "$($Ctx.OdSmtp) is already a site collection admin." }
            Set-SPOUser -Site $url -LoginName $Ctx.OdSmtp -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
            return New-StepResult 'Done' "$($Ctx.OdSmtp) granted site collection admin."
        }

        'NotifyOneDrive' {
            if (-not $Ctx.OdSmtp) { return New-StepResult 'Skipped' 'No OneDrive owner to notify.' }
            $url  = Get-OneDriveUrl -Smtp $Ctx.Smtp
            $body = Get-OneDriveNotificationBody -Type $Ctx.Type -Smtp $Ctx.Smtp -Url $url

            # The DEPARTING user's address: it identifies which OneDrive.
            $subject     = "OneDrive Ownership Change Notification - $($Ctx.Smtp)"
            $copySuffix  = if ($Ctx.Type -eq 'HQ') { 'HQ' } else { 'Fran' }
            $copySubject = "$subject - $copySuffix"

            Send-MailMessage -To $Ctx.OdSmtp -Subject $subject -BodyAsHtml $body `
                -From $script:DoNotReplyAddress -SmtpServer $script:SmtpServer -ErrorAction Stop

            # No configured operator address means NO copy -- never fall back to a
            # default, or somebody who did not do the work gets the mail.
            if ($script:NotificationCopyTo) {
                Send-MailMessage -To $script:NotificationCopyTo -Subject $copySubject -BodyAsHtml $body `
                    -From $script:DoNotReplyAddress -SmtpServer $script:SmtpServer -ErrorAction Stop
                return New-StepResult 'Done' "Notified $($Ctx.OdSmtp); copy to $($script:NotificationCopyTo) [$copySuffix]."
            }
            return New-StepResult 'Done' "Notified $($Ctx.OdSmtp). No copy sent -- set 'Your email' in ToolManager > Settings."
        }

        'DisableForwarding' {
            $rules = @()
            try {
                $rules = @(Invoke-EXOCommand -Name 'Get-InboxRule' -Parameters @{
                    Mailbox = $Ctx.Smtp; IncludeHidden = $true; ErrorAction = 'Stop' })
            } catch {
                return New-StepResult 'Failed' "Could not read inbox rules: $($_.Exception.Message)"
            }
            $fwd = @($rules | Where-Object {
                ($_.ForwardTo -or $_.RedirectTo -or $_.ForwardAsAttachmentTo) -and $_.Enabled })
            if ($fwd.Count -eq 0) {
                return New-StepResult 'Skipped' 'No enabled forwarding rules.'
            }
            $names = @()
            foreach ($rule in $fwd) {
                # -Identity ONLY: Exchange Online rejects Identity and Mailbox
                # together, and the Identity from Get-InboxRule is already scoped.
                Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters @{
                    Identity = $rule.Identity; Confirm = $false; ErrorAction = 'Stop' }
                $names += "$($rule.Name)"
            }
            return New-StepResult 'Done' "Disabled $($fwd.Count) forwarding rule(s): $($names -join ', ')"
        }

        'ValidateLicence' {
            # Re-read from AD: the point of this step is to confirm the removal landed.
            $fresh = Get-ADUser -Identity $Ctx.Sam -Server $Ctx.AdServer -Properties MemberOf -ErrorAction Stop
            $dnPart = "CN=$($script:LicenseGroupName)"
            if (@($fresh.MemberOf) | Where-Object { $_ -like "$dnPart,*" }) {
                return New-StepResult 'Failed' "STILL a member of '$($script:LicenseGroupName)' after removal."
            }
            return New-StepResult 'Done' "Confirmed removed from '$($script:LicenseGroupName)'."
        }

        default { return New-StepResult 'Failed' "Unknown step '$Step'." }
    }
}

#endregion

#region Preflight ------------------------------------------------------------

function Test-DisablePrerequisite {
    $map = Get-CapabilityStateMap -Refresh

    $needed = @(
        [PSCustomObject]@{ Id = 'EMS'; Why = 'converting the mailbox, hiding it from the GAL' }
        [PSCustomObject]@{ Id = 'AD';  Why = 'finding the account and licence group membership' }
        [PSCustomObject]@{ Id = 'ExO'; Why = 'delegate access and forwarding rules' }
        [PSCustomObject]@{ Id = 'SPO'; Why = 'OneDrive ownership' }
    )

    $missing = @()
    foreach ($n in $needed) {
        $state = $map[$n.Id]
        if (-not $state -or -not $state.Available) {
            $detail = if ($state) { "$($state.Detail)" } else { 'not probed' }
            $missing += [PSCustomObject]@{ Id = $n.Id; Why = $n.Why; Detail = $detail }
        }
    }
    return [PSCustomObject]@{ Ok = ($missing.Count -eq 0); Missing = @($missing); Map = $map }
}

function Update-ConnectionLabel {
    param($Check)
    if ($Check.Ok) {
        $lblConn.Text      = 'Connected: on-prem Exchange, AD, Exchange Online, SharePoint Online.'
        $lblConn.ForeColor = [System.Drawing.Color]::FromArgb(0, 110, 0)
        return
    }
    $lblConn.Text      = "NOT connected: $((@($Check.Missing | ForEach-Object { $_.Id })) -join ', ')  --  click Check Connections for detail."
    $lblConn.ForeColor = [System.Drawing.Color]::Firebrick
}

#endregion

#region Execution engine -----------------------------------------------------

function Invoke-DisableUser {
    param($Record, [scriptblock]$Progress, [string]$RunId)

    # ArrayList, not List[object] -- see Import-RetryQueue and ARCHITECTURE 14a.
    $results = New-Object System.Collections.ArrayList
    $stepsDone = @($Record.StepsDone)

    $addResult = {
        param($Step, $Status, $Message)
        [void]$results.Add([PSCustomObject]@{
            User    = $Record.MB
            Type    = $Record.Type
            Step    = $Step
            Label   = $(if ($script:StepLabels.ContainsKey($Step)) { $script:StepLabels[$Step] } else { $Step })
            Status  = $Status
            Message = $Message
        })
    }

    if ($Progress) { & $Progress "Resolving $($Record.MB)..." }
    $target = Resolve-DisableTarget -Record $Record
    if (-not $target.Ok) {
        & $addResult 'Resolve' 'Failed' "$($target.Reason) $($target.Fix)"
        Write-CompletedLog -Action 'Disable' -Target $Record.MB -Result 'Failed' `
            -Environment $Record.Type -RunId $RunId `
            -Detail "Resolve -- $($target.Reason)"
        Update-RetryQueue -Record $Record -Outcome 'Failed' -FailedStep 'Resolve' `
            -LastError $target.Reason -StepsDone $stepsDone
        return [PSCustomObject]@{ Outcome = 'Failed'; Results = @($results); Smtp = '' }
    }

    $ctx = [PSCustomObject]@{
        Type         = $Record.Type
        Sam          = $target.AdUser.SamAccountName
        Smtp         = $target.Smtp
        AdUser       = $target.AdUser
        AdServer     = (Get-DisableADServer -Type $Record.Type)
        DelegateSmtp = $(if ($Record.Delegate) { "$($Record.Delegate)@$($script:MailDomain)" } else { '' })
        OdSmtp       = $(if ($Record.OD)       { "$($Record.OD)@$($script:MailDomain)" }       else { '' })
    }

    $plan = @($script:StepPlan[$Record.Type])
    $failedStep = ''
    $lastError  = ''

    foreach ($step in $plan) {

        if ($script:StateTrackedSteps -contains $step -and $stepsDone -contains $step) {
            & $addResult $step 'Skipped' 'Already sent on a previous attempt -- not resent.'
            continue
        }

        if ($Progress) { & $Progress "$($Record.MB): $($script:StepLabels[$step])" }

        try {
            $r = Invoke-DisableStep -Step $step -Ctx $ctx
        } catch {
            $r = New-StepResult 'Failed' $_.Exception.Message
        }

        & $addResult $step $r.Status $r.Message

        if ($r.Status -eq 'Failed') {
            if (-not $failedStep) { $failedStep = $step }
            $lastError = $r.Message
        } elseif ($stepsDone -notcontains $step) {
            $stepsDone += $step
        }
    }

    $failed = @($results | Where-Object { $_.Status -eq 'Failed' })
    $outcome = if ($failed.Count -eq 0) { 'Success' } else { 'Partial' }

    $done    = @($results | Where-Object { $_.Status -eq 'Done' }).Count
    $skipped = @($results | Where-Object { $_.Status -eq 'Skipped' }).Count
    $detail  = "$done applied, $skipped already done, $($failed.Count) failed"
    if ($failed.Count) { $detail += " -- failed: $((@($failed | ForEach-Object { $_.Step })) -join ', ')" }

    Write-CompletedLog -Action 'Disable' -Target $ctx.Smtp -Result $outcome `
        -Environment $Record.Type -RunId $RunId -Detail $detail

    Update-RetryQueue -Record $Record -Outcome $outcome -FailedStep $failedStep `
        -LastError $lastError -StepsDone $stepsDone

    return [PSCustomObject]@{ Outcome = $outcome; Results = @($results); Smtp = $ctx.Smtp }
}

#endregion

#region UI -------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Disable User'
$form.Size = New-Object System.Drawing.Size(840, 958)
$form.StartPosition = 'Manual'
$form.MinimumSize = New-Object System.Drawing.Size(840, 700)
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblConn = New-Object System.Windows.Forms.Label
$lblConn.Text = 'Checking connections...'
$lblConn.Location = New-Object System.Drawing.Point(12, 10)
$lblConn.Size = New-Object System.Drawing.Size(614, 18)
$lblConn.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblConn.AutoEllipsis = $true
$lblConn.Tag = 'NoTheme'
$form.Controls.Add($lblConn)

$btnCheckConn = New-Object System.Windows.Forms.Button
$btnCheckConn.Text = 'Check Connections'
$btnCheckConn.Location = New-Object System.Drawing.Point(634, 6)
$btnCheckConn.Size = New-Object System.Drawing.Size(130, 24)
$form.Controls.Add($btnCheckConn)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, 34)
$tabs.Size = New-Object System.Drawing.Size(800, 400)
$tabs.Anchor = 'Top,Left,Right'
$form.Controls.Add($tabs)

# --- Tab 1: new requests ----------------------------------------------------
$tabNew = New-Object System.Windows.Forms.TabPage
$tabNew.Text = 'New Requests'
$tabs.TabPages.Add($tabNew)

$lblHQ = New-Object System.Windows.Forms.Label
$lblHQ.Text = 'HQ request (paste SNOW description):'
$lblHQ.Location = New-Object System.Drawing.Point(10, 8)
$lblHQ.Size = New-Object System.Drawing.Size(320, 16)
$tabNew.Controls.Add($lblHQ)

$txtHQ = New-Object System.Windows.Forms.TextBox
$txtHQ.Multiline = $true; $txtHQ.ScrollBars = 'Vertical'
$txtHQ.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtHQ.Location = New-Object System.Drawing.Point(10, 26)
$txtHQ.Size = New-Object System.Drawing.Size(766, 92)
$tabNew.Controls.Add($txtHQ)

$lblHQMgr = New-Object System.Windows.Forms.Label
$lblHQMgr.Text = 'HQ Manager:'
$lblHQMgr.Location = New-Object System.Drawing.Point(10, 127)
$lblHQMgr.Size = New-Object System.Drawing.Size(78, 16)
$tabNew.Controls.Add($lblHQMgr)

$txtHQMgr = New-Object System.Windows.Forms.TextBox
$txtHQMgr.Location = New-Object System.Drawing.Point(90, 124)
$txtHQMgr.Size = New-Object System.Drawing.Size(200, 22)
$tabNew.Controls.Add($txtHQMgr)

$lblFran = New-Object System.Windows.Forms.Label
$lblFran.Text = 'Franchise requests (Logon Name / Delegate / OneDrive):'
$lblFran.Location = New-Object System.Drawing.Point(10, 154)
$lblFran.Size = New-Object System.Drawing.Size(420, 16)
$tabNew.Controls.Add($lblFran)

$txtFran = New-Object System.Windows.Forms.TextBox
$txtFran.Multiline = $true; $txtFran.ScrollBars = 'Vertical'
$txtFran.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtFran.Location = New-Object System.Drawing.Point(10, 172)
$txtFran.Size = New-Object System.Drawing.Size(766, 110)
$tabNew.Controls.Add($txtFran)

$btnParse = New-Object System.Windows.Forms.Button
$btnParse.Text = 'Process Requests'
$btnParse.Location = New-Object System.Drawing.Point(10, 290)
$btnParse.Size = New-Object System.Drawing.Size(130, 26)
$tabNew.Controls.Add($btnParse)

$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Text = 'Import CSV...'
$btnImport.Location = New-Object System.Drawing.Point(148, 290)
$btnImport.Size = New-Object System.Drawing.Size(110, 26)
$tabNew.Controls.Add($btnImport)

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = 'Either box may be left empty. A Delegate or OneDrive line after the last Logon Name applies to everyone in the ticket. You confirm before anything is applied.'
$lblNote.Location = New-Object System.Drawing.Point(10, 322)
$lblNote.Size = New-Object System.Drawing.Size(766, 30)
$tabNew.Controls.Add($lblNote)

# --- Tab 2: retry queue -----------------------------------------------------
$tabRetry = New-Object System.Windows.Forms.TabPage
$tabRetry.Text = 'Retry Queue'
$tabs.TabPages.Add($tabRetry)

# Buttons declared BEFORE the grid: earlier-added controls sit at a LOWER
# z-index in WinForms, so a grid added first draws OVER its own buttons.
$btnRetryPlan = New-Object System.Windows.Forms.Button
$btnRetryPlan.Text = 'Retry Selected'
$btnRetryPlan.Location = New-Object System.Drawing.Point(10, 8)
$btnRetryPlan.Size = New-Object System.Drawing.Size(120, 26)
$tabRetry.Controls.Add($btnRetryPlan)

$btnRetryRefresh = New-Object System.Windows.Forms.Button
$btnRetryRefresh.Text = 'Refresh'
$btnRetryRefresh.Location = New-Object System.Drawing.Point(138, 8)
$btnRetryRefresh.Size = New-Object System.Drawing.Size(80, 26)
$tabRetry.Controls.Add($btnRetryRefresh)

$btnRetryDrop = New-Object System.Windows.Forms.Button
$btnRetryDrop.Text = 'Remove Selected'
$btnRetryDrop.Location = New-Object System.Drawing.Point(226, 8)
$btnRetryDrop.Size = New-Object System.Drawing.Size(125, 26)
$tabRetry.Controls.Add($btnRetryDrop)

$btnRetryAll = New-Object System.Windows.Forms.Button
$btnRetryAll.Text = 'Select All'
$btnRetryAll.Location = New-Object System.Drawing.Point(359, 8)
$btnRetryAll.Size = New-Object System.Drawing.Size(85, 26)
$tabRetry.Controls.Add($btnRetryAll)

$lblRetryNote = New-Object System.Windows.Forms.Label
$lblRetryNote.Text = 'Fix the underlying problem first, then tick rows and press Retry Selected.'
$lblRetryNote.Location = New-Object System.Drawing.Point(10, 40)
$lblRetryNote.Size = New-Object System.Drawing.Size(600, 16)
$tabRetry.Controls.Add($lblRetryNote)

$gridRetry = New-Object System.Windows.Forms.DataGridView
$gridRetry.Location = New-Object System.Drawing.Point(10, 60)
$gridRetry.Size = New-Object System.Drawing.Size(766, 158)
$gridRetry.AllowUserToAddRows = $false
$gridRetry.AllowUserToDeleteRows = $false
$gridRetry.RowHeadersVisible = $false
$gridRetry.SelectionMode = 'FullRowSelect'
$gridRetry.MultiSelect = $false
$tabRetry.Controls.Add($gridRetry)

$cSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$cSel.Name = 'Sel'; $cSel.HeaderText = ''; $cSel.Width = 28
[void]$gridRetry.Columns.Add($cSel)
foreach ($c in @(
    @{N='Type';W=70}, @{N='MB';W=115}, @{N='Attempts';W=60},
    @{N='FailedStep';W=125}, @{N='LastAttempt';W=110}, @{N='LastError';W=240})) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.N; $col.HeaderText = $c.N; $col.Width = $c.W; $col.ReadOnly = $true
    [void]$gridRetry.Columns.Add($col)
}

$lblRetryDetail = New-Object System.Windows.Forms.Label
$lblRetryDetail.Text = 'Selected row:'
$lblRetryDetail.Location = New-Object System.Drawing.Point(10, 224)
$lblRetryDetail.Size = New-Object System.Drawing.Size(300, 16)
$tabRetry.Controls.Add($lblRetryDetail)

$txtRetryDetail = New-Object System.Windows.Forms.TextBox
$txtRetryDetail.Location = New-Object System.Drawing.Point(10, 242)
$txtRetryDetail.Size = New-Object System.Drawing.Size(766, 108)
$txtRetryDetail.Multiline = $true
$txtRetryDetail.ReadOnly = $true
$txtRetryDetail.ScrollBars = 'Vertical'
$txtRetryDetail.Font = New-Object System.Drawing.Font('Consolas', 9)
$tabRetry.Controls.Add($txtRetryDetail)

# --- plan -------------------------------------------------------------------
$lblPlan = New-Object System.Windows.Forms.Label
$lblPlan.Text = 'Plan:'
$lblPlan.Location = New-Object System.Drawing.Point(12, 442)
$lblPlan.Size = New-Object System.Drawing.Size(800, 16)
$form.Controls.Add($lblPlan)

$gridPlan = New-Object System.Windows.Forms.DataGridView
$gridPlan.Location = New-Object System.Drawing.Point(12, 460)
$gridPlan.Size = New-Object System.Drawing.Size(800, 88)
$gridPlan.Anchor = 'Top,Left,Right'
$gridPlan.AllowUserToAddRows = $false
$gridPlan.AllowUserToDeleteRows = $false
$gridPlan.RowHeadersVisible = $false
$gridPlan.SelectionMode = 'FullRowSelect'
$gridPlan.ReadOnly = $true
$form.Controls.Add($gridPlan)
foreach ($c in @(@{N='Type';W=80}, @{N='User';W=150}, @{N='Delegate';W=265}, @{N='OneDrive';W=265})) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.N; $col.HeaderText = $c.N; $col.Width = $c.W
    [void]$gridPlan.Columns.Add($col)
}

$btnClearPlan = New-Object System.Windows.Forms.Button
$btnClearPlan.Text = 'Clear'
$btnClearPlan.Location = New-Object System.Drawing.Point(12, 556)
$btnClearPlan.Size = New-Object System.Drawing.Size(80, 26)
$form.Controls.Add($btnClearPlan)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(100, 561)
$lblStatus.Size = New-Object System.Drawing.Size(420, 16)
$form.Controls.Add($lblStatus)

$pbar = New-Object System.Windows.Forms.ProgressBar
$pbar.Location = New-Object System.Drawing.Point(528, 558)
$pbar.Size = New-Object System.Drawing.Size(284, 20)
$pbar.Anchor = 'Top,Right'
$form.Controls.Add($pbar)

# --- results ----------------------------------------------------------------
$lblResults = New-Object System.Windows.Forms.Label
$lblResults.Text = 'Results:'
$lblResults.Location = New-Object System.Drawing.Point(12, 588)
$lblResults.Size = New-Object System.Drawing.Size(400, 16)
$form.Controls.Add($lblResults)

$gridRes = New-Object System.Windows.Forms.DataGridView
$gridRes.Location = New-Object System.Drawing.Point(12, 606)
$gridRes.Size = New-Object System.Drawing.Size(800, 262)
$gridRes.Anchor = 'Top,Left,Right,Bottom'
$gridRes.AllowUserToAddRows = $false
$gridRes.AllowUserToDeleteRows = $false
$gridRes.RowHeadersVisible = $false
$gridRes.SelectionMode = 'FullRowSelect'
$gridRes.ReadOnly = $true
$form.Controls.Add($gridRes)
foreach ($c in @(@{N='User';W=120}, @{N='Step';W=195}, @{N='Status';W=65}, @{N='Message';W=400})) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.N; $col.HeaderText = $c.N; $col.Width = $c.W
    [void]$gridRes.Columns.Add($col)
}

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(732, 876)
$btnClose.Size = New-Object System.Drawing.Size(80, 26)
$btnClose.Anchor = 'Right,Bottom'
$form.Controls.Add($btnClose)

#endregion

#region Wire-up --------------------------------------------------------------

function Update-PlanGrid {
    $gridPlan.Rows.Clear()
    foreach ($r in $script:PlanRecords) {
        $del = if ($r.Delegate) { "$($r.Delegate)@$($script:MailDomain)" } else { '(none)' }
        $od  = if ($r.OD)       { "$($r.OD)@$($script:MailDomain)" }       else { '(none)' }
        [void]$gridPlan.Rows.Add($r.Type, $r.MB, $del, $od)
    }
    $n = @($script:PlanRecords).Count
    $done = @($script:PlanRecords | Where-Object { @($_.StepsDone).Count -gt 0 }).Count
    if ($done -gt 0) {
        $lblPlan.Text = "Plan:  $n user(s).  $done carry completed steps that will not repeat."
    } else {
        $lblPlan.Text = "Plan:  $n user(s)."
    }
}

function Update-RetryGrid {
    $gridRetry.Rows.Clear()
    foreach ($r in (Import-RetryQueue)) {
        [void]$gridRetry.Rows.Add($false, $r.Type, $r.MB, $r.Attempts, $r.FailedStep,
                                  $r.LastAttempt, $r.LastError)
    }
    $tabRetry.Text = if ($gridRetry.Rows.Count -gt 0) { "Retry Queue ($($gridRetry.Rows.Count))" } else { 'Retry Queue' }
}

function Invoke-PlanRun {
    if (@($script:PlanRecords).Count -eq 0) { return }

    $check = Test-DisablePrerequisite
    Update-ConnectionLabel -Check $check
    if (-not $check.Ok) {
        $detail = foreach ($m in $check.Missing) { "  $($m.Id) -- needed for $($m.Why)$([Environment]::NewLine)      $($m.Detail)" }
        [void][System.Windows.Forms.MessageBox]::Show(
            ('Not all connections are available, so this run would fail for every user:' +
             [Environment]::NewLine + [Environment]::NewLine + ($detail -join [Environment]::NewLine) +
             [Environment]::NewLine + [Environment]::NewLine +
             'Connect from the console window that launched this tool, then press Check Connections.'),
            'Cannot run', 'OK', 'Warning')
        return
    }

    $lines = foreach ($r in $script:PlanRecords) {
        $d = if ($r.Delegate) { $r.Delegate } else { '-' }
        $o = if ($r.OD)       { $r.OD }       else { '-' }
        "   $($r.Type.PadRight(10)) $($r.MB.PadRight(14)) delegate=$d  onedrive=$o"
    }
    $ans = [System.Windows.Forms.MessageBox]::Show(
        ("This will apply changes to $(@($script:PlanRecords).Count) user(s):" +
         [Environment]::NewLine + [Environment]::NewLine + ($lines -join [Environment]::NewLine) +
         [Environment]::NewLine + [Environment]::NewLine + 'Proceed?'),
        'Confirm disable', 'YesNo', 'Warning', 'Button2')
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $btnParse.Enabled = $false; $btnRetryPlan.Enabled = $false
    $prev = $form.Cursor; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $gridRes.Rows.Clear()

    $script:RunId = [guid]::NewGuid().ToString('N').Substring(0, 12)
    $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Disable'

    $total = @($script:PlanRecords).Count
    $i = 0
    $summary = @{ Success = 0; Partial = 0; Failed = 0 }

    try {
        foreach ($rec in @($script:PlanRecords)) {
            $i++
            & $progress "[$i/$total] $($rec.MB)" ([int](100 * ($i - 1) / $total))

            # One user's defect must not abort the batch.
            try {
                $res = Invoke-DisableUser -Record $rec -Progress $progress -RunId $script:RunId
            } catch {
                $res = [PSCustomObject]@{
                    Outcome = 'Failed'
                    Smtp    = ''
                    Results = @([PSCustomObject]@{
                        User = $rec.MB; Type = $rec.Type; Step = 'Engine'
                        Label = 'Unhandled error'; Status = 'Failed'
                        Message = "$($_.Exception.Message) [line $($_.InvocationInfo.ScriptLineNumber)]"
                    })
                }
            }
            $summary[$res.Outcome] = [int]$summary[$res.Outcome] + 1

            foreach ($r in $res.Results) {
                $idx = $gridRes.Rows.Add($r.User, $r.Label, $r.Status, $r.Message)
                switch ($r.Status) {
                    'Failed'  { $gridRes.Rows[$idx].Cells['Status'].Style.ForeColor = [System.Drawing.Color]::Firebrick }
                    'Skipped' { $gridRes.Rows[$idx].Cells['Status'].Style.ForeColor = [System.Drawing.Color]::DarkOrange }
                }
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
    } finally {
        Complete-ProgressReporter -Activity 'Disable' -Bar $pbar
        $form.Cursor = $prev
        $btnParse.Enabled = $true; $btnRetryPlan.Enabled = $true
    }

    $lblStatus.Text = "Done: $($summary.Success) complete, $($summary.Partial) partial, $($summary.Failed) failed."
    Update-RetryGrid
    if ($summary.Partial -or $summary.Failed) {
        [void][System.Windows.Forms.MessageBox]::Show(
            ("$($summary.Partial) partial, $($summary.Failed) failed." + [Environment]::NewLine + [Environment]::NewLine +
             'They are in the Retry Queue tab with the reason and what to do about it.'),
            'Finished with problems', 'OK', 'Warning')
    }
}

$btnParse.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($txtHQ.Text) -and [string]::IsNullOrWhiteSpace($txtFran.Text)) {
            Show-ParseError 'Both boxes are empty -- nothing to plan.'; return
        }
        $rHQ = ConvertFrom-HQText -Text $txtHQ.Text
        if (-not $rHQ.Ok) { return }

        $mgr = ConvertTo-NormalManager -Value $txtHQMgr.Text
        if ($null -eq $mgr) {
            Show-ParseError "HQ Manager value is not valid: '$($txtHQMgr.Text.Trim())'`r`nEnter as 'bob smith' or 'bob.smith'."
            return
        }
        foreach ($r in $rHQ.Records) { $r.Delegate = $mgr; $r.OD = $mgr }

        $rFran = ConvertFrom-FranchiseText -Text $txtFran.Text
        if (-not $rFran.Ok) { return }

        $script:PlanRecords = @($rFran.Records) + @($rHQ.Records)
        if ($script:PlanRecords.Count -eq 0) {
            Show-ParseError 'Nothing recognisable was found in either box.'; return
        }
        Update-PlanGrid
        Invoke-PlanRun
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Process Requests' }
})

$btnImport.Add_Click({
    try {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV (*.csv)|*.csv'
        if (Test-Path -LiteralPath $script:RequestFolder) { $dlg.InitialDirectory = $script:RequestFolder }
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $recs = @(Import-DisableListCsv -Path $dlg.FileName)
        if ($recs.Count -eq 0) { Show-ParseError 'No usable rows found in that file.'; return }
        $script:PlanRecords = $recs
        Update-PlanGrid
        Invoke-PlanRun
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Import CSV' }
})

$btnClearPlan.Add_Click({
    try {
        $script:PlanRecords = @()
        $gridPlan.Rows.Clear(); $gridRes.Rows.Clear()
        $lblPlan.Text = 'Plan:'; $lblStatus.Text = ''; $pbar.Value = 0
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Clear' }
})

$gridRetry.Add_SelectionChanged({
    try {
        if ($gridRetry.SelectedRows.Count -eq 0) { $txtRetryDetail.Text = ''; return }
        $row = $gridRetry.SelectedRows[0]
        $mb  = "$($row.Cells['MB'].Value)"
        $ty  = "$($row.Cells['Type'].Value)"
        $rec = @(Import-RetryQueue) | Where-Object { $_.MB -eq $mb -and $_.Type -eq $ty } | Select-Object -First 1
        if (-not $rec) { $txtRetryDetail.Text = ''; return }

        $nl = [Environment]::NewLine
        $steps = @($rec.StepsDone)
        $plan  = @($script:StepPlan[$rec.Type])
        $todo  = @($plan | Where-Object { $steps -notcontains $_ })

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append("User        : $($rec.MB)  ($($rec.Type))$nl")
        if ($rec.Employee) { [void]$sb.Append("Name        : $($rec.Employee)$nl") }
        [void]$sb.Append("Delegate    : $(if ($rec.Delegate) { $rec.Delegate } else { '(none)' })$nl")
        [void]$sb.Append("OneDrive    : $(if ($rec.OD) { $rec.OD } else { '(none)' })$nl")
        [void]$sb.Append("Attempts    : $($rec.Attempts)   first failed $($rec.FirstFailed), last $($rec.LastAttempt)$nl")
        [void]$sb.Append("Completed   : $(if ($steps.Count) { $steps -join ', ' } else { '(nothing yet)' })$nl")
        [void]$sb.Append("Will run    : $(if ($todo.Count) { $todo -join ', ' } else { '(nothing outstanding)' })$nl")
        [void]$sb.Append("Failed at   : $($rec.FailedStep)$nl")
        [void]$sb.Append("Last error  : $($rec.LastError)$nl")
        $txtRetryDetail.Text = $sb.ToString()
    } catch { $txtRetryDetail.Text = "Could not read detail: $($_.Exception.Message)" }
})

$btnRetryRefresh.Add_Click({
    try { Update-RetryGrid }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Refresh Retry Queue' }
})

$btnRetryPlan.Add_Click({
    try {
        $picked = @()
        $queue = @(Import-RetryQueue)
        for ($i = 0; $i -lt $gridRetry.Rows.Count; $i++) {
            if ($gridRetry.Rows[$i].Cells['Sel'].Value -eq $true) {
                $mb = "$($gridRetry.Rows[$i].Cells['MB'].Value)"
                $ty = "$($gridRetry.Rows[$i].Cells['Type'].Value)"
                $match = $queue | Where-Object { $_.MB -eq $mb -and $_.Type -eq $ty } | Select-Object -First 1
                if ($match) { $picked += $match }
            }
        }
        if ($picked.Count -eq 0) { Show-ParseError 'No rows ticked.'; return }
        $script:PlanRecords = $picked
        Update-PlanGrid
        Invoke-PlanRun
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Retry Selected' }
})

$btnRetryAll.Add_Click({
    try {
        $any = $false
        for ($i = 0; $i -lt $gridRetry.Rows.Count; $i++) {
            if ($gridRetry.Rows[$i].Cells['Sel'].Value -ne $true) { $any = $true; break }
        }
        for ($i = 0; $i -lt $gridRetry.Rows.Count; $i++) {
            $gridRetry.Rows[$i].Cells['Sel'].Value = $any
        }
        $btnRetryAll.Text = if ($any) { 'Select None' } else { 'Select All' }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Select All' }
})

$btnRetryDrop.Add_Click({
    try {
        $drop = @()
        for ($i = 0; $i -lt $gridRetry.Rows.Count; $i++) {
            if ($gridRetry.Rows[$i].Cells['Sel'].Value -eq $true) {
                $drop += "$($gridRetry.Rows[$i].Cells['Type'].Value)|$($gridRetry.Rows[$i].Cells['MB'].Value)"
            }
        }
        if ($drop.Count -eq 0) { Show-ParseError 'No rows ticked.'; return }
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($drop.Count) row(s) from the retry queue?`r`n`r`nThis does not undo anything -- it only stops tracking them.",
            'Remove from queue', 'YesNo', 'Question', 'Button2')
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $keep = @(Import-RetryQueue | Where-Object { $drop -notcontains "$($_.Type)|$($_.MB)" })
        Export-RetryQueue -Records $keep
        Update-RetryGrid
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove From Queue' }
})

$btnCheckConn.Add_Click({
    try {
        $prev = $form.Cursor; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try { Update-ConnectionLabel -Check (Test-DisablePrerequisite) }
        finally { $form.Cursor = $prev }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Check Connections' }
})

$btnClose.Add_Click({ $form.Close() })

# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $script:RequestFolder)) {
    Write-Warning "Disable-request folder not found: '$($script:RequestFolder)'. Set it in ToolManager -> Settings..."
}
Update-RetryGrid
Update-ConnectionLabel -Check (Test-DisablePrerequisite)

$script:Theme = Get-ConsoleTheme
if ($script:Theme) {
    Set-ConsoleTheme -Form $form -Theme $script:Theme
}
Add-FormLogo -Form $form -Corner 'TopRight' -MaxHeight 56 -Margin 6 | Out-Null
# BottomLeft: TopLeft belongs to AdminConsole.ps1 itself.
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

# Consistent '?' in the top-right of every tool. Added after the theme so
# it is not restyled into a normal button, and BringToFront'd inside the
# helper because WinForms draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Term Requests' | Out-Null

[void]$form.ShowDialog()

#endregion
