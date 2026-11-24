#Requires -Version 5.1
<#
    Disable-User.ps1
    Termination / disable processing for HQ and Franchise, in one tool.

    Replaces Disable-Intake.ps1, Disable-New.ps1 and Disable-New-HQ.ps1.

    The reasoning -- why one script, why per-user rather than per-step, why
    check-then-act, why notifications are the one exception -- is in
    ScriptNotes/Disable-User.ps1.notes.md and ARCHITECTURE.md 21.

    Author   : Glenn Armbruster
    Revised  : 2026-10-09  (HQ removes and verifies the licence group too)
#>
# ALWAYS OPENS. The scanner reads the cmdlets below and would otherwise
# require all four services at once, so one missing connection hid the whole
# leaver process -- including the retry queue and the plan, which need no
# connection at all. The tool now states its own readiness in the banner at
# the top of the window instead of being hidden.
# TOOLKIT-CAPABILITY-OPTIONAL: ExO, SPO, AD, EMS
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
# FILLING IT IN -- COPY STRAIGHT FROM THE TICKET
#
# Two boxes. Use whichever matches where the person works. You can fill in
# both at once and they are planned together.
#
# FRANCHISE -- paste these lines from the ticket:
#
#     Logon Name: BXRubble
#     Email Address: Barney.Rubble@Contoso.com
#     Delegate: Fred.Flintstone@Contoso.com
#     OneDrive: Fred.Flintstone@Contoso.com
#
#   * Logon Name is the one that matters. It must be the sAMAccountName --
#     letters and numbers only. No dots, no hyphens, not an email address.
#   * Email Address is IGNORED. Paste it anyway if that saves you editing
#     the block; the tool works from the logon name.
#   * Delegate and OneDrive accept an email address or a bare name. The
#     domain is stripped, so Fred.Flintstone@Contoso.com and
#     Fred.Flintstone mean the same thing.
#   * For "nobody", any of these work and case does not matter:
#     none, NA, N/A, null, no, or a single dash.
#   * OneDrive can also be written One Drive.
#
#   MORE THAN ONE PERSON: paste several Logon Name blocks. EACH BLOCK STANDS
#   ALONE -- a Delegate or OneDrive line applies ONLY to the Logon Name above
#   it, and a block without one grants nobody anything. Nothing is inherited
#   from another block.
#
#   If two people share a delegate, write the delegate on both blocks. That
#   used to be automatic and it granted a delegate access to a second
#   departing employee's mailbox and OneDrive, because "None" and "not
#   mentioned" were treated as the same thing.
#
# HQ -- paste the ServiceNow description exactly as it arrives:
#
#     Employee Name : Betty Rubble
#     UPN : Betty.Rubble@Contoso.com
#     Last Working Day : 2026-08-11
#     Manager requires access to email and OneDrive data. Also convert to
#     Shared Mailbox.
#
#   * UPN is required. It must be a full email address, or a plain
#     alphanumeric logon name. Betty.Rubble on its own will NOT do -- it
#     has a dot but no domain, and is rejected.
#   * Employee Name and Last Working Day are read. Everything else in the
#     description is ignored, including the sentence about the manager --
#     that sentence is not how the manager gets set. The HQ Manager box is.
#   * Spaces around the colons are fine. Do not tidy the paste.
#
# HQ MANAGER
#
# Who receives the mailbox and the OneDrive. Fill this in for every HQ
# request -- nothing in the pasted description supplies it.
#
#   * Wilma.Flintstone, Wilma Flintstone and Wilma.Flintstone@Contoso.com
#     are all accepted. A space becomes a dot and the domain is dropped.
#   * ONE MANAGER FOR THE WHOLE HQ BOX. Whatever is here is applied to
#     every HQ request pasted above it. Two HQ requests with different
#     managers are two separate runs, not one.
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
#   * ENABLE MAILBOX (2026-10-06). The commonest failure -- "exists but has
#     no remote mailbox" -- is fixed from the queue itself: tick the rows,
#     press Enable Mailbox, and the same Enable-RemoteMailbox + archive that
#     New User runs is done for each, on its own domain controller. Wait for
#     a directory sync (~30 min), then Retry Selected.
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

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

#region Configuration --------------------------------------------------------

# The licence group is NOT a constant here any more.
#
# It was 'E5 License Fran - Full' -- one tenant's group name, in one directory,
# written into a script that ships to six. It is resolved per record now from
# Tenant.json via Resolve-TenantLicenseGroupName, because which group applies
# depends on which directory the account is in.
$script:OneDriveBaseUrl   = "$((Get-TenantConfig).SharePoint.MySiteHost)/personal"
$script:SmtpServer        = "smtp.$(Get-MailDomain)"
$script:DoNotReplyAddress = "DO_NOT_REPLY@$(Get-MailDomain)"

$script:NotificationCopyTo = Get-OperatorEmail

# ImportCsvPath was removed, not overlooked. It pointed at a fixed
# disablelist.csv in this folder, and import has gone through a file dialog
# since the rewrite -- the variable was assigned and never read again.
$script:RetryFolder  = Get-RetryQueuePath
$script:RetryCsvPath = Join-Path $script:RetryFolder 'disable-retry.csv'

$script:StepPlan = @{
    Franchise = @(
        'ConvertShared', 'HideFromGAL', 'GrantDelegate', 'RemoveLicence',
        'GrantOneDrive', 'NotifyOneDrive', 'DisableForwarding', 'ValidateLicence'
    )
    # SAME AS FRANCHISE SINCE 2026-10-09: the HQ licence groups are flagged
    # 'Remove on term' in Tenant Manager now, so the two licence steps run
    # for HQ leavers too. Same step order, same code -- the directory the
    # row belongs to picks the group.
    HQ = @(
        'ConvertShared', 'HideFromGAL', 'GrantDelegate', 'RemoveLicence',
        'GrantOneDrive', 'NotifyOneDrive', 'DisableForwarding', 'ValidateLicence'
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
    $null = Show-ToolDialog($Message, 'Input Error', 'OK', 'Error')
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

    # NO INHERITANCE BETWEEN BLOCKS. EVERY BLOCK STANDS ALONE.
    #
    # This used to fill a blank Delegate or OneDrive from the LAST record in the
    # ticket, so that "one delegate at the bottom covers everybody above it".
    # It granted access nobody asked for, on a real termination:
    #
    #     Logon Name: CLHeid      Delegate: None
    #     Logon Name: LAGaytan    Delegate: Brett.Clark
    #
    # ConvertTo-NormalValue maps 'None' to '', which made an EXPLICIT REFUSAL
    # indistinguishable from an omission. The blank then inherited Brett, and he
    # was granted FullAccess and OneDrive on a second departing employee's data.
    # Nothing errored. The grid showed it and it read like what was typed.
    #
    # TWO DEFECTS, ONE REMOVAL. A value written inside a block belongs to that
    # block, and 'None' means nobody -- not "go and ask the record below".
    #
    # If several people really do share one delegate, write the delegate on each
    # block. Repetition is cheap; a silent extra grant is not.
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

    # Create the default folder, never a configured one. See Get-RetryQueuePath:
    # silently creating a mistyped path produces an empty queue in the wrong
    # place, and an empty queue reads as good news. Losing the rows is worse
    # than either, so a configured-but-missing folder throws rather than
    # dropping the write.
    $q = Test-RetryQueueConfigured
    if (-not $q.Exists) {
        if ($q.MayBeCreated) {
            New-Item -Path $q.Path -ItemType Directory -Force | Out-Null
        } else {
            throw "Retry queue folder '$($q.Path)' does not exist. It is set in ToolManager -> Settings and is not created automatically, because an empty folder there would look like an empty queue. $(@($Records).Count) row(s) were NOT saved."
        }
    }

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
    # Whatever directory Id the caller holds -- was a hardcoded two-way branch
    # naming one organisation's directories.
    if ($Type) { return (Get-TenantDirectoryServer -Id $Type) }
    return (Get-TenantDirectoryServer)
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

    # "Could not ask" is not "not there", and the difference is the whole
    # message.
    #
    # This was 'catch { $ad = $null }', so an unreachable domain controller, an
    # expired session or a permissions error all arrived at the same place as a
    # genuine miss -- and the operator was told "most likely a typo in the
    # ticket". They would then check the ticket, find the identity correct,
    # check it again, and have nothing else to go on, because the one fact that
    # would have explained it was discarded one line earlier.
    #
    # Same shape as the catch that reported "version unavailable" for what was
    # really a control-ordering bug. A catch that names the wrong cause costs
    # more than no catch at all.
    $adError = $null
    try {
        if ($Record.MB -match '@') {
            $upn = ConvertTo-LdapSafe $Record.MB
            $ad = Get-ADUser -LDAPFilter "(userPrincipalName=$upn)" -Server $server -Properties MemberOf, mail -ErrorAction Stop |
                  Select-Object -First 1
        } else {
            $ad = Get-ADUser -Identity $Record.MB -Server $server -Properties MemberOf, mail -ErrorAction Stop
        }
    } catch {
        $ad = $null
        # Get-ADUser -Identity throws ADIdentityNotFoundException when the
        # account genuinely is not there. That one IS a miss; anything else is
        # a failure to look.
        if ("$($_.Exception.GetType().FullName)" -notmatch 'ADIdentityNotFoundException') {
            $adError = $_
        }
    }

    if (-not $ad) {
        if ($adError) {
            $out.Reason = "Could not query $server for '$($Record.MB)': $($adError.Exception.Message)"
            $out.Fix    = ("The directory did not answer, so whether this account exists is UNKNOWN. " +
                           "Check the domain controller is reachable and that your session still has " +
                           "rights, then retry. Do not remove the row -- nothing has been ruled out.")
        } else {
            $out.Reason = "No AD account found for '$($Record.MB)' in $server."
            $out.Fix    = 'Most likely a typo in the ticket. Verify the identity and correct it, or remove the row.'
        }
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
    <#
        THE RECIPIENT IS NAMED IN THE BODY, not just in the To: header.

        The operator used to get a SEPARATE copy of this mail, addressed to
        themselves, so the To: line on their copy said their own name and the
        body never said whose OneDrive access had been handed to whom. Working
        out who was actually notified meant going back to the ticket.

        There is one send now, to the manager, with the operator on BCC -- and
        a BCC recipient sees the To: header of the original, so the address is
        visible either way. Restating it in the body is belt and braces: it
        survives forwarding, quoting, and mail clients that collapse headers.
    #>
    param([string]$Type, [string]$Smtp, [string]$Url, [string]$Manager)

    # Rendered only when known. An empty paragraph reading "Sent to:" with
    # nothing after it looks like a bug in the mail rather than a missing
    # value, and this function is called before the caller has validated much.
    $sentTo = ''
    if ("$Manager".Trim()) {
        $sentTo = "<p><i>This notification was sent to: $Manager</i></p>"
    }

    if ($Type -eq 'HQ') {
        return @"
<html>
<body>
$sentTo
<p>To access the user's OneDrive for Business, you can use the following link: <a href='$Url'>OneDrive for Business</a></p>
</body>
</html>
"@
    }

    return @"
<html>
<body>
<p>Greetings,</p>
$sentTo
<p>The user account $Smtp has been disabled in the Active Directory. However, please note that their OneDrive for Business data will be preserved for 30 days. During this period, you are the temporary owner of all the documents saved to their OneDrive for Business.</p>
<p>If you wish to save any important documents beyond the 30-day retention period, you have the option to copy them to another location. Additionally, you can contact the Assistance Center to reassign ownership to another OneDrive for Business owner. Please be aware that after the 30-day period, the user's OneDrive for Business data will be permanently deleted.</p>
<p>To access the user's OneDrive for Business, you can use the following link: <a href='$Url'>OneDrive for Business</a></p>
<p>If you have any questions or need further assistance, please contact the Assistance Center.</p>
</body>
</html>
"@
}

function Resolve-LicenceGroupMembership {
    <#
        Which of the configured licence groups is this account actually in --
        decided by DISTINGUISHED NAME, not by the name as typed.

        FOUND 2026-10-09 on the first HQ run: the step reported "Not a member
        of E5 License HQ - Full" for an account that ADUC showed in that very
        group, and Verify then "confirmed" the removal that never happened.
        The old test was  MemberOf -like "CN=<typed name>,*"  -- one stray
        character between Tenant Manager and the real group name (an en dash
        for a hyphen, a double space, a comma the DN escapes) and the match
        fails silently in the safe-looking direction.

        So each configured name is looked up as a group on the row's own
        controller; a name that does not resolve is an ERROR (the setting is
        wrong), and membership is the group's DN against the account's
        MemberOf, which cannot drift.
    #>
    param($Ctx, [string[]]$Names)
    $r = [PSCustomObject]@{ Groups = @(); Missing = @(); InGroup = @() }
    foreach ($n in $Names) {
        $g = $null
        try { $g = Get-ADGroup -Identity $n -Server $Ctx.AdServer -ErrorAction Stop }
        catch {
            # The name may differ only by dash or spacing: try a filter that
            # ignores both before giving up.
            $loose = ("$n" -replace '[\u2010-\u2015\u2212]', '-') -replace '\s+', ' '
            $pat = ($loose -replace '[-\s]', '*')
            try {
                $cands = @(Get-ADGroup -Filter "Name -like '$pat'" -Server $Ctx.AdServer -ErrorAction Stop)
                if ($cands.Count -eq 1) { $g = $cands[0] }
            } catch { }
        }
        if ($g) { $r.Groups += $g } else { $r.Missing += $n }
    }
    $mine = @($Ctx.AdUser.MemberOf | ForEach-Object { "$_".ToLower() })
    $r.InGroup = @($r.Groups | Where-Object { $mine -contains "$($_.DistinguishedName)".ToLower() })
    return $r
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
            # ONLY the groups flagged RemoveOnTerm in Tenant.json.
            #
            # This resolved to a SINGLE name and got whichever licence group was
            # first in the file. At Express that is 'E1 License Fran', which
            # exists so the Disabled/Licensed report can find people who are
            # disabled but still licensed -- a reporting input, not a licence a
            # leaver holds. So the step removed terminated staff from a group
            # they were never in, reported Done, and left the E5 licence
            # assigned and billing.
            #
            # Selecting "all groups for the directory" would be just as wrong in
            # the other direction: it would strip the E1 membership the report
            # depends on. Neither guess is available, so the config states it.
            $licNames = @(Get-TenantTerminationLicenseGroup -DirectoryId $Ctx.Type)
            if ($licNames.Count -eq 0) {
                # FAILED, not Skipped, and no fallback. "Not a member" would be
                # a claim this step is not entitled to make when it checked
                # nothing, and a green run that removed no licence is the
                # outcome that ships and turns up on an invoice.
                return New-StepResult 'Failed' ("No licence group for directory '$($Ctx.Type)' is " +
                    "flagged 'Remove on term' in Tenant Manager. The licence has NOT been " +
                    "removed. Tick the leaver's licence group -- not the reporting-only ones.")
            }

            $m = Resolve-LicenceGroupMembership -Ctx $Ctx -Names $licNames
            if ($m.Missing.Count) {
                return New-StepResult 'Failed' ("Licence group(s) not found in $($Ctx.AdServer): $($m.Missing -join ', ') -- " +
                    "the name in Tenant Manager does not match the AD group name (check dashes and spaces). Nothing was removed.")
            }
            if ($m.InGroup.Count -eq 0) {
                # SAYS WHAT IT COMPARED. If the account really is in the group
                # and this still says Skipped, the MemberOf count and the
                # resolved DNs are the two facts needed to see why.
                $nMine = @($Ctx.AdUser.MemberOf).Count
                return New-StepResult 'Skipped' ("Not a member of $(@($m.Groups | ForEach-Object { $_.Name }) -join ', '). " +
                    "(Account read from $($Ctx.AdServer) lists $nMine group(s); resolved group DN: $(@($m.Groups | ForEach-Object { $_.DistinguishedName }) -join ' | '))")
            }

            $removed = @()
            foreach ($grp in $m.InGroup) {
                Remove-ADGroupMember -Identity $grp -Members $Ctx.Sam -Server $Ctx.AdServer -Confirm:$false -ErrorAction Stop
                $removed += $grp.Name
                Write-CompletedLog -Action 'Revoke' -Target $Ctx.Smtp -Result 'Success' `
                    -Environment "$($Ctx.Type)" -Detail "Removed from licence group '$($grp.Name)'."
            }
            return New-StepResult 'Done' "Removed from $($removed -join ', ')."
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
            $body = Get-OneDriveNotificationBody -Type $Ctx.Type -Smtp $Ctx.Smtp -Url $url -Manager $Ctx.OdSmtp

            # The DEPARTING user's address: it identifies which OneDrive.
            #
            # The '- HQ' / '- Fran' suffix that used to be appended to the
            # operator's copy is gone. It distinguished two messages that no
            # longer exist as two messages, and the directory is already
            # evident from the addresses in the mail.
            $subject = "OneDrive Ownership Change Notification - $($Ctx.Smtp)"

            # ONE SEND, WITH THE OPERATOR ON BCC -- not two sends.
            #
            # Two sends meant two chances to fail, and they failed
            # independently: the manager could be notified while the operator's
            # copy bounced, or the reverse, with a single step result covering
            # both. Worse, the second send happened AFTER the first had already
            # gone out, so a throw there reported the whole step Failed when the
            # notification the ticket cares about had actually been delivered.
            #
            # BCC also removes the operator's address from what the manager
            # sees, which the separate copy never did.
            # BodyAsHtml IS A SWITCH. The body goes in -Body.
            #
            # The original line read '-Subject $subject -BodyAsHtml $body' and
            # worked, which is why this was easy to get wrong: -BodyAsHtml is a
            # switch so it consumed nothing, and $body then bound POSITIONALLY
            # to -Body, which is position 2 once -To and -Subject are named.
            #
            # Splatting has no positional binding. Carrying that line into a
            # hashtable verbatim put a string into a switch, and the step failed
            # after every other step in the termination had already succeeded.
            $mail = @{
                To          = $Ctx.OdSmtp
                Subject     = $subject
                Body        = $body
                BodyAsHtml  = $true
                From        = $script:DoNotReplyAddress
                SmtpServer  = $script:SmtpServer
                ErrorAction = 'Stop'
            }

            # No configured operator address means NO copy -- never fall back to
            # a default, or somebody who did not do the work gets the mail.
            if ($script:NotificationCopyTo) { $mail['Bcc'] = $script:NotificationCopyTo }

            Send-MailMessage @mail

            if ($script:NotificationCopyTo) {
                return New-StepResult 'Done' "Notified $($Ctx.OdSmtp); blind copy to $($script:NotificationCopyTo)."
            }
            return New-StepResult 'Done' "Notified $($Ctx.OdSmtp). No copy sent -- set 'Your email' in ToolManager > Settings."
        }

        'DisableForwarding' {
            $rules = @()
            try {
                # THE SLOWEST STEP IN THE PIPELINE, ROUTINELY 20-60 SECONDS,
                # AND MINUTES ON A MAILBOX WITH CORRUPT RULES.
                #
                # Get-InboxRule with -IncludeHidden reads the whole rule
                # collection out of the mailbox rather than a directory
                # attribute, and it does not stream. A mailbox carrying rules
                # left by old clients, or referring to folders and addresses
                # that no longer exist, is slower again -- the service works
                # through every one of them before answering.
                #
                # It is not stuck. Say so before making the call, because the
                # status line is the only thing telling the operator which it
                # is.
                if ($Progress) { & $Progress "$($Ctx.Smtp): reading inbox rules (slow -- up to a minute)" }
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
            $names  = @()
            $failed  = @()
            foreach ($rule in $fwd) {
                # ONE BAD RULE MUST NOT LEAVE THE REST FORWARDING.
                #
                # This loop had no inner try, so the first rule that would not
                # disable threw out of the whole step -- and every rule after it
                # was never attempted. On a mailbox with corrupt rules that is
                # the worst possible outcome: the tool reports a failure, the
                # operator sees one error, and mail is STILL LEAVING through the
                # rules it never reached. Corrupt rules are exactly the ones that
                # refuse to disable, so the case that breaks this is the case
                # that most needs it to work.
                #
                # -Identity ONLY: Exchange Online rejects Identity and Mailbox
                # together, and the Identity from Get-InboxRule is already scoped.
                try {
                    Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters @{
                        Identity = $rule.Identity; Confirm = $false; ErrorAction = 'Stop' }
                    $names += "$($rule.Name)"
                } catch {
                    # A corrupt rule often has no usable Name. Fall back to the
                    # Identity so the row names something you can search for.
                    $label = if ("$($rule.Name)".Trim()) { "$($rule.Name)" }
                             else { "(unnamed rule: $($rule.Identity))" }
                    $failed += "$label -- $($_.Exception.Message)"
                }
            }
            if ($failed.Count -eq 0) {
                return New-StepResult 'Done' "Disabled $($names.Count) forwarding rule(s): $($names -join ', ')"
            }
            # FAILED, NOT DONE, EVEN THOUGH MOST OF THEM WORKED. The step's job
            # is to stop mail leaving, and it has not finished doing that. Failed
            # here makes the user Partial at run level and puts them in the retry
            # queue, which is the only thing that brings somebody back to it.
            return New-StepResult 'Failed' (
                "Disabled $($names.Count) of $($fwd.Count) forwarding rule(s). " +
                "STILL FORWARDING -- $($failed.Count) could not be disabled: $($failed -join ' | ')")
        }

        'ValidateLicence' {
            # Re-read from AD: the point of this step is to confirm the removal landed.
            # Verifies the SAME set the removal acted on. Checking a different
            # group than the one removed is how a verification step ends up
            # confirming something nobody did.
            $licNames = @(Get-TenantTerminationLicenseGroup -DirectoryId $Ctx.Type)
            if ($licNames.Count -eq 0) {
                return New-StepResult 'Failed' ("No licence group for directory '$($Ctx.Type)' is " +
                    "flagged 'Remove on term', so removal cannot be verified.")
            }
            $fresh = Get-ADUser -Identity $Ctx.Sam -Server $Ctx.AdServer -Properties MemberOf -ErrorAction Stop
            $freshCtx = [PSCustomObject]@{ AdUser = $fresh; AdServer = $Ctx.AdServer }
            $m = Resolve-LicenceGroupMembership -Ctx $freshCtx -Names $licNames
            if ($m.Missing.Count) {
                return New-StepResult 'Failed' "Licence group(s) not found in $($Ctx.AdServer): $($m.Missing -join ', ') -- cannot verify."
            }
            if ($m.InGroup.Count -gt 0) {
                return New-StepResult 'Failed' "STILL a member of $(@($m.InGroup | ForEach-Object { $_.Name }) -join ', ') after removal."
            }
            return New-StepResult 'Done' "Confirmed removed from $(@($m.Groups | ForEach-Object { $_.Name }) -join ', ')."
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
    <#
        READY: quiet. The banner blends into the window, because a green bar
        shouting at you on every normal launch is how a red one gets ignored.

        NOT READY: bright red on black, and it says WHICH STEPS will not run.
        'NOT connected: SPO' is only useful to somebody who already knows what
        SPO does for a leaver; naming the step is what makes it actionable at
        the moment it matters.

        The black is set explicitly rather than left to the theme -- see the
        'NoTheme' tag on the control.
    #>
    param($Check)
    if ($Check.Ok) {
        $lblConn.Text        = 'Ready: on-prem Exchange, AD, Exchange Online and SharePoint Online are all connected.'
        $lblConn.BackColor   = $form.BackColor
        $lblConn.ForeColor   = [System.Drawing.Color]::FromArgb(0, 130, 0)
        $lblConn.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        return
    }
    $ids  = (@($Check.Missing | ForEach-Object { $_.Id })) -join ', '
    $whys = (@($Check.Missing | ForEach-Object { $_.Why })) -join '; '
    $onlySpo = -not @($Check.Missing | Where-Object { $_.Id -ne 'SPO' }).Count
    $lblConn.Text = $(if ($onlySpo) {
        "NOT CONNECTED: SPO`r`nOneDrive steps go to the Retry Queue; everything else runs.  Connect SPO in the console window to include them."
    } else {
        "NOT CONNECTED: $ids`r`nThis will fail: $whys.  Connect in the console window, then return here."
    })
    $lblConn.BackColor   = [System.Drawing.Color]::Black
    $lblConn.ForeColor   = [System.Drawing.Color]::FromArgb(255, 64, 64)
    $lblConn.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
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
        DelegateSmtp = $(if ($Record.Delegate) { "$($Record.Delegate)@$($(Get-MailDomain))" } else { '' })
        OdSmtp       = $(if ($Record.OD)       { "$($Record.OD)@$($(Get-MailDomain))" }       else { '' })
    }

    $plan = @($script:StepPlan[$Record.Type])
    $failedStep = ''
    $lastError  = ''

    foreach ($step in $plan) {

        if ($script:StateTrackedSteps -contains $step -and $stepsDone -contains $step) {
            & $addResult $step 'Skipped' 'Already sent on a previous attempt -- not resent.'
            continue
        }

        # SHAREPOINT ONLINE DOWN = ONEDRIVE WAITS, EVERYTHING ELSE RUNS
        # (2026-09-29). SPO is used by the two OneDrive steps and nothing
        # else, so its absence no longer blocks the run. For a user WITH a
        # OneDrive owner, the grant is recorded as failed -- which is what puts
        # the user in the Retry Queue -- and the notification is held back,
        # NOT marked done: it is state-tracked, and marking it done would mean
        # the retry never sends it. A user with no OneDrive owner is unaffected.
        #
        # Set by Invoke-PlanRun from the live probe. $null (not set) means
        # "not known to be down" and changes nothing.
        if ($script:SpoDown -and $Ctx.OdSmtp -and ($step -eq 'GrantOneDrive' -or $step -eq 'NotifyOneDrive')) {
            if ($step -eq 'GrantOneDrive') {
                & $addResult $step 'Failed' 'SharePoint Online is not connected -- OneDrive access left for the Retry Queue.'
                if (-not $failedStep) { $failedStep = $step }
                $lastError = 'SharePoint Online not connected -- OneDrive steps not attempted.'
            } else {
                & $addResult $step 'Skipped' 'Not sent yet -- OneDrive access has not been granted. It is sent when the Retry Queue completes the grant.'
            }
            continue
        }

        if ($Progress) { & $Progress "$($Record.MB): $($script:StepLabels[$step])" }

        # A FAILED STEP GETS ITS OWN LOG ROW.
        #
        # The run already writes one summary row at the end, which records that
        # something failed and names the step -- but not what the service
        # actually said. When a termination half-applies, the question is never
        # "did it fail", it is "which call, against which object, with what
        # error", and that has to be a row you can find by searching the target.
        #
        # Both paths are covered on purpose: a step can THROW, or it can return
        # a Failed result without throwing. Logging only the throw would leave
        # the quieter half of the failures with no row at all -- and the quiet
        # ones are the ones nobody notices.
        $thrown = $null
        try {
            $r = Invoke-DisableStep -Step $step -Ctx $ctx
        } catch {
            $thrown = $_
            $r = New-StepResult 'Failed' $_.Exception.Message
        }

        & $addResult $step $r.Status $r.Message

        if ($r.Status -eq 'Failed') {
            if (-not $failedStep) { $failedStep = $step }
            $lastError = $r.Message
            Write-FailedLog -Action $step -Target $ctx.Smtp -ErrorRecord $thrown `
                -Environment $Record.Type -RunId $RunId `
                -Detail "$($script:StepLabels[$step]) failed$(if (-not $thrown) { ": $($r.Message)" })"
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
$form.Size = New-Object System.Drawing.Size(840, 984)
$form.StartPosition = 'Manual'
$form.MinimumSize = New-Object System.Drawing.Size(840, 726)
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

# THE READINESS BANNER. Tag 'NoTheme' keeps the theme off its COLOURS -- the
# colour here carries meaning and a theme must never restyle it into
# decoration. The theme still re-faces it, which is wanted: Set-ThemeFontOn
# preserves size and style, so 12pt Bold survives in whatever family the
# theme uses.
#
# THE HEIGHT IS RESERVED WHETHER OR NOT THERE IS A PROBLEM. Growing the label
# only when something is missing would mean moving every control below it at
# run time, and a bottom-anchored grid re-measures its own distance the moment
# its bounds change -- the 80px travel documented in Add-ToolTargetBar. Fixed
# space, changing colour, is the version with no moving parts.
$lblConn = New-Object System.Windows.Forms.Label
$lblConn.Text = 'Checking connections...'
$lblConn.Location = New-Object System.Drawing.Point(12, 10)
$lblConn.Size = New-Object System.Drawing.Size(800, 44)
$lblConn.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12, [System.Drawing.FontStyle]::Bold)
$lblConn.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblConn.Padding = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$lblConn.Tag = 'NoTheme'
$form.Controls.Add($lblConn)

# The 'Check Connections' button that used to sit here has been removed.
#
# No other tool has one, and a control unique to one window reads as an
# inconsistency rather than a feature. It also asked the operator to do the
# tool's job: you connect in the console window, come back here, and the label
# is stale until you remember to press something.
#
# The refresh moved to Form.Activated instead -- returning to this window IS
# the signal that something may have changed. Same pattern as the report file
# controls in Common.ps1, which dropped their Refresh button for the same
# reason.
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, 60)
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
$lblNote.Text = 'Either box may be left empty. A Delegate or OneDrive line applies ONLY to the Logon Name above it -- nothing is shared between blocks. Check the Delegate and OneDrive columns in the plan before you press Process.'
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

$btnRetryEnable = New-Object System.Windows.Forms.Button
$btnRetryEnable.Text = 'Enable Mailbox'
$btnRetryEnable.Location = New-Object System.Drawing.Point(452, 8)
$btnRetryEnable.Size = New-Object System.Drawing.Size(120, 26)
$tabRetry.Controls.Add($btnRetryEnable)

$lblRetryNote = New-Object System.Windows.Forms.Label
$lblRetryNote.Text = 'Fix the underlying problem first, then tick rows and press Retry Selected. "No remote mailbox": tick the rows, Enable Mailbox, wait for a sync, then Retry Selected.'
$lblRetryNote.Location = New-Object System.Drawing.Point(10, 40)
$lblRetryNote.Size = New-Object System.Drawing.Size(766, 16)
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
$lblPlan.Location = New-Object System.Drawing.Point(12, 468)
$lblPlan.Size = New-Object System.Drawing.Size(800, 16)
$form.Controls.Add($lblPlan)

$gridPlan = New-Object System.Windows.Forms.DataGridView
$gridPlan.Location = New-Object System.Drawing.Point(12, 486)
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
$btnClearPlan.Location = New-Object System.Drawing.Point(12, 582)
$btnClearPlan.Size = New-Object System.Drawing.Size(80, 26)
$form.Controls.Add($btnClearPlan)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(100, 587)
$lblStatus.Size = New-Object System.Drawing.Size(420, 16)
$form.Controls.Add($lblStatus)

$pbar = New-Object System.Windows.Forms.ProgressBar
$pbar.Location = New-Object System.Drawing.Point(528, 584)
$pbar.Size = New-Object System.Drawing.Size(284, 20)
$pbar.Anchor = 'Top,Right'
$form.Controls.Add($pbar)

# --- results ----------------------------------------------------------------
$lblResults = New-Object System.Windows.Forms.Label
$lblResults.Text = 'Results:'
$lblResults.Location = New-Object System.Drawing.Point(12, 614)
$lblResults.Size = New-Object System.Drawing.Size(400, 16)
$form.Controls.Add($lblResults)

$gridRes = New-Object System.Windows.Forms.DataGridView
$gridRes.Location = New-Object System.Drawing.Point(12, 632)
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
$btnClose.Location = New-Object System.Drawing.Point(732, 902)
$btnClose.Size = New-Object System.Drawing.Size(80, 26)
$btnClose.Anchor = 'Right,Bottom'
$form.Controls.Add($btnClose)

#endregion

#region Wire-up --------------------------------------------------------------

function Update-PlanGrid {
    $gridPlan.Rows.Clear()
    foreach ($r in $script:PlanRecords) {
        $del = if ($r.Delegate) { "$($r.Delegate)@$($(Get-MailDomain))" } else { '(none)' }
        $od  = if ($r.OD)       { "$($r.OD)@$($(Get-MailDomain))" }       else { '(none)' }
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
    # ONLY WHAT EVERY USER NEEDS BLOCKS THE RUN. On-premises Exchange, AD and
    # Exchange Online are used for every user; SharePoint Online only for a
    # user with a OneDrive owner, and those go to the Retry Queue instead.
    $blocking = @($check.Missing | Where-Object { $_.Id -ne 'SPO' })
    $script:SpoDown = [bool](@($check.Missing | Where-Object { $_.Id -eq 'SPO' }).Count)
    if ($blocking.Count) {
        $detail = foreach ($m in $blocking) { "  $($m.Id) -- needed for $($m.Why)$([Environment]::NewLine)      $($m.Detail)" }
        $null = Show-ToolDialog(
            ('Not all connections are available, so this run would fail for every user:' +
             [Environment]::NewLine + [Environment]::NewLine + ($detail -join [Environment]::NewLine) +
             [Environment]::NewLine + [Environment]::NewLine +
             'Connect from the console window that launched this tool, then come back to this window -- the banner at the top refreshes on its own.'),
            'Cannot run', 'OK', 'Warning')
        return
    }
    $spoNote = ''
    if ($script:SpoDown) {
        $withOd = @($script:PlanRecords | Where-Object { "$($_.OD)".Trim() }).Count
        $spoNote = $(if ($withOd) {
            [Environment]::NewLine + [Environment]::NewLine +
            "SharePoint Online is NOT connected. Everything else runs now; the OneDrive steps for the $withOd user(s) with a OneDrive owner go to the Retry Queue -- connect SharePoint Online and retry them from there."
        } else {
            [Environment]::NewLine + [Environment]::NewLine +
            'SharePoint Online is not connected, but no one in this run has a OneDrive owner, so nothing needs it.'
        })
    }

    $lines = foreach ($r in $script:PlanRecords) {
        $d = if ($r.Delegate) { $r.Delegate } else { '-' }
        $o = if ($r.OD)       { $r.OD }       else { '-' }
        "   $($r.Type.PadRight(10)) $($r.MB.PadRight(14)) delegate=$d  onedrive=$o"
    }
    $ans = Show-ToolDialog(
        ("This will apply changes to $(@($script:PlanRecords).Count) user(s):" +
         [Environment]::NewLine + [Environment]::NewLine + ($lines -join [Environment]::NewLine) +
         $spoNote +
         [Environment]::NewLine + [Environment]::NewLine + 'Proceed?'),
        'Confirm disable', 'YesNo', 'Warning', 'Button2')
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    # EVERY interactive control, not two buttons. Clicks queue while the run
    # works and are delivered afterwards; a disabled control banks nothing at
    # all. Close is the one that mattered -- see New-ProgressReporter.
    $lockedControls = @(
        $btnParse, $btnImport, $btnClearPlan, $btnClose,
        $btnRetryPlan, $btnRetryRefresh, $btnRetryDrop, $btnRetryAll, $btnRetryEnable,
        $tabs, $txtHQ, $txtHQMgr, $txtFran, $gridPlan, $gridRetry
    )
    # EVERY ONE OF THESE IS ENABLED AGAIN AFTERWARDS, unconditionally. The
    # previous version remembered each control's prior state in a hashtable
    # keyed by the control object and put that back; if that lookup ever
    # missed, [bool]$null re-enabled nothing and the window was dead until it
    # was reopened (reported 2026-10-06). None of these is ever legitimately
    # disabled outside a run, so there is no prior state worth preserving.
    foreach ($c in $lockedControls) {
        if ($c) { $c.Enabled = $false }
    }
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
            # REPAINT, DO NOT PUMP. The results grid is the only thing that
            # changed, and Refresh() shows the new rows without delivering a
            # queued Close into the middle of the batch.
            if (-not $gridRes.IsDisposed) { $gridRes.Refresh() }
        }
    } finally {
        # GUARDED. If the form went away underneath this -- which it no longer
        # can, but this block is the last thing to run either way -- touching a
        # disposed control here throws from a finally, which loses whatever the
        # original failure was.
        if (-not $form.IsDisposed) {
            Complete-ProgressReporter -Activity 'Disable' -Bar $pbar
            $form.Cursor = $prev
            foreach ($c in $lockedControls) {
                if ($c -and -not $c.IsDisposed) { $c.Enabled = $true }
            }
        }
    }

    $lblStatus.Text = "Done: $($summary.Success) complete, $($summary.Partial) partial, $($summary.Failed) failed."
    Update-RetryGrid
    if ($summary.Partial -or $summary.Failed) {
        $null = Show-ToolDialog(
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
        # No InitialDirectory. This used to open in the disable-request folder
        # back when that folder held the list being imported. It now holds only
        # retry queues written by the toolkit, so starting there would point the
        # operator at the one place their import file will never be. Windows
        # remembers the last folder used, which is a better guess than ours.
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

$btnRetryEnable.Add_Click({
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

        # THE SAME GATE THE RUN USES: on-premises Exchange and AD, nothing else.
        $check = Test-DisablePrerequisite
        Update-ConnectionLabel -Check $check
        $need = @($check.Missing | Where-Object { $_.Id -in @('EMS', 'AD') })
        if ($need.Count) {
            $null = Show-ToolDialog(
                ('Enable Mailbox needs ' + (($need | ForEach-Object { $_.Id }) -join ' and ') + ' connected in the console window that launched this tool.'),
                'Cannot enable', 'OK', 'Warning')
            return
        }

        # RESOLVE EVERY ROW FIRST. The AD account must exist and must have NO
        # remote mailbox; anything else is reported and not touched. The
        # controller is the one Resolve-DirectoryDomainController picks for
        # the row's directory -- a real server, never the domain name, which
        # Exchange rejects on -DomainController.
        $plan = @(); $skip = @()
        foreach ($rec in $picked) {
            $t = Resolve-DisableTarget -Record $rec
            if (-not $t.AdUser) { $skip += "$($rec.MB)  --  $($t.Reason)"; continue }
            if ($t.Ok -or $t.Reason -notmatch 'no remote mailbox') { $skip += "$($rec.MB)  --  already has a mailbox: $($t.Reason)"; continue }
            $ad = $null
            $dc = ''
            try { $dc = Resolve-DirectoryDomainController -DirectoryId "$($rec.Type)" } catch { $dc = '' }
            $adp = @{ Identity = "$($t.AdUser.DistinguishedName)"; Properties = @('mailNickname', 'sAMAccountName'); ErrorAction = 'Stop' }
            if ($dc) { $adp['Server'] = $dc }
            try { $ad = Get-ADUser @adp } catch { $skip += "$($rec.MB)  --  could not re-read the account: $($_.Exception.Message)"; continue }
            $alias = "$($ad.mailNickname)".Trim(); if (-not $alias) { $alias = "$($ad.sAMAccountName)".Trim() }
            if (-not $alias) { $skip += "$($rec.MB)  --  no alias or logon name to build the routing address from"; continue }
            $plan += [PSCustomObject]@{ Record = $rec; Dn = "$($ad.DistinguishedName)"; Alias = $alias; Dc = $dc; Name = "$($ad.Name)" }
        }
        if ($plan.Count -eq 0) {
            $null = Show-ToolDialog(("Nothing to enable:" + [Environment]::NewLine + [Environment]::NewLine + ($skip -join [Environment]::NewLine)), 'Enable Mailbox', 'OK', 'Information')
            return
        }
        $routingDomain = Get-TenantRemoteRoutingDomain
        $lines = @($plan | ForEach-Object { "   $($_.Record.Type.PadRight(10)) $($_.Record.MB.PadRight(14)) -> $($_.Alias)@$routingDomain   [via $(if ($_.Dc) { $_.Dc } else { 'default DC' })]" })
        $ans = Show-ToolDialog(
            ("Enable a remote mailbox (and online archive) for $($plan.Count) account(s)?" + [Environment]::NewLine + [Environment]::NewLine +
             ($lines -join [Environment]::NewLine) +
             $(if ($skip.Count) { [Environment]::NewLine + [Environment]::NewLine + 'Not touched:' + [Environment]::NewLine + ($skip -join [Environment]::NewLine) } else { '' }) +
             [Environment]::NewLine + [Environment]::NewLine +
             'This is the same Enable-RemoteMailbox + archive that New User runs. The mailbox reaches Exchange Online on the next directory sync (~30 min); THEN press Retry Selected for these rows.'),
            'Enable Mailbox', 'YesNo', 'Question', 'Button2')
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $prev = $form.Cursor; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $out = @(); $ok = 0; $bad = 0
        try {
            foreach ($p in $plan) {
                try {
                    $res = Enable-ToolkitRemoteMailbox -Identity $p.Dn -Alias $p.Alias -DomainController $p.Dc
                    if ($res.Ok) {
                        $ok++
                        $out += "OK      $($p.Record.MB)  $($res.RoutingAddress)  archive=$($res.ArchiveOk)"
                        Write-CompletedLog -Action 'EnableRemoteMailbox' -Target (Get-LogTarget $p.Alias $p.Name) -Result 'Success' -Environment 'OnPrem' `
                            -Detail "Routing address $($res.RoutingAddress); archive = $($res.ArchiveOk) (Term Requests, Retry Queue)"
                    } else {
                        $bad++
                        $out += "FAILED  $($p.Record.MB)  $($res.Error)"
                        Write-FailedLog -Action 'EnableRemoteMailbox' -Target $p.Alias -Environment 'OnPrem' -Detail "Routing address $($res.RoutingAddress): $($res.Error)"
                    }
                } catch {
                    $bad++
                    $out += "FAILED  $($p.Record.MB)  $($_.Exception.Message)"
                    Write-FailedLog -Action 'EnableRemoteMailbox' -Target $p.Alias -ErrorRecord $_ -Environment 'OnPrem'
                }
            }
        } finally { $form.Cursor = $prev }
        foreach ($sk in $skip) { $out += "SKIP    $sk" }
        $null = Show-ToolDialog(
            (($out -join [Environment]::NewLine) + [Environment]::NewLine + [Environment]::NewLine +
             "$ok enabled, $bad failed." + $(if ($ok) { ' Wait for a directory sync (~30 minutes), then tick the same rows and press Retry Selected.' } else { '' })),
            "Enable Mailbox -- $ok enabled, $bad failed", 'OK', $(if ($bad) { 'Warning' } else { 'Information' }))
        $lblStatus.Text = "Enable Mailbox: $ok enabled, $bad failed."
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Enable Mailbox' }
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
        $ans = Show-ToolDialog(
            "Remove $($drop.Count) row(s) from the retry queue?`r`n`r`nThis does not undo anything -- it only stops tracking them.",
            'Remove from queue', 'YesNo', 'Question', 'Button2')
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $keep = @(Import-RetryQueue | Where-Object { $drop -notcontains "$($_.Type)|$($_.MB)" })
        Export-RetryQueue -Records $keep
        Update-RetryGrid
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Remove From Queue' }
})

# THROTTLED, because Activated fires on every focus change.
#
# Test-DisablePrerequisite calls Get-CapabilityStateMap -Refresh, which is four
# live probes against Exchange, AD, Exchange Online and SharePoint. Running
# that every time the window regains focus would stall on each alt-tab, with no
# visible cause -- the classic "why is this slow" that nobody traces back to a
# focus handler.
#
# Ten seconds is long enough that clicking between windows costs nothing, and
# short enough that connecting in the console and coming back shows the truth.
$script:LastConnCheck = [datetime]::MinValue

$form.Add_Activated({
    try {
        if (([datetime]::UtcNow - $script:LastConnCheck).TotalSeconds -lt 10) { return }
        $script:LastConnCheck = [datetime]::UtcNow

        $prev = $form.Cursor; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try { Update-ConnectionLabel -Check (Test-DisablePrerequisite) }
        finally { $form.Cursor = $prev }
    } catch {
        # Deliberately silent. This runs on focus, so a message box here would
        # be a popup the operator never asked for, appearing every time they
        # touch the window. The label simply keeps its last known state, and
        # the run-time check still blocks a run that cannot succeed.
    }
})

$btnClose.Add_Click({ $form.Close() })

# ---------------------------------------------------------------------------
# Only a CONFIGURED folder that is missing is worth a warning. The default is
# created when something is first queued, so warning about it would be noise on
# every clean install -- and a warning that fires when nothing is wrong is a
# warning people stop reading.
$script:RetryQueueState = Test-RetryQueueConfigured
if (-not $script:RetryQueueState.Exists -and -not $script:RetryQueueState.MayBeCreated) {
    Write-Warning "Retry queue folder not found: '$($script:RetryQueueState.Path)'. Set it in ToolManager -> Settings, or clear the box to use RetryQueue beside the toolkit."
}
Update-RetryGrid
Update-ConnectionLabel -Check (Test-DisablePrerequisite)

$script:Theme = Get-ConsoleTheme
if ($script:Theme) {
    Set-ConsoleTheme -Form $form -Theme $script:Theme
}
# NO LOGO HERE. The logo belongs to AdminConsole.ps1 alone -- it marks the
# front door, and a badge repeated on every window stops meaning anything.
# Add-FormLogo still exists in Common.ps1 and is still called by the console.
#
# Removing it also puts the '?' back where every other tool has it: the help
# button positions itself LEFT of a logo when one is present, so a window with
# a logo had its '?' inset and every other window did not.
#
# BottomLeft: TopLeft belongs to AdminConsole.ps1 itself.
# AFTER THE THEME. The ready state paints its background from $form.BackColor,
# and the theme sets that -- so the first call (before the theme) would leave a
# pale strip on a dark theme.
Update-ConnectionLabel -Check (Test-DisablePrerequisite)

Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

# Consistent '?' in the top-right of every tool. Added after the theme so
# it is not restyled into a normal button, and BringToFront'd inside the
# helper because WinForms draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Term Requests' | Out-Null

[void]$form.ShowDialog()

#endregion
