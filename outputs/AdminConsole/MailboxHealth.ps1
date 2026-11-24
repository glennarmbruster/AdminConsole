#Requires -Version 5.1
<#
    MailboxHealth.ps1
    Standalone Mailbox Health tool -- Layer 2 reference template for the
    launcher/tool-script architecture refactor.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# A read-only snapshot of one mailbox -- the first place to look when someone
# says "my email is doing something strange".
#
# Shows size and quota, the mailbox type, forwarding, delegates and
# permissions, auto-reply state, and the litigation hold and archive status
# in one view, so you are not running six commands to answer one question.
#
# Read-only. Nothing here changes anything -- it is for finding out what is
# true before you decide what to do.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

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

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. Resolve-ToolTarget is not used at all now:
# its job was to decide which DIRECTORY an address belongs to, and this tool
# reads Exchange Online, which does not need to be told.
#
# An empty address is fine. The window opens and asks.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Actions - mailbox health (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1556-1851)
function Get-MailboxHealthSnapshot {
    param([string]$Smtp)
    $health = [ordered]@{
        MailboxType                = $null
        RecipientTypeDetails       = $null
        LicenseState               = $null
        LastLogonTime              = $null
        ItemCount                  = $null
        MailboxSize                = $null
        ProhibitSendQuota          = $null
        ProhibitSendReceiveQuota   = $null
        ArchiveEnabled             = $null
        ArchiveSize                = $null
        RetentionPolicy            = $null
        AccountDisabled            = $null
        CloudExchangeGuid          = $null
        CloudArchiveGuid           = $null
        OnPremExchangeGuid         = $null
        OnPremArchiveGuid          = $null
        OnPremChecked              = $false
        OnPremFound                = $false
        ForwardingSmtpAddress      = $null
        ForwardingAddress          = $null
        DeliverToMailboxAndForward = $null
        LitigationHoldEnabled      = $null
        RetentionHoldEnabled       = $null
        InPlaceHoldsCount          = $null
        FullAccessDelegates        = @()
        SendAsDelegates            = @()
        RuleCount                  = $null
        SuspiciousRules            = @()
        OWAEnabled                 = $null
        ActiveSyncEnabled          = $null
        PopEnabled                 = $null
        ImapEnabled                = $null
        MAPIEnabled                = $null
        EwsEnabled                 = $null
        Errors                     = @()
    }
    try {
        Write-ToolStep 'Reading mailbox statistics ...'
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.LastLogonTime = $stats.LastLogonTime
        $health.ItemCount     = $stats.ItemCount
        $health.MailboxSize   = $stats.TotalItemSize
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (health)')
    }
    try {
        Write-ToolStep 'Reading mailbox settings ...'
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }

        # FREE -- this came back on the Get-Mailbox call that was already being
        # made. RecipientTypeDetails is kept alongside the simplified label
        # because the raw value distinguishes RoomMailbox from EquipmentMailbox
        # and SchedulingMailbox, and an operator chasing a booking problem needs
        # to know which one it is.
        $health.RecipientTypeDetails = "$($mbx.RecipientTypeDetails)"
        $health.MailboxType          = Get-SimpleMailboxType -RecipientTypeDetails $health.RecipientTypeDetails

        $health.ArchiveEnabled             = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)

        # FREE, same call. RetentionPolicy is the name of the policy assigned,
        # not its contents -- what it actually deletes or archives lives in the
        # policy's tags. Blank means NO policy, which on a tenant that relies on
        # one is a finding rather than a default.
        $health.RetentionPolicy            = "$($mbx.RetentionPolicy)"

        # These are half of the hybrid GUID check below.
        $health.CloudExchangeGuid          = "$($mbx.ExchangeGuid)"
        $health.CloudArchiveGuid           = "$($mbx.ArchiveGuid)"
        $health.AccountDisabled            = $mbx.AccountDisabled
        $health.ProhibitSendQuota          = $mbx.ProhibitSendQuota
        $health.ProhibitSendReceiveQuota   = $mbx.ProhibitSendReceiveQuota
        $health.ForwardingSmtpAddress      = $mbx.ForwardingSmtpAddress
        $health.ForwardingAddress          = $mbx.ForwardingAddress
        $health.DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
        $health.LitigationHoldEnabled      = $mbx.LitigationHoldEnabled
        $health.RetentionHoldEnabled       = $mbx.RetentionHoldEnabled
        $health.InPlaceHoldsCount          = if ($mbx.InPlaceHolds) { @($mbx.InPlaceHolds).Count } else { 0 }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (quota/forwarding/hold/account status)')
    }
    if ($health.ArchiveEnabled) {
        try {
            Write-ToolStep 'Reading archive statistics ...'
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $health.ArchiveSize = $astats.TotalItemSize
        } catch {
            $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (health)')
        }
    }
    try {
        Write-ToolStep 'Reading permissions ...'
        $fa = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.User -notlike 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false -and $_.Deny -eq $false -and ($_.AccessRights -contains 'FullAccess') })
        $health.FullAccessDelegates = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access delegates)')
    }
    try {
        $sa = @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.Trustee -notlike 'NT AUTHORITY\SELF' -and $_.AccessControlType -eq 'Allow' })
        $health.SendAsDelegates = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As delegates)')
    }
    try {
        $rules = Get-InboxRulesForMailbox -Smtp $Smtp
        if ($rules.Error) {
            $health.Errors += $rules.Error
        } else {
            $health.RuleCount = $rules.Rules.Count
            $health.SuspiciousRules = @(
                $rules.Rules | Where-Object {
                    $_.Enabled -and (
                        $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or
                        $_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage
                    )
                } | ForEach-Object {
                    $actions = @()
                    if ($_.ForwardTo -or $_.ForwardAsAttachmentTo) { $actions += 'Forward' }
                    if ($_.RedirectTo) { $actions += 'Redirect' }
                    if ($_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage) { $actions += 'Delete' }
                    "$($_.Name)  [$($actions -join ', ')]"
                }
            )
        }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule (rule count / suspicious rule scan)')
    }
    # LAST, AND IN ITS OWN TRY. This is the only field on the form that asks a
    # DIRECTORY rather than Exchange Online, so it is the only one that can fail
    # for a completely unrelated reason. Putting it at the end means an
    # unreachable domain controller costs the licence line and nothing else --
    # the whole Exchange-side snapshot is already built by the time it runs.
    # HYBRID GUID CHECK -- THE ON-PREMISES HALF.
    #
    # In a hybrid the on-premises remote-mailbox stub and the cloud mailbox must
    # agree on ExchangeGuid, and on ArchiveGuid if there is an archive. When they
    # drift, mail keeps flowing and nothing errors -- but the mailbox cannot be
    # moved, the archive stops being reachable, and a later migration fails with a
    # message that names neither GUID. It is invisible until it is expensive.
    #
    # Get-RemoteMailbox has NO Exchange Online counterpart, so there is no name to
    # collide with in a process holding both sessions (C1). It still goes through
    # Invoke-OnPremCommand so the intent is explicit in the code.
    #
    # SKIPPED ENTIRELY WITHOUT EMS. A cloud-only session cannot answer this, and a
    # blank pair of fields would read as "they do not match".
    if (Test-EmsAvailable) {
        $health.OnPremChecked = $true
        try {
            $rm = Invoke-OnPremCommand -Name 'Get-RemoteMailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
            if ($rm) {
                $health.OnPremFound        = $true
                $health.OnPremExchangeGuid = "$($rm.ExchangeGuid)"
                $health.OnPremArchiveGuid  = "$($rm.ArchiveGuid)"
            }
        } catch {
            $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RemoteMailbox (hybrid GUID check)')
        }
    }
    try {
        $health.LicenseState = Get-UserLicenseState -Smtp $Smtp
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UserLicenseState (licence group membership)')
    }
    try {
        Write-ToolStep 'Reading protocol settings ...'
        $cas = Invoke-EXOCommand -Name 'Get-CASMailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.OWAEnabled        = $cas.OWAEnabled
        $health.ActiveSyncEnabled = $cas.ActiveSyncEnabled
        $health.PopEnabled        = $cas.PopEnabled
        $health.ImapEnabled       = $cas.ImapEnabled
        $health.MAPIEnabled       = $cas.MAPIEnabled
        $health.EwsEnabled        = $cas.EwsEnabled
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-CASMailbox (protocol access)')
    }
    return [PSCustomObject]$health
}
# A READABLE RED ON BOTH A DARK AND A LIGHT THEME. Pure Color::Red is unreadable
# on the dark palettes and shouts on the light ones.
$script:HealthAlertColour = [System.Drawing.Color]::FromArgb(232, 78, 78)

function Set-HealthLineColour {
    <#
        Colour every line of a RichTextBox that contains $Contains.

        MUST RUN AFTER .Text IS ASSIGNED. Assigning .Text discards every colour
        run in the control, so colouring first and filling second silently does
        nothing. The same is true of Set-ConsoleTheme, which sets .ForeColor --
        which is why the first render happens after the theme, not before it.
    #>
    param($Box, [string]$Contains, $Colour)
    for ($n = 0; $n -lt $Box.Lines.Count; $n++) {
        if ($Box.Lines[$n] -like "*$Contains*") {
            $Box.Select($Box.GetFirstCharIndexFromLine($n), $Box.Lines[$n].Length)
            $Box.SelectionColor = $Colour
        }
    }
    $Box.Select(0, 0)
    $Box.SelectionColor = $Box.ForeColor
}

function Show-MailboxHealthForm {
    # NO -Smtp PARAMETER ANY MORE. The address can change while the window is
    # open, so it lives in $script:Target rather than in a parameter that was
    # fixed the moment the form was built.
    $healthForm = New-Object System.Windows.Forms.Form
    $healthForm.Text = 'Mailbox Health'
    # TALL ON PURPOSE. A mailbox with holds, delegates, suspicious rules and a
    # hybrid GUID block runs to about 65 lines, and 710px of Consolas 9 showed
    # roughly two thirds of it. Set-FormScreenCorner clamps this to WorkingArea
    # -- not Bounds (B4) -- so a smaller screen gets the tallest window that
    # fits rather than one hanging off the bottom.
    $healthForm.ClientSize = New-Object System.Drawing.Size(624, 960)
    $healthForm.StartPosition = 'Manual'
    $healthForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    # RICHTEXTBOX, NOT TEXTBOX. The hybrid GUID lines go red when they disagree,
    # and a plain TextBox has one colour for all of its text.
    $txtHealth = New-Object System.Windows.Forms.RichTextBox
    $txtHealth.Multiline = $true
    $txtHealth.ReadOnly = $true
    $txtHealth.DetectUrls = $false
    $txtHealth.ScrollBars = "Vertical"
    $txtHealth.Location = New-Object System.Drawing.Point(12, 12)
    $txtHealth.Size = New-Object System.Drawing.Size(600, 900)
    $txtHealth.Anchor = 'Top,Left,Right,Bottom'
    $txtHealth.Font = New-Object System.Drawing.Font("Consolas", 9)
    $healthForm.Controls.Add($txtHealth)
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 920)
    $btnRefresh.Anchor = 'Left,Bottom'
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnRefresh)
    $btnCopyHealth = New-Object System.Windows.Forms.Button
    $btnCopyHealth.Text = "Copy"
    $btnCopyHealth.Location = New-Object System.Drawing.Point(120, 920)
    $btnCopyHealth.Anchor = 'Left,Bottom'
    $btnCopyHealth.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnCopyHealth)
    $btnHealthClose = New-Object System.Windows.Forms.Button
    $btnHealthClose.Text = "Close"
    $btnHealthClose.Location = New-Object System.Drawing.Point(228, 920)
    $btnHealthClose.Anchor = 'Left,Bottom'
    $btnHealthClose.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnHealthClose)
    $copyResetTimer = New-Object System.Windows.Forms.Timer
    $copyResetTimer.Interval = 1200
    $copyResetTimer.Add_Tick({
        $btnCopyHealth.Text = "Copy"
        $copyResetTimer.Stop()
    })
    function Update-HealthFromCurrent {
        if (-not "$script:Target".Trim()) {
            $txtHealth.Text = 'Enter a mailbox address above and press Look Up.'
            $healthForm.Text = 'Mailbox Health'
            return
        }
        $healthForm.Text = "Mailbox Health: $script:Target"
        $btnRefresh.Enabled = $false
        $txtHealth.Text = "Loading mailbox health..."
        $healthForm.Refresh()
        try {
            $h = Get-MailboxHealthSnapshot -Smtp $script:Target
            $sb = New-Object System.Text.StringBuilder
            $issues = @()

            # A GUID PAIR ONLY DISAGREES IF BOTH SIDES ANSWERED. No EMS, or no
            # on-premises stub, is "not checked" -- never "mismatch".
            $exOn = "$($h.OnPremExchangeGuid)"; $exCl = "$($h.CloudExchangeGuid)"
            $arOn = "$($h.OnPremArchiveGuid)";  $arCl = "$($h.CloudArchiveGuid)"
            $exMismatch = ($h.OnPremFound -and $exOn -and $exCl -and $exOn -ne $exCl)
            $arMismatch = ($h.OnPremFound -and $arOn -ne $arCl)
            if ($exMismatch) {
                $issues += ("HYBRID MISMATCH: on-premises ExchangeGuid does not match Exchange Online. " +
                            "Mail still flows, but the mailbox cannot be moved and a future migration will fail. " +
                            "See the Hybrid GUIDs section below.")
            }
            if ($arMismatch) {
                $issues += ("HYBRID MISMATCH: on-premises ArchiveGuid does not match Exchange Online. " +
                            "The archive is usually unreachable from Outlook when this is wrong.")
            }
            if ($h.AccountDisabled -eq $true) { $issues += "Account sign-in is disabled (AccountDisabled=True) -- unverified as a Get-Mailbox field, confirm before relying on it." }
            if ($h.ForwardingSmtpAddress -or $h.ForwardingAddress) {
                $fwdTarget = if ($h.ForwardingSmtpAddress) { $h.ForwardingSmtpAddress } else { $h.ForwardingAddress }
                $fwdMode = if ($h.DeliverToMailboxAndForward) { "copy kept in mailbox AND forwarded" } else { "forwarded ONLY -- nothing kept in this mailbox" }
                $issues += "Mail forwarding is configured to '$fwdTarget' ($fwdMode)."
            }
            # Only raised when the licence question was actually ANSWERED.
            # Get-UserLicenseState returns Known = $false for a cloud-only
            # tenant or an unreachable DC, and "unknown" must not be reported
            # as "wrong" -- an issue list that cries wolf stops being read.
            if ($h.LicenseState -and $h.LicenseState.Known) {
                if ($h.MailboxType -eq 'Shared' -and $h.LicenseState.Licensed) {
                    $issues += ("Shared mailbox IS licensed ($($h.LicenseState.Groups -join ', ')). " +
                                "Shared mailboxes under 50 GB with no archive do not need one -- check before removing, " +
                                "since a licence here is sometimes deliberate.")
                }
                if ($h.MailboxType -eq 'User' -and -not $h.LicenseState.Licensed) {
                    $issues += ("User mailbox is NOT in any licence group. If the account is still active this will " +
                                "stop mail flow when the grace period ends.")
                }
            }
            if ($h.LitigationHoldEnabled -eq $true) { $issues += "Litigation Hold is enabled." }
            if ($h.RetentionHoldEnabled -eq $true) { $issues += "Retention Hold is enabled." }
            if ($h.InPlaceHoldsCount -gt 0) { $issues += "$($h.InPlaceHoldsCount) In-Place/eDiscovery hold(s) applied." }
            if ($h.SuspiciousRules.Count -gt 0) { $issues += "$($h.SuspiciousRules.Count) enabled Inbox rule(s) forward, redirect, or delete mail -- see 'Suspicious Inbox Rules' below." }
            [void]$sb.AppendLine("Mailbox                : $script:Target")
            [void]$sb.AppendLine("=== Potential Issues (review before escalating) ===")
            if ($issues.Count -eq 0) { [void]$sb.AppendLine("  None of the checks below flagged anything.") }
            else { $issues | ForEach-Object { [void]$sb.AppendLine("  - $_") } }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Basics ---")
            # FIRST IN THE SECTION, above the quotas and the logon date.
            #
            # Almost every question that reaches this form is answered
            # differently for a shared or resource mailbox than for a user's --
            # "why can't they sign in", "why is there no licence", "why did the
            # quota not go up". Reading that at the bottom is reading it after
            # the wrong assumption has already been made.
            [void]$sb.AppendLine("Mailbox Type            : $($h.MailboxType)  [$($h.RecipientTypeDetails)]")
            if ($h.LicenseState) {
                [void]$sb.AppendLine("Licensed                : $(Format-UserLicenseState -State $h.LicenseState)")
            } else {
                [void]$sb.AppendLine("Licensed                : Not checked -- see field errors below.")
            }
            [void]$sb.AppendLine("Last Logon             : $($h.LastLogonTime)")
            [void]$sb.AppendLine("Item Count              : $($h.ItemCount)")
            [void]$sb.AppendLine("Mailbox Size            : $($h.MailboxSize)")
            [void]$sb.AppendLine("ProhibitSendQuota       : $($h.ProhibitSendQuota)  (compare to Mailbox Size above)")
            [void]$sb.AppendLine("ProhibitSendReceiveQuota: $($h.ProhibitSendReceiveQuota)")
            [void]$sb.AppendLine("Archive Enabled         : $($h.ArchiveEnabled)")
            if ($h.ArchiveEnabled) { [void]$sb.AppendLine("Archive Size            : $($h.ArchiveSize)") }
            [void]$sb.AppendLine("Account Disabled        : $($h.AccountDisabled)  (unverified Get-Mailbox field -- see script header)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Hybrid GUIDs (on-premises stub vs cloud mailbox) ---")
            if (-not $h.OnPremChecked) {
                [void]$sb.AppendLine("  Not checked -- no on-premises Exchange session in this process.")
            } elseif (-not $h.OnPremFound) {
                [void]$sb.AppendLine("  No on-premises remote mailbox found for this address.")
                [void]$sb.AppendLine("  Normal for a cloud-only mailbox; unexpected in a hybrid one.")
            } else {
                [void]$sb.AppendLine("On-prem Exchange GUID   : $exOn")
                [void]$sb.AppendLine("On-prem Archive GUID    : $arOn")
                [void]$sb.AppendLine("ExO     Exchange GUID   : $exCl")
                [void]$sb.AppendLine("ExO     Archive GUID    : $arCl")
                if (-not $exMismatch -and -not $arMismatch) {
                    [void]$sb.AppendLine("  Both pairs match.")
                }
            }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Forwarding & Holds ---")
            [void]$sb.AppendLine("Forwarding SMTP Address : $($h.ForwardingSmtpAddress)")
            [void]$sb.AppendLine("Forwarding Address      : $($h.ForwardingAddress)")
            [void]$sb.AppendLine("Deliver And Forward     : $($h.DeliverToMailboxAndForward)")
            [void]$sb.AppendLine("Retention Policy        : $(if ("$($h.RetentionPolicy)".Trim()) { $h.RetentionPolicy } else { '(none assigned)' })")
            [void]$sb.AppendLine("Litigation Hold         : $($h.LitigationHoldEnabled)")
            [void]$sb.AppendLine("Retention Hold          : $($h.RetentionHoldEnabled)")
            [void]$sb.AppendLine("In-Place Holds          : $($h.InPlaceHoldsCount)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Delegates ---")
            [void]$sb.AppendLine("Full Access grantees ($($h.FullAccessDelegates.Count)):")
            if ($h.FullAccessDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.FullAccessDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("Send As grantees ($($h.SendAsDelegates.Count)):")
            if ($h.SendAsDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.SendAsDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("  (Distribution List memberships are not scanned here -- use the DL Memberships tool, since a full-tenant scan is too slow to run on every Health click.)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Inbox Rules ---")
            [void]$sb.AppendLine("Inbox Rule Count        : $($h.RuleCount)")
            [void]$sb.AppendLine("Suspicious Inbox Rules (enabled, forward/redirect/delete):")
            if ($h.SuspiciousRules.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.SuspiciousRules | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Protocol Access (enabled/disabled -- not traffic volume) ---")
            [void]$sb.AppendLine("  OWA        : $($h.OWAEnabled)")
            [void]$sb.AppendLine("  ActiveSync : $($h.ActiveSyncEnabled)")
            [void]$sb.AppendLine("  POP        : $($h.PopEnabled)")
            [void]$sb.AppendLine("  IMAP       : $($h.ImapEnabled)")
            [void]$sb.AppendLine("  MAPI       : $($h.MAPIEnabled)")
            [void]$sb.AppendLine("  EWS        : $($h.EwsEnabled)")
            if ($h.Errors.Count -gt 0) {
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("--- Field errors (partial data above is still valid) ---")
                $h.Errors | ForEach-Object { [void]$sb.AppendLine($_) }
            }
            $txtHealth.Text = $sb.ToString()

            # AFTER .Text, ALWAYS. See Set-HealthLineColour. Only the ON-PREM line
            # of a disagreeing pair is reddened -- the cloud value is the one that
            # is right, so colouring both would say "one of these is wrong" and
            # leave the operator to work out which.
            if ($exMismatch) { Set-HealthLineColour -Box $txtHealth -Contains 'On-prem Exchange GUID' -Colour $script:HealthAlertColour }
            if ($arMismatch) { Set-HealthLineColour -Box $txtHealth -Contains 'On-prem Archive GUID'  -Colour $script:HealthAlertColour }
        } catch {
            $txtHealth.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Health'
        } finally {
            $btnRefresh.Enabled = $true
        }
    }
    $btnRefresh.Add_Click({ Update-HealthFromCurrent })
    $btnCopyHealth.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($txtHealth.Text)
            $btnCopyHealth.Text = "Copied!"
            $copyResetTimer.Stop()
            $copyResetTimer.Start()
        } catch {
            Show-ToolDialog(
                "Copy to clipboard failed: $($_.Exception.Message)",
                "Copy", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
    })
    $btnHealthClose.Add_Click({ $healthForm.Close() })
    # BEFORE Add-ToolHelpButton and before the theme: it shifts every existing
    # control down, so anything added after it would be left at the old offset.
    $bar = Add-ToolTargetBar -Form $healthForm -InitialValue $script:Target

    # WIRED HERE, NOT PASSED INTO THE HELPER. Update-HealthFromCurrent is nested
    # in this function, so the handler has to be created in this scope to see it
    # -- exactly like $btnRefresh above.
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-HealthFromCurrent
    })

    $healthForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $healthForm
    Set-FormScreenCorner -Form $healthForm -Corner 'BottomLeft'
    # A VIEWER, SO IT RESIZES. Even at 960 the longest mailboxes scroll a
    # little, and this is the tool most likely to be read next to a ticket.
    # MinimumSize becomes the designed size, so it can only grow -- nobody has
    # ever wanted a smaller Mailbox Health. Called after the size is final,
    # per the note on Set-FormResizable.
    Set-FormResizable -Form $healthForm
        # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled into a normal button, and BringToFront'd inside the
    # helper because WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $healthForm -ScriptPath $PSCommandPath -Title 'Mailbox Health' | Out-Null

    # FIRST RENDER GOES LAST, AFTER Set-ConsoleTheme. The theme assigns
    # .ForeColor, which resets every colour run in a RichTextBox -- so rendering
    # before it would drop the red on the very first lookup and show it only
    # after a Refresh. The form is not visible until ShowDialog, so nothing
    # flickers.
    Update-HealthFromCurrent

    [void]$healthForm.ShowDialog()
}
#endregion

Show-MailboxHealthForm
