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

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'Mailbox Health -- Enter Target'
if (-not $target) { return }

#region Actions - mailbox health (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1556-1851)
function Get-MailboxHealthSnapshot {
    param([string]$Smtp)
    $health = [ordered]@{
        LastLogonTime              = $null
        ItemCount                  = $null
        MailboxSize                = $null
        ProhibitSendQuota          = $null
        ProhibitSendReceiveQuota   = $null
        ArchiveEnabled             = $null
        ArchiveSize                = $null
        AccountDisabled            = $null
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
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.LastLogonTime = $stats.LastLogonTime
        $health.ItemCount     = $stats.ItemCount
        $health.MailboxSize   = $stats.TotalItemSize
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (health)')
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.ArchiveEnabled             = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
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
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $health.ArchiveSize = $astats.TotalItemSize
        } catch {
            $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (health)')
        }
    }
    try {
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
    try {
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
function Show-MailboxHealthForm {
    param([string]$Smtp)
    $healthForm = New-Object System.Windows.Forms.Form
    $healthForm.Text = "Mailbox Health: $Smtp"
    $healthForm.Size = New-Object System.Drawing.Size(640, 660)
    $healthForm.StartPosition = 'Manual'
    $healthForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $txtHealth = New-Object System.Windows.Forms.TextBox
    $txtHealth.Multiline = $true
    $txtHealth.ReadOnly = $true
    $txtHealth.ScrollBars = "Vertical"
    $txtHealth.Location = New-Object System.Drawing.Point(12, 12)
    $txtHealth.Size = New-Object System.Drawing.Size(600, 560)
    $txtHealth.Font = New-Object System.Drawing.Font("Consolas", 9)
    $healthForm.Controls.Add($txtHealth)
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 580)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnRefresh)
    $btnCopyHealth = New-Object System.Windows.Forms.Button
    $btnCopyHealth.Text = "Copy"
    $btnCopyHealth.Location = New-Object System.Drawing.Point(120, 580)
    $btnCopyHealth.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnCopyHealth)
    $btnHealthClose = New-Object System.Windows.Forms.Button
    $btnHealthClose.Text = "Close"
    $btnHealthClose.Location = New-Object System.Drawing.Point(228, 580)
    $btnHealthClose.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnHealthClose)
    $copyResetTimer = New-Object System.Windows.Forms.Timer
    $copyResetTimer.Interval = 1200
    $copyResetTimer.Add_Tick({
        $btnCopyHealth.Text = "Copy"
        $copyResetTimer.Stop()
    })
    function Update-HealthFromCurrent {
        $btnRefresh.Enabled = $false
        $txtHealth.Text = "Loading mailbox health..."
        $healthForm.Refresh()
        try {
            $h = Get-MailboxHealthSnapshot -Smtp $Smtp
            $sb = New-Object System.Text.StringBuilder
            $issues = @()
            if ($h.AccountDisabled -eq $true) { $issues += "Account sign-in is disabled (AccountDisabled=True) -- unverified as a Get-Mailbox field, confirm before relying on it." }
            if ($h.ForwardingSmtpAddress -or $h.ForwardingAddress) {
                $fwdTarget = if ($h.ForwardingSmtpAddress) { $h.ForwardingSmtpAddress } else { $h.ForwardingAddress }
                $fwdMode = if ($h.DeliverToMailboxAndForward) { "copy kept in mailbox AND forwarded" } else { "forwarded ONLY -- nothing kept in this mailbox" }
                $issues += "Mail forwarding is configured to '$fwdTarget' ($fwdMode)."
            }
            if ($h.LitigationHoldEnabled -eq $true) { $issues += "Litigation Hold is enabled." }
            if ($h.RetentionHoldEnabled -eq $true) { $issues += "Retention Hold is enabled." }
            if ($h.InPlaceHoldsCount -gt 0) { $issues += "$($h.InPlaceHoldsCount) In-Place/eDiscovery hold(s) applied." }
            if ($h.SuspiciousRules.Count -gt 0) { $issues += "$($h.SuspiciousRules.Count) enabled Inbox rule(s) forward, redirect, or delete mail -- see 'Suspicious Inbox Rules' below." }
            [void]$sb.AppendLine("Mailbox                : $Smtp")
            [void]$sb.AppendLine("=== Potential Issues (review before escalating) ===")
            if ($issues.Count -eq 0) { [void]$sb.AppendLine("  None of the checks below flagged anything.") }
            else { $issues | ForEach-Object { [void]$sb.AppendLine("  - $_") } }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Basics ---")
            [void]$sb.AppendLine("Last Logon             : $($h.LastLogonTime)")
            [void]$sb.AppendLine("Item Count              : $($h.ItemCount)")
            [void]$sb.AppendLine("Mailbox Size            : $($h.MailboxSize)")
            [void]$sb.AppendLine("ProhibitSendQuota       : $($h.ProhibitSendQuota)  (compare to Mailbox Size above)")
            [void]$sb.AppendLine("ProhibitSendReceiveQuota: $($h.ProhibitSendReceiveQuota)")
            [void]$sb.AppendLine("Archive Enabled         : $($h.ArchiveEnabled)")
            if ($h.ArchiveEnabled) { [void]$sb.AppendLine("Archive Size            : $($h.ArchiveSize)") }
            [void]$sb.AppendLine("Account Disabled        : $($h.AccountDisabled)  (unverified Get-Mailbox field -- see script header)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Forwarding & Holds ---")
            [void]$sb.AppendLine("Forwarding SMTP Address : $($h.ForwardingSmtpAddress)")
            [void]$sb.AppendLine("Forwarding Address      : $($h.ForwardingAddress)")
            [void]$sb.AppendLine("Deliver And Forward     : $($h.DeliverToMailboxAndForward)")
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
        } catch {
            $txtHealth.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Health'
        } finally {
            $btnRefresh.Enabled = $true
        }
    }
    Update-HealthFromCurrent
    $btnRefresh.Add_Click({ Update-HealthFromCurrent })
    $btnCopyHealth.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($txtHealth.Text)
            $btnCopyHealth.Text = "Copied!"
            $copyResetTimer.Stop()
            $copyResetTimer.Start()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Copy to clipboard failed: $($_.Exception.Message)",
                "Copy", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
    })
    $btnHealthClose.Add_Click({ $healthForm.Close() })
    $healthForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $healthForm
    Set-FormScreenCorner -Form $healthForm -Corner 'BottomLeft'
        # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled into a normal button, and BringToFront'd inside the
    # helper because WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $healthForm -ScriptPath $PSCommandPath -Title 'Mailbox Health' | Out-Null

    [void]$healthForm.ShowDialog()
}
#endregion

Show-MailboxHealthForm -Smtp $target.Smtp
