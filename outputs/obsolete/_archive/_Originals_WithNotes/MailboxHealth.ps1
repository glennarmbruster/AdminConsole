#Requires -Version 5.1
<#
.SYNOPSIS
    Standalone Mailbox Health tool -- Layer 2 reference template for the
    launcher/tool-script architecture refactor.
.DESCRIPTION
    Fully functional on its own:
      .\MailboxHealth.ps1
    prompts for mailbox / Environment / Domain Controller via
    Common.ps1's Show-TargetPromptDialog, then opens the Health window.

    Fully functional from AdminConsole.ps1 too:
      .\MailboxHealth.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Health window for
    that mailbox.

    Get-MailboxHealthSnapshot and Show-MailboxHealthForm below are MOVED,
    UNCHANGED, from the original monolithic Mailbox-Group-Inspector.ps1
    (original lines 1556-1851). The only edit anywhere in this file that
    touches original tool behavior is mechanical: Show-MailboxHealthForm's
    closing call changed from

        [void]$healthForm.ShowDialog($form)

    to

        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin

        # Console, which opens top-left. Helper is in Common.ps1.

        Set-FormScreenCorner -Form $healthForm -Corner 'BottomLeft'

        [void]$healthForm.ShowDialog()

    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. Every other line of both functions,
    including all comments, is byte-for-byte the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 --
              see REFACTOR-DESIGN.md, "Phase 2: extract MailboxHealth.ps1")
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'Mailbox Health -- Enter Target'
if (-not $target) { return }   # operator cancelled the prompt

#region Actions - mailbox health (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1556-1851)
function Get-MailboxHealthSnapshot {
    # Read-only composite view. Reuses Get-InboxRulesForMailbox (now in
    # Common.ps1 -- see that file's header for why) for the rule list/count
    # rather than querying it twice.
    #
    # Interpretation / verification notes (documented since these terms are
    # not single, unambiguous Exchange fields, and per standards Section 2
    # every cmdlet/property used for a new capability is checked against
    # Microsoft Learn before delivery):
    #  - "Protocol usage" is shown as Get-CASMailbox protocol ENABLEMENT
    #    flags (OWA/ActiveSync/POP/IMAP/MAPI/EWS), not actual traffic volume.
    #  - "Delegates" now lists Full Access and Send As grantees by name
    #    (each from its own cmdlet), not just a combined count. Does NOT
    #    include Outlook's separate calendar/folder delegate list.
    #  - DL memberships are NOT part of this snapshot -- see DLMemberships.ps1
    #    instead. That scan iterates every DL in the tenant and can take a
    #    while; running it automatically inside Health made the whole tool
    #    appear to hang (2026-07-09 field report), so it's a separate
    #    on-demand tool with its own progress display.
    #  - Quota (ProhibitSendQuota, ProhibitSendReceiveQuota), Forwarding
    #    (ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward),
    #    Litigation Hold (LitigationHoldEnabled, InPlaceHolds), and Retention
    #    Hold (RetentionHoldEnabled) are all confirmed Get-Mailbox output
    #    properties per Microsoft Learn's "Property sets in Exchange Online
    #    PowerShell module cmdlets" (Quota / Delivery / Hold / Retention
    #    property sets respectively). Quota values are shown as raw
    #    ByteQuantifiedSize text next to Mailbox Size for manual comparison
    #    rather than computed as a percentage -- the ByteQuantifiedSize/
    #    Unlimited<T> numeric conversion API was not verified against
    #    Microsoft Learn, so this script does not guess at it.
    #  - AccountDisabled: Microsoft Learn confirms -AccountDisabled as a
    #    Set-Mailbox parameter but this script did NOT find Microsoft Learn
    #    confirmation that Get-Mailbox returns an AccountDisabled OUTPUT
    #    property (Microsoft's own Get-Mailbox page recommends Get-EXOMailbox
    #    instead, whose documented property sets don't list it either).
    #    Flagged here as UNVERIFIED per standards Section 2 ("if
    #    verification is unavailable, clearly say so before delivery") --
    #    confirm this field is meaningful on first use; if the property
    #    doesn't exist on the returned object, PowerShell returns $null here
    #    rather than throwing, so a wrong/missing value fails quietly. Treat
    #    a blank value as "unknown," not as "account is enabled."
    #  - "Suspicious Inbox Rules" flags ENABLED rules with a forwarding,
    #    redirect, or delete action (ForwardTo/ForwardAsAttachmentTo/
    #    RedirectTo/DeleteMessage/PermanentDelete/SoftDeleteMessage -- all
    #    confirmed Get-InboxRule/Set-InboxRule property names on Microsoft
    #    Learn). This is a flag for a human to review as a first-check item
    #    on compromised-account or missing-mail tickets, not a judgment that
    #    the rule is malicious.
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
    # DL membership is intentionally NOT scanned here -- see DLMemberships.ps1.
    return [PSCustomObject]$health
}
function Show-MailboxHealthForm {
    param([string]$Smtp)
    $healthForm = New-Object System.Windows.Forms.Form
    $healthForm.Text = "Mailbox Health: $Smtp"
    $healthForm.Size = New-Object System.Drawing.Size(640, 660)
    $healthForm.StartPosition = "CenterScreen"
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
    # One-shot timer so "Copied!" reverts to "Copy" without needing a second
    # click -- local to this form (closure-captured), same pattern as every
    # other control reference in this function.
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
            # --- Potential Issues summary first -- the point of this button is
            # "what do I need to know before I escalate this ticket", so lead
            # with anything that looks like a likely cause rather than making
            # the reader hunt for it further down.
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
    # CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the
    # monolith's shared parent window. This script has no parent window, so
    # it shows itself as the top-level dialog instead. Nothing else in this
    # function changed.
    [void]$healthForm.ShowDialog()
}
#endregion

Show-MailboxHealthForm -Smtp $target.Smtp
