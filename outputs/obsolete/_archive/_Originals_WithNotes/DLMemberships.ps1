#Requires -Version 5.1
<#
.SYNOPSIS
    Standalone DL Memberships tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\DLMemberships.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the DL Memberships window.

    Fully functional from AdminConsole.ps1 too:
        .\DLMemberships.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the DL Memberships window
    for that mailbox.

    Get-MailboxDLMemberships and Show-DLMembershipsForm below are MOVED,
    UNCHANGED, from the original monolithic Mailbox-Group-Inspector.ps1
    (original lines 1458-1555 and 1852-1954, region "Actions - mailbox
    health"), including every revision-history comment. The only edit
    anywhere in this file that touches original tool behavior is mechanical:
    Show-DLMembershipsForm's closing call changed from
        [void]$dlForm.ShowDialog($form)
    to
        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
        # Console, which opens top-left. Helper is in Common.ps1.
        Set-FormScreenCorner -Form $dlForm -Corner 'BottomLeft'
        [void]$dlForm.ShowDialog()
    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. Every other line, including all comments, is
    byte-for-byte the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'DL Memberships -- Enter Target'
if (-not $target) { return }   # operator cancelled the prompt

#region Actions - mailbox health / DL memberships (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1458-1555, 1852-1954)
function Get-MailboxDLMemberships {
    # REWRITTEN 2026-07-09c (field report: the original version iterated
    # every DL in the tenant with Get-DistributionGroupMember and froze the
    # UI -- see the two 2026-07-09 revision notes in the header. The
    # operator then asked whether the EAC "Group membership" pane does the
    # same full-tenant scan; it does not.)
    #
    # EAC's Group Membership pane, and Set-Mailbox et al under the hood, get
    # this from Active Directory's own "memberOf" attribute on the user
    # object -- a backlink that AD itself maintains automatically every time
    # a group's "member" attribute changes, so reading it is a single
    # indexed object lookup, not a search across every group. This function
    # now does the same thing via Get-ADUser -Properties memberOf (both the
    # -Properties parameter and its behavior confirmed against Microsoft
    # Learn), matched to the mailbox via its "mail" attribute the same way
    # Resolve-Target already matches on-prem GROUPS by mail elsewhere in
    # this script. Cost is now one Get-ADUser call plus one Get-ADGroup call
    # per group the user actually belongs to (typically a handful to a few
    # dozen) -- bounded by the user's real membership count, not by how many
    # DLs exist in the tenant.
    #
    # Two deliberate scope limits, both consistent with how this tool
    # already treats on-prem groups elsewhere:
    #  - Direct membership only, not nested/recursive. memberOf lists groups
    #    the user is a DIRECT member of; if that group is itself nested
    #    inside another group, the outer group is not included. This is what
    #    EAC's Group Membership pane and the on-prem AD "Member Of" tab both
    #    show. Get-ADPrincipalGroupMembership resolves nested membership too
    #    but requires a Global Catalog and behaves differently in
    #    multi-domain forests (per Microsoft Learn) -- not used here to keep
    #    behavior simple and predictable across the Franchise/HQ environment
    #    split this tool already has.
    #  - On-prem-synced Distribution / Mail-Enabled Security groups only,
    #    via the ActiveDirectory module -- Microsoft 365 Group membership is
    #    a separate object type handled by this tool's own Group lookup
    #    path (Get-UnifiedGroupLinks), not included here.
    #  - A mailbox with no matching on-prem AD user (cloud-only mailbox,
    #    never synced) has no memberOf to read -- $result.NoADAccount is set
    #    rather than silently returning an empty list, so that case isn't
    #    mistaken for "this user is in zero groups."
    #
    # -ProgressAction, if supplied, is invoked once per group-DN being
    # resolved as & $ProgressAction $currentIndex $totalCount $groupDN.
    # -ShouldCancel, if supplied, is checked after each one; returning
    # $true stops early and sets $result.Cancelled, keeping whatever partial
    # results were already resolved. Both are vestigial now that this is a
    # bounded, typically-fast operation, but kept since a handful of memberOf
    # entries can still point at cross-domain groups that are slow to
    # resolve individually.
    param(
        [string]$Smtp,
        [string]$ADServer,
        [scriptblock]$ProgressAction,
        [scriptblock]$ShouldCancel
    )
    $result = [ordered]@{
        Memberships = @()
        ScanErrors  = @()
        FatalError  = $null
        Cancelled   = $false
        NoADAccount = $false
    }
    try {
        $adUsers = @(Get-ADUser -Filter "mail -eq '$Smtp'" -Properties memberOf -Server $ADServer -ErrorAction Stop)
    } catch {
        $result.FatalError = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADUser (DL membership lookup)')
        return [PSCustomObject]$result
    }
    if ($adUsers.Count -eq 0) {
        $result.NoADAccount = $true
        return [PSCustomObject]$result
    }
    if ($adUsers.Count -gt 1) {
        $result.FatalError = "$($adUsers.Count) on-prem AD users matched mail '$Smtp'. Expected 0 or 1 -- resolve the duplicate mail attribute before using this feature."
        return [PSCustomObject]$result
    }
    $dns = @($adUsers[0].memberOf)
    $found = @()
    $i = 0
    foreach ($dn in $dns) {
        $i++
        if ($ProgressAction) {
            try { & $ProgressAction $i $dns.Count $dn } catch { }
        }
        try {
            $g = Get-ADGroup -Identity $dn -Server $ADServer -ErrorAction Stop
            $found += $g.Name
        } catch {
            $result.ScanErrors += "$($dn): $($_.Exception.Message)"
        }
        if ($ShouldCancel -and (& $ShouldCancel)) {
            $result.Cancelled = $true
            break
        }
    }
    $result.Memberships = $found
    return [PSCustomObject]$result
}
function Show-DLMembershipsForm {
    # Split out of the Health panel on 2026-07-09, then REWRITTEN 2026-07-09c
    # to read AD's memberOf backlink instead of scanning every DL -- see the
    # Get-MailboxDLMemberships header comment and the header .NOTES revision
    # entries for the full story. Now a single Get-ADUser call plus one
    # Get-ADGroup call per actual membership (typically fast), not a
    # full-tenant scan. Scan/Cancel/progress-label plumbing is kept since a
    # cross-domain group DN can still occasionally be slow to resolve, but
    # this should complete quickly in the common case.
    param([string]$Smtp, [string]$ADServer)
    $dlForm = New-Object System.Windows.Forms.Form
    $dlForm.Text = "DL Memberships: $Smtp"
    $dlForm.Size = New-Object System.Drawing.Size(560, 520)
    $dlForm.StartPosition = "CenterScreen"
    $dlForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads this mailbox's on-prem AD account (matched by mail attribute) and its memberOf attribute -- the same backlink EAC's Group Membership pane uses. Direct memberships only, classic DLs / mail-enabled security groups only (not Microsoft 365 Groups)."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(520, 46)
    $dlForm.Controls.Add($lblNote)
    $lblProgress = New-Object System.Windows.Forms.Label
    $lblProgress.Text = "Not started."
    $lblProgress.Location = New-Object System.Drawing.Point(12, 62)
    $lblProgress.Size = New-Object System.Drawing.Size(520, 20)
    $lblProgress.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $dlForm.Controls.Add($lblProgress)
    $txtDL = New-Object System.Windows.Forms.TextBox
    $txtDL.Multiline = $true
    $txtDL.ReadOnly = $true
    $txtDL.ScrollBars = "Vertical"
    $txtDL.Location = New-Object System.Drawing.Point(12, 88)
    $txtDL.Size = New-Object System.Drawing.Size(520, 340)
    $txtDL.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlForm.Controls.Add($txtDL)
    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "Scan"
    $btnScan.Location = New-Object System.Drawing.Point(12, 436)
    $btnScan.Size = New-Object System.Drawing.Size(100, 28)
    $dlForm.Controls.Add($btnScan)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(120, 436)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 28)
    $btnCancel.Enabled = $false
    $dlForm.Controls.Add($btnCancel)
    $btnDLClose = New-Object System.Windows.Forms.Button
    $btnDLClose.Text = "Close"
    $btnDLClose.Location = New-Object System.Drawing.Point(432, 436)
    $btnDLClose.Size = New-Object System.Drawing.Size(100, 28)
    $dlForm.Controls.Add($btnDLClose)
    $script:DLScanCancelRequested = $false
    $btnScan.Add_Click({
        $btnScan.Enabled = $false
        $btnCancel.Enabled = $true
        $txtDL.Text = ""
        $script:DLScanCancelRequested = $false
        $lblProgress.Text = "Looking up on-prem AD account..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $progressCB = {
                param($i, $total, $dn)
                $lblProgress.Text = "Resolving group $i of $total`: $dn"
                [System.Windows.Forms.Application]::DoEvents()
            }
            $cancelCB = { return $script:DLScanCancelRequested }
            $result = Get-MailboxDLMemberships -Smtp $Smtp -ADServer $ADServer -ProgressAction $progressCB -ShouldCancel $cancelCB
            if ($result.FatalError) {
                $lblProgress.Text = "Failed."
                $txtDL.Text = $result.FatalError
            } elseif ($result.NoADAccount) {
                $lblProgress.Text = "No on-prem AD account found."
                $txtDL.Text = "No on-prem AD user account has a mail attribute matching '$Smtp'. This mailbox may be cloud-only (never synced from on-prem AD) -- DL membership can't be read this way for it. If you believe this account IS synced, double-check the Environment/Domain Controller selection at the top of the main window."
            } else {
                $status = if ($result.Cancelled) { "Cancelled -- partial results shown below." } else { "Done." }
                $lblProgress.Text = $status
                $sb = New-Object System.Text.StringBuilder
                [void]$sb.AppendLine("$status")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("Memberships ($($result.Memberships.Count)):")
                if ($result.Memberships.Count -eq 0) { [void]$sb.AppendLine("  (none found)") }
                else { $result.Memberships | Sort-Object | ForEach-Object { [void]$sb.AppendLine("  $_") } }
                if ($result.ScanErrors.Count -gt 0) {
                    [void]$sb.AppendLine("")
                    [void]$sb.AppendLine("Groups that could not be resolved ($($result.ScanErrors.Count)) -- membership above may be incomplete:")
                    $result.ScanErrors | ForEach-Object { [void]$sb.AppendLine("  $_") }
                }
                $txtDL.Text = $sb.ToString()
            }
        } catch {
            $lblProgress.Text = "Failed."
            $txtDL.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'DL Membership Lookup'
        } finally {
            $btnScan.Enabled = $true
            $btnCancel.Enabled = $false
        }
    })
    $btnCancel.Add_Click({
        $script:DLScanCancelRequested = $true
        $btnCancel.Enabled = $false
        $lblProgress.Text = "$($lblProgress.Text)  (cancelling after the current group...)"
    })
    $btnDLClose.Add_Click({ $dlForm.Close() })
    # CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the
    # monolith's shared parent window. This script has no parent window, so
    # it shows itself as the top-level dialog instead. Nothing else in this
    # function changed.
    [void]$dlForm.ShowDialog()
}
#endregion

Show-DLMembershipsForm -Smtp $target.Smtp -ADServer $target.ADServer
