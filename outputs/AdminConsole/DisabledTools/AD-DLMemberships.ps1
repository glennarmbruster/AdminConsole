#Requires -Version 5.1
# TOOLKIT-VARIANT: OnPremAD
# ---------------------------------------------------------------------------
# AD- PREFIX: this is the ACTIVE DIRECTORY version of DLMemberships.
#
# It reads and writes raw AD attributes -- mailNickname, proxyAddresses,
# targetAddress -- through Get-ADUser/Get-ADGroup, and needs a reachable domain
# controller. DLMemberships.ps1 beside it does the same job against Exchange Online
# properties (Alias, EmailAddresses) and needs no directory at all.
#
# TWO FILES RATHER THAN TWO CODE PATHS. The alternative was one tool branching
# on Test-TenantCapability, which would have given the six most attribute-heavy
# tools in the toolkit a second path each, both needing testing. Prefixing the
# names keeps each file doing one thing, and Publish-Toolkit lists which set to
# strip for a given tenant.
#
# Ship AD-* where Tenant.json has OnPremAD true. Ship the unprefixed ones
# everywhere. Shipping both is not harmful -- the wrong one greys itself -- but
# it is two buttons doing one job.
# ---------------------------------------------------------------------------
<#
    DLMemberships.ps1
    Standalone DL Memberships tool -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# Shows every distribution list a person belongs to.
#
# Answers "why is this person getting this mail" and "what will they stop
# receiving when they leave" -- the second one is worth checking before a
# termination rather than after.
#
# Read-only. To change membership, use DL Management.
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

# THE ADDRESS BOX IS ON THE FORM -- see Add-ToolTargetBar in Common.ps1.
#
# Resolve-ToolTarget is KEPT here, unlike in the cloud tools. This one reads a
# domain controller, so it needs to know which directory holds the account, and
# an address it cannot place is an address this tool cannot serve. It is called
# per lookup instead of once at startup -- see Resolve-ToolTargetOnDemand.
$script:Target   = ''
$script:ADServer = $null

function Set-DLTarget {
    # Returns nothing. Both values are cleared first, so a failed lookup can
    # never leave the previous account's controller behind for the next scan.
    param([string]$Text)
    $script:Target   = ''
    $script:ADServer = $null
    $t = Resolve-ToolTargetOnDemand -Text $Text -DomainController $DomainController -ToolTitle 'DL Memberships -- Enter Target'
    if (-not $t) { return }
    $script:Target   = $t.Smtp
    $script:ADServer = $t.ADServer
}

Set-DLTarget -Text $SmtpAddress

#region Actions - mailbox health / DL memberships (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1458-1555, 1852-1954)
function Get-MailboxDLMemberships {
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
    # NO PARAMETERS. The account can change while the window is open, so the
    # address and its domain controller live in $script:Target and
    # $script:ADServer rather than in parameters fixed at build time.
    $dlForm = New-Object System.Windows.Forms.Form
    $dlForm.Text = 'DL Memberships'
    $dlForm.ClientSize = New-Object System.Drawing.Size(544, 476)
    $dlForm.StartPosition = 'Manual'
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

    function Update-DLTitle {
        $dlForm.Text = if ("$script:Target".Trim()) { "DL Memberships: $script:Target" } else { 'DL Memberships' }
    }
    Update-DLTitle

    $btnScan.Add_Click({
        if (-not "$script:Target".Trim()) {
            $lblProgress.Text = 'No target.'
            $txtDL.Text = 'Enter a mailbox address above and press Look Up first.'
            return
        }
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
            $result = Get-MailboxDLMemberships -Smtp $script:Target -ADServer $script:ADServer -ProgressAction $progressCB -ShouldCancel $cancelCB
            if ($result.FatalError) {
                $lblProgress.Text = "Failed."
                $txtDL.Text = $result.FatalError
            } elseif ($result.NoADAccount) {
                $lblProgress.Text = "No on-prem AD account found."
                $txtDL.Text = "No on-prem AD user account has a mail attribute matching '$script:Target'. This mailbox may be cloud-only (never synced from on-prem AD) -- DL membership can't be read this way for it. If you believe this account IS synced, double-check the Environment/Domain Controller selection at the top of the main window."
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

    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down and
    # grows the form, and the help button places itself from ClientSize at run
    # time. Before Set-FormScreenCorner too, which positions from the size.
    $bar = Add-ToolTargetBar -Form $dlForm -InitialValue $script:Target

    # THE SCAN IS NOT RE-RUN AUTOMATICALLY. It walks the directory and can take
    # a while, so Look Up resolves the account and stops; Scan is still the
    # deliberate act it was before.
    $bar.Button.Add_Click({
        Set-DLTarget -Text $bar.TextBox.Text
        Update-DLTitle
        $txtDL.Text = ''
        $lblProgress.Text = if ("$script:Target".Trim()) {
            "Ready -- press Scan to read group memberships for $script:Target."
        } else {
            'Not started.'
        }
    })

    $dlForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $dlForm
    Set-FormScreenCorner -Form $dlForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $dlForm -ScriptPath $PSCommandPath -Title 'DL Memberships' | Out-Null

    [void]$dlForm.ShowDialog()
}
#endregion

Show-DLMembershipsForm
