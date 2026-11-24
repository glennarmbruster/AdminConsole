#Requires -Version 5.1
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

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'DL Memberships -- Enter Target'
if (-not $target) { return }

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
    param([string]$Smtp, [string]$ADServer)
    $dlForm = New-Object System.Windows.Forms.Form
    $dlForm.Text = "DL Memberships: $Smtp"
    $dlForm.Size = New-Object System.Drawing.Size(560, 520)
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

Show-DLMembershipsForm -Smtp $target.Smtp -ADServer $target.ADServer
