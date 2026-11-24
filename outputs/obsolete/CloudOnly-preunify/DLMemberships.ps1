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
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')


$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -ToolTitle 'DL Memberships -- Enter Target'
if (-not $target) { return }

#region Actions - mailbox health / DL memberships (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1458-1555, 1852-1954)
function Get-MailboxDLMemberships {
    <#
        Which distribution groups this address belongs to.

        Exchange Online has no "what groups is this recipient in" cmdlet, so the
        only honest answer is to enumerate groups and test each one. That is why
        this reports progress and can take a while on a large tenant -- it is a
        scan, not a lookup, and pretending otherwise would just look hung.

        Both distribution groups and mail-enabled security groups are checked,
        because a user is just as surprised by mail from one as the other.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [scriptblock]$Progress,
        [scriptblock]$ShouldCancel
    )

    $result = [PSCustomObject]@{
        Groups     = @()
        Scanned    = 0
        ScanErrors = @()
        FatalError = ''
        Cancelled  = $false
    }

    try {
        $groups = @(Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
            ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
    } catch {
        $result.FatalError = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup')
        return $result
    }

    $found = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($g in $groups) {
        $i++
        if ($ShouldCancel -and (& $ShouldCancel)) {
            # Stop between groups, never mid-group. Partial results from a
            # whole number of groups are true; half a group's membership is a
            # wrong answer that looks like a right one.
            $result.Cancelled = $true
            break
        }
        if ($Progress -and ($i % 5 -eq 0)) {
            & $Progress $i $groups.Count "$($g.DisplayName)"
        }
        try {
            $members = @(Invoke-EXOCommand -Name 'Get-DistributionGroupMember' -Parameters @{
                Identity = $g.Identity; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
        } catch {
            $result.ScanErrors += "$($g.DisplayName): $($_.Exception.Message)"
            continue
        }
        $hit = $members | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $Smtp }
        if ($hit) {
            [void]$found.Add([PSCustomObject]@{
                DisplayName        = "$($g.DisplayName)"
                PrimarySmtpAddress = "$($g.PrimarySmtpAddress)"
                GroupType          = "$($g.RecipientTypeDetails)"
                ManagedBy          = (@($g.ManagedBy) -join '; ')
                HiddenFromGAL      = [bool]$g.HiddenFromAddressListsEnabled
                MemberCount        = @($members).Count
            })
        }
    }
    $result.Scanned = $i
    $result.Groups  = @($found)
    return $result
}
function Show-DLMembershipsForm {
    param([Parameter(Mandatory)][string]$Smtp)
    $dlForm = New-Object System.Windows.Forms.Form
    $dlForm.Text = "DL Memberships: $Smtp"
    $dlForm.Size = New-Object System.Drawing.Size(560, 520)
    $dlForm.StartPosition = 'Manual'
    $dlForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Exchange Online has no reverse-membership cmdlet, so this enumerates every distribution group and mail-enabled security group and tests each one. It is a scan, not a lookup -- expect it to take a while on a large tenant, and use Cancel for partial results. Direct memberships only; Microsoft 365 groups are not included."
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
        $lblProgress.Text = "Enumerating groups in Exchange Online..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $progressCB = {
                param($i, $total, $dn)
                $lblProgress.Text = "Checking group $i of $total`: $dn"
                [System.Windows.Forms.Application]::DoEvents()
            }
            $cancelCB = { return $script:DLScanCancelRequested }
            $result = Get-MailboxDLMemberships -Smtp $Smtp -Progress $progressCB -ShouldCancel $cancelCB
            if ($result.FatalError) {
                $lblProgress.Text = "Failed."
                $txtDL.Text = $result.FatalError
            } else {
                $status = if ($result.Cancelled) { "Cancelled after $($result.Scanned) group(s) -- partial results below." } else { "Done -- $($result.Scanned) group(s) checked." }
                $lblProgress.Text = $status
                $sb = New-Object System.Text.StringBuilder
                [void]$sb.AppendLine("$status")
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("Memberships ($(@($result.Groups).Count)):")
                if (@($result.Groups).Count -eq 0) { [void]$sb.AppendLine("  (none found)") }
                else {
                    $result.Groups | Sort-Object DisplayName | ForEach-Object {
                        $own = if ($_.ManagedBy) { $_.ManagedBy } else { '(no owner)' }
                        $gal = if ($_.HiddenFromGAL) { ', hidden from GAL' } else { '' }
                        [void]$sb.AppendLine("  $($_.DisplayName) <$($_.PrimarySmtpAddress)>")
                        [void]$sb.AppendLine("      $($_.GroupType), $($_.MemberCount) member(s)$gal -- owner: $own")
                    }
                }
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
    Add-ToolHelpButton -Form $dlForm -ScriptPath $PSCommandPath -Title 'DL Memberships' | Out-Null
    Set-FormScreenCorner -Form $dlForm -Corner 'BottomLeft'
    [void]$dlForm.ShowDialog()
}
#endregion

Show-DLMembershipsForm -Smtp $target.Smtp
