#Requires -Version 5.1
<#
.SYNOPSIS
    Tenant-agnostic WinForms Mailbox Health tool -- now a pure UI layer over
    M365Logic.psm1.
.DESCRIPTION
    Fully functional on its own:
        .\M365_MailboxHealth.ps1
    prompts for a mailbox via M365Common.ps1's Show-TargetPromptDialog, then
    opens the Health window.

    Also launchable with a target already known:
        .\M365_MailboxHealth.ps1 -SmtpAddress user@anydomain.example
    skips the prompt and opens straight to the Health window.

    REVISED 2026-07-24 -- converted to the "logic module + UI layer" model.
    Get-MailboxHealthSnapshot and the report-building code that used to live
    in this file have MOVED to M365Logic.psm1 (as
    Get-MailboxHealthSnapshot and Format-MailboxHealthReport), unchanged.
    This file kept ONLY the WinForms window: build controls, call those two
    functions, put the returned text in a TextBox.

    WHAT THIS BUYS, concretely -- the identical two calls this window makes:
        $h = Get-MailboxHealthSnapshot -Smtp $Smtp
        Format-MailboxHealthReport -Smtp $Smtp -H $h
    are now the same two calls M365_MailboxHealth_Web.ps1 makes for its
    browser page, and the same two a scheduled job or an ad-hoc console
    one-liner would make. One implementation, four consumers -- so a fix or
    a new health field lands everywhere at once instead of being copied into
    each tool and drifting. That was the whole point of the split.

    This tool remains cloud-only and has no -Environment /
    -DomainController parameters -- see M365Logic.psm1's header on the
    on-prem module boundary.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24 (ported from ANZMailboxHealth.ps1, made
              tenant-agnostic)
    Revised : 2026-07-24, later -- logic extracted to M365Logic.psm1; this
              file is now UI only. See REVISED note above.
#>
param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Dot-sourcing M365Common.ps1 also imports M365Logic.psm1 -- see that file's
# header. So Get-MailboxHealthSnapshot / Format-MailboxHealthReport /
# Format-ErrorDetail are all available below without importing anything else
# here.
. (Join-Path $PSScriptRoot 'M365Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -ToolTitle 'Mailbox Health -- Enter Target'
if (-not $target) { return }   # operator cancelled the prompt

#region WinForms UI -- the only thing left in this file
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
    # click -- local to this form (closure-captured).
    $copyResetTimer = New-Object System.Windows.Forms.Timer
    $copyResetTimer.Interval = 1200
    $copyResetTimer.Add_Tick({
        $btnCopyHealth.Text = "Copy"
        $copyResetTimer.Stop()
    })

    # The whole body of this function is now two module calls plus error
    # handling -- everything that used to be ~120 lines of snapshot and
    # report building here lives in M365Logic.psm1.
    function Update-HealthFromCurrent {
        $btnRefresh.Enabled = $false
        $txtHealth.Text = "Loading mailbox health..."
        $healthForm.Refresh()
        try {
            $h = Get-MailboxHealthSnapshot -Smtp $Smtp
            $txtHealth.Text = Format-MailboxHealthReport -Smtp $Smtp -H $h
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
    [void]$healthForm.ShowDialog()
}
#endregion

Show-MailboxHealthForm -Smtp $target.Smtp
