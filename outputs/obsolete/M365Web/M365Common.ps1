#Requires -Version 5.1
<#
    M365Common.ps1
    WinForms UI-layer helpers for the M365_ toolkit. Imports M365Logic (all
    the actual logic) and adds only the WinForms-specific pieces on top.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

#region Logic module import -------------------------------------------------------
$script:M365LogicManifest = Join-Path $PSScriptRoot 'M365Logic.psd1'
if (Test-Path -LiteralPath $script:M365LogicManifest) {
    Import-Module $script:M365LogicManifest -Force -ErrorAction Stop
} else {
    $script:M365LogicModule = Join-Path $PSScriptRoot 'M365Logic.psm1'
    if (-not (Test-Path -LiteralPath $script:M365LogicModule)) {
        throw "M365Logic.psd1/.psm1 not found in '$PSScriptRoot'. The M365_ tools need the logic module alongside this file -- copy M365Logic.psm1 and M365Logic.psd1 into the same folder."
    }
    Import-Module $script:M365LogicModule -Force -ErrorAction Stop
}
#endregion

#region "Enter a target" prompt for standalone M365_ tool launches ---------------
function Show-TargetPromptDialog {
    param(
        [string]$InitialSmtp,
        [string]$Title = 'Enter Target'
    )
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(420, 156)
    $dlg.StartPosition = 'CenterScreen'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = "Mailbox or group (email address -- domain can be omitted once Exchange Online is connected):"
    $lblUser.Location = New-Object System.Drawing.Point(12, 12)
    $lblUser.Size = New-Object System.Drawing.Size(380, 18)
    $dlg.Controls.Add($lblUser)
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(12, 32)
    $txtUser.Size = New-Object System.Drawing.Size(380, 22)
    $txtUser.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtUser.Text = $InitialSmtp
    $dlg.Controls.Add($txtUser)

    $lblErr = New-Object System.Windows.Forms.Label
    $lblErr.Location = New-Object System.Drawing.Point(12, 60)
    $lblErr.Size = New-Object System.Drawing.Size(380, 34)
    $lblErr.ForeColor = [System.Drawing.Color]::Red
    $dlg.Controls.Add($lblErr)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(220, 100)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Location = New-Object System.Drawing.Point(308, 100)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk

    $btnOk.Add_Click({
        $raw = $txtUser.Text.Trim()
        if ($raw -eq '') { $lblErr.Text = 'Enter a mailbox or group email address.'; return }
        $smtp = Get-TargetSmtp -Text $raw
        if (-not $smtp) {
            $lblErr.Text = "Could not determine a domain automatically -- type the full email address, or connect Exchange Online first (Connect-ExchangeOnline) and try again."
            return
        }
        $dlg.Tag = [PSCustomObject]@{ Smtp = $smtp }
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })

    $result = $dlg.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $dlg.Tag
}

function Resolve-ToolTarget {
    param(
        [string]$SmtpAddress,
        [string]$ToolTitle = 'Enter Target'
    )
    $smtp = Get-TargetSmtp -Text $SmtpAddress
    if (-not $smtp) {
        $picked = Show-TargetPromptDialog -InitialSmtp $SmtpAddress -Title $ToolTitle
        if (-not $picked) { return $null }
        return $picked
    }
    return [PSCustomObject]@{ Smtp = $smtp }
}
#endregion
