#Requires -Version 5.1
<#
.SYNOPSIS
    WinForms UI-layer helpers for the M365_ toolkit. Imports M365Logic (all
    the actual logic) and adds only the WinForms-specific pieces on top.
.DESCRIPTION
    REVISED 2026-07-24 -- GUTTED, deliberately. This file used to hold both
    the logic AND the UI helpers for the tenant-agnostic toolkit. Per the
    operator's "logic module + UI layer" direction, every piece of logic
    that was here has MOVED to M365Logic.psm1, unchanged:

        Get-EXOCommand, Invoke-EXOCommand, Format-ErrorDetail,
        Get-M365DefaultDomain, Get-TargetSmtp, Resolve-Target,
        Get-InboxRulesForMailbox, Write-CompletedLog

    Those are all still available to any script that dot-sources THIS file,
    because this file imports the module -- so no existing tool's calls had
    to change. What's left below is only what is genuinely WinForms and
    genuinely UI:

        Show-TargetPromptDialog  -- the modal "enter a mailbox/group" dialog
        Resolve-ToolTarget       -- decides whether that dialog is needed

    WHY THOSE TWO STAYED BEHIND: both build a System.Windows.Forms.Form.
    They cannot run in the PS7 browser host (WinForms is Windows-only), they
    cannot run in a scheduled job (nothing there to click OK), and they'd
    deadlock a web request if a server thread ever called them. Keeping them
    out of M365Logic.psm1 is what lets that module load and work identically
    in WinForms, PS7, and headless contexts.

    Note the split inside what used to be one concept: Get-TargetSmtp (pure
    string normalization -- append the tenant's default domain if the
    operator didn't type one) is LOGIC and moved to the module.
    Resolve-ToolTarget (if we have no usable address, pop a dialog and ask)
    is UI and stayed here. It calls the module's Get-TargetSmtp.

    HOW TO USE (unchanged for every existing WinForms tool):
        . (Join-Path $PSScriptRoot 'M365Common.ps1')
    That single dot-source still yields everything a WinForms tool needs --
    the module's functions plus the two dialog helpers.

    The web layer does NOT dot-source this file. M365WebCommon.ps1 imports
    M365Logic directly and never touches WinForms at all.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-24 (as the tenant-agnostic sibling of Common.ps1 /
              SRGCommon.ps1 / ANZCommon.ps1)
    Revised : 2026-07-24, later -- logic extracted to M365Logic.psm1; this
              file is now the WinForms UI layer only. See REVISED note.
#>

#region Logic module import -------------------------------------------------------
# -Force so edits to the .psm1 are picked up without restarting the process
# (without it, a second Import-Module in the same session is a no-op).
# Importing the .psd1 manifest rather than the .psm1 directly so the
# manifest's explicit FunctionsToExport applies.
$script:M365LogicManifest = Join-Path $PSScriptRoot 'M365Logic.psd1'
if (Test-Path -LiteralPath $script:M365LogicManifest) {
    Import-Module $script:M365LogicManifest -Force -ErrorAction Stop
} else {
    # Fall back to the raw module if the manifest isn't alongside it, so a
    # partially-copied folder fails with a clear message rather than a
    # confusing "command not found" later.
    $script:M365LogicModule = Join-Path $PSScriptRoot 'M365Logic.psm1'
    if (-not (Test-Path -LiteralPath $script:M365LogicModule)) {
        throw "M365Logic.psd1/.psm1 not found in '$PSScriptRoot'. The M365_ tools need the logic module alongside this file -- copy M365Logic.psm1 and M365Logic.psd1 into the same folder."
    }
    Import-Module $script:M365LogicModule -Force -ErrorAction Stop
}
#endregion

#region "Enter a target" prompt for standalone M365_ tool launches ---------------
function Show-TargetPromptDialog {
    <#
    .SYNOPSIS
        Modal WinForms dialog asking for a mailbox or group address.
    .DESCRIPTION
        WinForms-only, interactive-only -- see this file's header for why
        this did not move into M365Logic.psm1.

        No Environment radio group and no Domain Controller ComboBox, unlike
        the on-prem suite's equivalent: this toolkit is cloud-only, so there
        is nothing for them to select. The returned object carries only
        .Smtp.

        The label deliberately names no specific domain -- it can't, since
        which domain applies isn't known until Connect-ExchangeOnline has
        run against some tenant. It says the domain can be omitted once
        connected instead.
    .OUTPUTS
        PSCustomObject with a .Smtp property, or $null if cancelled.
    #>
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
        # Get-TargetSmtp now comes from M365Logic (imported above).
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
    <#
    .SYNOPSIS
        Returns a usable target for a WinForms tool, prompting only if
        needed.
    .DESCRIPTION
        UI-layer decision function -- it may pop a modal dialog, so it
        belongs here rather than in M365Logic.psm1.

        When a tool is launched with a usable -SmtpAddress (e.g. from a
        console that already looked one up), this returns immediately with no
        UI at all. When it's launched with nothing -- double-clicked
        directly, or handed a bare name with no tenant connected yet to
        resolve the domain against -- it pops Show-TargetPromptDialog.
    .OUTPUTS
        PSCustomObject with a .Smtp property, or $null if the operator
        cancelled.
    #>
    param(
        [string]$SmtpAddress,
        [string]$ToolTitle = 'Enter Target'
    )
    $smtp = Get-TargetSmtp -Text $SmtpAddress
    if (-not $smtp) {
        $picked = Show-TargetPromptDialog -InitialSmtp $SmtpAddress -Title $ToolTitle
        if (-not $picked) { return $null }   # operator cancelled
        return $picked
    }
    return [PSCustomObject]@{ Smtp = $smtp }
}
#endregion
