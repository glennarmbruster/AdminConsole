#Requires -Version 5.1
<#
    console.ps1
    PowerShell prompt using the console's live connections.

    Author   : Glenn Armbruster
    Revised  : 2026-07-28
#>
# TOOLKIT-HELP-BEGIN
# A PowerShell prompt that shares the console's live connections.
#
# For the times a tool does not cover what you need and you want to run cmdlets
# by hand, without connecting to everything again in a new window.
#
# Whatever the console is connected to, this is connected to. That includes the
# ability to change things, with none of the previews or confirmations the tools
# give you.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Data layer -----------------------------------------------------------------
function ConvertTo-LogSafeCommandText {
    param([string]$Text)
    $flat = ($Text -replace "`r`n", ' | ' -replace "`n", ' | ')
    $flat = $flat -replace '"', "'"
    return $flat
}
#endregion

#region Main window ----------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell Console (ad hoc -- logged, no guardrails)"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Runs whatever you type IMMEDIATELY, with no confirmation and no guardrails -- exactly as typed, in this session's live connection. Every command is logged before it runs. Ctrl+Enter also runs."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(860, 34)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblCmd = New-Object System.Windows.Forms.Label
$lblCmd.Text = "Command (one-liner or multi-line):"
$lblCmd.Location = New-Object System.Drawing.Point(12, 48)
$lblCmd.AutoSize = $true
$form.Controls.Add($lblCmd)
$txtCmd = New-Object System.Windows.Forms.TextBox
$txtCmd.Multiline = $true
$txtCmd.ScrollBars = "Vertical"
$txtCmd.AcceptsReturn = $true
$txtCmd.AcceptsTab = $true
$txtCmd.Location = New-Object System.Drawing.Point(12, 68)
$txtCmd.Size = New-Object System.Drawing.Size(860, 100)
$txtCmd.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($txtCmd)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run"
$btnRun.Location = New-Object System.Drawing.Point(12, 176)
$btnRun.Size = New-Object System.Drawing.Size(100, 30)
$btnRun.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)
$btnClearCmd = New-Object System.Windows.Forms.Button
$btnClearCmd.Text = "Clear Command"
$btnClearCmd.Location = New-Object System.Drawing.Point(120, 176)
$btnClearCmd.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnClearCmd)
$btnClearOutput = New-Object System.Windows.Forms.Button
$btnClearOutput.Text = "Clear Output"
$btnClearOutput.Location = New-Object System.Drawing.Point(248, 176)
$btnClearOutput.Size = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnClearOutput)

$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Output (appends each run -- newest at the bottom):"
$lblOutput.Location = New-Object System.Drawing.Point(12, 216)
$lblOutput.AutoSize = $true
$form.Controls.Add($lblOutput)
$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Multiline = $true
$txtOutput.ReadOnly = $true
$txtOutput.ScrollBars = "Both"
$txtOutput.WordWrap = $false
$txtOutput.Location = New-Object System.Drawing.Point(12, 236)
$txtOutput.Size = New-Object System.Drawing.Size(860, 380)
$txtOutput.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtOutput.BackColor = [System.Drawing.Color]::Black
$txtOutput.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($txtOutput)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(772, 626)
$btnClose.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnClose)

function Invoke-AdHocCommand {
    $cmdText = $txtCmd.Text
    if (-not $cmdText -or $cmdText.Trim() -eq '') { return }
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-CompletedLog -Task "Ran ad hoc PowerShell (Console): $(ConvertTo-LogSafeCommandText -Text $cmdText)"
    $txtOutput.AppendText("`r`n> [$timestamp]`r`n$cmdText`r`n")
    try {
        $result = Invoke-Expression $cmdText 2>&1 | Out-String
        if ("$result".Trim() -eq '') { $result = "(no output)`r`n" }
        $txtOutput.AppendText($result)
    } catch {
        $txtOutput.AppendText((Format-ErrorDetail -ErrorRecord $_ -Context 'Ad hoc PowerShell command') + "`r`n")
    }
    $txtOutput.SelectionStart = $txtOutput.TextLength
    $txtOutput.ScrollToCaret()
}

$btnRun.Add_Click({ Invoke-AdHocCommand })
$btnClearCmd.Add_Click({ $txtCmd.Clear(); $txtCmd.Focus() })
$btnClearOutput.Add_Click({ $txtOutput.Clear() })
$btnClose.Add_Click({ $form.Close() })

$txtCmd.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Return) {
        $e.SuppressKeyPress = $true
        Invoke-AdHocCommand
    }
})

$txtCmd.Focus()
$form.StartPosition = 'Manual'
Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Console' | Out-Null

[void]$form.ShowDialog()
#endregion
