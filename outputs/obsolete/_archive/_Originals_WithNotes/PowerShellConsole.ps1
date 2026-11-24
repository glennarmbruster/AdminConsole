#Requires -Version 5.1
<#
.SYNOPSIS
    PowerShell Console -- ad hoc command execution, in-session, logged. Meant
    for ONE console only, never distributed.
.DESCRIPTION
    NOT part of the standard tiered-delegation model this suite otherwise
    follows. Every other tool script has guardrails, confirm dialogs, or at
    minimum Write-CompletedLog on its own writes -- this one has none of that
    by design. It runs exactly what's typed, exactly as typed, with the full
    privileges of whatever is connected in this PowerShell session. That's
    the entire point of it, and also exactly why the operator who requested
    this was explicit that it installs on their own console only and is never
    copied into a Tools.json handed to another admin. Nothing in AdminConsole.
    ps1, ToolManager.ps1, or Tools.json needed to change to support that --
    Start-ToolScript already handles a Tools.json entry pointing at a script
    that isn't present in a given folder gracefully (a "Tool Script Not
    Found" message, not a crash -- see AdminConsole.ps1's own header). Simply
    don't include this file, or this file's Tools.json entry, in any copy of
    this suite handed to anyone else.

    Runs IN-SESSION via the same call-operator mechanism Start-ToolScript
    already uses for every tool in this suite -- inherits whatever Exchange
    Online connection / ActiveDirectory module are already loaded here,
    exactly like every other tool. Nothing special was needed to make that
    work: module state and a live EXO connection are process-level, not
    scoped to a particular script or function, so they're visible here
    automatically.

    -SmtpAddress / -Environment / -DomainController are declared only because
    Start-ToolScript always calls every tool with those three named
    parameters -- omitting them here would make every launch from
    AdminConsole throw a parameter-binding error. This tool ignores all
    three: it isn't object-scoped, never calls Resolve-ToolTarget, and never
    prompts for a target the way every other tool does.

    MANDATORY LOGGING, NOT OPTIONAL: every non-blank command is written to
    Write-CompletedLog BEFORE it's executed, not after -- so even a command
    that crashes this window, hangs it, or takes AdminConsole down with it
    (a real risk here more than anywhere else in this suite -- see below)
    still leaves a record that it was attempted. This was the one hard
    requirement from the design discussion that led to building this script.
    Do not add a way to skip or disable it.

    Write-CompletedLog (Common.ps1) wraps its Task string in literal double
    quotes with no CSV-standard escaping for embedded quotes -- perfectly
    fine for every other tool's short, internally-built log strings, which
    never contain a stray " or a line break, but a real risk for arbitrary
    typed PowerShell, which very commonly contains its own quotes and can
    span multiple lines. ConvertTo-LogSafeCommandText below flattens
    multi-line input to one line and swaps embedded double quotes for single
    quotes before logging -- a workaround local to this script, NOT a fix to
    Write-CompletedLog itself. That remains the broader, still-open
    "enterprise logging" gap flagged separately (operator identity, CSV
    escaping, centralized location) -- out of scope here.

    Execution uses Invoke-Expression, deliberately. Most PowerShell security
    guidance flags Invoke-Expression as something to avoid -- but that
    guidance is specifically about not running strings sourced from untrusted
    or remote input through it. Here the "untrusted string" is quite
    literally typed by the operator, by hand, into a box on their own
    console, for the express purpose of running it immediately. That's the
    textbook legitimate use case for Invoke-Expression, not the anti-pattern
    the usual warnings are about.

    Output -- both success output and errors, captured together via 2>&1 --
    is APPENDED to a running transcript in this window rather than cleared
    on every run, so a whole session's worth of ad hoc commands and their
    results stays visible and scrollable. Clear Output resets it.

    Trade-off worth naming plainly, carried over from AdminConsole's own
    documented reasoning for running tools in-session: an unhandled or
    interactive-prompting command typed here can hang or crash this window,
    and because this runs in AdminConsole's own process (not a separate one),
    that can take AdminConsole down with it too. That risk exists for every
    tool in this suite already -- it's simply far more exposed here, since
    arbitrary typed commands are exactly the kind of thing most likely to
    trigger it. No guardrails, no confirmation dialogs, no undo. A destructive
    command run here runs immediately, exactly as typed.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-18
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Data layer -----------------------------------------------------------------
# Write-CompletedLog has no CSV-standard escaping (no doubled "" for embedded
# quotes) -- see header. Flattens to one line and swaps " for ' before it
# ever reaches Write-CompletedLog, so a typed command containing quotes or
# line breaks can't corrupt CompletedTasks.csv.
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
    # Logged BEFORE execution, unconditionally -- see header. Do not move
    # this after the try/catch below, and do not make it conditional.
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

# Ctrl+Enter runs without adding a newline first -- plain Enter still just
# adds a line, matching how a multi-line command box should behave.
$txtCmd.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::Return) {
        $e.SuppressKeyPress = $true
        Invoke-AdHocCommand
    }
})

$txtCmd.Focus()
# ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
# Console, which opens top-left. Helper is in Common.ps1.
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
[void]$form.ShowDialog()
#endregion
