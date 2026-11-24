#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-RUNS: nothing against the tenant. Writes one timestamp to your local profile.
# TOOLKIT-HELP-BEGIN
# Starts the PIM countdown in the console header, without opening the portal.
#
# The jump box cannot reach the PIM activation portal, so activation happens
# elsewhere. This button is the "I have activated" note: press it once your
# access is active and the header starts counting down from the tenant's
# activation length (Pim.DurationHours in Tenant.json -- 9 hours here),
# going amber under 30 minutes and red when it runs out.
#
# It is a NOTE, not a live reading. Entra is not consulted. If a countdown is
# already running you are asked before it is reset, so a second click does
# not quietly hand you nine hours you do not have.
#
# Press it again after re-activating. Nothing else to configure.
# TOOLKIT-HELP-END
<#
    Set-PimTimer.ps1
    One click: record "activated now" and let the console header count down.

    Author   : Glenn Armbruster
    Revised  : 2026-09-12

    NO WINDOW. The console launches it, it records the moment, confirms, and
    returns; the console refreshes its own header the instant it comes back.
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms

. (Join-Path $PSScriptRoot 'Common.ps1')
if (-not (Test-ConsoleLaunch)) { return }

try {
    $hours = Get-PimDurationHours

    # A countdown already running is worth a question. Resetting it is the
    # right answer after a fresh activation and the wrong answer after a
    # mis-click, and only the operator knows which this is.
    $left = Get-PimRemaining
    if ($null -ne $left -and $left.TotalSeconds -gt 0) {
        $answer = Show-ToolDialog(
            "A PIM countdown is already running with $(Format-PimSpan -Span $left) left.`r`n`r`nReset it to $hours hours from now?",
            'PIM timer', 'YesNo', 'Question')
        if ("$answer" -ne 'Yes') { return }
    }

    $at    = Set-PimActivated
    $until = $at.AddHours($hours).ToLocalTime().ToString('HH:mm')

    try {
        Write-CompletedLog -Action 'PIM Activate' -Target $env:USERNAME -Result 'Success' `
            -Detail "Timer set from the console; $($hours)h assumed, expires ~$until" -Environment 'Cloud'
    } catch { }

    $null = Show-ToolDialog(
        "PIM timer set for $hours hours.`r`n`r`nExpires about $until. The countdown is in the console header.",
        'PIM timer', 'OK', 'Information')
} catch {
    Show-HandlerError -ErrorRecord $_ -Context 'Set PIM timer'
}
