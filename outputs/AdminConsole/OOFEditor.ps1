#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    OOFEditor.ps1
    Out of Office -- rich-text editor with templates and preview.

    Author   : Glenn Armbruster
    Revised  : 2026-09-09  (rich editor, templates, preview, date pickers)
#>
# TOOLKIT-HELP-RUNS: reads the auto-reply configuration. Save changes it and asks first.
# TOOLKIT-HELP-BEGIN
# Read and set someone's out-of-office message without needing their mailbox
# open. Both messages -- internal, and the one external senders receive --
# and the state, audience and schedule.
#
# THE EDITOR IS A REAL ONE. Exchange stores auto-reply text as HTML, so the
# message boxes are HTML editors, not plain text: font, size, colour,
# highlight, bold/italic/underline, bullets, alignment, links -- the same
# things a person would do in Outlook. Ctrl+B / Ctrl+I / Ctrl+U work. What
# you see is what the sender gets.
#
# TEMPLATES. Three standard messages (back on a date, extended leave, left
# the company). Pick one, press Insert, fill in the two or three blanks --
# return date comes from the End picker if it is set, the name from the
# mailbox -- and edit from there. A template REPLACES the tab you are on,
# and asks first if there is something in it.
#
# PREVIEW shows each message exactly as it will render, with the HTML
# source alongside for the times something looks wrong and you need to see
# why. Copy -> External copies the internal message over the external one.
#
# Scheduled needs both a start and an end; the pickers are only live in that
# state. External audience: None (external senders get nothing), Known
# (contacts only), All.
#
# Exchange Online only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:Target      = Get-TargetSmtp -Text $SmtpAddress
$script:DisplayName = ''

#region Window ------------------------------------------------------------------
<#
    EVERYTHING THIS TOOL DOES NOW LIVES IN Common.ps1.

    The editor, the templates, the preview and the save moved there on
    2026-09-18 so the Mailbox Editor's Automatic Replies tab could use the same
    one. Its tab used to hold plain text boxes, which threw away the HTML
    formatting of whatever message was already set -- two editors for one
    field, one of them destructive.

    What is left here is the thing a shared function cannot be: a tool with a
    name, a place in the console, and its own help text. Deleting this file and
    telling people to open the Mailbox Editor would be a worse trade -- setting
    one person's out-of-office is a job of its own and does not need eight other
    tabs in front of it.
#>
Show-OutOfOfficeForm -Smtp $script:Target -HelpScriptPath $PSCommandPath
#endregion
