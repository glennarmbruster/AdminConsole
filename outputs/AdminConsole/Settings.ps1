#Requires -Version 5.1
<#
    Settings.ps1
    Your paths, your email, your browser profiles. Per-engineer, not per-tenant.

    Reasoning: ScriptNotes/Settings.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-20
#>
# TOOLKIT-HELP-BEGIN
# The handful of things that are YOURS rather than the tenant's.
#
#   Your email          where a copy of a Feature Request or Bug Fix is sent,
#                       and the address the OneDrive notice comes from. Blank
#                       and those tools tell you they cannot run.
#   Reports / Logs      where this toolkit writes. Blank means "beside the
#                       toolkit", which always exists.
#   Retry queue         where a failed Disable User run is parked for another go.
#   SharePoint admin    your tenant's -admin URL, if the derived one is wrong.
#   Exchange server     the on-premises endpoint Connect ExSE uses.
#   Edge profiles       which profile QuickLaunch opens admin vs standard links in.
#
# NONE OF THIS TRAVELS. Deploy leaves Config\Settings.json out on purpose --
# it is the one file in Config\ that is different on every machine. Everything
# else in there is tenant fact and is the same for all of us.
#
# BLANK IS A REAL ANSWER for every path. A drive letter typed here works on the
# machine you typed it on and fails silently on one built differently, at the
# moment something tries to write.
#
# The console reads this at startup. Restart it to see changes.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: Writes Config\Settings.json on this machine. Connects to nothing.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

# THE WHOLE TOOL. Show-SettingsEditor lives in Common.ps1 precisely so this file
# can be four lines rather than a second copy of a 220-line dialog that would
# drift from ToolManager's.
#
# It is modal and returns $true when it wrote. Nothing is done with that here:
# the dialog reports its own outcome, and a second confirmation box saying the
# same thing is a click for no information.
[void](Show-SettingsEditor)
