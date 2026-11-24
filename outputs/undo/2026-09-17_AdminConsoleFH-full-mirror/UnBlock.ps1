#Requires -Version 5.1
# TOOLKIT-HELP-RUNS: nothing. It only touches files already on this machine.
# TOOLKIT-VARIANT: Any
# TOOLKIT-HELP-BEGIN
# Clears the mark-of-the-web from every file in the toolkit folder.
#
# WHAT IT IS FOR. A file that arrived by download, email or a copy from another
# machine carries a flag saying so. PowerShell then refuses to dot-source it,
# or warns on every launch -- which on a toolkit of a hundred scripts means a
# hundred warnings, or a console that will not start.
#
# RUN IT ONCE after copying a new build in. Nothing else needs doing.
#
# IT WORKS ON THE FOLDER IT IS IN, found from its own location -- so it cannot
# unblock the wrong copy of the toolkit, which is what the previous version
# did when it was pointed at one hardcoded path.
#
# READ-ONLY as far as any mailbox, group or directory is concerned. It changes
# a file attribute and nothing else.
# TOOLKIT-HELP-END
<#
    UnBlock.ps1
    Clears mark-of-the-web from downloaded script files.

    Author   : Glenn Armbruster
    Revised  : 2026-09-13

    2026-09-13  Path comes from $PSScriptRoot instead of a hardcoded drive
                letter.
#>

# $PSScriptRoot, NOT A DRIVE LETTER. This ran against one hardcoded path,
# so a copy of the toolkit anywhere else unblocked the wrong folder -- or,
# on a machine without that drive, silently nothing.
Get-ChildItem -LiteralPath $PSScriptRoot -File -Recurse |
    Unblock-File
