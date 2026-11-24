#Requires -Version 5.1
<#
    UnBlock.ps1
    Clears mark-of-the-web from downloaded script files.

    Author   : Glenn Armbruster
    Revised  : 2026-07-28
#>

Get-ChildItem "E:\Scripts\M365\AdminConsole" -File -Recurse |
    Unblock-File
