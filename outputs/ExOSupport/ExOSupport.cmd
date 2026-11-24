@echo off
rem Starts the ExO Support console from the folder this file is in.
rem
rem The black PowerShell window minimises itself once the console is up, and
rem comes back only while you sign in (Connect ExO). Do NOT add /min here or
rem -WindowStyle Minimized: Windows applies that to the console form instead,
rem and the console opens minimised. Closing the PowerShell window closes
rem ExO Support.
start "ExO Support" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0ExOSupport.ps1"
