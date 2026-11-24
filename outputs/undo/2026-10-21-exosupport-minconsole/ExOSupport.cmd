@echo off
rem Starts the ExO Support console from the folder this file is in.
rem
rem THE POWERSHELL WINDOW STAYS VISIBLE ON PURPOSE. Sign-in runs there while
rem the console is closed, and a hidden or minimised PowerShell window costs
rem the sign-in prompt its foreground (the Admin Console learned this the hard
rem way). Leave it open behind the console; closing it closes ExO Support.
start "ExO Support" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0ExOSupport.ps1"
