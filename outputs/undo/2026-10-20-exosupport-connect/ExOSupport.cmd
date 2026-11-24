@echo off
rem Starts the ExO Support console from the folder this file is in.
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0ExOSupport.ps1"
