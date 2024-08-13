@echo off
REM -- depends on -Version 2.0 for now; albeit *did* work with PowerShell 3.0 once
REM -- for debugging, leave out `-WindowStyle Hidden` of course
cd %~dp0
powershell.exe -Version 2.0 -STA -WindowStyle Hidden -File ./modules/starter.ps1
