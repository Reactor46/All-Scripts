@ECHO OFF
SET LOGS="C:\Scripts\WinSysChecklist\Logs"
SET SCRIPTS="C:\Scripts\WinSysChecklist"
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%a-%%b-%%c)

PowerShell.exe -File "%SCRIPTS%\chklst_DCDiag-HTML.ps1" -DomainName "USON.LOCAL" -HTMLFileName %LOGS%\Report.%mydate%.html

pause

