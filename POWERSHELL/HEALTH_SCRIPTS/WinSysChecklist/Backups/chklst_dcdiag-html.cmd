@ECHO OFF
SET LOGS="C:\LazyWinAdmin\WinSysChecklist\Logs"
SET SCRIPTS="C:\LazyWinAdmin\WinSysChecklist"
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%a-%%b-%%c)

PowerShell.exe -File "%SCRIPTS%\chklst_DCDiag-HTML.ps1" -DomainName "Contoso.corp" -HTMLFileName %LOGS%\Report.%mydate%.html
REM PowerShell.exe -File "%SCRIPTS%\chklst_DCDiag-HTML.ps1" -DomainName "phx.Contoso.corp" -HTMLFileName %LOGS%\Report.phx.%mydate%.html
pause

