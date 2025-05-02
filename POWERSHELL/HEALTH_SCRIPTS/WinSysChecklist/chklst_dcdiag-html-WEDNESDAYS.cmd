@ECHO OFF
SET LOGS="C:\LazyWinAdmin\WinSysChecklist\Logs"
SET SCRIPTS="C:\LazyWinAdmin\WinSysChecklist"
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)

PowerShell.exe -File "%SCRIPTS%\chklst_DCDiag-HTML.ps1" -DomainName "creditoneapp.biz" -HTMLFileName %LOGS%\Report.CreditOneApp.biz.%mydate%.html
REM PowerShell.exe -File "%SCRIPTS%\chklst_DCDiag-HTML.ps1" -DomainName "lasauth02.creditoneapp.biz" -HTMLFileName %LOGS%\Report.CreditOneApp.biz.%mydate%.html
REM PowerShell.exe -File "%SCRIPTS%\chklst_DCDiag-HTML.ps1" -DomainName "phxauth01.creditoneapp.biz" -HTMLFileName %LOGS%\Report.PHX.CreditOneApp.biz.%mydate%.html
pause

 

