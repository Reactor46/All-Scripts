@ECHO ON

SET LOGS=C:\LazyWinAdmin\BATCH\Logs
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)

psexec \\lasdc01 repadmin /syncall /AdeP >> %LOGS%\DomainSync.%mydate%.txt
psexec \\lasdc02 repadmin /syncall /AdeP >> %LOGS%\DomainSync.%mydate%.txt
psexec \\lasdc05 repadmin /syncall /AdeP >> %LOGS%\DomainSync.%mydate%.txt