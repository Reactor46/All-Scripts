@ECHO ON
SET LOGS=C:\Scripts\WinSysChecklist\LOGS
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)

dcdiag /c /v /skip:OutBoundSecureChannels /s:USONVSVRDC01 /f:%LOGS%\USONVSVRDC01.%mydate%.log
dcdiag /c /v /skip:OutBoundSecureChannels /s:USONVSVRDC02 /f:%LOGS%\USONVSVRDC02.%mydate%.log
dcdiag /c /v /skip:OutBoundSecureChannels /s:USONVSVRDC03 /f:%LOGS%\USONVSVRDC03.%mydate%.log
dcdiag /c /v /skip:OutBoundSecureChannels /s:MSODC04 /f:%LOGS%\MSODC04.%mydate%.log

notepad %LOGS%\USONVSVRDC01.%mydate%.log
notepad %LOGS%\USONVSVRDC02.%mydate%.log
notepad %LOGS%\USONVSVRDC03.%mydate%.log
notepad %LOGS%\MSODC04.%mydate%.log

