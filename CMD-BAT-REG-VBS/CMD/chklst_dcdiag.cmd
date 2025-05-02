@ECHO ON
SET LOGS=C:\LazyWinAdmin\WinSysChecklist\LOGS
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c-%%a-%%b)

dcdiag /c /v /skip:OutBoundSecureChannels /s:LASDC01 /f:%LOGS%\LASDC01.%mydate%.log
dcdiag /c /v /skip:OutBoundSecureChannels /s:LASDC02 /f:%LOGS%\LASDC02.%mydate%.log
dcdiag /c /v /skip:OutBoundSecureChannels /s:LASDC05 /f:%LOGS%\LASDC05.%mydate%.log
dcdiag /c /v /skip:OutBoundSecureChannels /s:PHXDC03 /f:%LOGS%\PHXDC03.%mydate%.log
dcdiag /c /v /skip:OutBoundSecureChannels /s:PHXDC04 /f:%LOGS%\PHXDC04.%mydate%.log

notepad++ %LOGS%\LASDC01.%mydate%.log
notepad++ %LOGS%\LASDC02.%mydate%.log
notepad++ %LOGS%\LASDC05.%mydate%.log
notepad++ %LOGS%\PHXDC03.%mydate%.log
notepad++ %LOGS%\PHXDC04.%mydate%.log