@ECHO On
SET LOGS=C:\LazyWinAdmin\WinSysChecklist\Logs
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%a-%%b-%%c)

net time \\LASDC01 >> %LOGS%\Time_Chk.%mydate%.txt
net time \\LASDC02 >> %LOGS%\Time_Chk.%mydate%.txt
net time \\LASDC05 >> %LOGS%\Time_Chk.%mydate%.txt
net time \\PHXDC03 >> %LOGS%\Time_Chk.%mydate%.txt
net time \\PHXDC04 >> %LOGS%\Time_Chk.%mydate%.txt