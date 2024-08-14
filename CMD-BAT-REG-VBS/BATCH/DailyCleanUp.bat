@echo off
forfiles.exe /p "c:\inetpub\logs\LogFiles" /s /m *.log /d -2 /c "cmd /c del @file"
ping 1.1.1.1 -n 1 -w 60000 > nul
forfiles.exe /p "c:\Program Files\Microsoft\Exchange Server\V15\Logging" /s /m *.log /d -2 /c "cmd /c del @file"
ping 1.1.1.1 -n 1 -w 60000 > nul
forfiles.exe /p "c:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\DailyPerformanceLogs" /s /m *.* /d -2 /c "cmd /c del @file"
ping 1.1.1.1 -n 1 -w 60000 > nul
Exit
