:Start
cls
@echo off
Echo.
Echo Remote Disk Drive Info
Echo.

Set /P Computer=Enter the Computer Name:
If  "%Computer%"==  "" goto BadName
pause

psexec \\%Computer% -h cmd /c reg query "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer" /v "svcVersion"
pause

Goto End

:BadName
Cls
Echo.
Echo You have entered an incorrect name or left this field blank
Echo Please enter a valid Name or press Ctr-C to exit.
Echo.
Pause
Goto Start

:End