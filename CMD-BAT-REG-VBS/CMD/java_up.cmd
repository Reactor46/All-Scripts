@ECHO OFF
cls
TITLE Uninstalling Java 5-7 and Java fx. . .

wmic product where "name like 'Java 7%%'" call uninstall /nointeractive
wmic product where "name like 'JavaFX%%'" call uninstall /nointeractive
wmic product where "name like 'Java(TM) 7%%'" call uninstall /nointeractive
wmic product where "name like 'Java(tm) 6%%'" call uninstall /nointeractive
wmic product where "name like 'J2SE Runtime Environment%%'" call uninstall /nointeractive
goto END

:END
pause
exit

@ECHO OFF
::http://www.java.com/en/download/manual.jsp
cls
title Installing Java
echo Installing Java. . .

set Java32="\\SERVER\FOLDER\jre-6u29-windows-i586-s.exe"
set Java64="\\SERVER\FOLDER\jre-6u29-windows-x64.exe"

if not exist %Java32% set missing=%Java32% & goto ERR1
if not exist %Java64% set missing=%Java64% & goto ERR1

IF NOT DEFINED PROCESSOR_ARCHITEW6432 (
IF %PROCESSOR_ARCHITECTURE% EQU x86 (
goto JA32
) ELSE (
goto JA64
)) ELSE (
goto JA64
)

:JA64
%Java64% /passive /norestart
goto JA32

:JA32
%Java32% /s REBOOT=Suppress
goto END

:ERR1
cls
echo Unable to locate: %missing%
echo.
pause
goto END

:END
exit