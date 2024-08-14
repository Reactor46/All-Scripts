@echo off

cd /d "E:\MailboxDatabaseLogs\GLGMBDB01"

echo.

forfiles /m *.log /c "cmd /c Del @PATH" /d -14 2> NUL

if %errorlevel% NEQ 0 echo No DB01 logs found older than 14 days !
if %errorlevel% EQU 0 echo Cleared DB01 logs older than 14 days !

echo.

pause