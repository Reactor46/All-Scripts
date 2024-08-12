_____________________________________________
schtasks_tool.bat
_____________________________________________

rem @echo off
cls
setlocal EnableDelayedExpansion

set runasUsername=domain\administrator	
set runasPassword=password

if %1. == export. call :export
if %1. == import. call :import
exit /b 0

:export
md tasks 2>nul

schtasks /query /fo csv | findstr /V /c:"TaskName" > tnlist.txt

for /F "delims=," %%T in (tnlist.txt) do (
set tn=%%T
set fn=!tn:\=#!
echo !tn!
schtasks /query /xml /TN !tn! > tasks\!fn!.xml
)

rem Windows 2008 tasks which should not be imported.
del tasks\#Microsoft*.xml
exit /b 0

:import
for %%f in (tasks\*.xml) do (
call :importfile "%%f"
)
exit /b 0

:importfile
set filename=%1

rem replace out the # symbol and .xml to derived the task name
set taskname=%filename:#=%
set taskname=%taskname:tasks\=%
set taskname=%taskname:.xml=%

schtasks /create /ru %runasUsername% /rp %runasPassword% /tn %taskname% /xml %filename%
echo.
echo.