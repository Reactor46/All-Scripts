@echo off
title Are You Backd00red?
setlocal enabledelayedexpansion

echo Backdoor Finder
echo.
echo Areyoubackdoored will check if any of your browsers are running.
echo Some backdoor programs use browsers to connect back.
echo You should close all browsers (don't kill the processes, though, just close the
echo windows) before you start, or else you would for sure find them running.
echo.
< nul set /p var=Press any key to start...
pause > nul

echo.
for /f "tokens=5" %%a in ('netstat -ano ^| find /i "established"') do (
    cls
    echo Please, wait...
    echo Now looking for ID %%a...
    for %%A in (chrome iexplore firefox opera safari tor) do (
        for /f %%b in ('tasklist ^| find "%%a"') do (
            for /f "delims=" %%j in ('echo %%b ^| find /i "%%A.exe"') do set check=%%j
        )
    )
)
cls
if not defined check goto nobackdoors
echo Done.
echo Your computer is probably backdoored:
echo process %check%looked open while it should have been killed.
echo.
< nul set /p var=Press any key to quit...
pause > nul
exit /b
:nobackdoors
echo Done.
echo No backdoors are using browsers on current computer.
echo.
< nul set /p var=Press any key to quit...
pause > nul