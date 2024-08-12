@echo off
:: Created by: Shawn Brink
:: http://www.sevenforums.com
:: Tutorial: http://www.sevenforums.com/tutorials/91738-windows-update-reset.html


set w=0
:wuauserv
set /a w=%w%+1
if %w% equ 3 (
   goto end
) 
net stop wuauserv
echo Checking the wuauserv service status.
sc query wuauserv | findstr /I /C:"STOPPED" 
if not %errorlevel%==0 ( 
    goto wuauserv 
)
goto Reset

:end
cls
echo.
echo Failed to reset Windows Update history due to wuauserv service failing to stop.
echo.
pause
goto Start

:Reset 
if exist "%SYSTEMROOT%\SoftwareDistribution.bak" rmdir /s /q "%SYSTEMROOT%\SoftwareDistribution.bak"
if exist "%SYSTEMROOT%\SoftwareDistribution" ( 
    attrib -r -s -h /s /d "%SYSTEMROOT%\SoftwareDistribution" 
    ren "%SYSTEMROOT%\SoftwareDistribution" SoftwareDistribution.bak 
) 

if exist "%SYSTEMROOT%\WindowsUpdate.log.bak" del /s /q /f "%SYSTEMROOT%\WindowsUpdate.log.bak" 
if exist "%SYSTEMROOT%\WindowsUpdate.log" ( 
    attrib -r -s -h /s /d "%SYSTEMROOT%\WindowsUpdate.log" 
    ren "%SYSTEMROOT%\WindowsUpdate.log" WindowsUpdate.log.bak 
) 
 
regsvr32 /s wuaueng.dll 
regsvr32 /s wuaueng1.dll 
regsvr32 /s atl.dll 
regsvr32 /s wups.dll 
regsvr32 /s wups2.dll 
regsvr32 /s wuweb.dll 
regsvr32 /s wucltui.dll 

:Start
net start wuauserv





