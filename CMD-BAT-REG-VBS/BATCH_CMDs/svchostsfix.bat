@echo off

:intro
color 1F
@echo off
cls

echo.
echo.
echo                            *******************
echo                          *                     *
echo                         *      Svhost.exe       *
echo                         *      corruption       *
echo                         *         Fix           *
echo                         *      created by       *
echo                         *       pcbutts1        *
echo                          *                     *
echo                            *******************
echo.
echo       This tool will Remove Corrupted Windows Update Files
echo       which can sometimes cause an increase in CPU activity
echo               Your system will reboot when done.
echo.
echo.
echo.
pause
echo.
echo.

@echo off


REGSVR32 /s WUAPI.DLL
REGSVR32 /s WUAUENG.DLL
REGSVR32 /s WUAUENG1.DLL
REGSVR32 /s ATL.DLL
REGSVR32 /s WUCLTUI.DLL
REGSVR32 /s WUPS.DLL
REGSVR32 /s WUPS2.DLL
REGSVR32 /s WUWEB.DLL
pause
net stop WuAuServ
pause
cd %windir%
pause
ren SoftwareDistribution SD_OLD
pause
net start WuAuServ
pause

echo.
echo      ##########################################################
echo     #                                                          #
echo     #     This removal tool has completed its functions.       #
echo     #                                                          #
echo     #           A Restart of your computer is needed           #
echo     #         Pressing any key will close this program         #
echo     #              and restart your computer.                  #
echo     #                                                          #
echo      ##########################################################
echo.
echo      ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

pause
shutdown -r -t 30
echo.
echo.
exit
