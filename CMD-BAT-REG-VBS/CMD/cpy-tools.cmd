@echo off
@call init-path.cmd

echo Copy of tools...
echo *****************************
xcopy .\toolkit\*.* %optdrv%%optpath%\toolkit\ /s /e /v /h /r /d /y
pause
echo *****************************
