@echo off
powershell.exe -noprofile -executionpolicy bypass -Command "& '%~dp0\WSUS_ServerOU.ps1'"
if %ERRORLEVEL% == 1 goto END
echo N | gpupdate /force
wuauclt /resetauthorization 
wuauclt /detectnow
wuauclt /reportnow
wuauclt /detectnow
wuauclt /reportnow

:END
