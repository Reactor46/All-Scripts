@echo off
set currdir=%~dp0
IF "%currdir:~-1%"=="\" SET currdir=%currdir:~0,-1%
cd /d "%currdir%"

set installpath=C:\solr

call solr-download-windows.bat
call nssm-download.bat

robocopy solr %installpath% /E

copy /Y "%currdir%\windows_service_create.bat" %installpath%
copy /Y "%currdir%\windows_service_delete.bat" %installpath%
copy /Y "%currdir%\windows_uninstall.bat" %installpath%
copy /Y "%currdir%\firewall_allow_solr.bat" %installpath%
copy /Y "%currdir%\firewall_delete_solr.bat" %installpath%
copy /Y "%currdir%\nssm.exe" %installpath%\bin

call windows_service_create.bat
call firewall_allow_solr.bat

:END
echo bye
timeout /t 3
