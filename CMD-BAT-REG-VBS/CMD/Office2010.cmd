::Installing Office 2010 SP2
setlocal
set ProductName=Office14.STANDARD
set DeployServer=%CD%\Apps\office\MS
set ConfigFile=%CD%\Apps\office\MS\Standard.WW\config.xml
set LogLocation=%CD%\Apps\office\MS\office2010Logs

IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)
:ARP64
reg query HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432NODE\Microsoft\Windows\CurrentVersion\Uninstall\%ProductName%
if NOT %errorlevel%==1 (goto End)
 
:ARP86
reg query HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%ProductName%
if %errorlevel%==1 (goto DeployOffice) else (goto End)

:DeployOffice
start /wait %DeployServer%\setup.exe /adminfile %DeployServer%\Updates\ACU.msp
echo %date% %time% Setup ended with error code %errorlevel%. >> %LogLocation%\%computername%.txt
::Activate Office
%PROGRAMFILES%\Microsoft Office\Office14>cscript OSPP.VBS /act

:End

Endlocal