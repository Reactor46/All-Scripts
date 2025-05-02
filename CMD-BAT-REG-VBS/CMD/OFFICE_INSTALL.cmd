@ECHO ON

setlocal 
set DEPLOYSCRIPT=\\branch1dc\deploy$


:Installing Office 2010 SP2
set ProductName=Office14.STANDARD
set ConfigFile=%DEPLOYSCRIPT%\Apps\office\MS\Standard.WW\config.xml
set LogLocation=%DEPLOYSCRIPT%\Apps\office\MS\office2010Logs

IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)
:ARP64
reg query HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432NODE\Microsoft\Windows\CurrentVersion\Uninstall\%ProductName%
if NOT %errorlevel%==1 (goto End)
 
:ARP86
reg query HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%ProductName%
if %errorlevel%==1 (goto DeployOffice) else (goto End)

:DeployOffice
start /wait %DEPLOYSCRIPT%\Apps\office\ms\setup.exe /adminfile %DEPLOYSCRIPT%\Apps\office\ms\Updates\ACU.msp

:Activate Office
%PROGRAMFILES%\Microsoft Office\Office14>cscript OSPP.VBS /act

:Disable-UAC 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v EnableLUA /t REG_DWORD /d 0 /f 

:Disable-Windows-Backup 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsBackup" /v DisableMonitoring /t REG_DWORD /d 0 /f 

:Reg-edits
regedit -s /i %DEPLOYSCRIPT%\Scripts\Reg\InstallTakeOwnership.reg 
regedit -s /i %DEPLOYSCRIPT%\Scripts\Reg\OEMBackground.reg
regedit -s /i %DEPLOYSCRIPT%\Scripts\Reg\WSUS.reg
regedit -s /i %DEPLOYSCRIPT%\Scripts\Reg\ACU_Default_Run.reg 
 


:Resetting-Automatic-Updates 
net stop "Windows Update"
net stop "Background Intelligent Transfer Service"
del /f /s /q %windir%\SoftwareDistribution\*.*
net start "Background Intelligent Transfer Service"
net start "Windows Update"
wuauclt.exe /resetauthorization /detectnow 

:Clearing-all-event-log 
%WINDIR%\clear_events.cmd 
:Policy Updates
gpupdate /force


endlocal