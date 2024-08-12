@ECHO ON

setlocal 
set DEPLOYSCRIPT=\\branch1dc\deploy$


:Set ACU Standards
IF NOT "C:\Windows\System32\oobe\info"=="" (goto Background) else (
rmdir /S /Q C:\Windows\System32\oobe\info\backgrounds )

:Background
mkdir C:\Windows\System32\oobe\info\backgrounds
copy %DEPLOYSCRIPT%\Apps\acu\backgroundDefault.jpg C:\Windows\System32\oobe\info\backgrounds 

:Sysinternals 
xcopy /h /i /c /k /o /r /e /y %DEPLOYSCRIPT%\Apps\acu\SysinternalsSuite\*.* C:\Windows\ 

copy %DEPLOYSCRIPT%\Apps\acu\clear_events.cmd C:\Windows\ 
copy %DEPLOYSCRIPT%\Apps\acu\NTP_SET_CLIENT.CMD C:\Windows\

:Importing ACU Registry Settings 
regedit -s /i %DEPLOYSCRIPT%\Apps\acu\InstallTakeOwnership.reg 
regedit -s /i %DEPLOYSCRIPT%\Apps\acu\OEMBackground.reg 
regedit -s /i %DEPLOYSCRIPT%\Apps\acu\ACU_Default_Run.reg

:Clear username from Win7 login 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v LastLoggedOnSAMUser /t REG_SZ /d "" /f 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v LastLoggedOnUser /t REG_SZ /d "" /f 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v DontDisplayLastUserName /t REG_DWORD /d 1 /f 

:Enable RDP
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f

:Disable Java Update Tab and also Updates and Notifications 
reg add "HKLM\SOFTWARE\JavaSoft\Java Update\Policy" /v EnableJavaUpdate /t REG_DWORD /d 00000000 /f 
reg add "HKLM\SOFTWARE\JavaSoft\Java Update\Policy" /v NotifyDownload /t REG_DWORD /d 00000000 /f 

:Block IE10 Install  
reg add "HKLM\SOFTWARE\Microsoft\Internet Explorer\Setup\10.0" /v DoNotAllowIE10 /t REG_DWORD /d 1 /f 

:Block IE11 Install  
reg add "HKLM\SOFTWARE\Microsoft\Internet Explorer\Setup\11.0" /v DoNotAllowIE11 /t REG_DWORD /d 1 /f 


:Disable UAC 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v EnableLUA /t REG_DWORD /d 0 /f 

:Disable Windows Backup 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsBackup" /v DisableMonitoring /t REG_DWORD /d 0 /f 

:Set Home Page
REG ADD "HKLM\Software\Microsoft\Internet Explorer\Main" /t REG_SZ /v Start page /d http://intranet/
REG ADD "HKCU\Software\Microsoft\Internet Explorer\Main" /t REG_SZ /v Start page /d https://www.aldergrovecu.ca/Personal/
REG ADD "HKCU\Software\Microsoft\Internet Explorer\Main" /t REG_SZ /v Default_Page_URL /d http://intranet/

:Adding Firewall Exceptions 
netsh advfirewall firewall set rule group="windows management instrumentation (WMI)" new enable=Yes 
netsh advfirewall firewall set rule group="Windows Remote Management" new enable=yes
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes
netsh advfirewall set currentprofile logging filename %systemroot%\system32\LogFiles\Firewall\pfirewall.log 
netsh advfirewall set currentprofile logging maxfilesize 4096 
netsh advfirewall set currentprofile logging droppedconnections enable 
netsh advfirewall set currentprofile logging allowedconnections enable  
netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol=icmpv4:8,any dir=in action=allow 
netsh advfirewall firewall add rule name="DameWare Mini Remote Control" dir=in action=allow program="C:\Windows\dwrcs\DWRCS.EXE" enable=yes profile=domain 
netsh advfirewall firewall add rule name="DameWare NT Utilities" dir=in action=allow program="C:\Windows\dwrcs\DWRCST.EXE" enable=yes profile=domain 
netsh advfirewall firewall add rule name="Dell KACE Agent" dir=in action=allow program="C:\Program Files\Dell\KACE\AMPAgent.exe" enable=yes profile=domain 
netsh advfirewall firewall add rule name="ESET Service" dir=in action=allow program="C:\Program Files\ESET\ESET Endpoint Antivirus\ekrn.exe" enable=yes profile=domain 


:W32Time
net stop w32time
w32tm /register
w32tm /unregister
w32tm /register
Echo "Configuring ALDERGROVE DOMAIN as update source"
w32tm /config /syncfromflags:domhier /update /reliable:yes
net start "w32time"
Echo "Updating"
w32tm /resync /rediscover
Echo "Check Peer list"
w32tm /query /peers
pause
Echo "Check status"
w32tm /query /status


:Resetting Automatic Updates 
net stop "Windows Update" 
del /f /s /q C:\Windows\SoftwareDistribution\*.* 
:Clearing all event log 
Start /wait %WINDIR%\clear_events.cmd 
:Restarting Windows Updates 
net start "Windows Update" 
wuauclt.exe /resetauthorization /detectnow 

:Installing Outlook 2010 Telus
start /wait %DEPLOYSCRIPT%\Apps\office\Outlook-Telus\setup.exe /adminfile %CD%\Apps\office\Outlook-Telus\Updates\ACU.MSP
::Activate Office
"%PROGRAMFILES%\Microsoft Office\Office14>cscript OSPP.VBS /act"

::Installing Dameware
IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)
:ARP64
start /wait msiexec.exe /i %DEPLOYSCRIPT%\Apps\dw\64.msi /qn
:ARP86
start /wait msiexec.exe /i %DEPLOYSCRIPT%\Apps\dw\32.msi /qn

::Installing ESET
IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)
:ARP64
msiexec.exe /i %DEPLOYSCRIPT%\Apps\eset\ESET_64_JULY_2013.msi /qn REBOOT="ReallySuppress"
:ARP86
msiexec.exe /i %DEPLOYSCRIPT%\Apps\eset\ESET_32_JULY_2013.msi /qn REBOOT="ReallySuppress"


:Chocolatey Install
@powershell -NoProfile -ExecutionPolicy unrestricted -Command "iex ((new-object net.webclient).DownloadString('https://chocolatey.org/install.ps1'))" && SET PATH=%PATH%;%systemdrive%\chocolatey\bin

:QuickApps
cinst notepadplusplus.install peazip javaruntime adobereader flashplayeractivex flashplayerplugin adobeshockwaveplayer AdobeAIR ccleaner libreoffice libreoffice-help
:Activate Office Install
%PROGRAMFILES%\Microsoft Office\Office14>cscript OSPP.VBS /act

endlocal