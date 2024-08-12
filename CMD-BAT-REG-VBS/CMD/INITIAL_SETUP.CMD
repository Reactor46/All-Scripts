@ECHO ON
setlocal 
set DEPLOYSCRIPT=\\SERVER\deploy$



:AddingFirewallExceptions
netsh advfirewall firewall set rule group="windows management instrumentation (WMI)" new enable=Yes
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes 
netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol=icmpv4:8,any dir=in action=allow
netsh advfirewall firewall add rule name="DameWare Mini Remote Control" dir=in action=allow program="%WINDIR%\dwrcs\DWRCS.EXE" enable=yes profile=domain
netsh advfirewall firewall add rule name="DameWare NT Utilities" dir=in action=allow program="%WINDIR%\dwrcs\DWRCST.EXE" enable=yes profile=domain
netsh advfirewall firewall add rule name="ESET Service" dir=in action=allow program="%PROGRAMFILES%\ESET\ESET Endpoint Antivirus\ekrn.exe" enable=yes profile=domain
netsh advfirewall set currentprofile logging filename %systemroot%\system32\LogFiles\Firewall\pfirewall.log
netsh advfirewall set currentprofile logging maxfilesize 4096
netsh advfirewall set currentprofile logging droppedconnections enable
netsh advfirewall set currentprofile logging allowedconnections enable

:Chocolatey Install
@powershell -NoProfile -ExecutionPolicy unrestricted -Command "iex ((new-object net.webclient).DownloadString('https://chocolatey.org/install.ps1'))" && SET PATH=%PATH%;%systemdrive%\chocolatey\bin

:QuickApps
cinst notepadplusplus.install peazip adobereader adobeshockwaveplayer flashplayerplugin flashplayeractivex AdobeAIR paint.net Silverlight k-litecodecpackmega 


:Set ACU Standards
IF NOT "C:\Windows\System32\oobe\info"=="" (goto Background) else (
rmdir /S /Q C:\Windows\System32\oobe\info\backgrounds )
:Background
mkdir C:\Windows\System32\oobe\info\backgrounds
copy %DEPLOYSCRIPT%\Apps\acu\backgroundDefault.jpg C:\Windows\System32\oobe\info\backgrounds
:Sysinternals 
xcopy /h /i /c /k /o /r /e /y %DEPLOYSCRIPT%\Apps\acu\SysinternalsSuite\*.* C:\Windows\ 
copy %DEPLOYSCRIPT%\Apps\scripts\clear_events.cmd C:\Windows\ 
copy %DEPLOYSCRIPT%\Apps\scripts\NTP_SET_CLIENT.CMD C:\Windows\

:Regedits
regedit -s /i %DEPLOYSCRIPT%\Apps\scripts\InstallTakeOwnership.reg 
regedit -s /i %DEPLOYSCRIPT%\Apps\scripts\OEMBackground.reg 
:Clear username from Win7 login 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v LastLoggedOnSAMUser /t REG_SZ /d "" /f 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v LastLoggedOnUser /t REG_SZ /d "" /f 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v DontDisplayLastUserName /t REG_DWORD /d 1 /f 

:Enable RDP
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f

:Disable UAC 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v EnableLUA /t REG_DWORD /d 0 /f 

:Disable Windows Backup 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsBackup" /v DisableMonitoring /t REG_DWORD /d 0 /f 

:Set Home Page
REG ADD "HKLM\Software\Microsoft\Internet Explorer\Main" /t REG_SZ /v Start page /d http://intranet/
REG ADD "HKCU\Software\Microsoft\Internet Explorer\Main" /t REG_SZ /v Default_Page_URL /d http://intranet/

:W32Time
net stop w32time
w32tm /register
w32tm /unregister
w32tm /register
Echo "Configuring DOMAIN as update source"
w32tm /config /syncfromflags:domhier /update /reliable:yes
net start "w32time"
Echo "Updating"
w32tm /resync /rediscover
Echo "Check Peer list"
w32tm /query /peers
Echo "Check status"
w32tm /query /status

:Windows Updates
net stop "Windows Update"
net stop "Background Intelligent Transfer Service"
del /f /s /q %windir%\SoftwareDistribution\*.*
net start "Background Intelligent Transfer Service"
net start "Windows Update"
wuauclt.exe /resetauthorization /detectnow 

:GPUpdate
gpupdate /force

