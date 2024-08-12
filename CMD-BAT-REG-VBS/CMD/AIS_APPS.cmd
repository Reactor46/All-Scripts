@echo on
setlocal 
net use Z: \\branch\deploy$ /user:DOMAIN\admin /persistent:no
set DEPLOYSCRIPT=Z:
set PRINTSCRIPT="C:\Windows\system32\Printing_Admin_Scripts\en-US"
CScript //H:CScript //S

:Installing Outlook 2010 Telus
set ProductName=Office14.OUTLOOK
IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)
:ARP64
reg query HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432NODE\Microsoft\Windows\CurrentVersion\Uninstall\%ProductName%
if NOT %errorlevel%==1 (goto Install_DW)
 
:ARP86
reg query HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%ProductName%
if %errorlevel%==1 (goto DeployOutlook) else (goto Install_DW)
:DeployOutlook
start /wait %DEPLOYSCRIPT%\Apps\office\Outlook-Telus\setup.exe /adminfile %DEPLOYSCRIPT%\Apps\office\Outlook-Telus\Updates\CUSTOM.MSP
:Activate Office
"%PROGRAMFILES%\Microsoft Office\Office14>cscript OSPP.VBS /act"

:Install_DW
IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)
:ARP64
start /wait msiexec.exe /i %DEPLOYSCRIPT%\Apps\dw\64.msi /qn
:ARP86
start /wait msiexec.exe /i %DEPLOYSCRIPT%\Apps\dw\32.msi /qn


:Set scripts Standards
IF NOT "C:\Windows\System32\oobe\info"=="" (goto Background) else (
rmdir /S /Q C:\Windows\System32\oobe\info\backgrounds )
:Background
mkdir C:\Windows\System32\oobe\info\backgrounds
copy %DEPLOYSCRIPT%\Apps\scripts\backgroundDefault.jpg C:\Windows\System32\oobe\info\backgrounds 
:Sysinternals 
xcopy /h /i /c /k /o /r /e /y %DEPLOYSCRIPT%\Apps\scripts\SysinternalsSuite\*.* C:\Windows\ 
copy %DEPLOYSCRIPT%\Apps\scripts\clear_events.cmd C:\Windows\ 
:Importing scripts Registry Settings 
regedit -s /i %DEPLOYSCRIPT%\Apps\scripts\InstallTakeOwnership.reg 
regedit -s /i %DEPLOYSCRIPT%\Apps\scripts\OEMBackground.reg 


:Printer Driver Install
IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)
:ARP64
%PRINTSCRIPT%\prndrvr.vbs -a -m "Canon iR1730/1740/1750 PCL6" -v 3 -e "Windows NT x64" -i %DEPLOYSCRIPT%\Apps\print\can\iR1750-x64\CNP60UA64.INF
%PRINTSCRIPT%\prndrvr.vbs -a -m "Canon iR1020/1024/1025 PCL6" -v 3 -e "Windows NT x64" -i %DEPLOYSCRIPT%\Apps\print\can\iR1025\P664USAL.INF
:ARP86
%PRINTSCRIPT%\prndrvr.vbs -a -m "Canon iR1730/1740/1750 PCL6" -v 3 -e "Windows NT x86" -i %DEPLOYSCRIPT%\Apps\print\can\iR1750\CNP60U.INF
%PRINTSCRIPT%\prndrvr.vbs -a -m "Canon iR1020/1024/1025 PCL6" -v 3 -e "Windows NT x86" -i %DEPLOYSCRIPT%\Apps\print\can\iR1025\P62KUSAL.INF

::AIS 1 Canon iR1025
%PRINTSCRIPT%\prnport.vbs -a -r AIS-1 -h 192.168.96.32 -o RAW -n 9100

::AIS 2 Canon iR1750iF
%PRINTSCRIPT%\prnport.vbs -a -r AIS-2 -h 192.168.94.31 -o RAW -n 9100

::AIS 6 Canon iR1750iF
%PRINTSCRIPT%\prnport.vbs -a -r AIS-6 -h 192.168.97.31 -o RAW -n 9100


:Adding Firewall Exceptions
netsh advfirewall firewall set rule group="windows management instrumentation (WMI)" new enable=Yes
netsh advfirewall firewall set rule group="remote administration" new enable=yes
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes 
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes 
netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol=icmpv4:8,any dir=in action=allow
netsh advfirewall firewall add rule name="DameWare Mini Remote Control" dir=in action=allow program="%WINDIR%\dwrcs\DWRCS.EXE" enable=yes profile=domain
netsh advfirewall firewall add rule name="DameWare NT Utilities" dir=in action=allow program="%WINDIR%\dwrcs\DWRCST.EXE" enable=yes profile=domain
netsh advfirewall firewall add rule name="Dell KACE Agent" dir=in action=allow program="%PROGRAMFILES%\Dell\KACE\AMPAgent.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="ESET Service" dir=in action=allow program="%PROGRAMFILES%\ESET\ESET Endpoint Antivirus\ekrn.exe" enable=yes profile=domain
netsh advfirewall set currentprofile logging filename %systemroot%\system32\LogFiles\Firewall\pfirewall.log
netsh advfirewall set currentprofile logging maxfilesize 4096
netsh advfirewall set currentprofile logging droppedconnections enable
netsh advfirewall set currentprofile logging allowedconnections enable
netsh firewall set service type=remoteadmin mode=enable
netsh advfirewall firewall set rule group="remote administration" new enable=yes


:Resetting-Automatic-Updates 
net stop "Windows Update" 
del /f /s /q C:\Windows\SoftwareDistribution\*.* 
:Clearing-all-event-log 
%WINDIR%\clear_events.cmd 
:Restarting-Windows-Updates 
net start "Windows Update" 
wuauclt.exe /resetauthorization /detectnow 

