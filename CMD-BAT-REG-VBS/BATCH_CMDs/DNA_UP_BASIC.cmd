@ECHO ON

setlocal 
set DEPLOYSCRIPT=\\branch1dc\deploy$
REM set DEPLOYSCRIPT=\\branch2dc\deploy$
REM set DEPLOYSCRIPT=\\branch3dc\deploy$
REM set DEPLOYSCRIPT=\\branch4dc\deploy$
REM set DEPLOYSCRIPT=\\branch5dc\deploy$
REM set DEPLOYSCRIPT=\\branch6dc\deploy$


:Sig
C:\Windows\SigPlus\UNWISE.EXE /S C:\Windows\SigPlus\INSTALL.LOG
RD /S /Q "C:\Windows\SigPlus" 
del /S /Q /F "C:\Windows\SigPlus.ini"
:Notepad++ 
Start /wait msiexec.exe /X {8EED4C6D-AA5E-4C75-AEB8-22CD48D438B3} /qn 
"C:\Program Files\Notepad++\uninstall.exe" /S 
wmic product where "name like 'Notepad++%%'" call uninstall /nointeractive 
RD /S /Q "C:\Program Files\Notepad++\"
:Peazip
"C:\Program Files\PeaZip\unins000.exe" /verysilent 
RD /S /Q "C:\Program Files\PeaZip\"

:JavaInstall 
START /WAIT %DEPLOYSCRIPT%\Apps\java\32.exe /s /v"/passive /norestart AUTOUPDATECHECK=0 JAVAUPDATE=0 JU=0 WEBSTARTICON=0 EULA=0 SYSTRAY=0" 

:eReceipts 
Start /wait %DEPLOYSCRIPT%\Apps\eReceipts\15.exe /s 

copy %DEPLOYSCRIPT%\Apps\ereceipts\global.ini "C:\Program Files\Open Solutions\DDS\eReceipts\" /y 
copy %DEPLOYSCRIPT%\Apps\ereceipts\workstation.ini "C:\Program Files\Open Solutions\DDS\eReceipts\" /y 
copy %DEPLOYSCRIPT%\Apps\ereceipts\logo.bmp "C:\Program Files\Open Solutions\DDS\eReceipts\" /y 
copy %DEPLOYSCRIPT%\Apps\ereceipts\eReceipts.lnk "%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Startup\" /y 
copy %DEPLOYSCRIPT%\Apps\ereceipts\eReceipts.lnk "C:\Users\Public\Desktop\" /y

:VerantID
msiexec.exe /i %DEPLOYSCRIPT%\Apps\VerantID\scanid552.msi /qn 

:Sig_Install 
Start /wait %DEPLOYSCRIPT%\Apps\sig\sig.exe /s 
Start /wait %DEPLOYSCRIPT%\Apps\sig\SigPlusAdjust.exe /s
copy %DEPLOYSCRIPT%\Apps\sig\SigPlus.ini C:\Windows /y

:NP_Install
start /wait %DEPLOYSCRIPT%\Apps\notepad++\np.exe /S 

:PZ_Install 
start /wait %DEPLOYSCRIPT%\Apps\peazip\pz5.exe /verysilent 

:Installing-Updating DameWare 
start /wait msiexec.exe /i %DEPLOYSCRIPT%\Apps\dw\32.msi /qn 
:DotNet 4.0
Start /wait %DEPLOYSCRIPT%\Apps\acu\dotNetFx40_Full_x86_x64.exe /passive /norestart
Start /wait %DEPLOYSCRIPT%\Apps\acu\dotNetFx45_Full_setup.exe /passive /norestart

:Set ACU Standards
IF NOT "C:\Windows\System32\oobe\info"=="" (goto Background) else (
rmdir /S /Q C:\Windows\System32\oobe\info\backgrounds )
:Background
mkdir C:\Windows\System32\oobe\info\backgrounds
copy %DEPLOYSCRIPT%\Apps\acu\backgroundDefault.jpg C:\Windows\System32\oobe\info\backgrounds 

:Sysinternals 
xcopy /h /i /c /k /o /r /e /y %DEPLOYSCRIPT%\Apps\acu\SysinternalsSuite\*.* C:\Windows\ 
copy %DEPLOYSCRIPT%\Apps\acu\clear_events.cmd C:\Windows\ 

:Importing ACU Registry Settings 
regedit -s /i %DEPLOYSCRIPT%\Apps\acu\InstallTakeOwnership.reg 
regedit -s /i %DEPLOYSCRIPT%\Apps\acu\OEMBackground.reg 
regedit -s /i %DEPLOYSCRIPT%\Apps\acu\IE_Settings.reg 
regedit -s /i %DEPLOYSCRIPT%\Apps\acu\ZoneMaps.reg 
regedit -s /i %DEPLOYSCRIPT%\Apps\acu\Win32TM.reg 

:Clear username from Win7 login 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v LastLoggedOnSAMUser /t REG_SZ /d "" /f 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v LastLoggedOnUser /t REG_SZ /d "" /f 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v DontDisplayLastUserName /t REG_DWORD /d 1 /f 

:Disable Java Update Tab and also Updates and Notifications 
reg add "HKLM\SOFTWARE\JavaSoft\Java Update\Policy" /v EnableJavaUpdate /t REG_DWORD /d 00000000 /f 
reg add "HKLM\SOFTWARE\JavaSoft\Java Update\Policy" /v NotifyDownload /t REG_DWORD /d 00000000 /f 
:Remove IE10 if Installed 
FORFILES /P C:\Windows\servicing\Packages /M Microsoft-Windows-InternetExplorer-*10.*.mum /c "cmd /c echo Uninstalling package @fname && start /w pkgmgr /up:@fname /quiet /norestart" 
::Block IE10 Install  
reg add "HKLM\SOFTWARE\Microsoft\Internet Explorer\Setup\10.0" /v DoNotAllowIE10 /t REG_DWORD /d 1 /f 

:Disable UAC 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v EnableLUA /t REG_DWORD /d 0 /f 

:Disable Windows Backup 
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsBackup" /v DisableMonitoring /t REG_DWORD /d 0 /f 

:Adding Firewall Exceptions 
netsh advfirewall firewall set rule group="windows management instrumentation (WMI)" new enable=Yes 
netsh advfirewall firewall set rule group="remote administration" new enable=yes 
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes 
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes 
netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol=icmpv4:8,any dir=in action=allow 
netsh advfirewall firewall add rule name="COWWWThreads" dir=in action=allow program="C:\Program Files\Open Solutions\eReceipts\COWWWReceiptThreads.exe" enable=yes profile=domain 
netsh advfirewall firewall add rule name="eReceipts" dir=in action=allow program="C:\Program Files\Open Solutions\eReceipts\eReceipts.exe" enable=yes profile=domain 
netsh advfirewall firewall add rule name="VerantID" dir=in action=allow program="C:\Program Files\VerantID\PIVS\ScanID.exe" enable=yes profile=domain 
netsh advfirewall firewall add rule name="Nexus ECU Remote" dir=in action=allow program="C:\Program Files\Nexus\Involve\Device Services\EcuRemote.exe" enable=yes profile=domain 
netsh advfirewall firewall add rule name="Nexus WOSA/XFS Service" dir=in action=allow program="C:\Program Files\Nexus\Involve\Device Services\WrmServ.exe" enable=yes profile=domain 
netsh advfirewall firewall add rule name="Nexus Trace Facility" dir=in action=allow program="C:\Program Files\Nexus\Involve\Device Services\ntfsvc.exe" enable=yes profile=domain 
netsh advfirewall firewall add rule name="DameWare Mini Remote Control" dir=in action=allow program="C:\Windows\dwrcs\DWRCS.EXE" enable=yes profile=domain 
netsh advfirewall firewall add rule name="DameWare NT Utilities" dir=in action=allow program="C:\Windows\dwrcs\DWRCST.EXE" enable=yes profile=domain 
netsh advfirewall firewall add rule name="Dell KACE Agent" dir=in action=allow program="C:\Program Files\Dell\KACE\AMPAgent.exe" enable=yes profile=domain 
netsh advfirewall firewall add rule name="ESET Service" dir=in action=allow program="C:\Program Files\ESET\ESET Endpoint Antivirus\ekrn.exe" enable=yes profile=domain 
netsh advfirewall set currentprofile logging filename %systemroot%\system32\LogFiles\Firewall\pfirewall.log 
netsh advfirewall set currentprofile logging maxfilesize 4096 
netsh advfirewall set currentprofile logging droppedconnections enable 
netsh advfirewall set currentprofile logging allowedconnections enable 


:List Products Installed 
wmic product get name,vendor,version >> %DEPLOYSCRIPT%\Logs\%COMPUTERNAME%-Apps.txt

:Resetting Automatic Updates 
net stop "Windows Update" 
del /f /s /q C:\Windows\SoftwareDistribution\*.* 
:Clearing all event log 
Start /wait clear_events.cmd 
:Restarting Windows Updates 
net start "Windows Update" 
Start /wait wuauclt.exe /a /detectnow /updatenow

endlocal

:Reboot
REM C:\Windows\system32\shutdown.exe" /r /t 300
