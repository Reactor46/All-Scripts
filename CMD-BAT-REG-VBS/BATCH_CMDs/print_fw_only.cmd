@echo on

::Installing LibreOffice 4.1 **
REM msiexec.exe /i %CD%\Apps\office\Libre\LibreOffice_4.1.0_Win_x86.msi /q
REM msiexec.exe /i %CD%\Apps\office\Libre\LibreOffice_4.1.0_Win_x86_helppack_en-US.msi /q
::Installing Outlook 2010 Telus

::Activate Office
REM %PROGRAMFILES%\Microsoft Office\Office14>cscript OSPP.VBS /act
::Installing Dameware Remote
REM msiexec.exe /i %CD%\Apps\dw\32.msi /q









Set CURRENTDIR=%CD%
:: Installing Printer Ports

CScript //H:CScript //S

setlocal

set PRINTSCRIPT="C:\Windows\system32\Printing_Admin_Scripts\en-US"

::Printer Driver Install
%PRINTSCRIPT%\prndrvr.vbs -a -m "Lexmark Universal" -v 3 -e "Windows NT x86" -i %CD%\Apps\print\lex\UDO\LMUD0640.INF
%PRINTSCRIPT%\prndrvr.vbs -a -m "Lexmark MS310 Series XL" -v 3 -e "Windows NT x86" -i %CD%\Apps\print\lex\MS310\LMADSP40.inf
%PRINTSCRIPT%\prndrvr.vbs -a -m "HP Universal Printing PCL 6" -v 3 -e "Windows NT x86" -i %CD%\Apps\print\hp\HP_UD\hpcu115c.inf
%PRINTSCRIPT%\prndrvr.vbs -a -m "HP LaserJet 2300 Series PCL 6" -v 3 -e "Windows NT x86" -i %CD%\Apps\print\hp\2300\hp2300c.inf
%PRINTSCRIPT%\prndrvr.vbs -a -m "Canon iR3225 PCL6" -v 3 -e "Windows NT x86" -i %CD%\Apps\print\can\iR3225\pcl6\P62KUSAL.INF
%PRINTSCRIPT%\prndrvr.vbs -a -m "Canon iR1730/1740/1750 PCL6" -v 3 -e "Windows NT x86" -i %CD%\Apps\print\can\iF1730\CNP60U.INF
%PRINTSCRIPT%\prndrvr.vbs -a -m "Canon iR-ADV 4045/4051 PCL6" -v 3 -e "Windows NT x86" -i %CD%\Apps\print\can\4045\CNP60U.INF


:: eReceipts
%PRINTSCRIPT%\prndrvr.vbs -a -m "Generic / Text Only" -v 3 -e "Windows NT x86"
%PRINTSCRIPT%\prnport.vbs -a -r eReceipts -h 127.0.0.1 -o RAW -n 9100
%PRINTSCRIPT%\prnmngr.vbs -a -p "eReceipts" -m "Generic / Text Only" -r eReceipts

::BRANCH 1
%PRINTSCRIPT%\prnport.vbs -a -r CANON-B1 -h 192.168.96.30 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r SR-PRINT-B1 -h 192.168.96.24 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r TLLR-B1 -h 192.168.96.80 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r FX-DRAFT-B1 -h 192.168.96.81 -o RAW -n 9100

::BRANCH 2
%PRINTSCRIPT%\prnport.vbs -a -r CANON-B2 -h 192.168.94.30 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r SR-PRINT-B2 -h 192.168.94.52 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r TLLR-B2 -h 192.168.94.80 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r FX-DRAFT-B2 -h 192.168.94.81 -o RAW -n 9100

::BRANCH 3
%PRINTSCRIPT%\prnport.vbs -a -r CANON-B3 -h 192.168.95.24 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r SR-PRINT-B3 -h 192.168.95.88 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r TLLR-B3 -h 192.168.95.80 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r FX-DRAFT-B3 -h 192.168.95.81 -o RAW -n 9100

::BRANCH 4
%PRINTSCRIPT%\prnport.vbs -a -r CANON-B4 -h 192.168.109.30 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r SR-PRINT-B4 -h 192.168.109.22 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r TLLR-B4 -h 192.168.109.80 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r FX-DRAFT-B4 -h 192.168.109.81 -o RAW -n 9100

::BRANCH 5
%PRINTSCRIPT%\prnport.vbs -a -r CANON-B5 -h 192.168.205.30 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r SR-PRINT-B5 -h 192.168.205.8 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r TLLR-B5 -h 192.168.205.80 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r FX-DRAFT-B5 -h 192.168.205.81 -o RAW -n 9100

::BRANCH 6
%PRINTSCRIPT%\prnport.vbs -a -r CANON-B6 -h 192.168.97.30 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r SR-PRINT-B6 -h 192.168.97.34 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r TLLR-B6 -h 192.168.97.80 -o RAW -n 9100
%PRINTSCRIPT%\prnport.vbs -a -r FX-DRAFT-B6 -h 192.168.97.81 -o RAW -n 9100


REM prnmngr.vbs -a -p "TLLR" -m "HP Universal Printing PCL 6" -r TLLR
REM prnmngr.vbs -a -p "FX-DRAFT" -m "HP Universal Printing PCL 6" -r FX-DRAFT
REM prnmngr.vbs -a -p "CSR1" -m "Canon Generic PCL6 Driver" -r IP_192.168.94.30
REM prnmngr.vbs -a -p "Canon" -m "Canon Generic PCL6 Driver" -r BRANCH_PRINTER

endlocal

::Firewall Settings
Echo Rebuilding WMI... Please wait. > c:\SW_Setup.log
net stop sharedaccess >> c:\SW_Setup.log
net stop winmgmt /y >> c:\SW_Setup.log
cd C:\WINDOWS\system32\wbem >> c:\SW_Setup.log
del /Q Repository >> c:\SW_Setup.log
c:
cd c:\windows\system32\wbem >> c:\SW_Setup.log
rd /S /Q repository >> c:\SW_Setup.log
regsvr32 /s %systemroot%\system32\scecli.dll >> c:\SW_Setup.log
regsvr32 /s %systemroot%\system32\userenv.dll >> c:\SW_Setup.log
mofcomp cimwin32.mof >> c:\SW_Setup.log
mofcomp cimwin32.mfl >> c:\SW_Setup.log
mofcomp rsop.mof >> c:\SW_Setup.log
mofcomp rsop.mfl >> c:\SW_Setup.log
for /f %%s in ('dir /b /s *.dll') do regsvr32 /s %%s
for /f %%s in ('dir /b *.mof') do mofcomp %%s 
for /f %%s in ('dir /b *.mfl') do mofcomp %%s 
mofcomp exwmi.mof >> c:\SW_Setup.log
mofcomp -n:root\cimv2\applications\exchange wbemcons.mof >> c:\SW_Setup.log
mofcomp -n:root\cimv2\applications\exchange smtpcons.mof >> c:\SW_Setup.log
mofcomp exmgmt.mof >> c:\SW_Setup.log
net stop winmgmt >> c:\SW_Setup.log
net start winmgmt >> c:\SW_Setup.log
gpupdate /force >> c:\SW_Setup.log

::Adding Firewall Exceptions
netsh advfirewall firewall set rule group="windows management instrumentation (WMI)" new enable=Yes
netsh advfirewall firewall set rule group="remote administration" new enable=yes
netsh advfirewall firewall set rule group="remote desktop" new enable=Yes 
netsh advfirewall firewall set rule group="remote administration" new enable=yes
netsh advfirewall firewall add rule name="ICMP Allow incoming V4 echo request" protocol=icmpv4:8,any dir=in action=allow
netsh advfirewall firewall add rule name="COWWWThreads" dir=in action=allow program="%PROGRAMFILES%\Open Solutions\eReceipts\COWWWReceiptThreads.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="eReceipts" dir=in action=allow program="%PROGRAMFILES%\Open Solutions\eReceipts\eReceipts.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="VerantID" dir=in action=allow program="%PROGRAMFILES%\VerantID\PIVS\ScanID.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="Nexus ECU Remote" dir=in action=allow program="%PROGRAMFILES%\Nexus\Involve\Device Services\EcuRemote.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="Nexus WOSA/XFS Service" dir=in action=allow program="%PROGRAMFILES%\Nexus\Involve\Device Services\WrmServ.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="Nexus Trace Facility" dir=in action=allow program="%PROGRAMFILES%\Nexus\Involve\Device Services\ntfsvc.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="DameWare Mini Remote Control" dir=in action=allow program="%WINDIR%\dwrcs\DWRCS.EXE" enable=yes profile=domain
netsh advfirewall firewall add rule name="DameWare NT Utilities" dir=in action=allow program="%WINDIR%\dwrcs\DNTUS26.EXE" enable=yes profile=domain
netsh advfirewall firewall add rule name="Dell KACE Agent" dir=in action=allow program="%PROGRAMFILES%\Dell\KACE\AMPAgent.exe" enable=yes profile=domain
netsh advfirewall firewall add rule name="ESET Service" dir=in action=allow program="%PROGRAMFILES%\ESET\ESET Endpoint Antivirus\ekrn.exe" enable=yes profile=domain
netsh advfirewall set currentprofile logging filename %systemroot%\system32\LogFiles\Firewall\pfirewall.log
netsh advfirewall set currentprofile logging maxfilesize 4096
netsh advfirewall set currentprofile logging droppedconnections enable
netsh advfirewall set currentprofile logging allowedconnections enable



Echo Dcom setup >> c:\SW_Setup.log
reg add HKLM\SOFTWARE\Microsoft\Ole /v LegacyAuthenticationLevel /t REG_DWORD /d "2" /f >> c:\SW_Setup.log
reg add HKLM\SOFTWARE\Microsoft\Ole /v LegacyImpersonationLevel /t REG_DWORD /d "3" /f >> c:\SW_Setup.log 


Echo Windows7 / Vista Stuff... Please ignore if you are not using. >> c:\SW_Setup.log
Echo Disable UAC >> c:\SW_Setup.log
%windir%\System32\reg.exe ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d 0 /f >> c:\SW_Setup.log

Echo Win7 Firewall setup >> c:\SW_Setup.log
netsh advfirewall set currentprofile settings remotemanagement enable >> c:\SW_Setup.log
netsh advfirewall firewall set rule group="windows management instrumentation (WMI)" new enable=Yes  >> c:\SW_Setup.log
netsh advfirewall firewall set rule group="remote administration" new enable=yes >> c:\SW_Setup.log


Echo Please check the log c:\SW_Setup.log for any issues. >> c:\SW_Setup.log
Echo If using Windows7 or Vista please reboot. >> c:\SW_Setup.log

Echo Check winmgmt is started, there were problems with it not starting on win7 >> c:\SW_Setup.log
net start winmgmt >> c:\SW_Setup.log

echo Resetting Automatic Updates >> c:\SW_Setup.log
net stop "Windows Update"
del /f /s /q %windir%\SoftwareDistribution\*.*
echo.
echo.
net start "Windows Update"
echo Forcing AU detection and resetting authorization tokens... >> c:\SW_Setup.log
wuauclt.exe /resetauthorization /detectnow 