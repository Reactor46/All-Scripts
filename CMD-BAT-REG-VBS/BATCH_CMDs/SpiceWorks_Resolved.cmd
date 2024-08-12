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

Echo Setting up fireall... Please wait. >> c:\SW_Setup.log
netsh advfirewall set service remoteadmin enable >> c:\SW_Setup.log
Echo Enable Ping >> c:\SW_Setup.log
netsh advfirewall set icmpsetting 8 >> c:\SW_Setup.log

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
net stop "Automatic Updates"
del /f /s /q %windir%\SoftwareDistribution\*.*
echo.
echo.
net start "Automatic Updates"
echo Forcing AU detection and resetting authorization tokens... >> c:\SW_Setup.log
wuauclt.exe /resetauthorization /detectnow 