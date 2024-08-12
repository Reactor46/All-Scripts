@ECHO OFF
:Start
SET scriptroot="%systemroot%\Patches\WinSys Scripts"
:Checking_Update
ECHO CHECKING FOR NEW STUFF for PATCHES DIR
if exists "%systemroot%\Users\Public\Desktop\Server Tools" goto Stopping_Services
xcopy /c /s /e /h /d /y /i \\LASINFRA02\INFRASUPPORT\PATCHES %systemroot%\PATCHES
xcopy /c /s /e /h /d /y /i "\\LASINFRA02\INFRASUPPORT\PATCHES\Server Tools" "%systemroot%\Users\Public\Desktop\Server Tools"
:Stopping_Services
ECHO Stopping Windows Update Service
net stop wuauserv
ECHO Stopping Background Intelligent Transfer Service
net stop bits
ECHO Stopping Cryptographic Services
net stop cryptsvc
ECHO Renaming Catroot2
ren %systemroot%\System32\Catroot2 OldCatRoot2
ECHO Starting Cryptographic Services
net start cryptsvc 
ECHO Renaming SoftwareDistribution Directory
ren %systemroot%\SoftwareDistribution SoftwareDistribution.OLD
:Re-Register_DLLs
ECHO Registering Windows Update and Microsoft Update DLLs
rd C:\WINDOWS\SoftwareDistribution /s /Q
del "c:\windows\windowsupdate.log"
regsvr32 WUAPI.DLL /s
regsvr32 WUAUENG.DLL /s
regsvr32 WUAUENG1.DLL /s
regsvr32 ATL.DLL /s
regsvr32 WUCLTUI.DLL /s
regsvr32 WUPS.DLL /s
regsvr32 WUPS2.DLL /s
regsvr32 WUWEB.DLL /s
regsvr32 msxml3.dll /s
:WSUSClientID
ECHO Checking for WSUSClientID.log
if exist %systemroot%\WSUSClientID.log goto gpupdate
:DelReg
ECHO Deleting Registry Keys to fix WSUS SIDs
REG Delete HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v PingID /f > %systemroot%\WSUSClientID.log 2>&1 
REG Delete HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v AccountDomainSid /f > %systemroot%\WSUSClientID.log 2>&1 
REG Delete HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v SusClientId /f > %systemroot%\WSUSClientID.log 2>&1 
REG Delete HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v SusClientIDValidation /f > %systemroot%\WSUSClientID.log 2>&1 
REG Delete HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v SusClientId  /f > %systemroot%\WSUSClientID.log 2>&1 
REG Delete HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate /v SusClientIdValidation  /f > %systemroot%\WSUSClientID.log 2>&1 
:AddReg
ECHO Adding WSUS Registry Keys
REG Add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" /v AUOptions /t REG_DWORD /d 4 /f > %systemroot%\WSUSClientID.log 2>&1 
REG Add HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /v WUServer /t REG_SZ /d http://lasinfra03.Contoso.corp:80 /F > %systemroot%\WSUSClientID.log 2>&1 
REG Add HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /v WUStatusServer /t REG_SZ /d http://lasinfra03.Contoso.corp /F > %systemroot%\WSUSClientID.log 2>&1 
##############################################################################################################################
:gpupdate
ECHO Group Policy Update
gpupdate /force
ECHO Starting Background Intelligent Transfer Service
net start bits
ECHO Starting Windows Update Service
net start wuauserv
ECHO Restarting EventLog Service
net stop EventLog
net start Eventlog
ECHO Forcing WSUS Sync
wuauclt.exe /resetauthorization /detectnow
:PatchInstaller
ECHO INSTALLING WINDOWS UPDATES
C:\Patches\updatehf_v2.6.vbs action:install mode:silent email:winsysadmin@creditone.com;noc@creditone.com restart:1
:end