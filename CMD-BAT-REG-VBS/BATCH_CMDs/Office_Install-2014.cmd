@ECHO ON

setlocal 
set DEPLOYSCRIPT=\\branch1dc\deploy$


:Check for LibreOffice Installing
:Libre Office
wmic product where "name like 'LibreOffice%%%%%%'" call uninstall /nointeractive
wmic product where "name like 'OpenOffice%%%%%%'" call uninstall /nointeractive

rmdir /S /Q "%PROGRAMFILES%\LibreOffice 4*"
rmdir /S /Q "%PROGRAMFILES%\LibreOffice*"
rmdir /S /Q "%PROGRAMFILES%\OpenOffice*"
rmdir /S /Q "%PROGRAMFILES%\OpenOffice.org 3*"


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
start /wait %DEPLOYSCRIPT%\Apps\office\ms\setup.exe /adminfile %DEPLOYSCRIPT%\Apps\office\ms\Updates\ACU_SR_AIS.msp

:Activate Office
IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)
:ARP64
CD "C;\Program Files (x86)\Microsoft Office\Office14\"
cscript OSPP.VBS /inpkey:BCTQH-CCX9G-DY6YF-37G8F-73HWC
cscript OSPP.VBS /act

:ARP86
CD "C:\Program Files\Microsoft Office\Office14\"
cscript OSPP.VBS /inpkey:BCTQH-CCX9G-DY6YF-37G8F-73HWC
cscript OSPP.VBS /act

:Policy Updates
gpupdate /force


endlocal