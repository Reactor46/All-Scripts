setlocal

REM *********************************************************************
REM Environment customization begins here. Modify variables below.
REM *********************************************************************

REM Get ProductName from the Office product's core Setup.xml file, and then add "office14." as a prefix. 
set ProductName=Office14.STANDARD

REM Set DeployServer to a network-accessible location containing the Office source files.
set DeployServer=\\branch1dc\o2k1032$

REM Set ConfigFile to the configuration file to be used for deployment (required)
REM set ConfigFile=\\branch1dc\o2k1032$\Standard.WW\config.xml

REM Set LogLocation to a central directory to collect log files.
set LogLocation=\\branch1dc\o2k1032$\office2010Logs

REM *********************************************************************
REM Deployment code begins here. Do not modify anything below this line.
REM *********************************************************************

IF NOT "%ProgramFiles(x86)%"=="" (goto ARP64) else (goto ARP86)

REM Operating system is X64. Check for 32 bit Office in emulated Wow6432 uninstall key
:ARP64
reg query HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432NODE\Microsoft\Windows\CurrentVersion\Uninstall\%ProductName%
if NOT %errorlevel%==1 (goto End)

REM Check for 32 and 64 bit versions of Office 2010 in regular uninstall key.(Office 64bit would also appear here on a 64bit OS) 
:ARP86
reg query HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%ProductName%
if %errorlevel%==1 (goto DeployOffice) else (goto End)

REM If 1 returned, the product was not found. Run setup here.
:DeployOffice
start /wait %DeployServer%\setup.exe /adminfile %DeployServer%\Updates\ACU.msp
echo %date% %time% Setup ended with error code %errorlevel%. >> %LogLocation%\%computername%.txt

REM If 0 or other was returned, the product was found or another error occurred. Do nothing.
:End

Endlocal