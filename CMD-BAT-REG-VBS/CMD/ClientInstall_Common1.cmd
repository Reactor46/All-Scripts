@echo off
::
:: Common components for client install scripts. Intended to be called only from those main scripts.
::

SET SECGUIDE=%CD%
SET SECGUIDELOGS=%SECGUIDE%\LOGS
MD "%SECGUIDELOGS%" 2> nul


:: Update this with help from MapGuidsToGpoNames.ps1:
::
SET GPO_Win10Computer={3C537678-BBC4-4F10-AACE-5D44C468CC6C}
SET GPO_Win10User={8E3D0A57-07DB-4723-B072-A6374FCC0779}
SET GPO_Win10BitLocker={2BE77E31-F6AC-479A-8DA4-093B6DA8F349}
SET GPO_WDAV={6D1D9215-0E34-4AD9-B4B9-B5ED8B8C83DA}
SET GPO_CredGuard={1CBF32AF-581B-400B-B3D7-304B50200C36}
SET GPO_IE11Computer={3BF11821-7EF8-43F9-9CBB-87D79F78D564}
SET GPO_IE11User={4B434505-3663-4D3E-80EE-5E5B3334D6D6}
SET GPO_DomainSec={BEEC0E5D-EEDC-44BC-9D28-41693B3CE82A}


ECHO Installing Windows 10 security settings and policies...
:: Create local directory for Exploit Protection and copy file locally
copy /y ConfigFiles\EP.xml %TEMP%
powershell.exe Set-ProcessMitigation -PolicyFilePath %TEMP%\EP.xml
del %TEMP%\EP.xml


ECHO Configuring Client Side Extensions...
:: Configure Client Side Extensions
Tools\LGPO.exe /v /e mitigation /e audit /e zone > "%SECGUIDELOGS%%\CSE-install.log"
echo Client Side Extensions Configured


:: Apply Windows 10 GPOs
Tools\LGPO.exe /v /g  ..\GPOs\%GPO_Win10Computer%  > "%SECGUIDELOGS%%\Win10-install.log"
Tools\LGPO.exe /v /g  ..\GPOs\%GPO_Win10User%     >> "%SECGUIDELOGS%%\Win10-install.log"
echo Windows 10 Local Policy Applied


ECHO Installing Internet Explorer 11 policies...
:: Apply Internet Explorer 11 Local Policy
Tools\LGPO.exe /v /g ..\GPOs\%GPO_IE11Computer%  > "%SECGUIDELOGS%%\IE_11-install.log"
Tools\LGPO.exe /v /g ..\GPOs\%GPO_IE11User%     >> "%SECGUIDELOGS%%\IE_11-install.log"
echo Internet Explorer 11 Local Policy Applied


ECHO Installing Windows Defender Antivirus policies...
:: Apply Windows Defender Local Policy
Tools\LGPO.exe /v /g ..\GPOs\%GPO_WDAV% > "%SECGUIDELOGS%%\Windows_DefenderAV-install.log"
echo Windows Defender Antivirus Local Policy Applied


ECHO Installing Windows Credential Guard policies...
:: Apply Windows Credential Guard Local Policy
Tools\LGPO.exe /v /g ..\GPOs\%GPO_CredGuard% > "%SECGUIDELOGS%%\Cred_Guard-install.log"
echo Windows Credential Guard Local Policy Applied


ECHO Installing Windows BitLocker policies...
:: Apply Windows BitLocker Local Policy
Tools\LGPO.exe /v /g ..\GPOs\%GPO_Win10BitLocker% > "%SECGUIDELOGS%%\BitLocker-install.log"
echo Windows BitLocker Local Policy Applied


ECHO Installing Domain Security policies...
:: Apply Domain Security Policy
Tools\LGPO.exe /v /g ..\GPOs\%GPO_DomainSec% > "%SECGUIDELOGS%%\Domain-install.log"
echo Domain Security Policy Applied


:: Copy Custom Administrative Templates
ECHO Copying custom administrative templates
copy /y ..\Templates\*.admx %SystemRoot%\PolicyDefinitions
copy /y ..\Templates\*.adml %SystemRoot%\PolicyDefinitions\en-US

:: Disabling Scheduled Tasks
SCHTASKS.EXE /Change /TN \Microsoft\XblGameSave\XblGameSaveTask      /DISABLE

