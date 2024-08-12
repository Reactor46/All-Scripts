@echo off
ECHO ===============================================================================
ECHO ===============================================================================
ECHO.
ECHO This script installs security baselines into local policy for Windows 10.
ECHO.
ECHO Press Ctrl+C to stop the installation, or press any other key to continue...
ECHO.
ECHO ===============================================================================
ECHO ===============================================================================
PAUSE > nul

:: Make the directory where this script lives the current dir.
PUSHD %~dp0
SET SECGUIDE=%CD%
SET SECGUIDELOGS=%SECGUIDE%\LOGS
MD "%SECGUIDELOGS%" 2> nul


ECHO Installing Windows 10 security settings and policies...
:: Apply Windows 10 Security
Tools\LGPO.exe /s Windows_10\GptTmpl.inf /m Windows_10\machine_registry.pol /u Windows_10\user_registry.pol /ac Windows_10\audit.csv /e mitigation /e audit /e zone > "%SECGUIDELOGS%%\Win10-install.log"
echo Windows 10 Local Policy Applied


ECHO Installing Internet Explorer 11 policies...
:: Apply Internet Explorer 11 Local Policy
Tools\LGPO.exe /m IE_11\machine_registry.pol /u IE_11\user_registry.pol > "%SECGUIDELOGS%%\IE_11-install.log"
echo Internet Explorer 11 Local Policy Applied


:: Copy Custom Administrative Templates
ECHO Copying custom administrative templates
copy /y Templates\*.admx %SystemRoot%\PolicyDefinitions
copy /y Templates\*.adml %SystemRoot%\PolicyDefinitions\en-US


::Display Notifications
ECHO.
ECHO.
ECHO ===============================================================================
ECHO ===============================================================================
ECHO.
ECHO In order to test properly, create a new non-administrative user account and
ECHO reboot.
ECHO.
ECHO Additionally, check log files located in this directory:
ECHO.    %SECGUIDELOGS%
ECHO.
ECHO Feedback can be directed to the following:  SecGuide@Microsoft.com
ECHO.
ECHO ===============================================================================
ECHO ===============================================================================
ECHO.
POPD