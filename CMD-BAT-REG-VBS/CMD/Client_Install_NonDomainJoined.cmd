@echo off
ECHO ===============================================================================
ECHO ===============================================================================
ECHO.
ECHO This script installs security baselines into local policy for Windows 10,
ECHO for a NON-DOMAIN-JOINED system.
ECHO.
ECHO Press Ctrl+C to stop the installation, or press any other key to continue...
ECHO.
ECHO ===============================================================================
ECHO ===============================================================================
PAUSE > nul

:: Make the directory where this script lives the current dir.
PUSHD %~dp0

:: Verify that LGPO.exe is present before continuing
IF EXIST Tools\LGPO.exe GOTO LgpoPresent
ECHO.
ECHO.
ECHO LGPO.exe must be downloaded and copied into the Tools directory before proceeding.
ECHO.
ECHO Exiting...
ECHO.
POPD
GOTO :EOF

:LgpoPresent
CALL Tools\ClientInstall_Common1.cmd

ECHO Adjusting some settings for non-domain-joined:
Tools\LGPO.exe /v /s ConfigFiles\DeltaForNonDomainJoined.inf /t ConfigFiles\DeltaForNonDomainJoined.txt > "%SECGUIDELOGS%%\NonDomainAdjustments.log"

CALL Tools\ClientInstall_Common2.cmd

POPD
