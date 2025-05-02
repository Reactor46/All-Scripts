@echo off

REM AuditPolicy.cmd
REM (c) 2006 Microsoft Corporation.  All rights reserved.
REM Sample Audit Script to deploy Windows Vista
REM Granular Audit Policy settings.

REM Should be run as a startup script from Group Policy

REM ###################################################
REM Declare Variables so that we only need to edit file
REM names/paths in one location in script
REM ###################################################

set AuditPolicyLog=%systemroot%\temp\auditpolicy.log
set OSVersionSwap=%systemroot%\temp\osversionwap.txt
set OsVersionTxt=%systemroot%\temp\osversion.txt
set MachineDomainTxt=%systemroot%\temp\machinedomain.txt
set MachineDomainSwap=%systemroot%\temp\machinedomainSwap.txt
set ApplyAuditPolicyCMD=applyauditpolicy.cmd
set AuditPolicyTxt=auditpolicy.txt

REM ###################################################
REM Clear Log & start fresh
REM ###################################################

if exist %AuditPolicyLog% del %AuditPolicyLog% /q /f
date /t > %AuditPolicyLog% & time /t >> %AuditPolicyLog%
echo.

REM ###################################################
REM Check OS Version
REM ###################################################

ver | findstr "[" > %OSVersionSwap%
for /f "tokens=2 delims=[" %%i in (%OSVersionSwap%) do echo %%i > %OsVersionTxt%
for /f "tokens=2 delims=] " %%i in (%OsVersionTxt%) do set osversion=%%i
echo OS Version=%osversion% >> %AuditPolicyLog%

REM ###################################################
REM Skip Pre-Vista
REM ###################################################

if "%osversion%" LSS "6.0" exit /b 1

REM ###################################################
REM Get Domain Name
REM ###################################################

WMIC /namespace:\\root\cimv2 path Win32_ComputerSystem get domain /format:list > %MachineDomainSwap%
find /i "Domain=" %MachineDomainSwap% > %MachineDomainTxt%
for /f "Tokens=2 Delims==" %%i in (%MachineDomainTxt%) do set machinedomain=%%i
echo Machine domain=%machinedomain% >> %AuditPolicyLog%

REM ###################################################
REM Copy Script & Policy to Local Directory or Terminate
REM ###################################################

xcopy \\%machinedomain%\netlogon\%ApplyAuditPolicyCMD% %systemroot%\temp\*.* /r /h /v /y
if %ERRORLEVEL% NEQ 0 (
    echo Could not read \\%machinedomain%\netlogon\%ApplyAuditPolicyCMD% >> %AuditPolicyLog%
    exit /b 1
) else (
    echo Copied \\%machinedomain%\netlogon\%ApplyAuditPolicyCMD% to %systemroot%\temp >> %AuditPolicyLog%
)

xcopy \\%machinedomain%\netlogon\%AuditPolicyTxt% %systemroot%\temp\*.* /r /h /v /y
if %ERRORLEVEL% NEQ 0 (
    echo Could not read \\%machinedomain%\netlogon\%AuditPolicyTxt% >> %AuditPolicyLog%
    exit /b 1
) else (
    echo Copied \\%machinedomain%\netlogon\%AuditPolicyTxt% to %systemroot%\temp >> %AuditPolicyLog%
)

REM ###################################################
REM Create Named Scheduled Task to Apply Policy
REM ###################################################

%systemroot%\system32\schtasks.exe /create /ru System /tn audit /sc hourly /mo 1 /f /rl highest /tr "%systemroot%\temp\%ApplyAuditPolicyCMD%"
if %ERRORLEVEL% NEQ 0 (
    echo Failed to create scheduled task for Audit >> %AuditPolicyLog%
    exit /b 1
) else (
    echo Created scheduled task for Audit >> %AuditPolicyLog%
)

REM ###################################################
REM Start Named Scheduled Task to Apply Policy
REM ###################################################

%systemroot%\system32\schtasks.exe /run /tn audit
if %ERRORLEVEL% NEQ 0 (
    Failed to execute scheduled task for Audit >> %AuditPolicyLog%
) else (
    echo Executed scheduled task for Audit >> %AuditPolicyLog%
)