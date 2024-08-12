@echo off

REM ApplyAuditPolicy.cmd
REM (c) 2006 Microsoft Corporation.  All rights reserved.
REM Sample Audit Script to deploy Windows Vista
REM Granular Audit Policy settings.


REM ###################################################
REM Declare Variables so that we only need to edit file
REM names/paths in one location in script
REM ###################################################

set DeleteAudit=DeleteAudit.txt
set AuditPolicyLog=%systemroot%\temp\AuditPolicy.log
set ApplyAuditPolicyLog=%systemroot%\temp\ApplyAuditPolicy.log
set OSVersionSwap=%systemroot%\temp\osversionwap.txt
set OsVersionTxt=%systemroot%\temp\osversion.txt
set MachineDomainTxt=%systemroot%\temp\machinedomain.txt
set MachineDomainSwap=%systemroot%\temp\machinedomainSwap.txt
set ApplyAuditPolicyCMD=ApplyAuditpolicy.cmd
set AuditPolicyTxt=AuditPolicy.txt

REM ###################################################
REM Clear Log & start fresh
REM ###################################################

if exist %ApplyAuditPolicyLog% del %ApplyAuditPolicyLog% /q /f
date /t > %ApplyAuditPolicyLog% & time /t >> %ApplyAuditPolicyLog%
echo.

REM ###################################################
REM Check OS Version
REM ###################################################

ver | findstr "[" > %OSVersionSwap%
for /f "tokens=2 delims=[" %%i in (%OSVersionSwap%) do echo %%i > %OsVersionTxt%
for /f "tokens=2 delims=] " %%i in (%OsVersionTxt%) do set osversion=%%i
echo OS Version=%osversion% >> %ApplyAuditPolicyLog%

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
echo Machine domain=%machinedomain% >> %ApplyAuditPolicyLog%

REM ###################################################
REM Delete Audit Task
REM Should only be used to remove the pseudo-policy from
REM client machines (designed for future Vista revisions
REM where this script will no longer be necessary, and this
REM script needs to be backed out).

REM to use, simply create a file in NETLOGON with a name
REM that matches the contents of DeleteAudit variable (above)
REM ###################################################

if exist \\%machinedomain%\netlogon\%DeleteAudit% (
    %systemroot%\system32\schtasks.exe /delete /tn "Audit" /F
    DEL %AuditPolicyLog%
    DEL %ApplyAuditPolicyLog%
    DEL %OSVersionSwap%
    DEL %OsVersionTxt%
    DEL %MachineDomainTxt%
    DEL %MachineDomainSwap%
    DEL %systemroot%\temp\%ApplyAuditPolicyCMD%
    DEL %systemroot%\temp\%AuditPolicyTxt%
    exit /b 1
) 

REM ###################################################
REM Copy Audit Policy to Local Directory
REM This is tolerant of failures since the copy is just
REM a "cache refresh".
REM ###################################################

xcopy \\%machinedomain%\netlogon\%AuditPolicyTxt% %systemroot%\temp\*.* /r /h /v /y
if %ERRORLEVEL% NEQ 0 (
    echo Could not read \\%machinedomain%\netlogon\%AuditPolicyTxt% so using previous cached copy>> %ApplyAuditPolicyLog%
) else (
    echo Copied \\%machinedomain%\netlogon\%AuditPolicyTxt% to %systemroot%\temp >> %ApplyAuditPolicyLog%
)

REM ###################################################
REM Apply Policy
REM ###################################################

%systemroot%\system32\auditpol.exe /restore /file:%systemroot%\temp\%AuditPolicyTxt%
if %ERRORLEVEL% NEQ 0 (
    Failed to apply audit settings >> %ApplyAuditPolicyLog%
) else (
    echo Successfully applied audit settings >> %ApplyAuditPolicyLog%
)