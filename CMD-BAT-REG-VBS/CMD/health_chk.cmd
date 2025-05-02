@echo off

REM HEALTH_CHK.CMD - Retrieves FRS-related state information from the specified DC or Server.
set HC_VERSION=v5.1a

SETLOCAL ENABLEEXTENSIONS

set toollist=reg repadmin ntfrsutl dcdiag netdiag

if NOT "%1" == "" goto :CHECK_TOOLS
:USAGE
    echo.
    echo Usage:   health_chk  result_dir   [target_computername]
    echo     Retrieve state info from the specified DC or Server running FRS.
    echo.
    echo     result_dir is created if it does not exist.  No trailing backslash.
    echo     Target_ComputerName is optional.  Default is current computer.
    echo         It can be a netbios name with no leading slashes or a full
    echo         dns name, xxx.yyy.zzz.com
    echo.
    echo     This script uses NTFRSUTL.EXE to gather data from FRS on the
    echo     Target computer.  This tool must be in your path and can be found
    echo     in the resource kit or Windows Server 2003 support tools folder.  
    echo     The NTFRS service must be running on the target computer.  
    echo.
    echo     Health_chk %HC_VERSION% is intended to be run on Windows Server 2003
    echo     systems. It can collect data from Windows 2000 or Windows Server 2003
    echo     systems.
    echo.
    echo     Health_chk %HC_VERSION% uses the following tools to gather information:
    echo       %toollist% eventquery.vbs
    echo ----
    echo.
    goto :QUIT


:CHECK_TOOLS

REM  see if we have ntfrsutl.exe
ntfrsutl > nul: 2> nul:
if  ERRORLEVEL 1  (
    echo ****** NTFRSUTL.EXE is not in your path.  This tool can be found in the 
    echo ****** Windows 2000 Resource kit or the Windows Server 2003 support tools folder.
    echo ****** This tool is needed to gather all the data.
    goto :USAGE
)

REM  see if we have repadmin.exe
repadmin > nul: 2> nul:
if  ERRORLEVEL 1  (
    echo ****** REPADMIN.EXE is not in your path.
    echo ****** This tool is needed to gather all the data.
    goto :USAGE
)

REM  see if we have dcdiag.exe
dcdiag > nul: 2> nul:
if  ERRORLEVEL 1  (
    echo ****** DCDIAG.EXE is not in your path.
    echo ****** This tool is needed to gather all the data.
    goto :USAGE
)

REM  see if we have reg.exe
reg HKEY_LOCAL_MACHINE\system\currentcontrolset\services\NtFrsxxx > nul: 2> nul:
if  ERRORLEVEL 2  (
    echo ****** REG.EXE is not in your path.
    echo ****** This tool is needed to gather all the data.
    goto :USAGE
)


if "%2" == "" (
    set CHKCOMP=%COMPUTERNAME%
) else (
    set CHKCOMP=%2
)

set QA=%1\%CHKCOMP%
if NOT EXIST %QA% (
    echo ****** Creating output directory: %QA%
    md %QA%
)
if NOT EXIST %QA% (
    echo ****** Failed to create output dir: %QA%
    echo ****** No data retrieved.
    goto :DONE
)


@echo Please WAIT....

for %%x in (ds stage errscan inlog outlog machine reg sets version config sysvol) do (
    del %QA%\ntfrs_%%x.txt 1>nul: 2>nul:
)

for %%x in (showreps showconn) do (
    del %QA%\ds_%%x.txt 1>nul: 2>nul:
)

del %QA%\evl_*.txt %QA%\net_view.txt %QA%\dcdiag.txt %QA%\netdiag.txt %QA%\netdiag.log 1>nul: 2>nul:


echo DateTime         : %DATE%, %TIME%               > %QA%\ntfrs_machine.txt
echo TargetComputer   : %2                          >> %QA%\ntfrs_machine.txt
echo LocalComputername: %COMPUTERNAME%              >> %QA%\ntfrs_machine.txt
echo LogonServer      : %LOGONSERVER%               >> %QA%\ntfrs_machine.txt
echo UserDomain       : %USERDOMAIN%                >> %QA%\ntfrs_machine.txt
echo UserName         : %USERNAME%                  >> %QA%\ntfrs_machine.txt
echo Health_ChkVersion: %HC_VERSION%                >> %QA%\ntfrs_machine.txt

if /I "%CHKCOMP%" EQU "%COMPUTERNAME%" (
    echo.>> %QA%\ntfrs_machine.txt
    echo Architecture     : %PROCESSOR_ARCHITECTURE%    >> %QA%\ntfrs_machine.txt
    echo NumberProcessors : %NUMBER_OF_PROCESSORS%      >> %QA%\ntfrs_machine.txt
    echo SystemRoot       : %SystemRoot%                >> %QA%\ntfrs_machine.txt

    REM Dump out free disk space on all drives.
    REM echo.>> %QA%\ntfrs_machine.txt
    REM (for %%i in (c: d: e: f: g: h: i: j: k: l: m: n: o: p: q: r: s: t: u: v: w: x: y: z:) do (
    REM     for /f "delims=" %%j in ('diruse /* /, /m %%i\') do (
    REM         @echo %%i %%j
    REM     )
    REM  )
    REM ) >> %QA%\ntfrs_machine.txt


    REM  It looks like we are running on the target machine so run netdiag.
    echo netdiag  /debug is running
    netdiag  /debug                                  > %QA%\netdiag.txt  2> nul:
    REM Cleanup
    del netdiag.log 1>nul: 2>nul:
)

echo  NTFRSUTL checks are running ...
ntfrsutl  version %CHKCOMP%  > %QA%\ntfrs_version.txt
findstr /c:"ERROR" %QA%\ntfrs_version.txt
if  NOT ERRORLEVEL  1  (
    echo ****** NTFRSUTL cannot access target computer "%CHKCOMP%".
    echo ****** You must be an admin or run HEALTH_CHK on the target computer to gather all the data.
    goto :GETREG
)

ntfrsutl  ds      %CHKCOMP%  > %QA%\ntfrs_ds.txt
ntfrsutl  sets    %CHKCOMP%  > %QA%\ntfrs_sets.txt
ntfrsutl  configtable  %CHKCOMP%  > %QA%\ntfrs_config.txt
ntfrsutl  stage   %CHKCOMP%  > %QA%\ntfrs_stage.txt

echo  Dumping FRS inbound and outbound logs ...
ntfrsutl  inlog   %CHKCOMP%  > %QA%\ntfrs_inlog.txt
ntfrsutl  outlog  %CHKCOMP%  > %QA%\ntfrs_outlog.txt


:GETREG
echo  Dumping FRS registry parameters ...
reg query \\%CHKCOMP%\HKLM\system\currentcontrolset\services\NtFrs /s > %QA%\ntfrs_reg.txt


echo SYSVOL check is running
dir \\%CHKCOMP%\sysvol /s > %QA%\ntfrs_sysvol.txt

echo SMB/DFS check is running
net view \\%CHKCOMP%      > %QA%\net_view.txt
echo. >> %QA%\net_view.txt
if /I "%CHKCOMP%" EQU "%COMPUTERNAME%" (
    echo 'net session' results: >> %QA%\net_view.txt
    net session >> %QA%\net_view.txt
    echo. >> %QA%\net_view.txt
    echo 'net file' results: >> %QA%\net_view.txt
    net file >> %QA%\net_view.txt
)
echo. >> %QA%\net_view.txt
echo 'dfsutil /domain:%userdomain% /view' results: >> %QA%\net_view.txt
dfsutil /domain:%userdomain% /view >> %QA%\net_view.txt


REM
REM For DS replication
REM
echo repadmin /showreps %CHKCOMP% is running
repadmin /showreps  %CHKCOMP% >  %QA%\ds_showreps.txt

echo repadmin /showconn %CHKCOMP% is running
repadmin /showconn %CHKCOMP%  >  %QA%\ds_showconn.txt

echo dcdiag /v /s:%CHKCOMP% is running
dcdiag /v /s:%CHKCOMP%        >  %QA%\dcdiag.txt

if EXIST %windir%\system32\eventquery.vbs (
    REM
    REM Do a simple dump of the 400 most recent records in the event logs.
    REM
    echo  Scanning eventlogs ...
    cscript %windir%\system32\eventquery.vbs  /r 400  /l "file replication service"       /s %CHKCOMP%  > %QA%\evl_ntfrs.txt
    cscript %windir%\system32\eventquery.vbs  /r 400  /l application /s %CHKCOMP%  > %QA%\evl_application.txt
    cscript %windir%\system32\eventquery.vbs  /r 400  /l system      /s %CHKCOMP%  > %QA%\evl_system.txt
    cscript %windir%\system32\eventquery.vbs  /r 400  /l "directory service"          /s %CHKCOMP%  > %QA%\evl_ds.txt
    cscript %windir%\system32\eventquery.vbs  /r 400  /l "DNS Server" /s %CHKCOMP%  > %QA%\evl_dns.txt
)

if NOT EXIST %windir%\system32\eventquery.vbs (
   echo ** NO eventquery.vbs - cannot obtain event records. 
   echo ** This tool is located in the %windir%\system32\ folder on Windows Server 2003 systems
)

echo  Scanning FRS debug logs for error/warning info ...


set XCOMP=%windir%\debug

if /i "%CHKCOMP%" EQU "%COMPUTERNAME%"     goto :FOUND

set XCOMP=\\%CHKCOMP%\debug
dir %XCOMP%\ntfrs_*.log 1>nul: 2>nul:
if NOT ERRORLEVEL 1     goto :FOUND

set XCOMP=\\%CHKCOMP%\admin$\debug
dir  %XCOMP%\ntfrs_*.log 1>nul: 2>nul:
if NOT ERRORLEVEL 1     goto :FOUND

set XCOMP=\\%CHKCOMP%\C$\winnt\debug

:FOUND

dir  %XCOMP%\ntfrs_*.log > %QA%\ntfrs_errscan.txt
if  ERRORLEVEL 1  (
    echo ****** HEALTH_CHK cannot access the FRS log files on "%XCOMP%".
    echo ****** You must be an admin or run HEALTH_CHK on the target computer to gather all the data.
    del %QA%\ntfrs_errscan.txt
    goto :DONE
)

findstr /i "S0: S1: :H: error invalid fail abort warn" %XCOMP%\ntfrs_*.log   |  findstr /v "IO_PEND ERROR_SUCCESS PrintThreadIds FrsErrorSuccess" >> %QA%\ntfrs_errscan.txt

:DONE
echo  Done ...


:QUIT






