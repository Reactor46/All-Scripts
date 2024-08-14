@ECHO OFF
REM Name: TaskRepair.CMD
REM Author: Daniel Sheehan
REM Requires: REG.EXE if it is not included with the OS and SetACL.EXE from http://sourceforge.net/projects/setacl/
REM Summary: Removes and repairs manually specified scheduled task entries so they will not generate "The task image is corrupt or has been tampered with" errors.
REM Summary: This scrpt was inspired by JimFlyer from the forum post http://social.technet.microsoft.com/Forums/en-US/w7itproinstall/thread/5e3849da-e186-40c3-acb5-238342c642b8/#fe9204ea-f938-4ad6-b160-90aa7e8ebe6e
REM Summary: The steps in this script following the instructions in the KB article - http://support.microsoft.com/kb/2305420

REM List all the scheduled tasks that have the error reflecting their folder membership under the "Task Scheduler Library".
CALL :LOOP "Microsoft\Windows\Active Directory Rights Management Services Client\AD RMS Rights Policy Template Management (Automated)"
CALL :LOOP "Microsoft\Windows\Active Directory Rights Management Services Client\AD RMS Rights Policy Template Management (Manual)"
CALL :LOOP "Microsoft\Windows\AppID\PolicyConverter"
CALL :LOOP "Microsoft\Windows\AppID\VerifiedPublisherCertStoreCheck"
CALL :LOOP "Microsoft\Windows\Application Experience\AitAgent"
CALL :LOOP "Microsoft\Windows\Application Experience\ProgramDataUpdater"
CALL :LOOP "Microsoft\Windows\Autochk\Proxy"
CALL :LOOP "Microsoft\Windows\CertificateServicesClient\SystemTask"
CALL :LOOP "Microsoft\Windows\CertificateServicesClient\UserTask"
CALL :LOOP "Microsoft\Windows\CertificateServicesClient\UserTask-Roam"
CALL :LOOP "Microsoft\Windows\Customer Experience Improvement Program\Consolidator"
CALL :LOOP "Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask"
CALL :LOOP "Microsoft\Windows\Customer Experience Improvement Program\UsbCeip"
CALL :LOOP "Microsoft\Windows\Customer Experience Improvement Program\Server\ServerRoleUsageCollector"
CALL :LOOP "Microsoft\Windows\Customer Experience Improvement Program\Server\ServerRoleCollector"
CALL :LOOP "Microsoft\Windows\Customer Experience Improvement Program\Server\ServerCeipAssistant"
CALL :LOOP "Microsoft\Windows\Defrag\ScheduledDefrag"
CALL :LOOP "Microsoft\Windows\MemoryDiagnostic\CorruptionDetector"
CALL :LOOP "Microsoft\Windows\MemoryDiagnostic\DecompressionFailureDetector"
CALL :LOOP "Microsoft\Windows\MUI\LPRemove"
CALL :LOOP "Microsoft\Windows\Multimedia\SystemSoundsService"
CALL :LOOP "Microsoft\Windows\NetTrace\GatherNetworkInfo"
CALL :LOOP "Microsoft\Windows\Power Efficiency Diagnostics\AnalyzeSystem"
CALL :LOOP "Microsoft\Windows\RAC\RacTask"
CALL :LOOP "Microsoft\Windows\Ras\MobilityManager"
CALL :LOOP "Microsoft\Windows\Registry\RegIdleBackup"
CALL :LOOP "Microsoft\Windows\Server Manager\ServerManager"
CALL :LOOP "Microsoft\Windows\SoftwareProtectionPlatform\SvcRestartTask"
CALL :LOOP "Microsoft\Windows\Task Manager\Interactive"
CALL :LOOP "Microsoft\Windows\Tcpip\IpAddressConflict1"
CALL :LOOP "Microsoft\Windows\Tcpip\IpAddressConflict2"
CALL :LOOP "Microsoft\Windows\termsrv\licensing\TlsWarning"
CALL :LOOP "Microsoft\Windows\TextServicesFramework\MsCtfMonitor"
CALL :LOOP "Microsoft\Windows\Time Synchronization\SynchronizeTime"
CALL :LOOP "Microsoft\Windows\UPnP\UPnPHostConfig"
CALL :LOOP "Microsoft\Windows\User Profile Service\HiveUploadTask"
CALL :LOOP "Microsoft\Windows\WDI\ResolutionHost"
CALL :LOOP "Microsoft\Windows\Windows Error Reporting\QueueReporting"
CALL :LOOP "Microsoft\Windows\Windows Filtering Platform\BfeOnServiceStartTypeChange"
CALL :LOOP "Microsoft\Windows\WindowsColorSystem\Calibration Loader"

ECHO.
ECHO All tasks have been repaired, and a reboot is now recommended.
ECHO Exiting the Task Repair script.
GOTO :EOF

:LOOP
REM Set the TASKNAME variable to the task name in quotes including the full folder path.
SET TASKNAME=%1
ECHO Grabbing the registry information for scheduled task %TASKNAME%.

REM Per the KB Step 1 sub-step 3 - Grab the GUID of the task from the registry.
FOR /F "tokens=2 delims={}" %%a IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\%TASKNAME:~1,-1%" /v Id') DO SET REGID={%%a}

REM Per the KB Step 1 sub-step 4 - determine which TaskCache key the GUID is listed in and record it to the REGCLEANUP variable.
REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\%REGID%" >Nul
IF %ERRORLEVEL%==0 SET REGCLEANUP="HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Plain\%REGID%"
REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\%REGID%" >Nul
IF %ERRORLEVEL%==0 SET REGCLEANUP="HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Logon\%REGID%"
REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\%REGID%" >Nul
IF %ERRORLEVEL%==0 SET REGCLEANUP="HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Boot\%REGID%"

ECHO Temporarily removing the task from the system.
REM Per the KB Step 2 - copy the task file to a temporary folder in the system designated TEMP folder.
ECHO F | XCOPY "%SYSTEMDRIVE%\Windows\System32\Tasks\%TASKNAME:~1,-1%" "%TEMP%\Tasks\%TASKNAME:~1,-1%">Nul
IF ERRORLEVEL 1 ECHO There was a problem copying the scheduled task file for %TASKNAME:~1,-1%, skipping this task.&ECHO.&GOTO :EOF

REM Assuming there were no issues copying the task file, per the KB Step 3 sub-step 1 remove it from the Tasks folder on the SYSTEMDRIVE.
DEL "%SYSTEMDRIVE%\Windows\System32\Tasks\%TASKNAME:~1,-1%" >Nul

REM Grant the local Administrators group ownership of the registry keys about to be deleted, otherwise the permissions can't be modified.
SetACL -on "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\%TASKNAME:~1,-1%" -ot reg -actn setowner -ownr "n:Administrators;s:N" >Nul
SetACL -on "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\%REGID%" -ot reg -actn setowner -ownr "n:Administrators;s:N" >Nul
SetACL -on %REGCLEANUP% -ot reg -actn setowner -ownr "n:Administrators;s:N" >Nul

REM Grant the local Administrators group full control on the registry keys about to be deleted.
SetACL -on "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\%TASKNAME:~1,-1%" -ot reg -actn ace -ace "n:Administrators;p:full" >Nul
SetACL -on "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\%REGID%" -ot reg -actn ace -ace "n:Administrators;p:full" >Nul
SetACL -on %REGCLEANUP% -ot reg -actn ace -ace "n:Administrators;p:full" >Nul

REM Per the KB Step 3 sub-steps 2-4 - remove the three registry keys associated with the task.
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\%TASKNAME:~1,-1%" /f >Nul
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks\%REGID%" /f >Nul
REG DELETE %REGCLEANUP% /f >Nul

REM Per the KB Step 4 - recreate the sceduled task from the temporary file in the TEMP folder.
Schtasks.exe /CREATE /TN %TASKNAME% /XML "%TEMP%\Tasks\%TASKNAME:~1,-1%"
ECHO.
