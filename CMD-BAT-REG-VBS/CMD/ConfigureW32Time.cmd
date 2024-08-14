set LogFile=%~dpn0.log
@echo %date% %time% Start: %* > %LogFile%

@rem
@rem Configure w32time service
@rem
@rem Stop it, set to auto start, delete triggers that cause w32time to change behavior
@rem on domain join/leave, set the polling interval to 4 hours (instead of 7 days), start, 
@rem enable logging, update internal config, resync, and change the weekly task that starts
@rem it (remove the task_started parameter as it tells w32time to run and immediately exit)
@rem

call :Run sc.exe stop w32time
ver | findstr -i 6.0
if %errorlevel% EQU 0 (goto :skipDeleteTriggerInfo)
call :Run sc.exe triggerinfo w32time Delete

:skipDeleteTriggerInfo
call :Run reg.exe add HKLM\system\controlset001\services\w32time\TimeProviders\NtpClient /v SpecialPollInterval /d 14400 /t REG_DWORD /f
call :Run sc.exe config w32time Start= auto
call :Run sc.exe start w32time
call :Run w32tm.exe /debug /enable /file:%windir%\w32time.log /size:1000000 /entries:0-300
call :Run w32tm.exe /config /update
call :Run w32tm.exe /resync /force
call :Run schtasks.exe /change /TN "\Microsoft\Windows\Time Synchronization\SynchronizeTime" /TR "%windir%\system32\sc.exe start w32time"

@echo %date% %time% End: %* >> %LogFile%
goto :eof

:Run
echo %date% %time% Execute: %* >> %LogFile%
start /w %*
echo %date% %time% Finished: %* - Errorlevel: %ERRORLEVEL% >> %LogFile%
goto :eof
