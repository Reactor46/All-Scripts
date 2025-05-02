:: ============================================================================
:: Script  : LockScreenInfo.cmd
:: Purpose : Update Logon or Lock Screen Background Image with technical informations
:: Must be scheduled on Boot and on SessionLock
:: ============================================================================
setlocal
set LOGFILE="%TEMP%\%~n0.log"
echo [%date:~-10% %time:~0,8%] [%~nx0] [info] === begin process. === >%LOGFILE% 2>>&1
pushd "%~dp0"
set LSICONF=%ProgramData%\LockScreenInfo
if exist "%LSICONF%\Themes\backgroundDefault.jpg" set SOURCEIMAGE="%LSICONF%\Themes\backgroundDefault.jpg"
set SOURCEPATH=%SystemRoot%\System32\oobe\info\backgrounds
set TARGETIMAGE="%SystemRoot%\System32\oobe\info\backgrounds\backgroundDefault.jpg"
set PS1FILE=%LSICONF%\LockScreenInfo.ps1
if exist "%~dp0LockScreenInfo.ps1" set PS1FILE=%~dp0LockScreenInfo.ps1
if NOT exist "%PS1FILE%" echo [%date:~-10% %time:~0,8%] [%~nx0] [Warning] Script "%PS1FILE%" Not Found! >>%LOGFILE% 2>>&1
echo [%date:~-10% %time:~0,8%] [%~nx0] [info] SOURCEIMAGE=%SOURCEIMAGE% >>%LOGFILE% 2>>&1
if NOT exist "%SystemRoot%\System32\oobe\info\backgrounds" md "%SystemRoot%\System32\oobe\info\backgrounds" >>%LOGFILE% 2>&1
if exist %TARGETIMAGE% del %TARGETIMAGE% >>%LOGFILE% 2>&1
if exist "%PS1FILE%" powershell.exe -ExecutionPolicy "UnRestricted" -File "%PS1FILE%" -SourcePath "%SOURCEPATH%" -TargetImage %TARGETIMAGE% >>%LOGFILE% 2>>&1
if NOT exist %TARGETIMAGE% if exist "%PS1FILE%" powershell.exe -ExecutionPolicy "UnRestricted" -File "%PS1FILE%" %SOURCEIMAGE% -TargetImage %TARGETIMAGE% -FitScreen >>%LOGFILE% 2>>&1
:End
echo [%date:~-10% %time:~0,8%] [%~nx0] [info] === exit process. === >>%LOGFILE% 2>>&1
popd
endlocal
exit /b 0