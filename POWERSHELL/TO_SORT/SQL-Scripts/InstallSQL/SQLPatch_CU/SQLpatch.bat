cd "C:\temp\InstallSQL\SQLPatch_CU"
echo %TIME%

call powershell -Command "&{set-executionpolicy remotesigned}" 

if %ERRORLEVEL% NEQ 0 GOTO END 

Set p=%~dp0

if %ERRORLEVEL% NEQ 0 GOTO END 
call powershell -File C:\temp\InstallSQL\SQLPatch_CU\SQLPatch_only.ps1 exit $LASTEXITCODE


:END
echo %TIME%
pause



