cd "C:\temp\InstallSQL\"
echo %TIME%

call powershell -Command "&{set-executionpolicy remotesigned}" 

if %ERRORLEVEL% NEQ 0 GOTO END 

Set p=%~dp0

call powershell -File C:\temp\InstallSQL\common.ps1 exit $LASTEXITCODE

if %ERRORLEVEL% NEQ 0 GOTO END 

call powershell -File C:\temp\InstallSQL\setenvvaribale.ps1 exit $LASTEXITCODE
call powershell -File C:\temp\InstallSQL\setuppolices.ps1 exit $LASTEXITCODE

if %ERRORLEVEL% NEQ 0 GOTO END 

call powershell -File C:\temp\InstallSQL\InstallSQL.ps1 exit $LASTEXITCODE

if %ERRORLEVEL% NEQ 0 GOTO END 
call powershell -File C:\temp\InstallSQL\SQLPatch.ps1 exit $LASTEXITCODE

if %ERRORLEVEL% NEQ 0 GOTO END 

call powershell -File C:\temp\InstallSQL\Infrastructurescripts.ps1 exit $LASTEXITCODE

if %ERRORLEVEL% NEQ 0 GOTO END 

call powershell -File C:\temp\InstallSQL\SSMS-Install.ps1 exit $LASTEXITCODE

if %ERRORLEVEL% NEQ 0 GOTO END 
call powershell -File C:\temp\InstallSQL\SSDT-Install.ps1 -Install SQL exit $LASTEXITCODE

if %ERRORLEVEL% NEQ 0 GOTO END 
call powershell -File C:\temp\InstallSQL\SQLCU.ps1 exit $LASTEXITCODE


:END
echo %TIME%
pause



