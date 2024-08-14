@Echo off
setlocal enabledelayedexpansion
SET Server=%1
SET DB=%2
SET workdir=%3
IF "x"%workdir%=="x" SET workdir=%cd%
SET curdir=%cd%
cd %workdir%
for /F %%x IN ('DIR /B *.sql') do (
echo Installing %workdir%\%%x ...
echo Installing %workdir%\%%x ... >> %curDir%\Deploymentlog.txt
SQLCMD -b -E -S %Server% -d %DB% -i%workdir%\%%x  >> %curDir%\Deploymentlog.txt 
IF !ERRORLEVEL!==1 (
ECHO Error installing %workdir%\%%x . Please refer log file for further details. 
ECHO Error installing %workdir%\%%x  >> %curDir%\Deploymentlog.txt 
GOTO ENDPRO
)

ECHO %workdir%\%%x Completed
ECHO %workdir%\%%x Completed >> %curDir%\Deploymentlog.txt



) 
:ENDPRO
cd %curdir%