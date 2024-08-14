set "OutFile=.\logs\%0%_MSS_Settings.log"
set "InParam=sceregvl-MSS.inf"
echo %time% > %OutFile% 2>&1
echo. >> %OutFile% 2>&1

echo *********** Configuring MSS-Settings *************** >> %OutFile% 2>&1

echo.  >> %OutFile% 2>&1
echo Taking Backup of %systemroot%\inf\sceregvl.inf as sceregvl-original.inf >> %OutFile% 2>&1
copy /Y %systemroot%\inf\sceregvl.inf logs\sceregvl-original.inf >> %OutFile% 2>&1

echo.  >> %OutFile% 2>&1
echo taking Ownership of the .inf file  >> %OutFile% 2>&1
takeown /f %systemroot%\inf\sceregvl.inf >> %OutFile% 2>&1

echo.  >> %OutFile% 2>&1
echo Set permissions to administrators group in .inf file  >> %OutFile% 2>&1
icacls  %systemroot%\inf\sceregvl.inf /grant administrators:f  >> %OutFile% 2>&1

echo.  >> %OutFile% 2>&1
echo Configure the .inf file for MSS settings  >> %OutFile% 2>&1
copy /Y %InParam% %systemroot%\inf\sceregvl.inf  >> %OutFile% 2>&1

echo.  >> %OutFile% 2>&1
echo Register the MSS Settings  >> %OutFile% 2>&1
regsvr32 scecli.dll

echo.  >> %OutFile% 2>&1
echo Copy the original .inf file back to System Root  >> %OutFile% 2>&1
copy /Y logs\sceregvl-original.inf %systemroot%\inf\sceregvl.inf  >> %OutFile% 2>&1

echo.  >> %OutFile% 2>&1
echo Revoke back the permissions and ownership  >> %OutFile% 2>&1
icacls  %systemroot%\inf\sceregvl.inf /setowner "NT SERVICE\TrustedInstaller" >> %OutFile% 2>&1
icacls  %systemroot%\inf\sceregvl.inf /remove administrators >> %OutFile% 2>&1


set "OutFile=.\logs\%0%_Security_Settings.log"
if [%1]==[] (
set "InParam=Baseline-Win2008.inf"
) else (
set "InParam=%1%"
)
echo %time% > %OutFile% 2>&1
echo. >> %OutFile% 2>&1