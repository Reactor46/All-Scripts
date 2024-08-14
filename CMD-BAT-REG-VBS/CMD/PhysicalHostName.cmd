@echo off
::
:: Get Physical Host Name from guest PC  (virtualPC or Hyper-V)
:: Gastone Canali
::
:: PhysicalHostName.cmd
::
set query=reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters\" /v "PhysicalHostNameFullyQualified"
(for /f "tokens=3" %%N in ('%query% ^|find "PhysicalHostNameFullyQualified"') do call set "name=%%N")1>nul 2>nul
REM
if "%NAME%"=="" (echo Not a Virtual machine) else (echo Physical Host name: %NAME%)
:: END
pause