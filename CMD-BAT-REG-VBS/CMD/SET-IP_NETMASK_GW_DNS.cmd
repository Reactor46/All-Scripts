@ECHO off
:: SET-IP_NETMASK_GW_DNS.cmd
:: By Gastone Canali
:: ver 0.4 16/05/2014
setlocal
set $null= 1^>nul 2^>nul


::**** Italian version
:: Connessione alla rete locale (LAN)
::
::
::**** English version
:: Local Area Connection
::
set $NIC_NAME=Connessione alla rete locale (LAN)

:: -----------------------
:: WINDOWS 8.1 version
:: -----------------------
SET $NIC_NAME=Wi-Fi

CALL :Get-IPv4 "%$NIC_NAME%" ip
echo The %$NIC_NAME% has the following IPv4: %ip%
echo.&echo.&echo.

CALL :Get-IPv6 "%$NIC_NAME%" ip
echo The %$NIC_NAME% has the following IPv6: %ip%
echo.&echo.&echo.

CALL :Get-Adapter-Index "%$NIC_NAME%" "$in"
ECHO "%$NIC_NAME%" Adapter index %$in%

CALL :Get-Nicinfo "Wireless Network Connection"
CALL :Get-Nicinfo "%$NIC_NAME%"

REM  no output
echo no output
CALL :Get-Nicinfo "%$NIC_NAME%" %$NULL%

REM ECHO -----&ECHO :ENA-DHCP "%$NIC_NAME%"
REM CALL :ENA-DHCP    "%$NIC_NAME%"
REM CALL :Get-Nicinfo "%$NIC_NAME%"

REM ECHO -----&ECHO :SET-FULL-IP "%$NIC_NAME%"  "2.2.2.2"  "255.255.255.0"  "2.2.2.254"
REM CALL :SET-FULL-IP "%$NIC_NAME%"  "2.2.2.2"  "255.255.255.0"  "2.2.2.254"
REM CALL :Get-Nicinfo "%$NIC_NAME%"

REM ECHO -----&ECHO :SET-GW  "%$NIC_NAME%" 222.1.2.1
REM CALL :SET-GW      "%$NIC_NAME%" 222.1.2.1
REM CALL :Get-Nicinfo "%$NIC_NAME%" 255.255.0.0

REM ECHO -----&ECHO :SET-MASK  "%$NIC_NAME%" 255.255.0.0
REM CALL :SET-MASK    "%$NIC_NAME%" 255.255.0.0
REM CALL :Get-Nicinfo "%$NIC_NAME%"

REM ECHO -----&ECHO :SET-DNS   "%$NIC_NAME%" 8.8.8.8 8.8.4.4
REM CALL :SET-DNS     "%$NIC_NAME%" 8.8.8.8 8.8.4.4
REM CALL :Get-Nicinfo "%$NIC_NAME%"


goto :EOF


:Get-Adapter-Index
REM  Adapter name "%~1"    Return variable "%~2"
  for /F %%I in ('wmic nicconfig  where "IPEnabled=TRUE" get index ^|findstr /r [0-9]') do (
   for /F "tokens=*" %%N in ('wmic nic where "deviceid=%%I" get NetConnectionID ^|findstr /v /r "^$ NetConnectionID" ^|findstr /i /c:"%~1"') do CALL set "%~2=%%I"
  )
goto :EOF

:Get-IPv4
REM *** Nic Name %~1 
  CALL :Get-Adapter-Index "%~1" "$I"
  set $PartialInfo=IPAddress
  for /f "delims={}, tokens=2"  %%A IN ('wmic nicconfig where index^=%$I% get %$PartialInfo% /format:list ^|findstr /V "^$"' %null%) do call set "%~2=%%~A"
goto :EOF

:Get-IPv6
REM *** Nic Name %~1 
  CALL :Get-Adapter-Index "%~1" "$I"
  set $PartialInfo=IPAddress
  for /f "delims={}, tokens=3"  %%A IN ('wmic nicconfig where index^=%$I% get %$PartialInfo% /format:list ^|findstr /V "^$"' %null%) do call set "%~2=%%~A"
goto :EOF

:Get-Nicinfo
REM *** Nic Name %~1 
  CALL :Get-Adapter-Index "%~1" "$I"
  ECHO ----------------------------------
  set $PartialInfo=Description,IPAddress,IPsubnet,DefaultIPGateway,Macaddress,DNSServerSearchOrder,DHCPEnabled,Index
  set $AllInfo=*
  wmic nicconfig where index=%$I% get %$PartialInfo% /format:list |findstr /V "^$"
goto :EOF

:SET-FULL-IP
REM *** Nic Name %~1   ip %~2   netmask  %~3   gateway %~4
  netsh interface ip set address name="%~1" source=static addr=%~2  mask=%~3  gateway=%~4 %$null%
 goto :EOF

:SET-MASK
REM *** Nic Name %~1  netmask %~2
  CALL :Get-Adapter-Index "%~1" "$I"
  for /f "delims={}" %%i in ('wmic nicconfig where "index=%$I%" get ipaddress ^|findstr /i /v  "^$ IPadd"') do (
      wmic nicconfig where index=%$I% CALL enablestatic %%i ^, "%~2"
  ) %$null%
goto :EOF

:ENA-DHCP
REM *** Nic Name %~1  
	set nicname=%~1
  netsh interface ipv4 set address name="%nicname%" source=dhcp %$null%
goto :EOF

:SET-DNS
REM *** Nic Name %~1 primary dns %~2  secondary dns %~3
  CALL :Get-Adapter-Index "%~1" "$I"
  if     "%~3"=="" wmic nicconfig where (index=%$I%) CALL SetDNSServerSearchOrder(%~2, %~2) %$null%
  if not "%~3"=="" wmic nicconfig where (index=%$I%) CALL SetDNSServerSearchOrder(%~2, %~3) %$null%
goto :EOF

:SET-GW
REM *** Nic Name %~1     DefaultIPGateway  %~2
  CALL :Get-Adapter-Index "%~1" "$I"
  wmic nicconfig where (index=%$I%) CALL SetGateways(%~2) %$null%
goto :EOF

:END