@ECHO OFF
SETLOCAL EnableDelayedExpansion

SET adapterName=

FOR /F "tokens=* delims=:" %%a IN ('IPCONFIG ^| FIND /I "ETHERNET ADAPTER"') DO (
SET adapterName=%%a

REM Removes "Ethernet adapter" from the front of the adapter name
SET adapterName=!adapterName:~17!

REM Removes the colon from the end of the adapter name
SET adapterName=!adapterName:~0,-1!

netsh interface ipv4 set dns name="!adapterName!" static 10.145.10.21 primary
netsh interface ipv4 add dns name="!adapterName!" 10.141.0.2 index=2
netsh interface ipv4 add dns name="!adapterName!" 10.141.0.1 index=3
)

ipconfig /flushdns

:EOF
