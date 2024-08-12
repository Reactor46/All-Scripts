@ECHO ON

REM Uninstall NXLog
REM MsiExec.exe /X{D5F06971-73E1-4B19-928F-50707ABCCC6E} /Q

REM Uninstall Zabbix
start /wait MsiExec.exe /I{B859296D-C606-49E7-9829-F2B028472682} /Q
start /wait MsiExec.exe /I{D85A0FFF-2354-48D3-A241-D76D53E116B1} /Q

REM Uninstall Graylog Sidecar
REM "C:\Program Files\Graylog\sidecar\uninstall.exe" /S

REM Uninstall NetWrix Auditor Agent
start /wait MsiExec.exe /X{A57A516F-19A2-4B21-95F1-19E9584E1696} /Q

REM Uninstall FireEye 31.28.8
REM MsiExec.exe /X{FF484B3E-5609-409E-A5A2-1DB9CABC6D41} /Q

REM Uninstall SolarWinds Client Components
"C:\ProgramData\{3006AC1F-535E-4666-9467-FA84DBDF978C}\wmiproviders_2019.4.0.125.exe" REMOVE=TRUE MODIFY=FALSE

************************************