REM 
REM Stopping Sophos Services
net stop "Sophos AutoUpdate Service"
net stop "Sophos Agent"
net stop "SAVService"
net stop "SAVAdminService"
net stop "Sophos Message Router"
net stop "Sophos Web Control Service"
net stop "swi_service"
net stop "swi_update"

REM 
REM Removing Sophos AutoUpdater
MsiExec.exe /X{15C418EB-7675-42be-B2B3-281952DA014D} /qn REBOOT=SUPPRESS /PASSIVE
MsiExec.exe /X{7CD26A0C-9B59-4E84-B5EE-B386B2F7AA16} /qn REBOOT=SUPPRESS /PASSIVE

REM
REM Removing Sophos Update Manager
MsiExec.exe /X{2C7A82DB-69BC-4198-AC26-BB862F1BE4D0} /qn REBOOT=SUPPRESS /PASSIVE

REM
REM Removing Sophos Remote Management System
MsiExec.exe /X{FED1005D-CBC8-45D5-A288-FFC7BB304121} /qn REBOOT=SUPPRESS /PASSIVE

REM
REM Removing Sophos Anti-Virus
MsiExec.exe /X{9ACB414D-9347-40B6-A453-5EFB2DB59DFA} /qn REBOOT=SUPPRESS /PASSIVE
MsiExec.exe /X{D929B3B5-56C6-46CC-B3A3-A1A784CBB8E4} /qn REBOOT=SUPRESS /PASSIVE