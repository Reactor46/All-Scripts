cls
@echo off
set /p Usern=Enter RunAs Username:
set userrunas=runas /user:Contosocorp\%usern%
%userrunas% "powershell \"\\Contosocorp\share\shared\IT\SupportServices\NOC\Scripts\test\CAPSEventandQueueBIG.ps1"\"


