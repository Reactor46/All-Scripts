cls
@echo off
set /p Usern=Enter RunAs Username:
set userrunas=runas /user:fnbmcorp\%usern%
%userrunas% "powershell \"\\fnbmcorp\share\shared\IT\SupportServices\NOC\Scripts\test\CAPSEventandQueueBIG.ps1"\"


