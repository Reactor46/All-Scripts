cls
@echo off
set /p host=Enter Server Name:
set /p Usern=Enter RunAs Username:
set userrunas=runas /user:Contosocorp\%usern%
%userrunas% "powershell \"\\Contosocorp\share\shared\IT\SupportServices\NOC\Scripts\test\Get-OS.ps1 %host%"\" 


