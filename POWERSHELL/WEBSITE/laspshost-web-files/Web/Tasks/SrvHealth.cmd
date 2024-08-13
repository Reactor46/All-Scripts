@echo off
cd\
cd C:\Scripts\Repository\jbattista\Web\reports\AdminScripts\ServerHealth\

powershell.exe ".\WinServ-Status-MODIFIED.ps1 -List C:\Scripts\Repository\jbattista\Web\reports\AdminScripts\ServerHealth\Configs\Servers.txt  -O C:\Scripts\Repository\jbattista\Web\reports\ -CpuAlertThreshold 80 -MemAlertThreshold 80 -RequestAlertThreshold 30 -ServiceMemThreshold 1800 -TotalWebThreshold 250"