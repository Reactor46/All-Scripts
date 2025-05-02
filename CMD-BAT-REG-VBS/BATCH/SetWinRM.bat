@echo on
echo %date% > "WinRM.txt"
Set startDate=%date%
Set sdy=%startDate:~10%
Set /a sdm=1%startDate:~4,2% - 100
Set /a sdd=1%startDate:~7,2% - 100
powershell winrm get winrm/config > c:\Utils\WinRM.txt
powershell set-executionpolicy unrestricted
powershell Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*'
powershell winrm quickconfig -q
powershell enable-psremoting -force
powershell restart-service winrm
rem "WinRM.txt" "WinRM%sdy%%sdm%%sdd%.txt"
pause
