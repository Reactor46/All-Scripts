echo
net stop trustedinstaller
sc config trustedinstaller start= demand
taskkill /f /im TrustedInstaller.exe
net stop wuauserv
sc config wuauserv start= demand
pause