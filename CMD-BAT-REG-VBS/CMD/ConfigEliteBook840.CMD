ECHO Set BIOS Password
%~dp0BiosConfigUtility64.exe /nspwdfile:"%~dp0passwordConfig.bin"

ECHO This command will run successfully if the password is set to the Password specified in the first steps
%~dp0BiosConfigUtility64.exe /cspwdfile:"%~dp0passwordConfig.bin" /set:"%~dp0ComputerModel.txt"

ECHO Exit Forcing Success Return Code of 0
exit /b 0

