rem suppres EUELA for current user
reg add "HKEY_CURRENT_USER\Software\Sysinternals\PsExec" /v EulaAccepted /t reg_dword /d "00000001" /f
rem suppres EUELA for SYSTEM accoun
reg add "HKEY_USERS\S-1-5-18\Software\Sysinternals\PsExec" /v EulaAccepted /t reg_dword /d "00000001" /f
rem install ESET product localy
PsExec.exe \\127.0.0.1  -s -d -c -f ERA_x64.exe --silent --accepteula --avr-disable 