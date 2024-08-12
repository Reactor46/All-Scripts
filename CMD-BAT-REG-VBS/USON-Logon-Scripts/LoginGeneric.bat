powercfg /SETACTIVE "Always On"

if exist "C:\Program Files (x86)\PDF24\pdf24.exe" (start "" "C:\Program Files (x86)\PDF24\pdf24.exe")



REM net use S: \\TempUSONFS01\company /persistent:yes

regedit.exe /s \\uson.local\SYSVOL\USON.LOCAL\scripts\adobebubble.reg

DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log" /F /Q 


EXIT

