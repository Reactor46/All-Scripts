powercfg /SETACTIVE "Always On"

cscript //nologo \\uson.local\netlogon\GFI.vbs

:Map Network Share Drive
REM net use S: \\TempUSONFS01\company /persistent:yes 

cscript //nologo \\uson.local\netlogon\Shoretel.vbs

:Standardized Outlook Signature
REM cscript //nologo \\uson.local\netlogon\EmailSignature\emailsigOPTUM.vbs
regedit.exe /s \\uson.local\SYSVOL\USON.LOCAL\scripts\adobebubble.reg

DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log" /F /Q
DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log.1" /F /Q
DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log.2" /F /Q
DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log.3" /F /Q
DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log.4" /F /Q
DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log.5" /F /Q 



EXIT

