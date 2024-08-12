powercfg /SETACTIVE "Always On"

cscript //nologo \\uson.local\netlogon\GFI.vbs

if exist "C:\Program Files (x86)\PDF24\pdf24.exe" (start "" "C:\Program Files (x86)\PDF24\pdf24.exe")

:PC Info
\\uson.local\netlogon\BGinfo\BGInfo.exe \\uson.local\netlogon\BGinfo\standard.bgi /timer:0 /silent /nolicprompt

:Map Network Share Drive
REM net use S: \\TempUSONFS01\company
net use U: \\usonvsvrfs01\USON /persistent:yes

:Standardized Outlook Signature 
REM cscript //nologo \\USON.local\NETLOGON\EmailSignature\emailsig.vbs
REM cscript //nologo \\USON.local\NETLOGON\EmailSignature\ROemailsig.vbs

cscript //nologo \\uson.local\netlogon\Shoretel.vbs

:Copy Dr. Text Signature
xcopy \\USON.local\NETLOGON\EmailSignature\DrText\*.* "%userprofile%\Application Data\Microsoft\Signatures" /Y /s /e

DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log" /F /Q
DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log.1" /F /Q
DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log.2" /F /Q
DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log.3" /F /Q
DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log.4" /F /Q
DEL "%USERPROFILE%\AppData\Roaming\ShoreWare Client\Logs\*.log.5" /F /Q 

EXIT

