powercfg /SETACTIVE "Always On"

if exist "C:\Program Files (x86)\PDF24\pdf24.exe" (start "" "C:\Program Files (x86)\PDF24\pdf24.exe")

:PC Info
\\uson.local\netlogon\BGinfo\BGInfo.exe \\uson.local\netlogon\BGinfo\standard.bgi /timer:0 /silent /nolicprompt
:Map Network Share Drive
REM net use S: \\TempUSONFS01\company
:Standardized Outlook Signature 
cscript //nologo \\uson.local\netlogon\Shoretel.vbs
REM cscript //nologo \\uson.local\netlogon\EmailSignature\DRemailsig.vbs
REM cscript //nologo \\USON.local\NETLOGON\EmailSignature\ROemailsig.vbs
:Copy Dr. Text Signature
xcopy \\uson.local\netlogon\EmailSignature\DrText\*.* "%userprofile%\Application Data\Microsoft\Signatures" /Y /s /e

EXIT