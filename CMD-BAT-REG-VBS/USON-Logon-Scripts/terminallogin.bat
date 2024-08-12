powercfg /SETACTIVE "Always On"
:PC Info
\\uson.local\netlogon\BGinfo\BGInfo.exe \\uson.local\netlogon\BGinfo\kiosk.bgi /timer:0 /silent /nolicprompt
cscript //nologo \\uson.local\netlogon\Shoretel.vbs
:Map Network Share Drive
REM net use S: \\TempUSONFS01\company
EXIT

