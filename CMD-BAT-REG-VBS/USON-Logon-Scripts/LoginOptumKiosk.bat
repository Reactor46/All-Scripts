powercfg /SETACTIVE "Always On"
%systemroot%\regedit.exe /s \\uson.local\netlogon\2xFarm.reg
:PC Info
\\uson.local\netlogon\BGinfo\BGInfo.exe \\uson.local\netlogon\BGinfo\OptumLogokisok.bgi /timer:0 /silent /nolicprompt
:Map Network Share Drive
REM net use S: \\TempUSONFS01\company /persistent:yes
:Standardized Outlook Signature 
cscript //nologo \\uson.local\netlogon\Shoretel.vbs
REM cscript //nologo \\uson.local\netlogon\EmailSignature\emailsig.vbs
:Copy Dr. Text Signature
EXIT

