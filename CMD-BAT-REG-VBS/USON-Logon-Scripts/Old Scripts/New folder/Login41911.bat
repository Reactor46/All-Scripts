powercfg /SETACTIVE "Always On"
net use S: \\servermain\company
regedit /s \\servermain\Company\OutlookConfig\OutlookConfig.reg
xcopy \\servermain\Address\Shortcuts\*.* "%userprofile%\desktop" /Y
regedit /S "\\ngdata\nextgenroot\prod\NextGenODBC.reg"
xcopy "\\ngdata\nextgenroot\prod\NGconfig.ini" "C:\windows" /Y
cscript //nologo \\servermain\sysvol\USON.local\scripts\EmailSignature\emailsig.vbs

