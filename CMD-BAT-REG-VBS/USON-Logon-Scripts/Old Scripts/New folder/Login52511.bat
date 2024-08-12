powercfg /SETACTIVE "Always On"
:Map Network Share Drive
net use S: \\servermain\company
:Outlook Config Reg Edit
regedit /s \\servermain\Company\OutlookConfig\OutlookConfig.reg
:Copy IE Shortcuts to User Desktop
xcopy \\servermain\Address\Shortcuts\*.* "%userprofile%\desktop" /Y
:SQL Native Client Install with regedit batch script
call \\ngdata\NextGenRoot\Install\SQLNativeClientDriver\NGSqlNativInstall.bat
:Nextgen ini file copy
xcopy "\\ngdata\nextgenroot\prod\NGconfig.ini" "C:\windows" /Y
:Standardized Outlook Signature 
cscript //nologo \\servermain\sysvol\USON.local\scripts\EmailSignature\emailsig.vbs
:TightVNC install
call \\servermain\sysvol\USON.local\scripts\TightVNC\TightVNCInstall.bat
EXIT

