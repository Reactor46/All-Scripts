powercfg /SETACTIVE "Always On"
:Map Network Share Drive
net use S: \\usonpsvrfpf\company
:Outlook Config Reg Edit
regedit /s \\usonpsvrfpf\Company\OutlookConfig\OutlookConfig.reg
:Copy IE Shortcuts to User Desktop
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\Shortcuts\DesktopFolder\*.* "%userprofile%\desktop" /Y /s /e
:Create IE Shortcuts in IE
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\Shortcuts\IEShortcuts\*.* "%userprofile%\Favorites" /Y /s /e
:SQL Native Client Install with regedit batch script
call \\usonvsvrdc\sysvol\USON.local\scripts\SQLNativeClientDriver\NGSqlNativInstall.bat
:Nextgen ini file copy
xcopy "\\usonvsvrsql1\NextGenRoot\prod\NGconfig.ini" "C:\windows" /Y
regedit /s \\usonvsvrsql1\NextGenRoot\Prod\NextGenODBC.reg /Y
:Standardized Outlook Signature 
cscript //nologo \\usonvsvrdc\sysvol\USON.local\scripts\EmailSignature\emailsig.vbs
:Copy Dr. Text Signature
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\EmailSignature\DrText\*.* "%userprofile%\Application Data\Microsoft\Signatures" /Y /s /e
:TightVNC install
call \\usonvsvrdc\sysvol\USON.local\scripts\TightVNC\TightVNCInstall.bat
:Change Local Administrator Password
cscript //nologo \\usonvsvrdc\sysvol\USON.local\scripts\VBscripts\ChangeLocalAdminPass.vbs
EXIT

