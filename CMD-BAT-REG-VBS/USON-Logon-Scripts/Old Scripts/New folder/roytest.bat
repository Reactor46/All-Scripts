powercfg /SETACTIVE "Always On"
:Map Network Share Drive
net use S: \\usonpsvrfpf\company
:Outlook Config Reg Edit
regedit /s \\usonpsvrfpf\Company\OutlookConfig\OutlookConfig.reg
:Copy IE Shortcuts to User Desktop
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\Shortcuts\DesktopFolder\*.* "%userprofile%\desktop" /Y /s /e
:Create IE Shortcuts in IE
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\Shortcuts\IEShortcuts\*.* "%userprofile%\Favorites" /Y /s /e
:Nextgen ini file copy
xcopy "\\ngdata\nextgenroot\prod\NGconfig.ini" "C:\windows" /Y
:Standardized Outlook Signature 
cscript //nologo \\usonvsvrdc\sysvol\USON.local\scripts\EmailSignature\emailsig.vbs
:Copy Dr. Text Signature
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\EmailSignature\DrText\*.* "%userprofile%\Application Data\Microsoft\Signatures" /Y /s /e
Exit

