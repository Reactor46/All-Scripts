powercfg /SETACTIVE "Always On"
net use S: \\usonpsvrfpf\company
regedit /s \\usonpsvrfpf\Company\OutlookConfig\OutlookConfig.reg
:Desktop Links Cleanup
call \\usonvsvrdc\sysvol\USON.local\scripts\Shortcuts\DesktopCleanup.bat
:Copy IE Shortcuts to User Desktop
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\Shortcuts\DesktopFolder\*.* "%userprofile%\desktop" /Y /s /e
:Create IE Shortcuts in IE
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\Shortcuts\IEShortcuts\*.* "%userprofile%\Favorites" /Y /s /e
regedit /S "\\usonvsvrsql1\nextgenroot\prod\NextGenODBC.reg"
xcopy "\\usonvsvrsql1\nextgenroot\prod\NGconfig.ini" "C:\windows" /Y
cscript //nologo \\usonvsvrdc\sysvol\USON.local\scripts\EmailSignature\emailsig.vbs
:Copy Dr. Text Signature
xcopy \\usonvsvrdc\sysvol\USON.local\scripts\EmailSignature\DrText\*.* "%userprofile%\Application Data\Microsoft\Signatures" /Y /s /e
:Change Local Administrator Password
cscript //nologo \\usonvsvrdc\sysvol\USON.local\scripts\VBscripts\ChangeLocalAdminPass.vbs
Exit