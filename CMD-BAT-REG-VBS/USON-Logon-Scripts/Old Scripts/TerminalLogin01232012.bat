powercfg /SETACTIVE "Always On"
:Nextgen ini file copy
xcopy "\\usonvsvrsql1\NextGenRoot\prod\NGconfig.ini" "C:\windows" /Y
regedit /s \\usonvsvrsql1\NextGenRoot\Prod\NextGenODBC.reg /Y
EXIT

