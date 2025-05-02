@ECHO ON
REM Uninstall KACE, SCCM, GFI, IE11 etc

start /wait "%WINDIR%\CCMSetup\ccmsetup /uninstall"
start /wait "%PROGRAMFILES%\Dell\KACE\AMPTools.exe uninstall all-kuid"
start /wait "%PROGRAMFILES(X86)%\Dell\KACE\AMPTools.exe uninstall all-kuid"
MsiExec.exe /X {A0707C59-4B32-48B8-94ED-73BB68E1C569} /QN