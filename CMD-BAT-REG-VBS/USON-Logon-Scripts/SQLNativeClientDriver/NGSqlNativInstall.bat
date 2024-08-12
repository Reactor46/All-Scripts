IF EXIST C:\sqlclientinstall.txt goto regedit else goto driverinstall 
:driverinstall
msiexec /qb /l* C:\sqlclientinstall.txt /i \\usonvsvrsql1\NextGenRoot\Install\SQLNativeClientDriver\sqlncli.msi PERL_PATH=Yes PERL_EXT=Yes
:regedit
regedit /S "\\usonvsvrsql1\NextGenRoot\Install\SQLNativeClientDriver\sqlncliRegEdit.reg"


