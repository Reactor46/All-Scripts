$Users = Get-Content 'D:\WinSys\SolarWinds Server Migration\jb-perm.txt'

ForEach($usr in $Users){
Invoke-Expression -Command ('takeown $usr /A /R /D')
Invoke-Expression -Command ('icacls $usr /remove FNBMCORP\battistaj /T /C')
Invoke-Expression -Command ('icacls $usr /setowner FNBMCORP\$usr /T /C')
Invoke-Expression -Command ('icacls $usr /inheritance:e /T /C ')
Invoke-Expression -Command ('icacls $usr')
}    