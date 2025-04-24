. ./function.ps1
#import-module "C:\SQLInstall\common.psm1"
################################################################################
# Step : Granting required privileges to Admins group 
################################################################################
$ServiceAccount =Get-serviceaccount

$date = Get-Date -format "yyyy-MM-dd HH:mm:ss" 
Write-Output "$date Granting required privileges to Administrators group"

.\ntrights.exe -u $ServiceAccount +r SeBatchLogonRight
.\ntrights.exe -u $ServiceAccount +r SeServiceLogonRight
.\ntrights.exe -u $ServiceAccount +r SeTcbPrivilege
.\ntrights.exe -u $ServiceAccount +r SeManageVolumePrivilege
.\ntrights.exe -u $ServiceAccount +r SeLockMemoryPrivilege


#C:\SQLInstall\InstallSQL2016.ps1 -installSQLserverSp exit $LASTEXITCODE
#if ($LASTEXITCODE -NEQ 0) GOTO END 
#call powershell -File C:\SQLInstall\InstallSQL2016.ps1 -setpolices exit $LASTEXITCODE

#:END

#setpolices 
