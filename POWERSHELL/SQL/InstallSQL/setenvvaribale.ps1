#import-module "C:\SQLInstall\common.psm1"
. ./function.ps1
$hostName = get-content env:computername
#$srvenv= read-host "enter environment like dev,test,dev02,test02"
#$instance=read-host "enter SQL server Instace Name like SQL01,SQl02"
$srvenv=Get-serverenvironment
$instance=Get-instacneName
$envSQlname=$hostName+'\'+$instance.ToUpper()
$envvariable1='SQL_SERVERNAME'
$envvariable2='SERVER_ENVIRONMENT'
$envvariable3='SQL_ENVIRONMENT'
$envexchangsrv='EXCHANGE_SERVER'
$envexchangsrvvalue='sadcrelay.pmmr.com'
write-host "Configuring the Server environment variables"
icm -comp $hostName {   Param ($envvariable1,$envSQlname)   [Environment]::SetEnvironmentVariable($envvariable1,$envSQlname,"Machine")   } -ArgumentList $envvariable1,$envSQlname
icm -comp $hostName {   Param ($envvariable2,$srvenv)   [Environment]::SetEnvironmentVariable($envvariable2,$srvenv,"Machine")   } -ArgumentList $envvariable2,$srvenv
icm -comp $hostName {   Param ($envvariable3,$srvenv)   [Environment]::SetEnvironmentVariable($envvariable3,$srvenv,"Machine")   } -ArgumentList $envvariable3,$srvenv
icm -comp $hostName {   Param ($envexchangsrv,$envexchangsrvvalue)   [Environment]::SetEnvironmentVariable($envexchangsrv,$envexchangsrvvalue,"Machine")   } -ArgumentList $envexchangsrv,$envexchangsrvvalue
