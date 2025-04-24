param ([string] $instance, [string] $collation,$ServiceAccount, $ServiceAccountPassword, $saPassword, $sourceDir, $servicePackExec, $instDrive, $userDbDrive, $userLogDrive, $tempDbDrive,  $tempLogDrive, $backupDrive , $port )


###############################
# install prerequisites for sqlversion
function installPrereqs ()
{
    #Import-Module ServerManager
    $dotnet=(Get-WindowsFeature -Name Net-Framework-Core).InstallState
    $message = ".NET Framework 3.5 is installed on the server"
    Write-Output $message $dotnet.InstallState
    #Add-WindowsFeature Application-Server,AS-NET-Framework,NET-Framework,NET-Framework-Core,WAS,WAS-Process-Model,WAS-NET-Environment,WAS-Config-APIs
    # get-windowsfeature | Where {$_.Installed} | Sort FeatureType,Parent,Name | Select Name,Displayname,FeatureType,Parent
}

###############################
# prepare the standard configuration file

function prepareConfigFile ([String]$version, $instance, [String]$collation, $instDrive, $userDbDrive, $userLogDrive, $tempDbDrive, $tempLogDrive, $backupDrive, $ServiceAccount ) 
{
    If($version -like "2019") {$ver="MSSQL15."}
    If($version -like "2017") {$ver="MSSQL14."}
    If($version -like "2016") {$ver="MSSQL13."}
    If($version -like "2014") {$ver="MSSQL12."}
    If($version -like "2012") {$ver="MSSQL11."}
    $config = "[OPTIONS]
    ACTION=""Install""
    FEATURES=SQLENGINE,IS,BC
    INSTANCENAME=""MSSQLSERVER""
    INSTANCEID=""MSSQLSERVER""
    INSTALLSHAREDDIR=""C:\Program Files\Microsoft SQL Server""
    INSTALLSHAREDWOWDIR=""C:\Program Files (x86)\Microsoft SQL Server""
    INSTANCEDIR=""C:\Program Files\Microsoft SQL Server""
    INSTALLSQLDATADIR="""+$instDrive+":\Program Files\Microsoft SQL Server""
    SQLUSERDBDIR="""+$userDbDrive+":\Program Files\Microsoft SQL Server\"+$ver+$instance+"\MSSQL\Data""
    SQLUSERDBLOGDIR="""+$userLogDrive+":\Program Files\Microsoft SQL Server\"+$ver+$instance+"\MSSQL\Data""
    SQLTEMPDBDIR="""+$tempDbDrive+":\Program Files\Microsoft SQL Server\"+$ver+$instance+"\MSSQL\Data""
    SQLTEMPDBFILECOUNT=""4""
    SQLTEMPDBFILESIZE=""1024""
    SQLTEMPDBFILEGROWTH=""0""
    SQLMAXDOP=""4""
    SQLTEMPDBLOGDIR="""+$tempLogDrive+":\Program Files\Microsoft SQL Server\"+$ver+$instance+"\MSSQL\Data""
    SQLTEMPDBLOGFILESIZE=""64""
    SQLTEMPDBLOGFILEGROWTH=""64""
    SQLBACKUPDIR="""+$backupDrive+":\Program Files\Microsoft SQL Server\"+$ver+$instance+"\MSSQL\Backup""
    FILESTREAMLEVEL=""0""
    TCPENABLED=""1""
    ENU=""True""
    NPENABLED=""0""
    SQLCOLLATION=""SQL_Latin1_General_CP1_CI_AS""
    SQLSVCINSTANTFILEINIT=""True""
    ADDCURRENTUSERASSQLADMIN=""True""
    SQLSVCACCOUNT=""NT Service\MSSQLSERVER""
    SQLSVCSTARTUPTYPE=""Automatic""
    SQLTELSVCACCT=""NT Service\SQLTELEMETRY""
    SQLTELSVCSTARTUPTYPE=""Automatic""
    ISTELSVCSTARTUPTYPE=""Automatic""
    ISTELSVCACCT=""NT Service\SSISTELEMETRY150""
    AGTSVCACCOUNT=""NT Service\SQLSERVERAGENT""
    AGTSVCSTARTUPTYPE=""Automatic""
    ISSVCACCOUNT=""NT Service\MsDtsServer150""
    ISSVCSTARTUPTYPE=""Disabled""
    BROWSERSVCSTARTUPTYPE=""Automatic""
    SQLSYSADMINACCOUNTS=""KELSEY-SEYBOLD\BMS DBAs"" ""KELSEY-SEYBOLD\sabackup"" ""KELSEY-SEYBOLD\ISAppDevs""
    IACCEPTPYTHONLICENSETERMS=""False""
    IACCEPTROPENLICENSETERMS=""False""
    SECURITYMODE=""SQL""
    SQMREPORTING=""False""
    SUPPRESSPRIVACYSTATEMENTNOTICE=""True""
    IACCEPTSQLSERVERLICENSETERMS=""True""
    SQLMINMEMORY=""0""
    SQLMAXMEMORY=""2147483647"""

    $config
}


######################
# Starting the main code from here #
######################


. ./function.ps1

$sourceDir1=get-SQlservermedia
#$sourceDir="'"+$sourceDir1+'\'+"'"
$sourceDir=$sourceDir1 -replace ' ', '` '
write-host $sourceDir
$version=get-SQlserverversion
$ServiceAccount =Get-serviceaccount
$syncSvcAccountPassword=get-Serviceaccountpassword
$instance=Get-instacneName
$saPassword=get-sapassword
$instDrive=Datadriveletter
$userDbDrive=Datadriveletter
$userLogDrive=logdriveletter
$tempDbDrive=tempdbdriveletter
$tempLogDrive=tempdbdriveletter
$backupDrive=Datadriveletter

$collation="SQL_Latin1_General_CP1_CI_AS"
$workDir = pwd


"Creating Ini File for Installation..."

$configFile = "$workDir\sql"+$version+"_"+$instance+"_install.ini"

prepareConfigFile $version $instance $collation $instDrive $userDbDrive $userLogDrive $tempDbDrive $tempLogDrive $backupDrive $ServiceAccount | Out-File $configFile

"Configuration File written to: "+$configFile

######################
# Installing Prereqs #
######################
"Installing Prerequisites (.Net, etc) ..."
installPrereqs 

#######################################
# Starting SQL Base Installation #
#######################################

#set-location $sourceDir

"Starting SQL Base Installation..."
$installCmd = $sourceDir+"\setup.exe /qs /SQLSVCPASSWORD='$syncSvcAccountPassword' /AGTSVCPASSWORD='$syncSvcAccountPassword' /ISSVCPASSWORD='$syncSvcAccountPassword' /SAPWD='$saPassword'  /ConfigurationFile=""$configFile"""

#write-host  $installCmd
Invoke-Expression $installCmd
