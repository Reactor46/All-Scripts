Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Configuration variables
$ServiceAppName = "State Service Application"
$ServiceAppProxyName ="State Service Application Proxy"
$DatabaseName ="State_Service DB"
 
#Create New State Service application
$StateServiceApp = New-SPStateServiceApplication -Name $ServiceAppName
 
#Create Database for State Service App
$Database = New-SPStateServiceDatabase -Name $DatabaseName -ServiceApplication $StateServiceApp 
 
#Create Proxy for State Service
New-SPStateServiceApplicationProxy -Name $ServiceAppProxyName -ServiceApplication $StateServiceApp -DefaultProxyGroup 
 
Initialize-SPStateServiceDatabase -Identity $Database
 
Write-host "State Service Application Created Successfully!"