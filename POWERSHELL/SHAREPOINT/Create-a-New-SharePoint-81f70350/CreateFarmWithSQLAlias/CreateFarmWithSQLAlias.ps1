$ver = $host | select version
if ($ver.Version.Major -gt 1)  {$Host.Runspace.ThreadOptions = "ReuseThread"}
Add-PsSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

##
#Set Script Variables
##

##
#Variables for the SQL Alias
##

Write-Progress -Activity "Configuring Script Variables" -Status "Setting SQL Alias Variables"
$SQLServer = "SQLServerName"
$SQLAlias = "AliasForSharePointSQL"

##
#Variables for the SharePoint Farm
##

Write-Progress -Activity "Configuring Script Variables" -Status "Setting Farm Variables"
$FarmAccountUserName = "Domain\User"
$FarmAccountPassword = "AccountPassword"
$PassPhrase = "FarmPassphrase"
$FarmConfigDatabase = "SPFarm_ConfigDatabase"
$AdminContentDatabase = "SPFarm_Admin_ContentDB"
$CentralAdminPort = 8080

##
#Begin Farm Creation Script
##

##
#Create the SQL Alias
##

Write-Progress -Activity "Configuring Environment" -Status "Creating SQL Alias Entries"

#Define the Registry Locations For the SQL Alias Locations 
$x86 = "HKLM:\Software\Microsoft\MSSQLServer\Client\ConnectTo"
$x64 = "HKLM:\Software\Wow6432Node\Microsoft\MSSQLServer\Client\ConnectTo"  


#Check to See if The ConnectTo Key Already Exists, Create it if it Doesn't. 
if ((test-path -path $x86) -ne $True) {
     write-host "$x86 doesn't exist"    
	New-Item $x86 
} 

if ((test-path -path $x64) -ne $True) {
     write-host "$x64 doesn't exist"    
	New-Item $x64 
}   

#Record the Alias Type 
$TCPAlias = "DBMSSOCN," + $SQLServer


#Create The TCP/IP Aliases 

New-ItemProperty -Path $x86 -Name $SQLAlias -PropertyType String -Value $TCPAlias 
New-ItemProperty -Path $x64 -Name $SQLAlias -PropertyType String -Value $TCPAlias

##
#Configure the SharePoint Farm
##

Write-Progress -Activity "SharePoint Farm Configuration" -Status "Creating A New SharePoint Farm"
$FarmCredentials = New-Object System.Management.Automation.PSCredential $FarmAccountUserName, (ConvertTo-SecureString $FarmAccountPassword -AsPlainText -Force)
$FarmPassphrase = (ConvertTo-SecureString $Passphrase -AsPlainText -force)
New-SPConfigurationDatabase -DatabaseServer $SQLAlias -DatabaseName $FarmConfigDatabase -AdministrationContentDatabaseName $AdminContentDatabase -Passphrase $FarmPassphrase -FarmCredentials $FarmCredentials

#Install Services and Features

Write-Progress -Activity "SharePoint Farm Configuration" -Status "Installing Services and Features"
Initialize-SPResourceSecurity
Install-SPService  
Install-SPFeature -AllExistingFeatures 

#Create Central Administration Web Application
Write-Progress -Activity "SharePoint Farm Configuration" -Status "Creating A New Central Administration Web Application"
New-SPCentralAdministration -Port $CentralAdminPort -WindowsAuthProvider NTLM 
Install-SPHelpCollection -All
Install-SPApplicationContent

#Register the Service Connection Point in Active Directory
#See Technet article: http://technet.microsoft.com/en-us/library/ff730261.aspx for more information
$ServiceConnectionPoint = get-SPTopologyServiceApplication | select URI
Set-SPFarmConfig -ServiceConnectionPointBindingInformation $ServiceConnectionPoint -Confirm:$False
