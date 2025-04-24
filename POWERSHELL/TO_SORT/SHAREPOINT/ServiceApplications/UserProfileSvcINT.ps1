################################################################################################# 
## Kelsey-Seybold SharePoint Server 2013 User Profile Service Application                      ## 
################################################################################################# 
Set-ExecutionPolicy Unrestricted
  
Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue 
  
## Settings you may want to change depending on the install## 
$databaseServerName = "SPINT\Internal"

#User Profile Service
$userAppPoolName = "SharePoint - User Profile Service"
$userAppPoolUserName = "kelsey-seybold\sa5SP2013_INT_UserPr"  

## Service Application Service Names ## 
$secureStoreSAName = "Secure Store Service"
$userProfileSAName = "User Profile Service Application"

# MySite Host Location
$mysite = "https://my.int.kscpulse.com"

##########################################################################################################################################  
Write-Host "Creating User Profile Service and Proxy..."
# Creating the User Profile Service Application Pool
# Process/prereqs to install/configure the User Profile Service
#1. Create a Web application to host My Sites 
#2. Create a managed path for My Sites 
#3. Create a My Site Host site collection 
#4. Create a User Profile service application 
#5. Enable NetBIOS domain names 
#6. Start the User Profile service 

$userAppPool = Get-SPServiceApplicationPool -Identity $userAppPoolName -EA 0 
if($userAppPool -eq $null) 
{ 
  Write-Host "Creating User Profile Service Application Pool..."
  
  $userAppPoolAccount = Get-SPManagedAccount -Identity $userAppPoolUserName -EA 0 
  
  if($userAppPoolAccount -eq $null) 
  { 
    Write-Host "Cannot create or find the User Profile Service managed account $userAppPoolUserName, please ensure the account exists."
    Exit -1 
  } 
  
  New-SPServiceApplicationPool -Name $userAppPoolName -Account $userAppPoolAccount -EA 0 > $null
      
} 
#Create the User Profile Service Application
$userProfileService = New-SPProfileServiceApplication -Name $userProfileSAName -ApplicationPool $userAppPoolName -ProfileDBServer $databaseServerName -ProfileDBName "SP1013_INT_KSC_Intranet_ProfileDB" -SocialDBServer $databaseServerName -SocialDBName "SP2013_INT_KSC_Intranet_SocialDB" -ProfileSyncDBServer $databaseServerName -ProfileSyncDBName "SP2013_INT_KSC_Intranet_SyncDB" -MySiteHostLocation $mysite
#Create the User Profile Service Application Proxy
New-SPProfileServiceApplicationProxy -Name "$userProfileSAName Proxy" -ServiceApplication $userProfileService -DefaultProxyGroup > $null

$ServiceApps = Get-SPServiceApplication
$UserProfileServiceApp = $userProfileSAName
foreach ($sa in $ServiceApps)
  {if ($sa.DisplayName -eq $userProfileSAName) 
    {$UserProfileServiceApp = $sa}
  }
$UserProfileServiceApp.NetBIOSDomainNamesEnabled = 1
$UserProfileServiceApp.Update()

Get-SPServiceInstance | where-object {$_.TypeName -eq "User Profile Service Application"} | Start-SPServiceInstance > $null

##########################################################################################################################################

#Remove-PSSnapin Microsoft.SharePoint.Powershell