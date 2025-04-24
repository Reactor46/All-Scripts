Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
   
#Configuration Parameters
$ServiceAppName = "User Profile Service Application"
$ServiceAppProxyName = "User Profile Service Application Proxy"
$AppPoolAccount = "AD\NPF1A0005b8SPAFarm"
$AppPoolName = "Service Application App Pool"
$DatabaseServer ="agl-sppre-d.ad.ing.net"
$UserProfileDB = "Profile_DB"
$UserProfileSyncDB = "Profile_Sync_DB"
$UserProfileSocialDB = "Profile_Social_DB"
$MySiteHostLocation="https://es-mysitepre.ad.ing.net"

 
Try {
    #Set the Error Action
    $ErrorActionPreference = "Stop"
    
    #Check if Managed account is registered already
    Write-Host -ForegroundColor Yellow "Checking if the Managed Accounts already exists"
    $AppPoolAccount = Get-SPManagedAccount -Identity $AppPoolAccount -ErrorAction SilentlyContinue
    if($AppPoolAccount -eq $null)
    {
        Write-Host "Please Enter the password for the Service Account..."
        $AppPoolCredentials = Get-Credential $AppPoolAccount
        $AppPoolAccount = New-SPManagedAccount -Credential $AppPoolCredentials
    }
    
    #Check if the application pool exists already
    Write-Host -ForegroundColor Yellow "Checking if the Application Pool already exists"
    $AppPool = Get-SPServiceApplicationPool -Identity $AppPoolName -ErrorAction SilentlyContinue
    if ($AppPool -eq $null)
    {
        Write-Host -ForegroundColor Green "Creating Application Pool..."
        $AppPool = New-SPServiceApplicationPool -Name $AppPoolName -Account $AppPoolAccount
    }
    
    #Check if the Service application exists already
    Write-Host -ForegroundColor Yellow "Checking if User Profile Service Application exists already"
    $ServiceApplication = Get-SPServiceApplication -Name $ServiceAppName -ErrorAction SilentlyContinue
    if ($ServiceApplication -eq $null)
    {
        Write-Host -ForegroundColor Green "Creating User Profile Service  Application..."
        $ServiceApplication =  New-SPProfileServiceApplication -Name $ServiceAppName -ApplicationPool $AppPoolName -ProfileDBName $UserProfileDB -ProfileSyncDBName $UserProfileSyncDB -SocialDBName $UserProfileSocialDB -MySiteHostLocation $MySiteHostLocation
    }
    #Check if the Service application Proxy exists already
    $ServiceAppProxy = Get-SPServiceApplicationProxy | where { $_.Name -eq $ServiceAppProxyName}
    if ($ServiceAppProxy -eq $null)
    {
        #Create Proxy
        $ServiceApplicationProxy = New-SPProfileServiceApplicationProxy -Name $ServiceAppProxyName -ServiceApplication $ServiceApplication -DefaultProxyGroup       
    }
    #Start service instance
    $ServiceInstance  = Get-SPServiceInstance | Where-Object { $_.TypeName -eq "User Profile Service" }
    
    #Check the Service status
    if ($ServiceInstance.Status -ne "Online")
    {
        Write-Host -ForegroundColor Yellow "Starting the User Profile Service Instance..."
        Start-SPServiceInstance $ServiceInstance
    }
    
    Write-Host -ForegroundColor Green "User Profile Service Application created successfully!"
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
 }
 finally {
    #Reset the Error Action to Default
    $ErrorActionPreference = "Continue"
} 