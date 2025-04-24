Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
   
#Configuration Parameters
$ServiceAppName = "User Profile Service Application"
$ServiceAppProxyName = "User Profile Service Application Proxy"
$AppPoolAccount = "AD\dev_shp_Srv"
$AppPoolName = "SharePoint Service Applications"
$UserProfileDB = "SP_Profile"
$UserProfileSyncDB = "SP_Sync"
$UserProfileSocialDB = "SP_Social"
$MySiteHostLocation = "http://dev-sharepoint:8080"
 
Try {
    #Set the Error Action
    $ErrorActionPreference = "Stop"
    
    #Check if Managed account is registered already
    Write-Host -ForegroundColor Yellow "Checking if the Managed Accounts already exists"
    $AppPoolAccount = Get-SPManagedAccount -Identity $AppPoolAccount -ErrorAction SilentlyContinue
    if ($null -eq $AppPoolAccount) {
        Write-Host "Please Enter the password for the Service Account..."
        $AppPoolCredentials = Get-Credential $AppPoolAccount
        $AppPoolAccount = New-SPManagedAccount -Credential $AppPoolCredentials
    }
    
    #Check if the application pool exists already
    Write-Host -ForegroundColor Yellow "Checking if the Application Pool already exists"
    $AppPool = Get-SPServiceApplicationPool -Identity $AppPoolName -ErrorAction SilentlyContinue
    if ($null -eq $AppPool) {
        Write-Host -ForegroundColor Green "Creating Application Pool..."
        $AppPool = New-SPServiceApplicationPool -Name $AppPoolName -Account $AppPoolAccount
    }
    
    #Check if the Service application exists already
    Write-Host -ForegroundColor Yellow "Checking if User Profile Service Application exists already"
    $ServiceApplication = Get-SPServiceApplication -Name $ServiceAppName -ErrorAction SilentlyContinue
    if ($null -eq $ServiceApplication) {
        Write-Host -ForegroundColor Green "Creating User Profile Service Application..."
        $ServiceApplication = New-SPProfileServiceApplication -Name $ServiceAppName -ApplicationPool $AppPoolName -ProfileDBName $UserProfileDB -ProfileSyncDBName $UserProfileSyncDB -SocialDBName $UserProfileSocialDB -MySiteHostLocation $MySiteHostLocation
    }
    #Check if the Service application Proxy exists already
    $ServiceAppProxy = Get-SPServiceApplicationProxy | Where-Object { $_.Name -eq $ServiceAppProxyName }
    if ($null -eq $ServiceAppProxy) {
        #Create Proxy
        $ServiceApplicationProxy = New-SPProfileServiceApplicationProxy -Name $ServiceAppProxyName -ServiceApplication $ServiceApplication -DefaultProxyGroup       
    }
    #Start service instance
    $ServiceInstance = Get-SPServiceInstance | Where-Object { $_.TypeName -eq "User Profile Service" }
    
    #Check the Service status
    if ($ServiceInstance.Status -ne "Online") {
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

#Read more: https://www.sharepointdiary.com/2017/08/powershell-to-create-user-profile-service-application-sharepoint-2016.html#ixzz7sAuaGNvX