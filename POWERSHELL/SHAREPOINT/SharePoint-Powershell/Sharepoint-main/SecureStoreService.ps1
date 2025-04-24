Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
  
#Configuration Parameters
$ServiceAppName = "Secure Store Service Application"
$ServiceAppProxyName = "Secure Store Service Application Proxy"
$AppPoolAccount = "AD\NPF1A0005b8SPAFarm"
$AppPoolName = "Service Application App Pool"
$DatabaseServer ="agl-sppre-d.ad.ing.net"
$DatabaseName = "Service_SecureStore DB"
 
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
    Write-Host -ForegroundColor Yellow "Checking if Secure Store Service Application exists already"
    $ServiceApplication = Get-SPServiceApplication -Name $ServiceAppName -ErrorAction SilentlyContinue
    if ($ServiceApplication -eq $null)
    {
        Write-Host -ForegroundColor Green "Creating Secure Store Service Application..."
        $ServiceApplication = New-SPSecureStoreServiceApplication -Name $ServiceAppName -ApplicationPool $AppPoolName -DatabaseName $DatabaseName -DatabaseServer $DatabaseServer -AuditingEnabled:$false
        $ServiceApplicationProxy = New-SPSecureStoreServiceApplicationProxy -Name $ServiceAppName" Proxy" -ServiceApplication $ServiceApplication -DefaultProxyGroup
    }
  
    #Start service instance 
    $ServiceInstance = Get-SPServiceInstance | Where-Object { $_.TypeName -like "*Secure Store Service*" }
 
    #Check the Service status
    if ($ServiceInstance.Status -ne "Online")
    {
        Write-Host -ForegroundColor Yellow "Starting the Secure Store Service Instance..."
        Start-SPServiceInstance $ServiceInstance
    }
  
    Write-Host -ForegroundColor Green "Secure Store Service Application created successfully!"
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
 }
 finally {
    #Reset the Error Action to Default
    $ErrorActionPreference = "Continue"
 }



 #Config parameters
$Passphrase = "Sharepoint_2022"
$ServiceAppProxyName="Secure Store Service Application Proxy"
 
#Get the Service App Proxy
$ServiceAppProxy = Get-SPServiceApplicationProxy | where { $_.Name -eq $ServiceAppProxyName}
 
#Create Master key
Update-SPSecureStoreMasterKey -ServiceApplicationProxy $ServiceAppProxy -Passphrase $Passphrase
