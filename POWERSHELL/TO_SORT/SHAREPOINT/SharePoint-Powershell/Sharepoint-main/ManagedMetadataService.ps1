Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
   
#Configuration Parameters
$ServiceAppName = "Managed Metadata Service Application"
$ServiceAppProxyName = "Managed Metadata Service Application Proxy"
$AppPoolAccount = "AD\NPF1A0005b8SPAFarm"
$AppPoolName = "Service Application App Pool"
$DatabaseServer ="agl-sppre-d.ad.ing.net"
$DatabaseName = "ManagedMetadataService"
$ContentTypeHub="https://es-intranetpre.wbspain.ad.ing.net/ctypehub"
  
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
    Write-Host -ForegroundColor Yellow "Checking if Managed Metadata Service Application exists already"
    $ServiceApplication = Get-SPServiceApplication -Name $ServiceAppName -ErrorAction SilentlyContinue
    if ($ServiceApplication -eq $null)
    {
        Write-Host -ForegroundColor Green "Creating Managed Metadata Service Application..."
        $ServiceApplication =  New-SPMetadataServiceApplication -Name $ServiceAppName -ApplicationPool $AppPoolName -DatabaseName $DatabaseName -DatabaseServer $DatabaseServer
    }
    #Check if the Service application Proxy exists already
    $ServiceAppProxy = Get-SPServiceApplicationProxy | where { $_.Name -eq $ServiceAppProxyName}
    if ($ServiceAppProxy -eq $null)
    {
        #Optional Parameters: 
        $ServiceApplicationProxy = New-SPMetadataServiceApplicationProxy -Name $ServiceAppProxyName -ServiceApplication $ServiceApplication -DefaultProxyGroup -ContentTypePushdownEnabled -DefaultKeywordTaxonomy -DefaultSiteCollectionTaxonomy -Uri $ContentTypeHub
    }
    #Start service instance 
    $ServiceInstance = Get-SPServiceInstance | Where-Object { $_.TypeName -eq "Managed Metadata Web Service" }
  
    #Check the Service status
    if ($ServiceInstance.Status -ne "Online")
    {
        Write-Host -ForegroundColor Yellow "Starting the Managed MetadataService Instance..."
        Start-SPServiceInstance $ServiceInstance
    }
   
    Write-Host -ForegroundColor Green "Managed Metadata Service Application created successfully!"
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
 }
 finally {
    #Reset the Error Action to Default
    $ErrorActionPreference = "Continue"
 }