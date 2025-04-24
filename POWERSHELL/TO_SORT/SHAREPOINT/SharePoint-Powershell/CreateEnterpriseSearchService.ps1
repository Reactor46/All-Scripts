#############################################################################################################
# Kelsey-Seybold Search Service Application Setup and Configuration                                         #
# Production Environment Environment Setup                                                                  #
#############################################################################################################
Set-ExecutionPolicy Unrestricted

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
# Specify the Settings for the Search Service Application
$ServerName = "FBV-SPWFE-P01"
$IndexLocation = "D:\SearchIndex"
$SearchServiceApplicationName = "Enterprise Search Service Application"
$SearchServiceApplicationProxyName = "Enterprise Search Service Application Proxy"
$SearchDatabaseServer = "SPProd\Internal"
$SearchServiceApplicationDatabase = "SP2013_Prod_KSC_Intranet_EnterpriseSearchService"
 
$SearchAppPoolName = "Entrprise Search Service Application pool"
$SearchAppPoolAccount =  Get-SPManagedAccount "kelsey-seybold\sa5SP2013_Pool"
 
#Check if Managed account is registered already
Write-Host -ForegroundColor Yellow "Checking if the Managed Accounts already exists"
$SearchAppPoolAccount = Get-SPManagedAccount -Identity $SearchAppPoolAccount -ErrorAction SilentlyContinue
If ($SearchAppPoolAccount -eq $null)
{
    Write-Host "Please Enter the password for the Service Account..."
    $AppPoolCredentials = Get-Credential $SearchAppPoolAccount
    $SearchAppPoolAccount = New-SPManagedAccount -Credential $AppPoolCredentials
}
 
#*** Step 1: Create Application Pool for Search Service Application ****
Write-Host -ForegroundColor Yellow "Step 1: Create Application Pool for Search Service Application"
#Get the existing Application Pool
$SearchServiceAppPool = Get-SPServiceApplicationPool -Identity $SearchAppPoolName -ErrorAction SilentlyContinue
#If Application pool Doesn't exists, Create it
if (!$SearchServiceAppPool)
{
    $SearchServiceAppPool = New-SPServiceApplicationPool -Name $SearchAppPoolName -Account $SearchAppPoolAccount
    write-host "Created New Application Pool" -ForegroundColor Green
}
 
#*** Step 2: Start Search Service Instances ***
Write-Host -ForegroundColor Yellow "Step 2: Start Search Service Instances"
Start-SPEnterpriseSearchServiceInstance $ServerName 
Start-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance $ServerName 
 
#*** Step 3: Create Search Service Application ****
Write-Host -ForegroundColor Yellow "Step 3: Create Search Service Application"
# Get the Search Service Application
$SearchServiceApplication = Get-SPEnterpriseSearchServiceApplication -Identity $SearchServiceApplicationName -ErrorAction SilentlyContinue
Read-Host
# Create the Search Service Application, If it doesn't exists!
if(!$SearchServiceApplication)
{
    Read-Host
    $SearchServiceApplication = New-SPEnterpriseSearchServiceApplication -Name $SearchServiceApplicationName -ApplicationPool $SearchServiceAppPool -DatabaseServer $SearchDatabaseServer -DatabaseName $SearchServiceApplicationDatabase
    write-host "Created New Search Service Application" -ForegroundColor Green
}
 

#*** Step 4: Create Search Service Application Proxy ****
 #Get the Search Service Application Proxy
 $SearchServiceAppProxy = Get-SPEnterpriseSearchServiceApplicationProxy -Identity $SearchServiceApplicationProxyName -ErrorAction SilentlyContinue
 # Create the Proxy, If it doesn't exists!
if(!$SearchServiceAppProxy)
{
    $SearchServiceAppProxy = New-SPEnterpriseSearchServiceApplicationProxy -Name $SearchServiceApplicationProxyName -SearchApplication $SearchServiceApplication
    write-host "Created New Search Service Application Proxy" -ForegroundColor Green
}
 
#*** Step 5: Create New Search Topology
$SearchServiceInstance = Get-SPEnterpriseSearchServiceInstance -Local
#To Get Search Service Instance on Other Servers: use - $SearchServiceAppSrv1 = Get-SPEnterpriseSearchServiceInstance -Identity "<Server Name>"
 
# Create New Search Topology
$SearchTopology =  New-SPEnterpriseSearchTopology -SearchApplication $SearchServiceApplication
 
#*** Step 6: Create Components of Search
 
New-SPEnterpriseSearchContentProcessingComponent -SearchTopology $SearchTopology -SearchServiceInstance $SearchServiceInstance
 
New-SPEnterpriseSearchAnalyticsProcessingComponent -SearchTopology $SearchTopology -SearchServiceInstance $SearchServiceInstance
 
New-SPEnterpriseSearchCrawlComponent -SearchTopology $SearchTopology -SearchServiceInstance $SearchServiceInstance
 
New-SPEnterpriseSearchAdminComponent -SearchTopology $SearchTopology -SearchServiceInstance $SearchServiceInstance
 
#Prepare Index Location
Remove-Item -Recurse -Force -LiteralPath $IndexLocation -ErrorAction SilentlyContinue
MKDIR -Path $IndexLocation -Force
 
#Create Index and Query Components
New-SPEnterpriseSearchIndexComponent -SearchTopology $SearchTopology -SearchServiceInstance $SearchServiceInstance -RootDirectory $IndexLocation
 
New-SPEnterpriseSearchQueryProcessingComponent -SearchTopology $SearchTopology -SearchServiceInstance $SearchServiceInstance
 
#*** Step 7: Activate the Toplogy for Search Service ***
$SearchTopology.Activate() # Or Use: Set-SPEnterpriseSearchTopology -Identity $SearchTopology


