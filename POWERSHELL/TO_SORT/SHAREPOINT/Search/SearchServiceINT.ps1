#############################################################################################################
# Kelsey-Seybold Search Service Application Setup and Configuration                                         #
# Integration Environment Environment Setup                                                                 #
#############################################################################################################
Set-ExecutionPolicy Unrestricted

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

#############################################################################################################
# 1.Setting up some initial variables.																		#
#############################################################################################################
write-host 1.Setting up the initial variables. 
$SSAName = "SharePoint - Search Service" 
$SVCAcct = "kelsey-seybold\sa5SP2013_INT_Search"
$SVCAppPool = "kelsey-seybold\sa5SP2013_INT_Pool" 
$SSI = get-spenterprisesearchserviceinstance -local 
$err = $null 
$databaseServerName = "SPINT\Internal"
$SSANameDB = "SP2013_INT_KSC_Intranet_SearchService"

# Based on scripts at http://www.harbar.net/articles/sp2013mt.aspx
# Thanks Spence!

#############################################################################################################
# Start Services search services for SSI																	#
#############################################################################################################
write-host Start Services - Search Services for SSI 
Start-SPEnterpriseSearchServiceInstance -Identity $SSI 

#############################################################################################################
# 2.Create an Application Pool.																				#
#############################################################################################################
write-host 2.Create an Application Pool. 
$AppPool = new-SPServiceApplicationPool -name $SSAName"-AppPool" -account $SVCAppPool 

# Get App Pool
$saAppPoolName = "SharePoint - Search Service-AppPool"

# Search Specifics, we are single server farm
$searchServerName = (Get-ChildItem env:computername).value
$serviceAppName = "Search Service"
$searchDBName = "SP2013_INT_KSC_Intranet_SearchService_DB"

# Grab the Appplication Pool for Service Application Endpoint
$saAppPool = Get-SPServiceApplicationPool $saAppPoolName

# Start Search Service Instances
Write-Host "Starting Search Service Instances..."
Start-SPEnterpriseSearchServiceInstance $searchServerName
Start-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance $searchServerName

# Create the Search Service Application and Proxy
Write-Host "Creating Search Service Application and Proxy..."
$searchServiceApp = New-SPEnterpriseSearchServiceApplication -Name $serviceAppName -ApplicationPool $saAppPoolName -DatabaseName $searchDBName
$searchProxy = New-SPEnterpriseSearchServiceApplicationProxy -Name "$serviceAppName Proxy" -SearchApplication $searchServiceApp

# Clone the default Topology (which is empty) and create a new one and then activate it
Write-Host "Configuring Search Component Topology..."
$clone = $searchServiceApp.ActiveTopology.Clone()
$searchServiceInstance = Get-SPEnterpriseSearchServiceInstance
New-SPEnterpriseSearchAdminComponent –SearchTopology $clone -SearchServiceInstance $searchServiceInstance
New-SPEnterpriseSearchContentProcessingComponent –SearchTopology $clone -SearchServiceInstance $searchServiceInstance
New-SPEnterpriseSearchAnalyticsProcessingComponent –SearchTopology $clone -SearchServiceInstance $searchServiceInstance 
New-SPEnterpriseSearchCrawlComponent –SearchTopology $clone -SearchServiceInstance $searchServiceInstance 
New-SPEnterpriseSearchIndexComponent –SearchTopology $clone -SearchServiceInstance $searchServiceInstance
New-SPEnterpriseSearchQueryProcessingComponent –SearchTopology $clone -SearchServiceInstance $searchServiceInstance
$clone.Activate()


Write-Host "Search Done!"
