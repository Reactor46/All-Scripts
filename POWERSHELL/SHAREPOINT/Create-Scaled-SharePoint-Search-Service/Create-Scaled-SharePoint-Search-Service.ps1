###############################################################################
#                 Create-Scaled-SharePoint-Search-Service.ps1                 #
#                                     by                                      #
#                             Cornelius J. van Dyk                            #
#                               http://aurl.to/C                              #
#                                                                             #
# This script will create a new SharePoint Search Service Application based   #
# on the Configuration Settings or any run time settings supplied by the user #
# at exectution time.  The script can be automated for zero user interaction  #
# by changing the $useUI setting to $false.  In such a case, the script will  #
# use the values from the Configuration Settings section for the creation of  #
# the service application.  Both Crawl servers and Query servers can be       #
# scaled to multiple servers by simply adding to the array or following the   #
# prompts at run time.                                                        #
###############################################################################

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Change this value to $false if you wish to use the defaults with NO user interaction.
$useUI = $true

# Configuration Settings
$databaseServerName = "SQL2012"
$databaseName = "SearchDB"
$searchAdminServerName = "SP-APP01"
$searchCrawlServerNames = @("SP-APP02", "SP-APP03")
$searchQueryServerNames = @("SP-WEB01", "SP-WEB02", "SP-WEB03", "SP-WEB04")
$appPoolName = "Search Service App Pool"
$appPoolUserName = "DOMAIN\Search-Service-Account"
$searchService = "SharePoint Search"

if ($useUI)
{
  $override = Read-Host "Do you wish to configure the Search architecture? Answering y will use the default settings. (y/n)"
}
else
{
  $override = "n"
}
if ($override -eq "y")
{
  # Get the Search Service name
  $searchService = ""
  do
  {
    Write-Host -ForegroundColor yellow "Please enter the name of the Search Service to create."
    $searchService = Read-Host "Search Service name?"
    if ($searchService -eq "")
    {
      Write-Host -ForegroundColor red "INVALID ENTRY!  Please enter the Search Service name."
    }
  }
  until ($searchService -ne "")
  
  # Get the App Pool name
  $appPoolName = ""
  do
  {
    Write-Host -ForegroundColor yellow "Please enter the name of the Search Service Application Pool.  If you don't have an app pool created for Search yet, don't worry, we'll create it for you."
    $appPoolName = Read-Host "Search Service App Pool Name?"
    if ($appPoolName -eq "")
    {
      Write-Host -ForegroundColor red "INVALID ENTRY!  Please enter the Application Pool name."
    }
  }
  until ($appPoolName -ne "")
  
  # Get the App Pool Username
  $appPoolUserName = ""
  do
  {
    Write-Host -ForegroundColor yellow "Please enter the Application Pool service account name, including the domain."
    $appPoolUserName = Read-Host "App Pool service account?"
    if ($appPoolUserName -eq "")
    {
      Write-Host -ForegroundColor red "INVALID ENTRY!  Please enter the Application Pool service account name."
    }
  }
  until ($appPoolUserName -ne "")
  
  # Get the Database server name
  $databaseServerName = ""
  do
  {
    Write-Host -ForegroundColor yellow "Please enter the name of the Database server.  The Database server needs to be a SQL Server.  It will be used to create the needed Search databases."
    $databaseServerName = Read-Host "Database Server Name?"
    if ($databaseServerName -eq "")
    {
      Write-Host -ForegroundColor red "INVALID ENTRY!  Please enter the target Database server name."
    }
  }
  until ($databaseServerName -ne "")
  
  # Get the Database name
  $databaseName = ""
  do
  {
    Write-Host -ForegroundColor yellow "Please enter the base name of the Database to use.  This name will be used in the creation of three databases on the [$databaseServerName] server."
    $databaseName = Read-Host "Database Name?"
    if ($databaseName -eq "")
    {
      Write-Host -ForegroundColor red "INVALID ENTRY!  Please enter the desired Database name."
    }
  }
  until ($databaseName -ne "")
  
  # Get the Admin server name
  $searchAdminServerName = ""
  do
  {
    Write-Host -ForegroundColor yellow "Please enter the name of the Search administration server.  The Admin server is usually an app server, but preferably not an app server that you intend to configure as a Crawl server."
    $searchAdminServerName = Read-Host "Admin Server Name?"
    if ($searchAdminServerName -eq "")
    {
      Write-Host -ForegroundColor red "INVALID ENTRY!  Please enter the target Admin server name."
    }
  }
  until ($searchAdminServerName -ne "")
  
  # Get the Crawl server name(s)
  $numberOfCrawl = 0
  $inputOK = $false
  do
  {
    try
    {
      Write-Host -ForegroundColor yellow "Please enter the number of servers you wish to configure as Crawl servers.  Crawl servers do the heavy lifting with indexing the content.  If possible, these should be application servers dedicated to Crawl.  These should never be web front end servers that serve content to users."
      [int]$numberOfCrawl = Read-Host "Number of Crawl Servers?"
      $inputOK = $true
    }
    catch
    {
      $inputOK = $false
      Write-Host -ForegroundColor red "INVALID ENTRY!  Please enter a numeric value."
    }
  }
  until ($inputOK)
  $searchCrawlServerNames = @()
  for ($i=1; $i -lt $numberOfCrawl + 1; $i++) 
  {
    $searchCrawlServerName = ""
    do
    {
      $searchCrawlServerName = Read-Host "Please enter the name for Crawl server number $i."
      if ($searchCrawlServerName -eq "")
      {
        Write-Host -ForegroundColor red "INVALID ENTRY!  Please enter the name of the Crawl server."
      }
    }
    until ($searchCrawlServerName -ne "")
	$searchCrawlServerNames += $searchCrawlServerName;
  }

  # Get the Query server name(s)
  $numberOfQuery = 0
  $inputOK = $false
  do
  {
    try
    {
      Write-Host -ForegroundColor yellow "Please enter the number of servers you wish to configure as Query servers.  Query servers contain the search indexes.  For redundancy, there should always be more than one of these servers.  Search continues to function for end users, even if the Crawl servers go down, but losing the last query server could force a recrawl of all content which can be very time consuming.  Web front end servers can be configured as Query servers due to the relatively low impact the Query component has.  Crawl servers should not be configured as Query servers so as to avoid high impact crawls from affecting query performance for end users."
      [int]$numberofQuery = Read-Host "Number of Query Servers?"
      $inputOK = $true
    }
    catch
    {
      $inputOK = $false
      Write-Host -ForegroundColor red "INVALID ENTRY!  Please enter a numeric value."
    }
  }
  until ($inputOK)
  $searchQueryServerNames = @()
  for ($i=1; $i -lt $numberOfQuery + 1; $i++) 
  {
    $searchQueryServerName = ""
    do
    {
      $searchQueryServerName = Read-Host "Please enter the name for Query server number $i."
      if ($searchQueryServerName -eq "")
      {
        Write-Host -ForegroundColor red "INVALID ENTRY!  Please enter the name of the Query server."
      }
    }
    until ($searchQueryServerName -ne "")
	$searchQueryServerNames += $searchQueryServerName;
  }
}
Write-Host
Write-Host -ForegroundColor white "Ready for Search Service creation using these values:"
Write-Host "Search Service Name: "$searchService
Write-Host "App Pool: "$appPoolName
Write-Host "App Pool Service Account: "$appPoolUserName
Write-Host "Database Server: "$databaseServerName
Write-Host "Database Name: "$databaseName
Write-Host "Admin Server: "$searchAdminServerName
Write-Host "Crawl Server(s): "
foreach ($server in $searchCrawlServerNames)
{
  Write-Host "  "$server
}
Write-Host "Query Server(s): "
foreach ($server in $searchQueryServerNames)
{
  Write-Host "  "$server
}
Write-Host
if ($useUI)
{
  $override = Read-Host "Proceed with Search Service creation. (y/n)"
  if ($override -ne "y")
  {
    Write-Host -ForegroundColor red "Search Service creation aborted by user request."
    Exit -1 
  }
}

Write-Host "Getting the Search Service Application Pool."
$appPool = Get-SPServiceApplicationPool -Identity $appPoolName -EA 0 
if($appPool -eq $null) 
{ 
  Write-Host "Search Service Application Pool not found.  Attempting to find the managed account for [$appPoolUserName]."
  $appPoolAccount = Get-SPManagedAccount -Identity $appPoolUserName -EA 0 
  if($appPoolAccount -eq $null) 
  { 
      Write-Host "Managed Account not found.  Please supply the password for [$appPoolUserName]."
      $appPoolCred = Get-Credential $appPoolUserName
      Write-Host "Creating a new Managed Account."
      $appPoolAccount = New-SPManagedAccount -Credential $appPoolCred -EA 0 
  } 
  $appPoolAccount = Get-SPManagedAccount -Identity $appPoolUserName -EA 0 
  if($appPoolAccount -eq $null) 
  { 
    Write-Host -ForegroundColor red "Cannot create or find the managed account [$appPoolUserName], please ensure the account exist and rerun this script."
    Exit -1 
  } 
  Write-Host "Creating a new Application Pool for the Search Service."
  New-SPServiceApplicationPool -Name $appPoolName -Account $appPoolAccount -EA 0 > $null
}

Write-Host "Creating the Search Service and Proxy."
Write-Host "  Starting Services on the Admin server [$searchAdminServerName]."
Start-SPEnterpriseSearchServiceInstance $searchAdminServerName
Start-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance $searchAdminServerName
foreach ($server in $searchCrawlServerNames)
{
  Write-Host "  Starting Services on the Crawl server [$server]."
  Start-SPEnterpriseSearchServiceInstance $server
  Start-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance $server
}
foreach ($server in $searchQueryServerNames)
{
  Write-Host "  Starting Services on the Query server [$server]."
  Start-SPEnterpriseSearchServiceInstance $server
  Start-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance $server
}

Write-Host "  Creating the Search Service Application."
$searchApp = New-SPEnterpriseSearchServiceApplication -Name $searchService -ApplicationPool $appPoolName -DatabaseServer $databaseServerName -DatabaseName $databaseName
$searchInstanceAdmin = Get-SPEnterpriseSearchServiceInstance $searchAdminServerName
$searchInstanceCrawl = @()
foreach ($server in $searchCrawlServerNames)
{
  $searchInstanceCrawl += Get-SPEnterpriseSearchServiceInstance $server
}
$searchInstanceQuery = @()
foreach ($server in $searchQueryServerNames)
{
  $searchInstanceQuery += Get-SPEnterpriseSearchServiceInstance $server
}

Write-Host "  Creating the Administration Component."
$searchApp | Get-SPEnterpriseSearchAdministrationComponent | Set-SPEnterpriseSearchAdministrationComponent -SearchServiceInstance $searchInstanceAdmin

Write-Host "  Creating the Crawler(s)."
Write-Host "    Getting the current Crawl Topology."
$oldCrawlTopology = $searchApp | Get-SPEnterpriseSearchCrawlTopology -Active
Write-Host "    Creating a new Crawl Topology."
$newCrawlTopology = $searchApp | New-SPEnterpriseSearchCrawlTopology
Write-Host "    Getting the Crawl Database."
$newCrawlDatabase = ([array]($searchApp | Get-SPEnterpriseSearchCrawlDatabase))[0] 
Write-Host "    Creating the new Crawl Component(s)."
$newCrawlComponents = @()
foreach ($instance in $searchInstanceCrawl)
{
  $newCrawlComponents += New-SPEnterpriseSearchCrawlComponent -CrawlTopology $newCrawlTopology -CrawlDatabase $newCrawlDatabase -SearchServiceInstance $instance
}
Write-Host "    Activating the new Crawl Topology."
$newCrawlTopology | Set-SPEnterpriseSearchCrawlTopology -Active
Write-Host -ForegroundColor white "    Waiting for the old crawl topology to become inactive" -NoNewline
do {write-host -NoNewline .;Start-Sleep 6;} while ($oldCrawlTopology.State -ne "Inactive") 
Write-Host "    Removing the old Crawl Topology."
$oldCrawlTopology | Remove-SPEnterpriseSearchCrawlTopology -Confirm:$false
Write-Host

Write-Host "  Creating Query Component(s)."
Write-Host "    Getting the old Query Topology."
$oldQueryTopology = $searchApp | Get-SPEnterpriseSearchQueryTopology -Active
Write-Host "    Creating the new Query Topology."
$newQueryTopology = $searchApp | New-SPEnterpriseSearchQueryTopology -Partitions 1
Write-Host "    Getting the Index Partition."
$newIndexPartition = (Get-SPEnterpriseSearchIndexPartition -QueryTopology $newQueryTopology) 
Write-Host "    Creating the new Query Component(s)."
$newQueryComponents = @()
foreach ($instance in $searchInstanceQuery)
{
  $newQueryComponents += New-SPEnterpriseSearchQuerycomponent -QueryTopology $newQueryTopology -IndexPartition $newIndexPartition -SearchServiceInstance $instance
}
Write-Host "    Getting the Property Database."
$newPropertyDatabase = ([array]($searchApp | Get-SPEnterpriseSearchPropertyDatabase))[0]  
Write-Host "    Setting the Index Partition."
$newIndexPartition | Set-SPEnterpriseSearchIndexPartition -PropertyDatabase $newPropertyDatabase
Write-Host "    Activating the new Query Topology."
$newQueryTopology | Set-SPEnterpriseSearchQueryTopology -Active
Write-Host -ForegroundColor white "    Waiting for the old query topology to become inactive" -NoNewline
do {write-host -NoNewline .;Start-Sleep 6;} while ($oldQueryTopology.State -ne "Inactive") 
Write-Host "    Removing the old Query Topology."
$oldQueryTopology | Remove-SPEnterpriseSearchQueryTopology -Confirm:$false
Write-Host

Write-Host "  Creating the Proxy."
$searchAppProxy = New-SPEnterpriseSearchServiceApplicationProxy -Name "$searchService Proxy" -SearchApplication $searchApp > $null
Write-Host "The Search Service Application [$searchService] was successfully created with the Admin component on [$searchAdminServerName]."
