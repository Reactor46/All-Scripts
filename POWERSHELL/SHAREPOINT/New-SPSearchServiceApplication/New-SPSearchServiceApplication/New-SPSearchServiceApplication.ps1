<##############################################################################
This script creates enterprise search service application in the current 
SharePoint 2013 farm. 

Assumptions: 1. Index partition will be created as locally on the server. So 
run this script on server, which will host search admin component.
2. The service account for application pool, has been registered under managed 
accounts in SP farm.
3. Current logged-in user has required privileges to create search service 
application
################################################################################>

#Specify Settings for the search service configuration.
#Specify the directory to store search index data. This directory should not contain any data.
$IndexLocation = "D:\SearchIndex” 

#Specify the search application pool name. 
#You can have a new pool or specify any existing one, too.
$SearchAppPoolName = "SearchSvcAppPool" 

#Specify the service account for application pool, in case you want to have a 
#n ew application pool. Else this value will not be  used.
$SearchAppPoolAccountName = "contoso\spsearchpool" 
$SearchServerName = (Get-ChildItem env:computername).value 

#Specify the name for search service application and application proxy
$SearchServiceName = "Contoso Search Service Application" 
$SearchServiceProxyName = "Contoso Search Service Application Proxy"

#Specify the prefix for database names. Pls note that there will be 4 databases created 
# starting with this as prefix. Also they will not have any GUID's in their names. 
$DatabaseName = "Contoso_SearchService_Dev" 

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

Try{
    Start-SPAssignment -Global

    #Check if application pool already exist. If not, create a new one.
    Write-Output "Checking if Search Application Pool exists..." 
    $SearchServiceAppPool = Get-SPServiceApplicationPool -Identity $SearchAppPoolName -ErrorAction SilentlyContinue
    if ($SearchServiceAppPool -eq $null){ 
        Write-Output "Creating Search Application Pool"
        $SearchServiceAppPool = New-SPServiceApplicationPool -Name $SearchAppPoolName -Account $SearchAppPoolAccountName -Verbose
        if(-not $?){
            throw "Failed to create service application pool. Pls check if this account is added to managed accounts and this console is running in admin mode."
        }
    }

    #Starts the Search Service Instance on the specified Server
    Write-Output "Start Search Service instance on this server..." 
    Start-SPEnterpriseSearchServiceInstance $SearchServerName
    Start-SPEnterpriseSearchQueryAndSiteSettingsServiceInstance $SearchServerName
    $Instance = Get-SPEnterpriseSearchServiceInstance -Identity $SearchServerName
    sleep 60
    if($Instance.Status -ne "Online"){
        Write-Output "Waiting for 60 seconds for instance to come online..."
        sleep 60
        for ($Count=1;$Count -lt 5;$Count++){
            $Instance = Get-SPEnterpriseSearchServiceInstance -Identity $SearchServerName
            if($Instance.Status -ne "Online"){
                Write-Output "Waiting for another 60 seconds..."
                sleep 60
            }
            else{
                $Count = 6
            }
        }
        if($Count -eq 5){
            throw "Failed to start search service instance on server. Pls start manually and run script again"
        }
    }

    #Check if Search Service Application already exist. If not, create a new one.
    Write-Output "Checking if Search Service Application exists" 
    $ServiceApplication = Get-SPEnterpriseSearchServiceApplication -Identity $SearchServiceName -ErrorAction SilentlyContinue
    if ($ServiceApplication -eq $null){ 
        Write-Output "Creating Search Service Application..." 
        $ServiceApplication = New-SPEnterpriseSearchServiceApplication -Partitioned -Name $SearchServiceName -ApplicationPool $SearchServiceAppPool.Name -DatabaseName $DatabaseName -Verbose
        if(-not $?){
            throw "Failed to create search service application. Pls check for errors manually."
        }

    }

    #Check if Search Service Application Proxy already exist. If not, create a new one.
    Write-Output "Checking if Search Service Application Proxy exists" 
    $Proxy = Get-SPEnterpriseSearchServiceApplicationProxy -Identity $SearchServiceProxyName -ErrorAction SilentlyContinue
    if ($Proxy -eq $null){ 
        Write-Output "Creating Search Service Application Proxy" 
        New-SPEnterpriseSearchServiceApplicationProxy -Partitioned -Name $SearchServiceProxyName -SearchApplication $ServiceApplication -Verbose
        if(-not $?){
            throw "Failed to create search service application. Pls check for errors manually."
        }
    }

    #Checks if index directory already exists. If yes, removes it and all sub-directories and files. If not, creates it.
    Remove-Item -Recurse -Force -LiteralPath $IndexLocation -ErrorAction SilentlyContinue 
    New-Item -ItemType Directory -Path $IndexLocation -Force

    # Clone the default Topology (which is empty) and create a new one and then activate it 
    Write-Output "Configuring Search Component Topology...." 
    $ClonedTopology = $ServiceApplication.ActiveTopology.Clone() 
    $SSI = Get-SPEnterpriseSearchServiceInstance -local 
    New-SPEnterpriseSearchAdminComponent –SearchTopology $ClonedTopology -SearchServiceInstance $SSI 
    New-SPEnterpriseSearchContentProcessingComponent –SearchTopology $ClonedTopology -SearchServiceInstance $SSI 
    New-SPEnterpriseSearchAnalyticsProcessingComponent –SearchTopology $ClonedTopology -SearchServiceInstance $SSI 
    New-SPEnterpriseSearchCrawlComponent –SearchTopology $ClonedTopology -SearchServiceInstance $SSI 
    New-SPEnterpriseSearchIndexComponent –SearchTopology $ClonedTopology -SearchServiceInstance $SSI -RootDirectory $IndexLocation 
    New-SPEnterpriseSearchQueryProcessingComponent –SearchTopology $ClonedTopology -SearchServiceInstance $SSI 
    
    #Sets new topology as active one
    Write-Output "Setting new search topology as active..."
    Set-SPEnterpriseSearchTopology -Identity $ClonedTopology

    #Displays Current topology components to user
    Write-Output ""
    Write-Output "Here are the current search components:"
    $ActiveToplogy = Get-SPEnterpriseSearchTopology -Active -SearchApplication $ServiceApplication
    Get-SPEnterpriseSearchComponent -SearchTopology $ActiveToplogy    
}
Catch{
    Write-Error "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Error "Exception Message: $($_.Exception.Message)"
}
finally{
    Stop-SPAssignment -Global

    Write-Output ""
    Write-Output "Script Execution finished"
}

<###############################################################################
End of Script
###############################################################################>