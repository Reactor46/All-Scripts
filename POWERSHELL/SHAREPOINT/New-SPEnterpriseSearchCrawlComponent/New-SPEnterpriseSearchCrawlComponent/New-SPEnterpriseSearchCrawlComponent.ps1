######################################################################################################
# Filename: New-SPEnterpriseSearchCrawlComponent.ps1
# Version : 1.0
# Description : This script creates a new crawl component on another server in existing 
#               enteprise search topology. This sript assumes that all running crawls have been stopped
#               already, as precautionary step. 
# Written by  : Mohit Goyal
######################################################################################################

#load SharePoint Snap-in
Add-PSSnapin Microsoft.SharePoint.PowerShell

#Mention server name on which you need to create an additional crawl component
$Server = "con-spappdev-02"

Try{

    #Gets the existing search service application
    Write-Output "Fetching search service application..."
    $SSA = Get-SPEnterpriseSearchServiceApplication
    if(!$?)
    {
        throw "Unable to fetch search service application. Pls make sure its up and running fine"
    }

    #Gets the Search Service Instance on the specified Server
    Write-Output "Fetching search service instance on server $Server..."
    $Instance = Get-SPEnterpriseSearchServiceInstance -Identity $Server
    if(!$?){
        throw "Unable to fetch search service instance on the server $Server"
    }

    #Starts Search Service Instance on the specified server.
    if($Instance.Status -ne "Online"){
        Write-Output "Starting service instance..."   
        Start-SPEnterpriseSearchServiceInstance $Instance
        Write-Output "Waiting for 60 seconds for instance to come online..."
        sleep 60
        for ($Count=1;$Count -lt 5;$Count++){
            $Instance = Get-SPEnterpriseSearchServiceInstance -Identity $Server
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

    #Clone existing topology
    Write-Output "Fetching active topology and cloning it..."
    $ActiveToplogy = Get-SPEnterpriseSearchTopology -Active -SearchApplication $SSA
    $ClonedTopology = New-SPEnterpriseSearchTopology -SearchApplication $SSA -Clone –SearchTopology $ActiveToplogy

    #Add crawl component for another server and set topology as active
    Write-Output "Adding crawl component for specified server..."
    New-SPEnterpriseSearchCrawlComponent -SearchTopology $ClonedTopology -SearchServiceInstance $Instance -SearchApplication $SSA
    Write-Output "Setting new search topology as active..."
    Set-SPEnterpriseSearchTopology -Identity $ClonedTopology

    #Displays Current topology components to user
    Write-Output ""
    Write-Output "Here are the current search components:"
    $ActiveToplogy = Get-SPEnterpriseSearchTopology -Active -SearchApplication $SSA
    Get-SPEnterpriseSearchComponent -SearchTopology $ActiveToplogy

    Write-Output ""
    Write-Output "Script Execution finished"
}

Catch{
    Write-Error "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Error "Exception Message: $($_.Exception.Message)"
    Write-Output ""
    Write-Output "Script Execution finished"
}

#######################################################################################################################
## End of Script
#######################################################################################################################