#####################################################################################################################################
# Filename: New-SPEnterpriseSearchIndexComponent.ps1
# Version : 1.0
# Description : This script creates new search index partition on another/same server in existing enteprise search topology. This 
#               script assumes that all running crawls have been stopped already, as precautionary step. 
# Written by  : Mohit Goyal
#####################################################################################################################################

#load SharePoint Snap-in
Add-PSSnapin Microsoft.SharePoint.PowerShell

#Mention server name where existing index component exists
$FromServer = "CON-SPAPPVDEV-01"
#Mention server name on which you need to create an additional index component
$Server = "CON-SPAPPVDEV-02"
#Mention directory location for storing search index data for new component
$Directory = "E:\SearchIndex"

Try{
    Start-SPAssignment -Global

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

    #Checks if index directory already exists. If yes, removes it and all sub-directories and files. If not, creates it.
    Remove-Item -Recurse -Force -Path $Directory -ErrorAction SilentlyContinue 
    New-Item -ItemType Directory -Path $Directory -Force

    #Add index component for another server
    Write-Output "Getting existing index component number for specified server..."
    $IndexPartition = Get-SPEnterpriseSearchComponent -SearchTopology $ClonedTopology | Where-Object {$_.servername -eq $FromServer -and $_.name -like "IndexComponent*"} 
    $IndexPartitionNo = $IndexPartition.IndexPartitionOrdinal
    $IndexComponentName = $IndexPartition.Name

    Write-Output "Adding index component number for specified server..."
    New-SPEnterpriseSearchIndexComponent -SearchTopology $ClonedTopology -RootDirectory $Directory -IndexPartition $IndexPartitionNo -SearchServiceInstance $Instance

    #Pause search service application for repartitioning
    Write-Output "Pausing search service application..."
    $SSA.PauseForIndexRepartitioning() 

    #Sets new search toplogy as active on the server
    Write-Output "Setting new search topology as active. This might take several hours."
    Set-SPEnterpriseSearchTopology -Identity $ClonedTopology

    #Monitors progress of index partitioning
    $SSA = Get-SPEnterpriseSearchServiceApplication
    $Status = Get-SPEnterpriseSearchStatus -SearchApplication $SSA -HealthReport -Component $IndexComponentName | Where-Object { ($_.name -match "repart") -or ( $_.name -match "splitting")}
    While($Status -ne $null){
        Write-Output "Waiting for 5 mins for new index partition to be completed..."
        Sleep 300
        $SSA = Get-SPEnterpriseSearchServiceApplication
        $Status = Get-SPEnterpriseSearchStatus -SearchApplication $SSA -HealthReport -Component $IndexComponentName | Where-Object { ($_.name -match "repart") -or ( $_.name -match "splitting")}
    }

    #Displays Current topology components to user
    Write-Output ""
    Write-Output "Here are the current search components:"
    $ActiveToplogy = Get-SPEnterpriseSearchTopology -Active -SearchApplication $SSA
    Get-SPEnterpriseSearchComponent -SearchTopology $ActiveToplogy

    #Resume search service application for repartitioning
    Write-Output "Resuming search service application..."
    $SSA.ResumeAfterIndexRepartitioning()

    Write-Output ""
    Write-Output "Script Execution finished"
}

Catch{
    Write-Error "Exception Type: $($_.Exception.GetType().FullName)"
    Write-Error "Exception Message: $($_.Exception.Message)"
    Write-Output ""
    Write-Output "Script Execution finished"
}
finally{
    Stop-SPAssignment -Global
}

#########################################################################################################################################
## End of Script
#########################################################################################################################################