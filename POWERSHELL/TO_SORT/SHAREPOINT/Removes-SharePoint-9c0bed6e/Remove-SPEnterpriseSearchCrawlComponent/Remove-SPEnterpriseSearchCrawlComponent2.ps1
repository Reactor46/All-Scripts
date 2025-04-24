#####################################################################################################################################
# Filename: Remove-SPEnterpriseSearchCrawlComponent2.ps1
# Version : 1.0
# Description : This script removes existing crawl component on a server in existing enteprise search topology. This 
#               script assumes that all running crawls have been stopped already, as precautionary step. 
# Written by  : Mohit Goyal
#####################################################################################################################################

[CmdletBinding()]
Param(
    [parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [alias("ServerName","MachineName","Name")]
    [ValidateNotNullOrEmpty()]
    [String]
    $ComputerName
)


Function Remove-SPSearchCrawlComponent([string] $Server)
{

    Try{
        #Load SharePoint Snap-in
        Add-PSSnapin Microsoft.SharePoint.PowerShell

        #Gets the existing search service application
        Write-Output "Fetching search service application..."
        $SSA = Get-SPEnterpriseSearchServiceApplication
        if(!$?)
        {
            throw "Unable to fetch search service application. Pls make sure its up and running fine"
        }

        #Clone existing topology
        Write-Output "Fetching active topology and cloning it..."
        $ActiveToplogy = Get-SPEnterpriseSearchTopology -Active -SearchApplication $SSA
        $ClonedTopology = New-SPEnterpriseSearchTopology -SearchApplication $SSA -Clone –SearchTopology $ActiveToplogy
        if(!$?){
            throw "Unable to clone existing search topology."
        }
    
        #Removes crawl component from specified server 
        Write-Output "Removing crawl component from specified server..."
        $Components = Get-SPEnterpriseSearchComponent -SearchTopology $ClonedTopology | Where-Object {$_.Name -like "crawlcomponent*" -and $_.Servername -eq $Server}
        $Components | Remove-SPEnterpriseSearchComponent -SearchTopology $ClonedTopology -Confirm:$false
    
        #Sets cloned topology as active
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
}

#Main Block Starts here
Remove-SPSearchCrawlComponent $ComputerName

#########################################################################################################################################
## End of Script
#########################################################################################################################################