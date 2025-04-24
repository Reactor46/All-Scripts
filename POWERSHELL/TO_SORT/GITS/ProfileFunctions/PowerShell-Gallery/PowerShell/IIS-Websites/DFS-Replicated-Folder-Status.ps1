#Restart a Web App Pool and IIS a Wesbsite
#CopyLeft SVET :)
#https://www.resilio.com/blog/dfs-replication-not-working
#
#Get-DfsrBacklog
#This command shows pending updates between two Windows-based file servers that participate in DFSRfile replication service.
#
#Get-DfsrState
#This command shows you the current replication state of DFSR in regard to its DFS replication group partners.

$allyesno = Read-Host -Prompt 'Do you want to check the prending replica updates? y/n'

if ($yesno -like 'y'){

Get-DfsrBacklog -SourceComputerName wh -DestinationComputerName wh1

Write-Host "The Window will close in 5 seconds"
Start-Sleep -Seconds 5
}

else {
    $replicatedfolder = Read-Host -Prompt 'Enter Replciated Web Site Fodler Name'
	Write-Host "`n$line"
	Get-DfsrBacklog -SourceComputerName wh -DestinationComputerName wh1 | findstr *$replicatedfolder*
}
