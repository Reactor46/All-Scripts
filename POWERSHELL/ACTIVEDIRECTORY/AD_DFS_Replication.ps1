$domains = "Lab.local","DRLab.local","contoso.local"
foreach ($doamin in $domains){
$DCs = (Get-ADForest $doamin).GlobalCatalogs
foreach ($DC in $DCs){
$DCCheck = Get-WmiObject -ComputerName $DC -Namespace "root\microsoftdfs" -Class dfsrreplicatedfolderinfo  | 
Select-Object -Property PSComputerName, Replicationgroupname, Replicatedfoldername, State
$DCCheck
if ($DCCheck.State -ne "4")
{
Write-Warning "AD DFS replication is not working on $DC will run repair command"
$DFSRep = Get-WmiObject -ComputerName $DC -Namespace "root\microsoftdfs" -Class dfsrVolumeConfig
$DFSRep.ResumeReplication()
}
        }
    }