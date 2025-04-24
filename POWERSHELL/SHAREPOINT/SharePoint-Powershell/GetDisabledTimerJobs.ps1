Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue

$farm = Get-SPFarm
$disabledTimers = $farm.TimerService.Instances | where {$_.Status -ne "Online"}

if ($disabledTimers -ne $null)
{
foreach ($timer in $disabledTimers)
{
Write-Host "Timer service instance on server " $timer.Server.Name " is not Online. Current status:" $timer.Status
Write-Host "Attempting to set the status of the service instance to online"
$timer.Status = [Microsoft.SharePoint.Administration.SPObjectStatus]::Online
$timer.Update()
}
}
else
{
Write-Host "All Timer Service Instances in the farm are online! No problems found"
}

#Get-SPTimerJob | ?{$_.schedule.description -eq "One-time"} |select displayname,server,locktype,lastruntime | fl


#$job = Get-SPTimerJob | ?{$_.schedule.description -eq "One-time"}
#$job[0].Delete()