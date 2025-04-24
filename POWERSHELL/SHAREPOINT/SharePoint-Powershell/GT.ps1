$farm  = Get-SPFarm 
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
