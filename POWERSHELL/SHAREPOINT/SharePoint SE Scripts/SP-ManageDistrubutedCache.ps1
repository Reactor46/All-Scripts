# Manage the Distributed Cache service in SharePoint Server

# To start the Distributed Cache service by using SharePoint Management Shell
$instanceName ="SPDistributedCacheService Name=SPCache"
$serviceInstance = Get-SPServiceInstance | ? {($_.service.tostring()) -eq $instanceName -and ($_.server.name) -eq $env:computername}
$serviceInstance.Provision()

# To stop the Distributed Cache service by using SharePoint Management Shell
$instanceName ="SPDistributedCacheService Name=SPCache"
$serviceInstance = Get-SPServiceInstance | ? {($_.service.tostring()) -eq $instanceName -and ($_.server.name) -eq $env:computername}
$serviceInstance.Unprovision()

# Change the memory allocation of the Distributed Cache
# (Optional) To check the existing memory allocation for the Distributed Cache service on a server,
# run the following command at the SharePoint Management Shell command prompt:
Use-SPCacheCluster
Get-SPCacheHostConfig -HostName $Env:ComputerName

# Stop the Distributed Cache service on all cache hosts. 
# To stop the Distributed Cache service, go to Services on Server in Central Administration,
# and Stop the Distributed Cache service on all cache hosts in the farm.
# To reconfigure the cache size of the Distributed Cache service, run the following command
# one time only on any cache host at the SharePoint Management Shell command prompt:
Update-SPDistributedCacheSize -CacheSizeInMB CacheSize

# Add a server to the cache cluster and starting the Distributed Cache service by using SharePoint Management Shell
# At the SharePoint Management Shell command prompt, run the following command:
Add-SPDistributedCacheServiceInstance

# Remove a server from the cache cluster by using SharePoint Management Shell
# At the SharePoint Management Shell command prompt, run the following command:
Remove-SPDistributedCacheServiceInstance

# Perform a graceful shutdown of the Distributed Cache service by using a PowerShell script

## Settings you may want to change for your scenario ##
$startTime = Get-Date
$currentTime = $startTime
$elapsedTime = $currentTime - $startTime
$timeOut = 900
Use-CacheCluster
try
{
    Write-Host "Shutting down distributed cache host."
 $hostInfo = Stop-CacheHost -Graceful -CachePort 22233 -ComputerName sp2016App.contoso.com
 while($elapsedTime.TotalSeconds -le $timeOut-and $hostInfo.Status -ne 'Down')
 {
     Write-Host "Host Status : [$($hostInfo.Status)]"
     Start-Sleep(5)
     $currentTime = Get-Date
     $elapsedTime = $currentTime - $startTime
     $hostInfo = Get-CacheHost -HostName SP2016app.contoso.com -CachePort 22233
 }
 Write-Host "Stopping distributed cache host was successful. Updating Service status in SharePoint."
 Stop-SPDistributedCacheServiceInstance
 Write-Host "To start service, please use Central Administration site."
}
catch [System.Exception]
{
 Write-Host "Unable to stop cache host within 15 minutes." 
}