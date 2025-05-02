[CmdletBinding()]

## Settings you may want to change for your scenario ##
$startTime = Get-Date
$currentTime = $startTime
$elapsedTime = $currentTime - $startTime
$timeOut = 600
$currentHostFQDN = (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain

try

{

Use-CacheCluster
Get-AFCacheClusterHealth

Write-Verbose "Shutting down distributed cache host."
$hostInfo = Stop-CacheHost -Graceful -CachePort 22233 -HostName $currentHostFQDN

while($elapsedTime.TotalSeconds -le $timeOut-and $hostInfo.Status -ne 'Down')
{
Write-Verbose "Host Status : [$($hostInfo.Status)]"
Start-Sleep(5)
$currentTime = Get-Date
$elapsedTime = $currentTime - $startTime
#Get-AFCacheClusterHealth
$hostInfo = Get-CacheHost -HostName $currentHostFQDN -CachePort 22233
}

Write-Verbose "Stopping distributed cache host was successful. Updating Service status in SharePoint."
Stop-SPDistributedCacheServiceInstance
Write-Verbose "To start service, please use Central Administration site."
}
catch [System.Exception]
{
Write-Error "Unable to stop cache host within $($timeOut/60) minutes."
} 