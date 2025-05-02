# Try  Remove-SPDistributedCacheServiceInstance and  Add-SPDistributedCacheServiceInstance first!
Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue
$instanceName ="SPDistributedCacheService Name=AppFabricCachingService"
$serviceInstance = Get-SPServiceInstance | ? {($_.service.tostring()) -eq $instanceName -and ($_.server.name) -eq $env:computername}
If($serviceInstance -ne $null)
{
$serviceInstance.Delete()
}