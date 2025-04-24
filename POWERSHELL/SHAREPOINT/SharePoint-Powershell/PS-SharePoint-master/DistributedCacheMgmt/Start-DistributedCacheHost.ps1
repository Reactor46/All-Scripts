[CmdletBinding()]
param ()
begin {
}

process {

    $instanceName ="SPDistributedCacheService Name=AppFabricCachingService" 
    $serviceInstance = Get-SPServiceInstance | ? {($_.service.tostring()) -eq $instanceName -and ($_.server.name) -eq $env:computername} 
    if ($serviceInstance)
    {
        $serviceInstance.Provision()
    }
    else
    {
        Write-Error "Distributed Cache service instance waas not found. Verify that $($env:computername) is added to Distributed Cache cluster before starting the service instance."
    }

}