function Get-SPRDisrtibutedCacheConfig
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPCacheClusterConfig,
    [Parameter(Mandatory = $true)]
    [object[]]$SPCacheHost
  )

foreach ($entry in $SPCacheClusterConfig.configuration.dataCache.hosts.host)
{
    foreach ($o2 in $SPCacheHost)
    {
        if ($entry.Name -eq $o2.HostName) {
            $properties = [ordered]@{
            'Host'          = $entry.Name
            'Size'          = $entry.Size
            'CachePort'     = $entry.CachePort
            'ClusterPort'   = $entry.ClusterPort
            'ReplicationPort' = $entry.ReplicationPort
            'ArbitrationPort' = $entry.ArbitrationPort
            'Status'        = $o2.status.tostring()
            }
        $output = New-Object -TypeName PSObject -Property $properties
        Write-Output -InputObject $output
        } 

    }

}

}