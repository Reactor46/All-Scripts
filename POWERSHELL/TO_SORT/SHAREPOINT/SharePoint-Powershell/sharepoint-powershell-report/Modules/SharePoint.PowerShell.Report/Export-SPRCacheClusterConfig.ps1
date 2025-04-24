function Export-SPRCacheClusterConfig
{
  param(
    [string]$Path
  )

  $file = '{0}\SPCacheClusterConfig.xml' -f $Path
  Start-Job -Name 'Cache Config' -ScriptBlock {

    Use-CacheCluster
    Export-CacheClusterConfig -Path c:\SPCacheClusterConfig.xml
    $object = [xml](get-content c:\SPCacheClusterConfig.xml)
    $object | Export-Clixml -Path $args[0]
    Remove-Item c:\SPCacheClusterConfig.xml
    
  } -ArgumentList $file
}
