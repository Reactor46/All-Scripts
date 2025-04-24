function Export-SPRCacheHost
{
  param(
    [string]$Path
  )

  $file = '{0}\SPCacheHost.xml' -f $Path
  Start-Job -Name 'Cache Host' -ScriptBlock {

    Use-CacheCluster
    $object=Get-CacheHost
    $object | Export-Clixml -Path $args[0]
    
  } -ArgumentList $file
}
