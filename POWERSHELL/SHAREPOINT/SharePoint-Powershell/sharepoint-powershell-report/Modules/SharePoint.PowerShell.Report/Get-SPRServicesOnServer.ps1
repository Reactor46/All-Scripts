function Get-SPRServicesOnServer
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPServiceInstance,
    [Parameter(Mandatory = $true)]
    [object[]]$SPServer
  )
  

  foreach ($serviceInstance in $SPServiceInstance)
  {
    $properties = [ordered]@{
      'ServiceName' = $serviceInstance.TypeName
    }
    
    
    foreach ($server in $SPServer)
    {
      $serverName = $server.Address
      $serviceStatus = $serviceInstance.Status
      
      $properties.Add($serverName,$serviceStatus)
    }
    
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}