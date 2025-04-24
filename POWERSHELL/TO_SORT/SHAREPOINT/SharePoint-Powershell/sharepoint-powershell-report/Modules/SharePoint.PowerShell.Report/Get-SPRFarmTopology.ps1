function Get-SPRFarmTopology
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$spserver,
    [Parameter(Mandatory = $true)]
    [object[]]$spserviceinstance
  )
	
  foreach ($server in $spserver)
  {
    $computerName = Get-SPRComputerName -Address $server.Address
    $ipAddress = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ComputerName $computerName |
    Where-Object -FilterScript {
      $_.IPAddress -ne $null
    } |
    Select-Object -ExpandProperty IPAddress
    $centralAdminServiceInstance = $spserviceinstance | Where-Object -FilterScript {
      $_.TypeName -eq 'Central Administration' -and (($_.Server|Select-Object -ExpandProperty Name) -eq $server.Address)
    }
    if ($centralAdminServiceInstance -ne $null)
    {
      $hostsCentralAdmin = $true
    }
    else
    {
      $hostsCentralAdmin = $false
    }
		
    $properties = [ordered]@{
      'Name'                = $computerName
      'IPAddress'           = $ipAddress
      'Role'                = $server.Role
      'CentralAdministration' = $hostsCentralAdmin
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}