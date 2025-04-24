function Get-SPRServersInFarm
{
  param
  (
    [parameter(Mandatory = $true)]
    [object[]]$spserver
  )
	
  foreach ($server in $spserver)
  {
    $computerName = Get-SPRComputerName -Address $server.Address
		
    $os = Get-CimInstance -ClassName CIM_OperatingSystem -ComputerName $computerName
    $memory = Get-CimInstance -ClassName CIM_PhysicalMemory -ComputerName $computerName
		
    $properties = [ordered]@{
      'ServerName'    = $server.Address
      'Role'          = $server.Role
      'OperatingSystem' = $os.Caption
      'Memory'        = $memory.Capacity
    }
		
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}