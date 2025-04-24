function Get-SPRServerProcessorInfo
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Server'        = $XXX
      'Name'          = $XXX
      'MaxClockSpeed' = $XXX
      'NumberOfCores' = $XXX
      'CPUManufacturer' = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}