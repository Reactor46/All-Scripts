function Get-SPRServerDiskList
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Server'   = $XXX
      'Name'     = $XXX
      'VolumeName' = $XXX
      'Size'     = $XXX
      'FreeSpace' = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}