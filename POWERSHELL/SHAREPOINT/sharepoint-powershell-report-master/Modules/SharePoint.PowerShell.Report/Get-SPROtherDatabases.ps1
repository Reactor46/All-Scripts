function Get-SPROtherDatabases
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Name'        = $XXX
      'Status'      = $XXX
      'Type'        = $XXX
      'Server'      = $XXX
      'RecoveryModel' = $XXX
      'Size'        = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}