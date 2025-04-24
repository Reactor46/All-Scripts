function Get-SPRServiceApplicationProxies
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'ServiceApplication' = $XXX
      'Name'             = $XXX
      'Type'             = $XXX
      'Status'           = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}