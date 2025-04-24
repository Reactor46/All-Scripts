function Get-SPRSearchManagedProperties
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
      'PropertyName'     = $XXX
      'Type'             = $XXX
      'MayBeDeleted'     = $XXX
      'UseInScopers'     = $XXX
      'Optimized'        = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}