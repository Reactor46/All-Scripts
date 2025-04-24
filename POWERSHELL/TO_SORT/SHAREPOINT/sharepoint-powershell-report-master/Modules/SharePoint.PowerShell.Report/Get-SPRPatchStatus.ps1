function Get-SPRPatchStatus
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Server' = $XXX
      'SkuName' = $XXX
      'Product' = $XXX
      'Version' = $XXX
      'Status' = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}