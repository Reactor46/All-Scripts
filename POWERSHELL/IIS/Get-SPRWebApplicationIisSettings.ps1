function Get-SPRWebApplicationIisSettings
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'WebApplication' = $XXX
      'Path'          = $XXX
      'Authentication' = $XXX
      'ApplicationPool' = $XXX
      'SSL'           = $XXX
      'ClaimsAuth'    = $XXX
      'CEIP'          = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}