function Get-SPRSqlServerCertificates
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
      'Property' = $XXX
      'Value'  = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}