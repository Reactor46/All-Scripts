function Get-SPRSqlServerAliases
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Alias' = $XXX
      'Server' = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}