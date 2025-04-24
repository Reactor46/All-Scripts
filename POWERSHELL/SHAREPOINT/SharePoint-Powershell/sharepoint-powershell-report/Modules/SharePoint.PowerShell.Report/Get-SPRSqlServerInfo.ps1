function Get-SPRSqlServerInfo
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Server'  = $XXX
      'Edition' = $XXX
      'Version' = $XXX
      'Release' = $XXX
      'SQLAccont' = $XXX
      'IsCluster' = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}