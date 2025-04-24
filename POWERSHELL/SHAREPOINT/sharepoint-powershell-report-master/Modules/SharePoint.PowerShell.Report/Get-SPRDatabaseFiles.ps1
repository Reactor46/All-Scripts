function Get-SPRDatabaseFiles
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$XXX
  )
	
  foreach ($XXX in $XXXs)
  {
    $properties = [ordered]@{
      'Online'        = $XXX
      'FileGrowth'    = $XXX
      'FileGrowthValue' = $XXX
      'FileSize'      = $XXX
      'MaxFileSize'   = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}