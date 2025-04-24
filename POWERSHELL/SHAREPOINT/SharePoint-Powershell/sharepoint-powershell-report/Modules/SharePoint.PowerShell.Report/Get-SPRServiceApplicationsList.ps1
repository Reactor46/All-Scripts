function Get-SPRServiceApplicationsList
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPServiceApplication
  )
	
  foreach ($serviceApplication in $SPServiceApplication)
  {
    $properties = [ordered]@{
      'Name'          = $serviceApplication.Name
      'Status'        = $serviceApplication.Status
      'ApplicationPool' = $serviceApplication.ApplicationPool.DisplayName
      'Databases'     = $XXX
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}