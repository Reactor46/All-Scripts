function Get-SPRSearchFederatedLocations
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPEnterpriseSearchServiceApplication
  )
	
  $output = @()
  foreach ($searchServiceApplication in $SPEnterpriseSearchServiceApplication)
  {
    foreach($locationConfiguration in $searchServiceApplication.LocationConfiguration){
        
    }

    $properties = [ordered]@{
      'ServiceApplication' = $searchServiceApplication.Name
      'DisplayName'      = $locationConfiguration.Name
      'Author'           = $locationConfiguration.Author
      'Version'          = $locationConfiguration.Version
      'LocationType'     = $locationConfiguration.Type
      'Trigger'          = $locationConfiguration.Visualizations
      'CreationDate'     = $locationConfiguration.CreationDate
    }
    $entry = New-Object -TypeName PSObject -Property $properties
		
    $output += $entry
  }
  Write-Output -InputObject $output
}