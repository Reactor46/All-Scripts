function Get-SPRSearchContentSources
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPEnterpriseSearchServiceApplication
  )
	$output = @()
  foreach ($searchServiceApplication in $SPEnterpriseSearchServiceApplication)
  {
    
    foreach ($contentSource in $searchServiceApplication.CrawlContentSources) {
            $properties = [ordered]@{
              'ServiceApplication' = $searchServiceApplication.Name
              'ContentSourceName'             = $contentSource.Name
              'ContentSourceType'             = $contentSource.Type
        }
        $entry = New-Object -TypeName PSObject -Property $properties
        $output += $entry
    }

    
		
    Write-Output -InputObject $output
  }
}