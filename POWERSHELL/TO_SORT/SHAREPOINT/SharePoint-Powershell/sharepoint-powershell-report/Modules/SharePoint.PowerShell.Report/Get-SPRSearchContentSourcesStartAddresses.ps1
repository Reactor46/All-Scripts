function Get-SPRSearchContentSourcesStartAddresses
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPEnterpriseSearchServiceApplication
  )
  $output = @()
  foreach ($searchServiceApplication in $SPEnterpriseSearchServiceApplication)
  {
    foreach ($contentSource in $searchServiceApplication.CrawlContentSources) 
    {
      foreach ($startAddress in $contentSource.startAddresses)
      {
        $properties = [ordered]@{
          'ServiceApplication'      = $searchServiceApplication.Name
          'ContentSourceName'       = $contentSource.Name
          'ContentSourceStartAddress' = $startAddress.AbsoluteUri
        }
        $entry = New-Object -TypeName PSObject -Property $properties
        $output += $entry
      }
    }

    Write-Output -InputObject $output
  }
}
