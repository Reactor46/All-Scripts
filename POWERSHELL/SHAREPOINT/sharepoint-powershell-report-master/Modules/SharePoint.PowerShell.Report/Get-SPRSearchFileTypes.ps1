function Get-SPRSearchFileTypes
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPEnterpriseSearchServiceApplication
  )
	
  foreach ($searchServiceApplication in $SPEnterpriseSearchServiceApplication)
  {
    $filetypes = $searchServiceApplication.CrawlFileTypes.FileExtension -join (',')
    $properties = [ordered]@{
      'ServiceApplication' = $searchServiceApplication.Name
      'FileTypes'        = $filetypes
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}