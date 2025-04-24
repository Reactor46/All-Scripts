function Get-SPRSearchServiceApplicationTopology
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPEnterpriseSearchComponent
  )
	
  $serverGroups = $SPEnterpriseSearchComponent | Group-Object -Property ServerID
	
	
  foreach ($g in $serverGroups) 
  {
    $serverName = $g.Group | Select-Object -First 1 -ExpandProperty ServerName
    $properties = [ordered]@{
      'ServerName'                 = $serverName
      'AdminComponent'             = ($g.Group | Where-Object -FilterScript {
          $_.Name -like 'AdminComponent*'
      }) -ne $null
      'AnalyticsProcessingComponent' = ($g.Group | Where-Object -FilterScript {
          $_.Name -like 'AnalyticsProcessingComponent*'
      }) -ne $null
      'ContentProcessingComponent' = ($g.Group | Where-Object -FilterScript {
          $_.Name -like 'ContentProcessingComponent*'
      }) -ne $null
      'CrawlComponent'             = ($g.Group | Where-Object -FilterScript {
          $_.Name -like 'CrawlComponent*'
      }) -ne $null
      'IndexComponent'             = ($g.Group | Where-Object -FilterScript {
          $_.Name -like 'IndexComponent*'
      }) -ne $null
      'QueryProcessingComponent'   = ($g.Group | Where-Object -FilterScript {
          $_.Name -like 'QueryProcessingComponent*'
      }) -ne $null
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}