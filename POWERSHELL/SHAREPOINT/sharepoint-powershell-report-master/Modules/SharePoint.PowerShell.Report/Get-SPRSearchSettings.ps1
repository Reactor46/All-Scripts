function Get-SPRSearchSettings
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPEnterpriseSearchServiceApplication,
    [Parameter(Mandatory = $true)]
    [object[]]$SPEnterpriseSearchServiceInstance
    
  )
	
  foreach ($searchServiceApplication in $SPEnterpriseSearchServiceApplication)
  {
    $properties = [ordered]@{
      'ServiceApplication'        = $searchServiceApplication.Name
      'DefaultContentAccessAccount' = $searchServiceApplication.Content.DefaultGatheringAccount
      'ContactEmailAddress'       = $XXX
      'IndexLocation'             = $SPEnterpriseSearchServiceInstance.CrawlComponents.IndexLocation
      'SearchAlertsStatus'        = $searchServiceApplication.AlertsEnabled
      'QueryLogging'              = $searchServiceApplication.QueryLoggingEnabled
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}
