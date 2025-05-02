function Get-SPRContentDatabases
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPDatabase
  )
	
  $contentDatabases = $spdatabase | Where-Object -FilterScript {
    $_.Type -eq 'Content Database'
  }
	
  foreach ($contentDatabase in $contentDatabases)
  {
  $DBConnectionString=$contentDatabase.DatabaseConnectionString
  if ($DBConnectionString -like "*Integrated Security=True*") {$AuthenticationType='Windows'}
  else {$AuthenticationType='SQL Authentication'}
    $properties = [ordered]@{
      'Name'             = $contentDatabase.Name
      'Status'           = $contentDatabase.Status.tostring()
      'Server'           = $contentDatabase.Server
      'WebApplication'   = $contentDatabase.WebApplication.Url
      'CurrentSiteCount' = $contentDatabase.CurrentSiteCount
      'WarningSiteCount' = $contentDatabase.WarningSiteCount
      'MaximumSiteCount' = $contentDatabase.MaximumSiteCount
      'AuthenticationType' = $AuthenticationType
      'ReadOnly'         = $contentDatabase.IsReadOnly
      'RecoveryModel'    = $xxx
      'TotalSize'        = $contentDatabase.DiskSizeRequired
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}