function Get-SPRSiteCollectionQuotas
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPSite,
    [Parameter(Mandatory = $true)]
    [object[]]$SPQuotaTemplate
  )
	
  foreach ($site in $SPSite)
  {
    $quotaID = $site.Quota.QuotaID
    $quotaName = 'Individual Quota'
    foreach ($quotaTemplate in $SPQuotaTemplate) {
        if ($quotaTemplate.QuotaID -eq $site.Quota.QuotaID) {
            $quotaName = $quotaTemplate.Name
        }
    }

    $properties = [ordered]@{
      'SiteCollection'    = $site.RootWeb.Title
      'URL'               = $site.RootWeb.Url
      'QuotaName'         = $quotaName
      'LockStatus'        = $site.IsReadLocked
      'StorageMaximumLevel' = $site.Quota.StorageMaximumLevel
      'StorageWarningLevel' = $site.Quota.StorageWarningLevel
      'UsageStorage'      = $site.Usage.Storage
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}