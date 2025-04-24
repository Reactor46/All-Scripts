function Get-SPRSiteTopology
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$spsite
  )
	
  foreach ($site in $spsite)
  {
    $properties = [ordered]@{
      'WebApplication' = '{0} - {1}' -f $site.WebApplication.DisplayName, $site.WebApplication.Url
      'SiteCollection' = '{0} ({1})' -f $site.RootWeb.Title, $site.Url
      'ContentDatabase' = $site.ContentDatabase.Name
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}