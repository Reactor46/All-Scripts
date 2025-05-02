function Get-SPRWebApplicationsAndSiteCollections
{
  param
  (
    [parameter(Mandatory = $true)]
    [object[]]$spsite
		
  )
	
  foreach ($site in $spsite)
  {
    $siteAdmins = $site.RootWeb.SiteAdministrators | ForEach-Object -Process {
      ConvertTo-SPRLoginName -Login $_
    }
    $webApplication = $site.WebApplication
		
    $properties = [ordered]@{
      'WebApplication' = $webApplication.DisplayName
      'SiteCollection' = $site.Url
      'SiteAdmins'   = $siteAdmins
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}