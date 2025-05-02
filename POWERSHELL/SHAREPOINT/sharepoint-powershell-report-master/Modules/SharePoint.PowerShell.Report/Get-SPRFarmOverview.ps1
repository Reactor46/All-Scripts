function Get-SPRFarmOverview
{
  param
  (
    [Parameter(Mandatory)]
    [object]$spfarm,
    [Parameter(Mandatory)]
    [object[]]$SPDatabases
  )
	
  $date = Get-Date
  $configurationDatabase = $SPDatabases | Where-Object -FilterScript {
    $_.Type -eq 'Configuration Database'
  }
  $patchLevel = Get-SPRSharePointBuild -BuildVersion $spfarm.BuildVersion
  $license = Get-SPRSharePointLicense -Products $spfarm.Products
	
  $properties = [ordered]@{
    'FarmName'            = $spfarm.Name
    'BuildVersion'        = $spfarm.BuildVersion.ToString()
    'PatchLevel'          = $patchLevel
    'SharePointLicense'   = $license
    'ConfigurationDatabase' = $configurationDatabase.Name
    'FarmServerCount'     = $spfarm.Servers.Count
    'ReportCreatedOn'     = $date.ToString()
  }
	
  $output = New-Object -TypeName PSObject -Property $properties
	
  Write-Output -InputObject $output
}