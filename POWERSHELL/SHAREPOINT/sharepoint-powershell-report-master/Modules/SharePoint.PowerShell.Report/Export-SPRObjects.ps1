function Export-SPRObjects
{
  [OutputType([void])]
  param
  (
    [Parameter(
        Mandatory = $true,
    Position = 0)]
    [string]$Path
  )

  if (!(Test-Path -PathType Container -Path $Path))
  {
    mkdir -Path $Path
  } 
	
  Add-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
  
  Export-SPRFarmConfiguration -Path $Path
  Export-SPRDatabaseConfiguration -Path $Path
  Export-SPRServerConfiguration -Path $Path
  Export-SPRWebApplicationConfiguration -Path $Path
  Export-SPRSite -Path $Path
  Export-SPRServiceInstanceConfiguration -Path $Path
  Export-SPREnterpriseSearchServiceApplicationConfiguration -Path $Path
  Export-SPREnterpriseSearchTopologyConfiguration -Path $Path
  Export-SPREnterpriseSearchComponent -Path $Path
  Export-SPRFeature -Path $Path
  Export-SPRManagedPath -Path $Path
  Export-SPRManagedAccount -Path $Path
  Export-SPRAlternateUrl -Path $Path
  Export-SPRAuthenticationProvider -Path $Path
  Export-SPRDiagnosticLogging -Path $Path
  Export-SPRContentDeploymentSettings -Path $Path
  Export-SPRCacheClusterConfig -Path $Path
  Export-SPRCacheHost -Path $Path
  Export-SPRLogLevel -Path $Path
  Export-SPRFarmAdmins -Path $Path
  Export-SPRWebServices -Path $Path
  Export-SPRWebServicesAdministration -Path $Path
  Export-SPRScheduledTasks -Path $Path
  Export-SPRUserProfileManager -Path $Path
  Export-SPRUserProfileServiceApplication -Path $Path
  Export-SPRQuotaTemplates -Path $Path
  Export-SPREnterpriseSearchServiceApplicationInstance -Path $Path
    
  Write-Host -Object 'Waiting on exports to complete...' -NoNewline
  $null = Get-Job | 
  Wait-Job
  Write-Host -Object 'Done.'

  Remove-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
}

