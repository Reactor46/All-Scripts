function Get-SPRContentDeploymentSettings
{
	
 param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPContentDeploymentSettings
  )


    $properties = [ordered]@{
      'AcceptIncomingJobs' = $SPContentDeploymentSettings.AcceptIncomingJobs
      'ExportWebServer'     = $SPContentDeploymentSettings.ExportWebServer
      'ImportWebServer'=$SPContentDeploymentSettings.ImportWebServer
      'TemporaryFolder'=$SPContentDeploymentSettings.TemporaryFolder
      'RequireSecureConnection'=$SPContentDeploymentSettings.RequiresSecureConnection
      }
   
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
}