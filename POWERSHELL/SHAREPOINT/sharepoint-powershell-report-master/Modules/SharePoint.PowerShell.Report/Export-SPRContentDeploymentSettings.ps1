function Export-SPRContentDeploymentSettings
{
  param(
    [string]$Path
  )

  $file = '{0}\SPContentDeploymentSettings.xml' -f $Path
  $object=[Microsoft.SharePoint.Publishing.Administration.ContentDeploymentConfiguration]::GetInstance()

  Start-Job -Name 'Content Deployment Settings'-ScriptBlock {
  $args[1] | Export-Clixml -Path $args[0]
  } -ArgumentList $file,$object

  

 

}
