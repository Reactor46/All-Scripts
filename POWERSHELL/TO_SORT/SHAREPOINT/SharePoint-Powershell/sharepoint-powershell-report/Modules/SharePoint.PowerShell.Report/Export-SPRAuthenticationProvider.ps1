function Export-SPRAuthenticationProvider
{
  param(
    [string]$Path
  )

  $file = '{0}\SPAuthenticationProvider.xml' -f $Path
  Start-Job -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

    $object = @()
    $webApplications = Get-SPWebApplication -IncludeCentralAdministration
    
    foreach ($webApplication in $webApplications)
    {
      $alternateUrls = Get-SPAlternateURL -WebApplication $webApplication
      foreach ($alternateUrl in $alternateUrls)
      {
        $sprAuthenticationProvider = Get-SPAuthenticationProvider -WebApplication $webApplication -Zone $alternateUrl.Zone
        $sprAuthenticationProvider | Add-Member -MemberType NoteProperty -Name WebApplicationDisplayName -Value $webApplication.DisplayName
        $sprAuthenticationProvider | Add-Member -MemberType NoteProperty -Name WebApplicationUrl -Value $webApplication.Url
        $sprAuthenticationProvider | Add-Member -MemberType NoteProperty -Name Zone -Value $alternateUrl.Zone
        $object += $sprAuthenticationProvider
      }
    }
    $object | Export-Clixml -Path $args[0]
  } -ArgumentList $file
}


