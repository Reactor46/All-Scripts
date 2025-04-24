function Export-SPRUserProfileManager
{
  param
  (
    [String]
    [Parameter(Mandatory)]
    $Path
  )
  
  $file = '{0}\SPUserProfileManager.xml' -f $Path
  Start-Job -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

    $webAppUrl = Get-SPWebApplication -IncludeCentralAdministration |
    Select-Object -First 1  |
    Select-Object -ExpandProperty URL
    $sc = Get-SPServiceContext($webAppUrl)
    $userProfileManager = New-Object -TypeName Microsoft.Office.Server.UserProfiles.UserProfileManager -ArgumentList ($sc)
        
    $userProfileManager | Export-Clixml -Path $args[0]
  } -ArgumentList $file
}
