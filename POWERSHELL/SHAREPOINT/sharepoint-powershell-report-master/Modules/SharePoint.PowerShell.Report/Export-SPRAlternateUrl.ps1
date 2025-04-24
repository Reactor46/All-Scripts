function Export-SPRAlternateUrl
{
  param(
    [string]$Path
  )

  $file = '{0}\SPAlternateURL.xml' -f $Path
  Start-Job -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

    $object = @()
    $webApplications = Get-SPWebApplication -IncludeCentralAdministration
    foreach ($webApplication in $webApplications)
    {
      $sprAlternateURL = Get-SPAlternateURL -WebApplication $webApplication
      $sprAlternateURL | Add-Member -MemberType NoteProperty -Name WebApplicationDisplayName -Value $webApplication.DisplayName
      $sprAlternateURL | Add-Member -MemberType NoteProperty -Name WebApplicationUrl -Value $webApplication.Url
      $object += $sprAlternateURL
    }

    $object | Export-Clixml -Path $args[0]
  } -ArgumentList $file
}
