function Export-SPRUserProfileServiceApplication
{
  param(
    [string]$Path
  )

  $file = '{0}\SPServiceApplicationUserProfiles.xml' -f $Path
  Start-Job -Name '' -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell
    $userProfileServiceApplications = Get-SPServiceApplication | Where-Object {$_.TypeName -like "*User Profile Service Application*"}
    
    $site = Get-SPWebApplication -IncludeCentralAdministration | Get-SPSite -Limit 1
    $context = Get-SPServiceContext $site; 
    $audienceManager = New-Object Microsoft.Office.Server.Audience.AudienceManager($context)

    $object = $()
    foreach ($userProfileServiceApplication in $userProfileServiceApplications) {
        $userProfileServiceApplication | Add-Member -MemberType NoteProperty -Name 'AudienceManager' -Value $audienceManager
        $object += $userProfileServiceApplication
    }
    $object | Export-Clixml -Path $args[0]
  } -ArgumentList $file
}
