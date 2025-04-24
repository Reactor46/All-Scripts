function Export-SPRFarmAdmins
{
  param(
    [string]$Path
  )

  $file = '{0}\SPFarmAdmins.xml' -f $Path
  Start-Job -Name 'Farm Admins' -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell  
    $adminwebapp = Get-SPwebapplication -includecentraladministration | where {$_.IsAdministrationWebApplication}
    $adminsite = Get-SPweb($adminwebapp.Url)
    $AdminGroupName = $adminsite.AssociatedOwnerGroup
    $farmAdministratorsGroup = $adminsite.SiteGroups[$AdminGroupName]
    $object = $farmAdministratorsGroup.users
    $object | Export-Clixml -Path $args[0]
    
  } -ArgumentList $file
}
