function Export-SPRWebServicesAdministration
{
  param(
    [string]$Path
  )

  $file = '{0}\SPWebServicesAdministration.xml' -f $Path
  Start-Job -Name 'Web Services Administration' -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell
    $Object = [Microsoft.SharePoint.Administration.SPWebService]::AdministrationService.Features

    $object | Export-Clixml -Path $args[0]
    
  } -ArgumentList $file
}
