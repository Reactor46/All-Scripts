function Export-SPRWebServices
{
  param(
    [string]$Path
  )

  $file = '{0}\SPWebServices.xml' -f $Path
  Start-Job -Name 'Web Services' -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell
    $Object = [Microsoft.SharePoint.Administration.SPWebService]::ContentService

    $object | Export-Clixml -Path $args[0]
    
  } -ArgumentList $file
}
