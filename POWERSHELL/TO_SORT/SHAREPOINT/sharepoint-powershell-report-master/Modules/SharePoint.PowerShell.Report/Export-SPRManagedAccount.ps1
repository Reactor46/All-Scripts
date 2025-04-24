function Export-SPRManagedAccount
{
  param(
    [string]$Path
  )

  $file = '{0}\SPManagedAccount.xml' -f $Path
  Start-Job -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

    $object = Get-SPManagedAccount
    $object | Export-Clixml -Path $args[0]
  } -ArgumentList $file
}
