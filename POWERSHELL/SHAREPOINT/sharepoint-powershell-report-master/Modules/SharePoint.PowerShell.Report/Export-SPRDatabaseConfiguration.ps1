function Export-SPRDatabaseConfiguration 
{
  param(
    [string]$Path
  )

  $file = '{0}\SPDatabase.xml' -f $Path
  Start-Job -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

    $object = Get-SPDatabase
    $object | Export-Clixml -Path $args[0]
  } -ArgumentList $file
}
