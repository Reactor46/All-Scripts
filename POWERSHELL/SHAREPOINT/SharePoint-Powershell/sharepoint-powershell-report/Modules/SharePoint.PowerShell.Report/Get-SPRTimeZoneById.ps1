function Get-SPRTimeZoneById 
{
  param(
    [int]$Id
  )

  [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SharePoint')
  $output = [Microsoft.SharePoint.SPregionalSettings]::GlobalTimeZones | Where-Object -FilterScript {
    $_.ID -eq $Id
  } | Select-Object -ExpandProperty Description
  if($output -eq $null)
  {
    $output = '(none)'
  }

  Write-Output -InputObject $output
}
