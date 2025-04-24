function Export-SPRWebApplicationConfiguration 
{
  param(
    [string]$Path
  )

  $file = '{0}\SPWebApplication.xml' -f $Path
  Start-Job -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell
    $object = Get-SPWebApplication -IncludeCentralAdministration
    
    foreach ($webApplication in $object)
    {
      $superUserHasFullControl = $null
      $superReaderHasFullRead = $null

      $sprWebApplication = $webApplication
      $superUserAccount = $webApplication.Properties['portalsuperuseraccount'] 
      $superUserHasFullControl = (($webApplication.Policies | Where-Object -FilterScript {
            $_.UserName -like "*$superUserAccount"
      } ) | Select-Object -First 1 -ExpandProperty PolicyRoleBindings ).Name -eq 'Full Control' 
      $sprWebApplication | Add-Member -MemberType NoteProperty -Name SuperUserHasFullControl -Value $superUserHasFullControl
      $superReaderAccount = $webApplication.Properties['portalsuperreaderaccount'] 
      $superReaderHasFullRead = (($webApplication.Policies | Where-Object -FilterScript {
            $_.UserName -like "*$superReaderAccount"
      } ) | Select-Object -First 1 -ExpandProperty PolicyRoleBindings ).Name -eq 'Full Read' 
      $sprWebApplication | Add-Member -MemberType NoteProperty -Name SuperReaderHasFullRead -Value $superReaderHasFullRead
    }
     
    $object | Export-Clixml -Path $args[0]
  } -ArgumentList $file
}
