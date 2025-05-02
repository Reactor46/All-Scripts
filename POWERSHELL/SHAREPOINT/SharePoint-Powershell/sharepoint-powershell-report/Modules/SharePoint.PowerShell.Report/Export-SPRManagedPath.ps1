function Export-SPRManagedPath
{
  param
  (
    [String]
    [Parameter(Mandatory)]
    $Path
  )
  
      $file = '{0}\SPManagedPath.xml' -f $Path
        Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $webApplications = Get-SPWebApplication -IncludeCentralAdministration
        
        $SPRManagedPath = @()
        
        foreach ($webApplication in $webApplications)
        {
          $managedPath = Get-SPManagedPath -WebApplication $webApplication
          $managedPath | Add-Member -MemberType NoteProperty 'WebApplicationUrl' -Value $webApplication.Url
          $SPRManagedPath += $managedPath
        }
        
        $SPRManagedPath | Export-Clixml -Path $args[0]
    } -ArgumentList $file
  
  

}