function Export-SPREnterpriseSearchTopologyConfiguration 
{
  param(
    [string]$Path
  )

  $file = '{0}\SPREnterpriseSearchTopologyConfiguration.xml' -f $Path
  Start-Job -ScriptBlock {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell
    $object = @()
    $searchServiceApplications = Get-SPEnterpriseSearchServiceApplication
    foreach ($searchServiceApplicateion in $searchServiceApplications)
    {
      $topologies = $searchServiceApplicateion | Get-SPEnterpriseSearchTopology
      foreach ($topology in $topologies)
      {
        $topology | Add-Member -MemberType NoteProperty -Name 'SearchServiceApplicationName' -Value $searchServiceApplicateion.Name
          
        $object += $topology
      }
    }
        
    $object | Export-Clixml -Path $args[0]
  } -ArgumentList $file
}
