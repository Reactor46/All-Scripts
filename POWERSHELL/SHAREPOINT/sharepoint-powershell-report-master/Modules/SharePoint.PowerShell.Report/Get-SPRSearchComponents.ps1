function Get-SPRSearchComponents
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SearchComponents
  )
	
  foreach ($SearchComponent in $SearchComponents)
  {
  Add-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

  $searchapp= Get-SPEnterpriseSearchServiceApplication -Identity 'e06695d8-4081-4fba-a64a-8a645d9fe2a6'
  $Topologies = Get-SPEnterpriseSearchTopology -SearchApplication $searchapp -active

    $properties = [ordered]@{
      'Name'             = $SearchComponent.Name
      'Topology'         = $SearchComponent.TopologyId
      'ServiceApplication' = $searchapp.Name
      'Server'           = $SearchComponent.ServerName
      'Type'             = $SearchComponent
      'IsActive'         = $Topologies.State
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}