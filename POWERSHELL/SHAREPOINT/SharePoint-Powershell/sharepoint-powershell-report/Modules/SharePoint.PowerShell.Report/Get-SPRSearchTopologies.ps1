function Get-SPRSearchTopologies
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$SPREnterpriseSearchTopologyConfiguration
  )


  foreach ($topology in $SPREnterpriseSearchTopologyConfiguration)
  {
    $properties = [ordered]@{
      'ServiceApplication' = $topology.SearchServiceApplicationName
      'Topology'         = $topology.TopologyId
      'CreationDate'     = $topology.CreationDate
      'State'            = $topology.State
      'ComponentCount'   = $topology.ComponentCount
      'IsActive'         = $topology.State
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}
