function Get-SPRSearchIndexPartitions
{
  param
  (
    [Parameter(Mandatory = $true)]
    [object[]]$IndexPartitions
  )
	
  foreach ($IndexPartition in $IndexPartitions)
  {
  Add-PSSnapin Microsoft.sharepoint.powershell
  $IndexPartition=Get-SPEnterpriseSearchServiceInstance -Local
  $searchapp= Get-SPEnterpriseSearchServiceApplication -Identity 'Search Service Application'
  $Topologies = Get-SPEnterpriseSearchTopology -SearchApplication $searchapp -active


    $properties = [ordered]@{
      'Name'             = $IndexPartition.Components.name
      'Topology'         = $IndexPartition.Topologies
      'ServiceApplication' = $IndexPartition.TypeName
      'Server'           = $IndexPartition.Server
      'RootDirectory'    = $IndexPartition.RootDirectory
      'IndexLocation'    = $IndexPartition.DefaultIndexLocation
      'Role'             = $IndexPartition.Role
      'Service'          = $IndexPartition.Service
      'Status'           = $IndexPartition.Status
    }
    $output = New-Object -TypeName PSObject -Property $properties
		
    Write-Output -InputObject $output
  }
}