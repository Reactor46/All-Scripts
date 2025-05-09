# Start of Settings 
# End of Settings 

Get-ADReplicationConnection -Filter * | Select Name, AutoGenerated, @{Name="From";e={($_.ReplicateFromDirectoryServer -split ",")[1].trim("CN=")}}, @{Name="To";e={($_.ReplicateToDirectoryServer -split ",")[0].trim("CN=")}}

$Title = "Replication Connection Information"
$Author = "John Sneddon"
$PluginVersion = 1.0
$Header = "Replication Connection Information"
$Comments = "List all replication connections"
$Display = "Table"
$PluginCategory = "ActiveDirectory"