$instanceName ="SPDistributedCacheService Name=AppFabricCachingService"
$serviceInstance = Get-SPServiceInstance | ? {($_.service.tostring()) -eq $instanceName -and ($_.server.name) -eq "FBV-SPWFE-T05"}
$serviceInstance.Unprovision()
$serviceInstance.Delete()
$serviceInstance = Get-SPServiceInstance | ? {($_.service.tostring()) -eq $instanceName -and ($_.server.name) -eq "FBV-SPAPP-T03"}
$serviceInstance.Unprovision()
$serviceInstance.Delete()

Unregister-CacheHost -HostName "FBV-SPWFE-T05" -provider "System.Data.SqlClient" -ConnectionString "Data Source=SPUAT;Initial Catalog=SP2013_UAT_CentralAdmin_Config;Integrated Security=True;Enlist=False"

Unregister-CacheHost -HostName "FBV-SPAPP-T03" -provider "System.Data.SqlClient" -ConnectionString "Data Source=SPUAT;Initial Catalog=SP2013_UAT_CentralAdmin_Config;Integrated Security=True;Enlist=False"


Unregister-CacheHost -HostName "FBV-SPAPP-T03.KSNET.COM" -ProviderType SPDistributedCacheClusterProvider -ConnectionString "\\FBV-SPAPP-T03.KSNET.COM"
Unregister-CacheHost -HostName "FBV-SPWFE-T05.KSNET.COM" -ProviderType SPDistributedCacheClusterProvider -ConnectionString "\\FBV-SPWFE-T05.KSNET.COM"