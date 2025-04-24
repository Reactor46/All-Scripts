$server = Get-SPServer FBV-SPAPP-T03
$server.ServiceInstances | where{$_.TypeName -like "*SharePoint Foundation Administration*"}
$serviceinstance = $server.ServiceInstances | where{$_.TypeName -like "*SharePoint Foundation Administration*"}
$serviceinstance.Provision()
$server.ServiceInstances | where{$_.TypeName -like "*SharePoint Foundation Administration*"}