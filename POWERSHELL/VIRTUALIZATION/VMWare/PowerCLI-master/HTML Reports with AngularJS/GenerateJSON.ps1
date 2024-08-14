#Variable declaration
#Connect-VIServer -Server "192.168.1.200" -User "administrator@vsphere.local" -Password "J@bb3rJ4w"
$vCenterIPorFQDN="192.168.1.200"
$vCenterUsername="Administrator@vsphere.local"
$vCenterPassword="J@bb3rJ4w"
$ClusterName="SuperNAP_Rack1" #Name of the cluster from which you want to retrieve VM infos
#Location where you want to place generated JSON Files.
#Please be aware that you should place them in the "data" folder in order to make WebPowerCLI read data from them
#$OutputPath="C:\Users\Paolo\Desktop\data" 
#$OutputPath="C:\LazyWinAdmin\VMWare\VMware Reports"
$OutputPath="C:\LazyWinAdmin\VMWare\PowerCLI-master\HTML Reports with AngularJS\data"

Write-Host "Depending on how many VMs you have in your cluster this script could take a while...please be patient" -foregroundcolor "magenta" 

Write-Host "Connecting to vCenter" -foregroundcolor "magenta" 
Connect-VIServer -Server $vCenterIPorFQDN -User $vCenterUsername -Password $vCenterPassword

$vms = Get-Cluster -Name $ClusterName | Get-VM

$vms | ConvertTo-Json -Depth 1 > $OutputPath\vms_all.json

foreach($vm in $vms){

Write-Host "Generating JSON file for VM:" $vm -foregroundcolor "magenta" 

Get-VM -Name $vm | Select * -ExcludeProperty ExtensionData | ConvertTo-Json -Depth 1 > $OutputPath\$($vm.Id).json

}

Write-Host "Disconnecting from vCenter" -foregroundcolor "magenta" 
Disconnect-VIServer * -Confirm:$false