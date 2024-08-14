$fromVMname is the VM that I just have cloned
$newVMname is the cloned VM

$NewMACAddr = Get-NetworkAdapter $fromVMname
# get the .NET view object of the VM
$viewTargetVM = 
Get-View -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"Name" = "^${newVMname}$"}
$deviceNIC = $viewTargetVM.Config.Hardware.Device | Where-Object {$_ -is [VMware.Vim.VirtualEthernetCard]}
$cardnumber = $NewMACAddr.MacAddress.Count
for ($i=1; $i -le $cardnumber; $i++) {
$j=$i-1
# get the NIC device (further operations assume that this VM has only one NIC)
$deviceNIC[$j].MacAddress = $NewMACAddr[$j].MacAddress
#set the MAC address type to manual
$deviceNIC[$j].addressType = "Manual"
# create the new VMConfigSpec
$specNewVMConfig = New-Object VMware.Vim.VirtualMachineConfigSpec -Property @{
# setup the deviceChange object
deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec -Property @{
# the kind of operation, from the given enumeration
operation = "edit"
# the device to change, with the desired settings
device = $deviceNIC[$j]
} # end New-Object
} # end New-Object
# Reconfiguration de la VM clone pour prendre en compte les nouveaux paramètres
$viewTargetVM.ReconfigVM($specNewVMConfig)
}