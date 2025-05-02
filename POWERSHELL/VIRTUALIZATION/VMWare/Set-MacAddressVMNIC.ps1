## script function:  set the MAC address of a VM's NIC
## Author: vNugglets.com
 
## the name of the VM whose NIC's MAC address to change
$strTargetVMName = "myVM01"
## the MAC address to use
$strNewMACAddr = "00:50:56:90:00:01"
 
## get the .NET view object of the VM
$viewTargetVM = Get-View -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"Name" = "^${strTargetVMName}$"}
## get the NIC device (further operations assume that this VM has only one NIC)
$deviceNIC = $viewTargetVM.Config.Hardware.Device | Where-Object {$_ -is [VMware.Vim.VirtualEthernetCard]}
## set the MAC address to the specified value
$deviceNIC.MacAddress = $strNewMACAddr
## set the MAC address type to manual
$deviceNIC.addressType = "Manual"
 
## create the new VMConfigSpec
$specNewVMConfig = New-Object VMware.Vim.VirtualMachineConfigSpec -Property @{
   ## setup the deviceChange object
   deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec -Property @{
       ## the kind of operation, from the given enumeration
       operation = "edit"
       ## the device to change, with the desired settings
       device = $deviceNIC
   } ## end New-Object
} ## end New-Object
 
## reconfigure the "clone" VM
$viewTargetVM.ReconfigVM($specNewVMConfig)