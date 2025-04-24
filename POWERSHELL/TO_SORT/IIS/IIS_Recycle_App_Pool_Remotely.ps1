## PowerShell: Script to Stop / Start / Recycle an IIS Application Pool Remotely

$pc = "fbv-uipath-d01" ##Provide your machine name here
 
## 1. List the app pools, note the __RELPATH of the one you want to target: 
Get-WMIObject IISApplicationPool -ComputerName $pc -Namespace root\MicrosoftIISv2 -Authentication PacketPrivacy  
 
## 2. Target a specific Application Pool: 
$Name = "W3SVC/APPPOOLS/AppPoolName"  ##This is the __RELPATH property for the IIsApplicationPool.Name determined from the query above
$Path = "IISApplicationPool.Name='$Name'" ##This is the __RELPATH 

## 3. Invoke the command against the remote Aplication Pool. The Invoke-WMIMethod accepts the following commands: Stop / Start / Recycle
#Invoke-WMIMethod Recycle -Path $Path -ComputerName $pc -Namespace root\MicrosoftIISv2 -Authentication PacketPrivacy