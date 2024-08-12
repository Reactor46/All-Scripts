$path = Read-Host "Where would you like your Hyper-V setup to reside? (ex: C:\Hyper-V -- no slash!)"

Write-Output "Creating Hyper-V Directories..."
mkdir C:\Hyper-V
mkdir C:\Hyper-V\VMs
mkdir C:\Hyper-V\VHDs

Write-Output "Installing Hyper-V roles and features..."
Install-WindowsFeature Hyper-V,Hyper-V-Tools,Hyper-V-PowerShell

Write-Output "Setting up adapter(s)..."
$adapters = Get-NetAdapter | Where Status -eq "Up"
ForEach ($adapter in $adapters) { New-VMSwitch -Name "External" -NetAdapterName $adapter.Name -AllowManagementOS $true }

Write-Output "Creating baseline VMs..."
New-VM -Name "Domain" -MemoryStartupBytes 2GB -NewVHDPath $path\VHDs\C-Domain.vhdx -NewVHDSizeBytes 250GB
New-VM -Name "Data" -MemoryStartupBytes 2GB -NewVHDPath $path\VHDs\C-Data.vhdx -NewVHDSizeBytes 250GB
New-VM -Name "Database" -MemoryStartupBytes 4GB -NewVHDPath $path\VHDs\C-Database.vhdx -NewVHDSizeBytes 250GB
New-VM -Name "Remote" -MemoryStartupBytes 2GB -NewVHDPath $path\VHDs\C-Remote.vhdx -NewVHDSizeBytes 250GB
New-VM -Name "AVG" -MemoryStartupBytes 4GB -NewVHDPath $path\VHDs\C-AVG.vhdx -NewVHDSizeBytes 250GB

Write-Output "Setting VMs CPU count..."
Get-VM | Set-VMProcessor -Count 2

Write-Output "Adding network adapters to VMs..."
Get-VM | Add-VMNetworkAdapter -SwitchName "External"

Write-Output "Enabling remoting..."
Enable-PSRemoting -Force

Write-Output "Enabling Guest Integration Services..."
ForEach ($vm in Get-VM) { Enable-VMIntegrationService -Name "Guest Service Interface" -VMName $vm.Name }

Restart-Computer