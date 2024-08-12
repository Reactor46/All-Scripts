Write-Output "Renaming computer..."
Rename-Computer -NewName "CORE"

Write-Output "Enabling RDP..."
set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server'-name "fDenyTSConnections" -Value 0
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -name "UserAuthentication" -Value 1   

Write-Output "Setting time zone..."
& "$env:windir\system32\tzutil.exe" /s "Eastern Standard Time"

Write-Output "Activating with SPLA licensing..."
slmgr -ipk T7NKP-FKBPW-V9X28-2W49B-QJY48

Write-Output "Turning off IESC..."
$AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
$UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
Set-ItemProperty -Path $AdminKey -Name “IsInstalled” -Value 0
Set-ItemProperty -Path $UserKey -Name “IsInstalled” -Value 0

Write-Output "Enabling automatic updates..."
$AUSettings = (New-Object -com "Microsoft.Update.AutoUpdate").Settings
$AUSettings.NotificationLevel = 4
$AUSettings.Save()

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
New-VM -Name "Domain" -MemoryStartupBytes 2GB -NewVHDPath C:\Hyper-V\VHDs\C-Domain.vhdx -NewVHDSizeBytes 250GB
New-VM -Name "Data" -MemoryStartupBytes 2GB -NewVHDPath C:\Hyper-V\VHDs\C-Data.vhdx -NewVHDSizeBytes 250GB
New-VM -Name "Database" -MemoryStartupBytes 4GB -NewVHDPath C:\Hyper-V\VHDs\C-Database.vhdx -NewVHDSizeBytes 250GB
New-VM -Name "Remote" -MemoryStartupBytes 2GB -NewVHDPath C:\Hyper-V\VHDs\C-Remote.vhdx -NewVHDSizeBytes 250GB
New-VM -Name "AVG" -MemoryStartupBytes 4GB -NewVHDPath C:\Hyper-V\VHDs\C-AVG.vhdx -NewVHDSizeBytes 250GB

Write-Output "Setting VMs CPU count..."
Get-VM | Set-VMProcessor -Count 2

Write-Output "Adding network adapters to VMs..."
Get-VM | Add-VMNetworkAdapter -SwitchName "External"

Write-Output "Enabling remoting..."
Enable-PSRemoting -Force

Write-Output "Enabling Guest Integration Services..."
ForEach ($vm in Get-VM) { Enable-VMIntegrationService -Name "Guest Service Interface" -VMName $vm.Name }

Restart-Computer