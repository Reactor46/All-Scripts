$ComputerName = Read-Host "Please enter a computer name"

Write-Output "Renaming computer to $ComputerName..."
Rename-Computer -NewName $ComputerName

Write-Output "Enabling RDP..."
set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server'-name "fDenyTSConnections" -Value 0
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -name "UserAuthentication" -Value 1   

Write-Output "Setting time zone..."
& "$env:windir\system32\tzutil.exe" /s "Eastern Standard Time"

Write-Output "Activating with MAPs licensing..."
slmgr -ipk PNQYD-2839T-JGKDT-7GBFM-CWDWM

Write-Output "Turning off IESC..."
$AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
$UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
Set-ItemProperty -Path $AdminKey -Name “IsInstalled” -Value 0
Set-ItemProperty -Path $UserKey -Name “IsInstalled” -Value 0

Write-Output "Enabling automatic updates..."
$AUSettings = (New-Object -com "Microsoft.Update.AutoUpdate").Settings
$AUSettings.NotificationLevel = 4
$AUSettings.Save()

Write-Output "Enabling remoting..."
Enable-PSRemoting -Force

Write-Output "Joining to the domain..."
$DomainName = Read-Host "Please enter your domain name"
Add-Computer -DomainName $DomainName

Write-Output "Restarting computer..."
Restart-Computer