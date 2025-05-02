

Import-Module -Name VMWare.PowerCLI
Connect-VIServer -Server LASVCENTER01 -User %username% -SaveCredentials


Stop-Computer -ComputerName LASAUTHTST01 -Force -Confirm:$false
Get-VM LASAUTHTST01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASAUTHTST01 -Confirm:$false

Stop-Computer -ComputerName LASCODETST01 -Force -Confirm:$false
Get-VM LASCODETST01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:0c:29:f0:d4:ce' -WakeOnLan:$true
Start-VM -Vm LASCODETST01 -Confirm:$false

Stop-Computer -ComputerName LASDEVTOOLS01 -Force -Confirm:$false
Get-VM LASDEVTOOLS01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASDEVTOOLS01 -Confirm:$false

Stop-Computer -ComputerName LASITSTST01 -Force -Confirm:$false
Get-VM LASITSTST01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -VM LASITSTST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASJIRATST01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASMDT01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASPROCAPPTST02 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASPROCWEB01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASPSHOST01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASSECSRV01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASSFSUPG01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASSTGDEV01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASTRITON01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASTSTINTRA04 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASWDS01 | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASJIRATST01 -Force -Confirm:$false
Get-VM LASPROCWEB02 | Get-NetworkAdapter -Name 'Network adapter 1' | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASJIRATST01 -Confirm:$false

Stop-Computer -ComputerName LASSQL02N01 -Force -Confirm:$false
Get-VM LASSQL02N01 | Get-NetworkAdapter -Name 'Network adapter 1' | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Get-VM LASSQL02N01 | Get-NetworkAdapter -Name 'Network adapter 2' | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASSQL02N01 -Confirm:$false

Stop-Computer -ComputerName LASPROCAPP01 -Force -Confirm:$false
Get-VM LASPROCAPP01 | Get-NetworkAdapter -Name 'Network adapter 1' | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Get-VM LASPROCAPP01 | Get-NetworkAdapter -Name 'Network adapter 2' | Set-NetworkAdapter -Type Vmxnet3 -MacAddress '00:50:56:aa:03:1a' -WakeOnLan:$true
Start-VM -Vm LASPROCAPP01 -Confirm:$false

Get-Content $PSScriptRoot\Check-Comp.txt | 
 ForEach { if (test-connection $_ -Count 3 -Delay 10 -quiet) { write-output "$_" |
  Out-File $PSScriptRoot\Alive.log -append
  } else { 
  write-output "$_" |
   Out-File $PSScriptRoot\Dead.log -append}}

Remove-Module -Name VMware* -Force