
Import-Module -Name VMWare.PowerCLI
Connect-VIServer -Server LASVCENTER01 -User %username% -SaveCredentials

$Computers = GC -Path $PSScriptRoot\Nic-Change.txt

ForEach($comp in $Computers){
Stop-Computer -ComputerName $_ -Force -Confirm:$false |
Get-VM -Name $_ | Get-NetworkAdapter -Name 'Network adapter 1' | Set-NetworkAdapter -Type Vmxnet3 |
Start-VM -VM $_ -Confirm:$false
}

Stop-Computer -ComputerName LASSQL02N01 -Force -Confirm:$false
Get-VM LASSQL02N01 | Get-NetworkAdapter -Name 'Network adapter 1' | Set-NetworkAdapter -Type Vmxnet3
Get-VM LASSQL02N01 | Get-NetworkAdapter -Name 'Network adapter 2' | Set-NetworkAdapter -Type Vmxnet3
Start-VM -Vm LASSQL02N01 -Confirm:$false

Stop-Computer -ComputerName LASPROCAPP01 -Force -Confirm:$false
Get-VM LASPROCAPP01 | Get-NetworkAdapter -Name 'Network adapter 1' | Set-NetworkAdapter -Type Vmxnet3
Get-VM LASPROCAPP01 | Get-NetworkAdapter -Name 'Network adapter 2' | Set-NetworkAdapter -Type Vmxnet3
Start-VM -Vm LASPROCAPP01 -Confirm:$false

Get-Content $PSScriptRoot\Check-Comp.txt | 
 ForEach { if (test-connection $_ -Count 3 -Delay 10 -quiet) { write-output "$_" |
  Out-File $PSScriptRoot\Alive.log -append
  } else { 
  write-output "$_" |
   Out-File $PSScriptRoot\Dead.log -append}}

Remove-Module -Name VMware* -Force