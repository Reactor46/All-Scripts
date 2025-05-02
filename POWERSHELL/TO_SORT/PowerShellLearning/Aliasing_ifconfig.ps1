#This command makes it so that 'ifconfig', the linux command for getting net Ip configuration, can be used in Powershell.
#ifconfig is easier to remember and does the same thing as Get-NetIPConfiguration
New-Alias -Name ifconfig -Value Get-NetIPConfiguration