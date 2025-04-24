param (  
	[int] $RAM = 8
) 
$factor = $RAM/4

# DedicatedFilterProcessMemoryQuota 
$DedicatedFilterProcessMemoryQuota = 104857600

$val = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office Server\14.0\Search\Global\Gathering Manager\" -Name DedicatedFilterProcessMemoryQuota
Write-Host -ForeGroundColor Yellow "Current DedicatedFilterProcessMemoryQuota: " + $val.DedicatedFilterProcessMemoryQuota

$newVal = $DedicatedFilterProcessMemoryQuota * $factor
Write-Host -ForeGroundColor Green "New DedicatedFilterProcessMemoryQuota: " + $newVal
Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office Server\14.0\Search\Global\Gathering Manager\" -Name DedicatedFilterProcessMemoryQuota -Value $newVal


# FilterProcessMemoryQuota 
$FilterProcessMemoryQuota = 104857600

$val = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office Server\14.0\Search\Global\Gathering Manager\" -Name FilterProcessMemoryQuota
Write-Host -ForeGroundColor Yellow "Current FilterProcessMemoryQuota: " + $val.FilterProcessMemoryQuota

$newVal = $FilterProcessMemoryQuota * $factor
Write-Host -ForeGroundColor Green "New FilterProcessMemoryQuota: " + $newVal
Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office Server\14.0\Search\Global\Gathering Manager\" -Name FilterProcessMemoryQuota -Value $newVal

Write-Host -ForegroundColor Red "You need to reboot the server"
Restart-Computer -Confirm 
