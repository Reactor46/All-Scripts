#  Script name:		SYSVOL Replication Fix
#  Created on:		07-21-2016
#  Author:        	Jack Musick
#  Purpose:       	Fixes SYSVOL replication. Note, after running this on each DC, run Start-Service NTFRS.

Start-Service "NTFRS"

$registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\NtFrs\Parameters\Backup/Restore\Process at Startup"
$name = "BurFlags"

$dcTypes = "Primary","Secondary"
Write-Host "DC Types:`n`n"

$dcIndex = 0

ForEach($type in $dcTypes) { Write-Host "$dcIndex : $type"; $dcIndex++ }
Write-Host ""

# Getting valid input
do {
    try {
        $validNumber = $true
        [int]$selectedIndex = Read-Host "Select an Index"
    }
    catch { ($validNumber = $false) }
}
until (($selectedIndex -le ($dcTypes.Count - 1)) -and $validNumber)

if([$selectedIndex]$dcTypes -eq "Primary") {
    $value = 212
}
else {
    $value = 210
}

New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force