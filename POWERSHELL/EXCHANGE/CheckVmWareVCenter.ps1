<#

.SYNOPSIS
Check all VMs in the remote VMWare vCenter and list those might now have WannaCry ransomware fix

.DESCRIPTION
This script connect to vCenter, finds all Windows machines that are powerd on and gets the last hotfix update date.
The machine is reported as "under risk" in case if the latest update is older than 15.03.2017 - which means, it definitely
has no WannaCry exploit patch installed.
The final list of such VMs is saved to the file MACHINES_UNDER_RISK.txt in the current folder. 

.PARAMETER ViServer
Remote ViServer to connect to

.PARAMETER Credential 
Credentials for connecting to ViServer and accessing VMs

#>

param(
    [string]$ViServer,
    $Credential = (Get-Credential)
)



Get-Module -ListAvailable VMWare* | Import-Module

Write-Host "Connecting to vSphere"
$server = Connect-VIServer -Server $ViServer -Protocol https -Credential $Credential | Out-Null

$patchDate = Get-Date -Date "2017-03-15T00:00:00.0000000" -F o 

Write-Host "Getting Windows powered on machines"
$allVms = Get-Vm | where { $_.PowerState -eq "PoweredOn" -and $_.ExtensionData.Guest.GuestFullName -like "*windows*" }

$VmsUnderRisk = @()

foreach ($vm in $allVms) {


    $command = "Get-Date -Date (Get-HotFix | sort InstalledOn -desc)[0].InstalledOn -Format o"

    Write-Host "Checking machine: $($vm)"

    $updateDate = Invoke-VMScript -VM $vm -GuestCredential $Credential -ScriptType Powershell -ScriptText $command -Server $server

    if (-Not $updateDate) {
        Write-Host "No info on hotfixes, considering just installed or unavailable for scan"
        $updateDate = Get-Date -Format o
    }

    $updateDate = Get-Date -Date $updateDate -Format o

    if ($updateDate -lt $currentDate)
    {
        Write-Host "Last update date is: $($updateDate)"
        Write-Host "$($VM) under risk!!! :O"
        $VmsUnderRisk += @($Vm.Name, $updateDate)
        write-host $VmsUnderRisk
    }
    else {
            Write-Host "$($VM) is ok!"
    }

}

$VmsUnderRisk | Out-file -FilePath "MACHINES_UNDER_RISK.txt"
