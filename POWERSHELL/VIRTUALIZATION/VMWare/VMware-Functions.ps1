$vms = Get-VM | where {$_.PowerState -eq "PoweredOff"}
$vmPoweredOff = $vms | %{$_.Name}
$events = Get-VIEvent -Start (Get-Date).AddDays(-30) -Entity $vms | where{$_.FullFormattedMessage -like "*is powered off"}
$lastMonthVM = $events | %{$_.Vm.Name}
$vmPoweredOff | where {!($lastMonthVM -contains $_)}

# Get a VM's last power on date based on the VM's events.
# Requires PowerCLI 4.0 and PowerShell v2.

function Get-LastPowerOn { 

param(
[Parameter(Mandatory=$true,
ValueFromPipeline=$true,
HelpMessage="VM")
]

[

VMware.VimAutomation.Types.VirtualMachine]
$VM
)
 

Process {

# Patterns that indicate an attempt to power a VM on. This differ
# across versions and likely across language. Please add your own
# if you find one missing.

$patterns = @(
"*Power On virtual machine*", # vCenter 4 English
"*is starting*" # ESX 4/3.5 English
)


$events = $VM | Get-VIEvent
$qualifiedEvents = @()

foreach ($pattern in $patterns) {
    $qualifiedEvents += $events | Where { $_.FullFormattedMessage -like $pattern }

}

 

$qualifiedEvents = $qualifiedEvents | Where { $_ -ne $null }

 

$sortedEvents = Sort-Object -InputObject $qualifiedEvents -Property CreatedTime -Descending

 

$event = $sortedEvents | select -First 1

 

$obj = New-Object PSObject

 

$obj | Add-Member -MemberType NoteProperty -Name VM -Value $_

 

$obj | Add-Member -MemberType NoteProperty -Name PowerState -Value $_.PowerState

 

$obj | Add-Member -MemberType NoteProperty -Name LastPoweron -Value $null

 

if ($event) {

 

$obj.LastPoweron = $event.CreatedTime

}

 

Write-Output $obj

}

}

#$Vms | Get-LastPowerOn | Export-Csv LastPowerOn.csv -NoTypeInformation -UseCulture