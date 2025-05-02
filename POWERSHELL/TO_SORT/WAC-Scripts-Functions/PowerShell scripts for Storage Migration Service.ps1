function Get-WACSMSCutoverSummary {
<#

.SYNOPSIS
Get Cutover Summary

.DESCRIPTION
Get Cutover Summary

.ROLE
Readers

#>
Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('CutoverSummary', $true)


$status=1
$exception = $null
try {
  $result = Get-SmsState @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSCutoverSummary ##
function Get-WACSMSInventoryConfigDetail {
<#

.SYNOPSIS
Get Inventory config detail

.DESCRIPTION
Get Inventory config detail

.ROLE
Readers

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName,

  [Parameter(Mandatory=$true)]
  [string]$computerName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('InventoryConfigDetail', $true)
$parameters.Add('ComputerName', $computerName)

$status=1
$exception = $null
try {
  $result = Get-SmsState @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSInventoryConfigDetail ##
function Get-WACSMSInventoryDFSNDetail {
<#

.SYNOPSIS
Get Inventory DFSN Detail

.DESCRIPTION
Get Inventory DFSN Detail

.ROLE
Readers

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName,
  [Parameter(Mandatory=$true)]
  [string]$computerName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{
  'ErrorAction' = 'Stop';
  'Name' = $jobName;
  'InventoryDFSNDetail' = $true;
  'ComputerName' = $computerName;
}

$status=1
$exception = $null
try {
  $result = Get-SmsState @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSInventoryDFSNDetail ##
function Get-WACSMSInventorySMBDetail {
<#

.SYNOPSIS
Get Inventory SMB Detail

.DESCRIPTION
Get Inventory SMB Detail

.ROLE
Readers

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName,
  [Parameter(Mandatory=$true)]
  [string]$computerName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{
  'ErrorAction' = 'Stop';
  'Name' = $jobName;
  'InventorySMBDetail' = $true;
  'ComputerName' = $computerName;
}

$status=1
$exception = $null
try {
  $result = Get-SmsState @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSInventorySMBDetail ##
function Get-WACSMSInventorySummary {
<#

.SYNOPSIS
Get Inventory Summary

.DESCRIPTION
Get Inventory Summary

.ROLE
Readers

#>
Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('InventorySummary', $true)


$status=1
$exception = $null
try {
  $result = Get-SmsState @parameters
  if($result -eq $null) {
    $result = "null" # ability to detect null is lost in the conversion to JSON
  }
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSInventorySummary ##
function Get-WACSMSSmbNetFirewallRule {
<#

.SYNOPSIS
Get SMB Net Firewall Rule

.DESCRIPTION
Returns the status of the SMB firewall rules
To enable: Get-NetFirewallRule -DisplayGroup 'File and Printer Sharing'|Set-NetFirewallRule -Profile 'Private, Domain' -Enabled true -PassThru

.ROLE
Readers

#>
Import-Module NetSecurity, Microsoft.PowerShell.Utility

 # Conversion to/from csv returns enums as text values, but does not retain Arrays or Complex Objects
 Get-NetFirewallRule -DisplayGroup 'File and Printer Sharing' | Microsoft.PowerShell.Utility\ConvertTo-Csv | Microsoft.PowerShell.Utility\ConvertFrom-Csv | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmbNetFirewallRule ##
function Get-WACSMSSmsCutover {
<#

.SYNOPSIS
Get Sms Cutover

.DESCRIPTION
Get Sms Cutover

.ROLE
Readers

#>
Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)


$status=1
$exception = $null
try {
  $result = Get-SmsCutover @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsCutover ##
function Get-WACSMSSmsCutoverPairing {
<#

.SYNOPSIS
Get SMS Cutover Pairing

.DESCRIPTION
Get SMS Cutover Pairing

.ROLE
Readers

#>
Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)

$status=1
$exception = $null
try {
  $result = Get-SmsCutoverPairing @parameters
  if($result -eq $null) {
    $result = "null" # ability to detect null is lost in the conversion to JSON
  }
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsCutoverPairing ##
function Get-WACSMSSmsDestinationConfig {
<#

.SYNOPSIS
Get Sms Destination config

.DESCRIPTION
Get Sms Destination config

.ROLE
Readers

#>
Param
(
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [Parameter(Mandatory = $true)]
    [string]$destinationComputerName,

    [Parameter(Mandatory = $false)]
    [string]$orchestratorComputerName,

    [Parameter(Mandatory = $false)]
    [int]$orchestratorPort
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('DestinationComputerName', $destinationComputerName)

if ($orchestatorComputerName) {
    $parameters.Add('OrchestratorComputerName', $orchestratorComputerName)
    $parameters.Add('OrchestratorPort', $orchestratorPort)
}

$status = 1
$exception = $null
try {
    $result = Get-SmsDestinationConfig @parameters
}
catch {
    $exception = $_ # the exception
    $status = 0
}
@{Result = $result; Status = $status; Error = $exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsDestinationConfig ##
function Get-WACSMSSmsFeature {
<#

.SYNOPSIS
Get Sms Feature

.DESCRIPTION
Get Sms Feature

.ROLE
Readers

#>

<#########################################################################################################
 # File: get-smsfeature.ps1
 #
 # .DESCRIPTION
 #
 #  invokes Get-WindowsFeature
 #
 #  Copyright (c) Microsoft Corp 2018.
 #
 #########################################################################################################>
Import-Module ServerManager

Get-WindowsFeature -Name 'SMS', 'SMS-PROXY'

}
## [END] Get-WACSMSSmsFeature ##
function Get-WACSMSSmsInventory {
<#

.SYNOPSIS
Get Sms Inventory

.DESCRIPTION
Get Sms Inventory

.ROLE
Readers

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)


$status=1
$exception = $null
try {
  $result = Get-SmsInventory @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsInventory ##
function Get-WACSMSSmsNasPrescan {
<#

.SYNOPSIS
Get Sms Nas Prescan

.DESCRIPTION
Get Sms Nas Prescan

.ROLE
Readers

#>
Param(
    [Parameter(Mandatory = $true)]
    [string]$jobName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{
    'Name' = $jobName;
}

$status=1
$exception = $null
try {
    $result = Get-SmsNasPrescan @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsNasPrescan ##
function Get-WACSMSSmsRequiredVolumeSize {
<#

.SYNOPSIS
Get Sms Required Volume Size

.DESCRIPTION
Get Sms Required Volume Size

.ROLE
Readers

#>
Param (
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [Parameter(Mandatory = $true)]
    [string]$srcComputerName,

    [Parameter(Mandatory = $true)]
    [array]$excludedShares,

    [Parameter(Mandatory = $true)]
    [bool]$anyExcludedShares
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$volumes = @{}
$volumeSizes = @{}
$volumesSorted = @{}
$shareSizes = @{}

# try {

    get-smsstate $jobName -ComputerName $srcComputerName -InventorySMBDetail -ErrorAction Ignore |
        foreach {
        if($_.Volume -eq $null){
          $shareVolume = $_.path.Substring(0,2)
        } else{
          $shareVolume = $_.Volume
        }
        $shareName = $_.Name
        $sharePath = $_.Path
        # echo "Doing $shareVolume - $shareName - $sharePath"
        if (!$volumes.containsKey($shareVolume)) {
            # echo "Did not contain $shareVolume"
            $sharePaths = @()
            if (!$anyExcludedShares -or !$excludedShares.Contains($sharePath)) {
                # echo "Not Excluded: $sharePath"
                $sharePaths += ($_.path.ToString())
                $volumes.Add($shareVolume, $sharePaths)
                $shareSizes.Add($_.path, $_.SizeTotal)
            }
            if (!$volumeSizes.containsKey($shareVolume)) {
                # echo "Adding to volumeSizes: $shareVolume"
                $volumeSizes.Add($shareVolume, 0)
            }
        }
        elseif (!$anyExcludedShares -or !$excludedShares.Contains($sharePath)) {
            # echo "Volume was in, adding share $shareVolume, $sharePath"
            # $sharePath = $volumes.$shareVolume
            $sharePaths = $volumes[$shareVolume]
            $sharePaths += ($_.path.ToString())
            $shareSizes.Add($_.path, $_.SizeTotal)
            $volumes.Set_Item($shareVolume, $sharePaths)
        }
    }

    $volumes.GetEnumerator() |
        foreach {
        $sortedSharePaths = $_.Value | Microsoft.PowerShell.Utility\Sort-Object
        $volumesSorted.Set_Item($_.Name, $sortedSharePaths)
    }

    $volumesSorted.GetEnumerator() |
        foreach {
        $currentVolume = $_.Name
        $currentShares = $_.Value
        if ($currentShares -is [array]) {
            ($currentShares.Count - 1)..0 |
                foreach {
                $currentShareName = $currentShares[$_].ToString()
                if ($shareSizes.ContainsKey($currentShareName)) {
                    if ($_ - 1 -ge 0) {
                        $contains = $true
                        $previousShareName =$currentShares[$_ - 1].ToString()
                        # if ($currentShareName -notcontains $previousShareNameToCompare) {
                        # echo "$previousShareName contains $currentShareName"

                        $path1Split = $previousShareName.split("\")
                        $path2Split = $currentShareName.split("\")

                        $path1Split = New-Object System.Collections.ArrayList(,$path1Split)
                        $path2Split = New-Object System.Collections.ArrayList(,$path2Split)
                        0..($path1Split.Count - 1) |
                            foreach {
                            $p1 = $path1Split[$_]
                            $p2 = $path2Split[$_]
                            # echo "$p1 equal $p2"
                            if(!($p1 -eq $p2)) {
                                $contains = 0
                            }
                        }
                        if($contains){
                            # echo "$currentShareName did contain $previousShareName"
                        } else {
                            # echo "$currentShareName did not contain $previousShareName"
                            $shareSize = $shareSizes.$currentShareName
                            $currentSize = $volumeSizes.$currentVolume
                            $volumeSizes.Set_Item($currentVolume, $currentSize + $shareSize)
                        }



                        # if((PathContains -path1 $currentShareName -path2 $previousShareName) -eq 1){
                        #     echo "$currentShareName did contain $previousShareName"
                        # } else {
                        #     echo "$currentShareName did not contain $previousShareName"
                        #     $shareSize = $shareSizes.$currentShareName
                        #     $currentSize = $volumeSizes.$currentVolume
                        #     $volumeSizes.Set_Item($currentVolume, $currentSize + $shareSize)
                        # }
                    }
                    else {
                        $shareSize = $shareSizes[$currentShareName]
                        $currentSize = $volumeSizes.$currentVolume
                        $volumeSizes.Set_Item($currentVolume, $currentSize + $shareSize)
                    }
                }
            }
        }
        else {
            if ($shareSizes.ContainsKey($currentShares)) {
                $shareSize = $shareSizes.$currentShares
                $currentSize = $volumeSizes.$currentVolume
                $volumeSizes.Set_Item($currentVolume, $currentSize + $shareSize)
            }
        }
    }
# }
# catch {

# }

$status = 1
@{Result = $volumeSizes; Status = $status; Error = $exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsRequiredVolumeSize ##
function Get-WACSMSSmsState {
<#

.SYNOPSIS
Get Sms State

.DESCRIPTION
Get Sms State

.ROLE
Readers

#>

Param
(
    [string]$jobName,
    [string]$nasController
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')

if ($jobName) {
    $parameters.Add('Name', $jobName)
    if ($nasController) {
        $parameters.Add('NasController', $nasController)
        $parameters.Add('GetNasPrescanResult', $true)
    }
}

$status = 1
$exception = $null
try {
    $result = Get-SmsState @parameters
}
catch {
    $exception = $_ # the exception
    $status = 0
}
@{Result = $result; Status = $status; Error = $exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsState ##
function Get-WACSMSSmsStats {
<#

.SYNOPSIS
Get Sms Stats

.DESCRIPTION
Get Sms Stats

.ROLE
Readers

#>

Param (
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

# Job list
$jobs = New-Object System.Collections.Generic.List[System.String]
get-smsstate | foreach { $jobs.Add($_.job) }

# Inventory
$inventoryDeviceCount = 0
$inventorySizeTotal = 0
$inventoryFilesTotal = 0
$inventoryJobsRunning = 0
$inventoryJobsPaused = 0
$inventoryJobsCompleted = 0


# Transfer
$transferDeviceCount = 0
$transferSizeTotal = 0
$transferFilesTotal = 0
$transferSizeTransferred = 0
$transferFilesTransferred = 0
$transferJobsRunning = 0
$transferJobsPaused = 0
$transferJobsCompleted = 0

# Cutover
$cutoverDeviceCount = 0
$cutoverJobsRunning = 0
$cutoverJobsPaused = 0
$cutoverJobsCompleted = 0

# JobSpecific
$runningJobStats = @()

# export enum SubOperationState {
#     NA,
#     NotStarted,
#     Running,
#     Succeeded,
#     Canceled,
#     Failed,
#     PartiallyFailed
# }
$SUBSTATE_NA = 0
$SUBSTATE_NOT_STARTED = 1
$SUBSTATE_RUNNING = 2
$SUBSTATE_SUCCEEDED = 3
$SUBSTATE_CANCELED = 4
$SUBSTATE_FAILED = 5
$SUBSTATE_PARTIALLYFAILED = 6
# export enum OperationType {
#     None = 0,
#     Inventory,
#     Transfer,
#     Cutover
# }
$OPERATION_NONE = 0
$OPERATION_INVENTORY = 1
$OPERATION_TRANSFER = 2
$OPERATION_CUTOVER = 3
# export enum OperationState {
#     None,
#     Idle,
#     Running,
#     Paused,
#     Canceled,
#     Failed,
#     Succeeded,
#     PartiallySucceeded
# }
$STATE_NONE = 0
$STATE_IDLE = 1
$STATE_RUNNING = 2
$STATE_PAUSED = 3
$STATE_CANCELED = 4
$STATE_FAILED = 5
$STATE_SUCCEEDED = 6
$STATE_PARTIALLYSUCCEEDED = 7

$NOT_RUNNING = 0
$INVENTORY_RUNNING = 1
$TRANSFER_RUNNING = 2
$CUTOVER_RUNNING = 3

# TODO: clean this up a bit. This feature was added last minute and went through 3 quick iterations so a bit messy
$debug = ''
try {
    foreach ($job in $jobs) {
        $jobIsRunning = $NOT_RUNNING

        #Determine if any phases running
        get-smsstate -name $job |
            foreach {
            if ($_.State.value__ -eq $STATE_RUNNING) {
                if ($_.LastOperation.value__ -eq $OPERATION_INVENTORY) {
                    $jobIsRunning = $INVENTORY_RUNNING
                }
                elseif ($_.LastOperation.value__ -eq $OPERATION_TRANSFER) {
                    $jobIsRunning = $TRANSFER_RUNNING
                }
                elseif ($_.LastOperation.value__ -eq $OPERATION_CUTOVER) {
                    $jobIsRunning = $CUTOVER_RUNNING
                }
            }
        }

        $tmpServerStats = @()
        $tmpJobStats = @{}
        # Inventory
        $tmpJobInventoryDevicesTotal = 0
        $tmpJobInventoryDevicesCompleted = 0
        get-smsstate -name $job -Inventorysummary -ErrorAction Ignore |
            foreach {
            if (!($_.InventoryState.value__ -eq $SUBSTATE_NA -or $_.InventoryState.value__ -eq $SUBSTATE_NOT_STARTED -or $_.InventoryState.value__ -eq $SUBSTATE_RUNNING)) {
                $tmpJobInventoryDevicesCompleted++;
            }
            if (!($_.InventoryState.value__ -eq $SUBSTATE_FAILED)) {
                $inventoryDeviceCount++
            }
            # if ($_.InventoryState.value__ -eq $SUBSTATE_RUNNING) {
            #     $inventoryJobsRunning++;
            # }
            if ($jobIsRunning -eq $INVENTORY_RUNNING) {
                $tmpStats = @{}
                $deviceTmp = @()
                $tmpStats.Add('Name', $_.SuppliedDeviceName)
                $tmpStats.Add('SubState', $_.InventoryState)

                # get-smsinventory -name $job -ErrorAction Ignore |
                #     foreach {
                #     $_.ComputerName | foreach {
                #         $deviceTmp += $_
                #     }
                # }
                # $tmp
                $tmpServerStats += $tmpStats
            }
            $inventorySizeTotal += $_.SizeTotal
            $inventoryFilesTotal += $_.FilesTotal

        }

        get-smsinventory -name $job -ErrorAction Ignore |
            foreach {
            $_.ComputerName | foreach {
                $tmpJobInventoryDevicesTotal++;
            }
        }

        # Transfer
        $tmpJobTransferSizeTotal = 0
        $tmpJobTransferSizeTransferred = 0
        get-smsstate -name $job -TransferSummary -ErrorAction Ignore |
            foreach {
            if (!($_.TransferState.value__ -eq $SUBSTATE_FAILED)) {
                $transferDeviceCount++
            }
            if ($jobIsRunning -eq $TRANSFER_RUNNING) {
                $tmpStats = @{}
                $deviceTmp = @()
                $tmpStats.Add('Name', $_.SourceDevice)
                $tmpStats.Add('SubState', $_.TransferState)
                $tmpStats.Add('Total', $_.SizeTotal)
                $tmpStats.Add('Complete', $_.SizeTransferred)
                $date = Get-Date
                $tmpStats.Add('Timestamp', $date)
                $tmpServerStats += $tmpStats
            }
            $transferSizeTotal += $_.SizeTotal
            $transferFilesTotal += $_.FilesTotal
            $transferSizeTransferred += $_.SizeTransferred
            $transferFilesTransferred += $_.FilesTransferred
            # $tmpJobTransferSizeTotal += $_.SizeTotal
            # $tmpJobTransferSizeTransferred += $_.SizeTransferred
        }

        # Cutover
        $tmpJobCutoverDevicesTotal = 0
        $tmpJobCutoverDevicesCompleted = 0
        get-smsstate -name $job -CutoverSummary -ErrorAction Ignore |
            foreach {
            $tmpJobCutoverDevicesTotal++
            if (!(($_.CutoverState.value__ -eq $SUBSTATE_NA) -Or ($_.CutoverState.value__ -eq $SUBSTATE_NOT_STARTED) -Or ($_.CutoverState.value__ -eq $SUBSTATE_RUNNING))) {
                $tmpJobCutoverDevicesCompleted++
            }
            if (!($_.CutoverState.value__ -eq $SUBSTATE_FAILED)) {
                $cutoverDeviceCount++
            }
            if ($jobIsRunning -eq $CUTOVER_RUNNING) {
                $tmpStats = @{}
                $deviceTmp = @()
                $tmpStats.Add('Name', $_.SourceDevice)
                $tmpStats.Add('SubState', $_.CutoverState)
                $tmpStats.Add('Total', 100)                  # Cutover has an internally calculated percentage out of 100%
                $tmpStats.Add('Complete', $_.CutoverProgress)
                $tmpServerStats += $tmpStats
            }
        }

        # Operation Totals
        $tmpJobName = ''
        get-smsstate -name $job |
            foreach {
            $tmpJobName = $_.job
            if ($_.LastOperation.value__ -eq $OPERATION_NONE) {
                # Placeholder
            }
            if ($_.LastOperation.value__ -eq $OPERATION_INVENTORY) {
                if ($_.State.value__ -eq $STATE_RUNNING) {
                    $tmpJobStats.Add('JobChartTotal', $tmpJobInventoryDevicesTotal)
                    $tmpJobStats.Add('JobChartCompleted', $tmpJobInventoryDevicesCompleted)
                    $tmpJobStats.Add('ChartType', 'inventoryChartDevices')
                    $tmpJobStats.Add('StatType', 'inventory')
                }
                switch ($_.State.value__) {
                    $STATE_RUNNING {$inventoryJobsRunning++}
                    $STATE_PAUSED {$inventoryJobsPaused++}
                    $STATE_PARTIALLYSUCCEEDED {$inventoryJobsCompleted++}
                    $STATE_SUCCEEDED {$inventoryJobsCompleted++}
                    $STATE_FAILED {$inventoryJobsCompleted++}
                }
            }
            elseif ($_.LastOperation.value__ -eq $OPERATION_TRANSFER) {
                $inventoryJobsCompleted++
                if ($_.State.value__ -eq $STATE_RUNNING) {
                    $tmpJobStats.Add('JobChartTotal', $tmpJobTransferSizeTotal)
                    $tmpJobStats.Add('JobChartCompleted', $tmpJobTransferSizeTransferred)
                    $tmpJobStats.Add('ChartType', 'transferChartSize')
                    $tmpJobStats.Add('StatType', 'transfer')
                }
                switch ($_.State.value__) {
                    $STATE_RUNNING {$transferJobsRunning++}
                    $STATE_PAUSED {$transferJobsPaused++}
                    $STATE_PARTIALLYSUCCEEDED {$transferJobsCompleted++}
                    $STATE_SUCCEEDED {$transferJobsCompleted++}
                    $STATE_FAILED {$transferJobsCompleted++}
                }
            }
            elseif ($_.LastOperation.value__ -eq $OPERATION_CUTOVER) {
                $inventoryJobsCompleted++
                $transferJobsCompleted++
                if ($_.State.value__ -eq $STATE_RUNNING) {
                    $tmpJobStats.Add('JobChartTotal', $tmpJobCutoverDevicesTotal)
                    $tmpJobStats.Add('JobChartCompleted', $tmpJobCutoverDevicesCompleted)
                    # $tmpJobStats.Add('CutoverDevicesRunning', $tmpJobCutoverDevicesRunning)
                    $tmpJobStats.Add('ChartType', 'cutoverChartDevices')
                    $tmpJobStats.Add('StatType', 'cutover')
                }
                switch ($_.State.value__) {
                    $STATE_RUNNING {$cutoverJobsRunning++}
                    $STATE_PAUSED {$cutoverJobsPaused++}
                    $STATE_PARTIALLYSUCCEEDED {$cutoverJobsCompleted++}
                    $STATE_SUCCEEDED {$cutoverJobsCompleted++}
                    $STATE_FAILED {$cutoverJobsCompleted++}
                }
            }
        }

        if ($tmpJobStats.Count -gt 0 -or !($jobIsRunning -eq $NOT_RUNNING)) {
            $tmpJobStats.Add('JobName', $tmpJobName)
            $tmpJobStats.Add('JobServerStats', $tmpServerStats)
            $runningJobStats += $tmpJobStats
        }
    }
}
catch {

}
$status = 1
$result = @{
    'InventoryDeviceCount'   = $inventoryDeviceCount;
    'InventorySizeTotal'     = $inventorySizeTotal;
    'InventoryFilesTotal'    = $inventoryFilesTotal;
    'InventoryJobsRunning'   = $inventoryJobsRunning;
    'InventoryJobsPaused'    = $inventoryJobsPaused;
    'InventoryJobsCompleted' = $inventoryJobsCompleted;
    'TransferDeviceCount'    = $transferDeviceCount;
    'TransferSizeTotal'      = $transferSizeTotal;
    'TransferFilesTotal'     = $transferFilesTotal;
    'TransferSizeTransferred'= $transferSizeTransferred;
    'TransferFilesTransferred' = $transferFilesTransferred;
    'TransferJobsRunning'    = $transferJobsRunning;
    'TransferJobsPaused'     = $transferJobsPaused;
    'TransferJobsCompleted'  = $transferJobsCompleted;
    'CutoverDeviceCount'     = $cutoverDeviceCount;
    'CutoverJobsRunning'     = $cutoverJobsRunning;
    'CutoverJobsPaused'      = $cutoverJobsPaused;
    'CutoverJobsCompleted'   = $cutoverJobsCompleted;
    'JobSpecificStats'       = $runningJobStats;
}

@{Result = $result; Status = $status; Error = $exception; Debug = $debug} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5


}
## [END] Get-WACSMSSmsStats ##
function Get-WACSMSSmsTransfer {
<#

.SYNOPSIS
Get Sms Transfer

.DESCRIPTION
Get Sms Transfer

.ROLE
Readers

#>
Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)


$status=1
$exception = $null
try {
  $result = Get-SmsTransfer @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsTransfer ##
function Get-WACSMSSmsTransferExcludedShares {
<#

.SYNOPSIS
Get Sms Transfer Exluded Shares

.DESCRIPTION
Get Sms Transfer Exluded Shares

.ROLE
Readers

#>
Param(
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [Parameter(Mandatory = $true)]
    [string]$computerName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{
    'Name'             = $jobName;
    'ComputerName'     = $computerName;
    'ExcludeSMBShares' = $true;
    'ErrorAction'      = 'Stop';
}

$status = 1
$exception = $null
try {
    $result = Get-SmsTransferPairing @parameters
}
catch {
    $exception = $_ # the exception
    $status = 0
}
@{Result = $result; Status = $status; Error = $exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsTransferExcludedShares ##
function Get-WACSMSSmsTransferExcludedSharesAndAFS {
<#

.SYNOPSIS
Get Sms Transfer Exluded Shares and AFS

.DESCRIPTION
Get Sms Transfer Exluded Shares and AFS

.ROLE
Readers

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [Parameter(Mandatory = $true)]
    [string]$computerName,

    [Parameter(Mandatory = $true)]
    [bool]$getAFSPairings
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$excludeParameters = @{
    'Name'             = $jobName;
    'ComputerName'     = $computerName;
    'ExcludeSMBShares' = $true;
    'ErrorAction'      = 'Stop';
}

$afsParameters = @{
    'Name'                    = $jobName;
    'ComputerName'            = $computerName;
    'TieredAFSVolumeSettings' = $true;
    'ErrorAction'             = 'Stop';
}

$status = 1
$exception = $null
$result = @{ }

try {
    function GetResultAsArray($outputCollection) {
        Import-Module Microsoft.PowerShell.Utility
        $resultArray = @()
        foreach ($item in $outputCollection) {
            $resultArray += $item
        }
        Write-Output -NoEnumerate $resultArray
    }

    $excludedShares = @()
    $excludedShares = Get-SmsTransferPairing @excludeParameters
    $excludedSharesAsArray = GetResultAsArray($excludedShares)
    $result += @{"ExcludeShares" = $excludedSharesAsArray }

    if($getAFSPairings){
        $afsPairings = @()
        $afsPairings = Get-SmsTransferPairing @afsParameters
        $afsPairingsAsArray = GetResultAsArray($afsPairings)
        $result += @{"AFSPairings" = $afsPairingsAsArray }
    }

}
catch {
    $exception = $_ # the exception
    $status = 0
}
@{Result = $result; Status = $status; Error = $exception } | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsTransferExcludedSharesAndAFS ##
function Get-WACSMSSmsTransferPairing {
<#

.SYNOPSIS
Get Sms Transfer Pairing

.DESCRIPTION
Get Sms Transfer Pairing

.ROLE
Readers

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)

$status=1
$exception = $null
try {
  $result = Get-SmsTransferPairing @parameters
  if($result -eq $null) {
    $result = "null" # ability to detect null is lost in the conversion to JSON
  }
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsTransferPairing ##
function Get-WACSMSSmsTransferVolumePairing {
<#

.SYNOPSIS
Get Sms Transfer Volume Pairing

.DESCRIPTION
Get Sms Transfer Volume Pairing

.ROLE
Readers

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName,

  [Parameter(Mandatory=$true)]
  [string]$computerName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('ComputerName', $computerName)
$parameters.Add('VolumePairings', $true)

$status=1
$exception = $null
try {
  $result = Get-SmsTransferPairing @parameters
  if($result -eq $null) {
    $result = "null" # ability to detect null is lost in the conversion to JSON
  }
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSSmsTransferVolumePairing ##
function Get-WACSMSSmsVersion {
<#

.SYNOPSIS
Get Sms Version

.DESCRIPTION
Get Sms Version

.ROLE
Readers

#>

Param
(
    [Parameter(Mandatory = $true)]
    [bool]$getSms,

    [Parameter(Mandatory = $true)]
    [bool]$getSmsPS,

    [Parameter(Mandatory = $true)]
    [bool]$getSmsProxy
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$includedTypes = @();

if ($getSms) {
    $includedTypes += 'Sms';
}

if ($getSmsPS) {
    $includedTypes += 'SmsPS';
}

if ($getSmsProxy) {
    $includedTypes += 'SmsProxy';
}


$status = 1;
$exception = $null;
try {
    $result = Get-SmsVersion -Type $includedTypes;
}
catch {
    $exception = $_; # the exception
    $status = 0;
}
@{Result = $result; Status = $status; Error = $exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5;

}
## [END] Get-WACSMSSmsVersion ##
function Get-WACSMSTemporaryFile {
<#

.SYNOPSIS
Get Temporary File

.DESCRIPTION
Get Temporary File

.ROLE
Readers

#>

Param
(
)
Import-Module Microsoft.PowerShell.Utility

$newTempFile = [System.IO.Path]::GetTempFileName() | Microsoft.PowerShell.Utility\ConvertTo-Json -Depth 5

Write-Output $newTempFile

}
## [END] Get-WACSMSTemporaryFile ##
function Get-WACSMSTransferDFSNDetail {
<#

.SYNOPSIS
Get Transfer DFSN Detail

.DESCRIPTION
Get Transfer DFSN Detail

.ROLE
Readers

#>

Param
(
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [Parameter(Mandatory = $true)]
    [string]$computerName,

    [Parameter(Mandatory = $false)]
    [bool]$pipeToFile,

    [Parameter(Mandatory = $false)]
    [string]$filename
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility


$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('TransferDFSNDetail', $true)
$parameters.Add('ComputerName', $computerName)


$status = 1
$exception = $null
try {
    if (!$pipeToFile) {
        $result = Get-SmsState @parameters
        if($result -eq $null) {
            $result = "null" # ability to detect null is lost in the conversion to JSON
          }
    }
    else {
      $result = Get-SmsState @parameters | Microsoft.PowerShell.Utility\ConvertTo-Csv | Out-File $filename
    }
}
catch {
    $exception = $_ # the exception
    $status = 0
}
@{Result = $result; Status = $status; Error = $exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSTransferDFSNDetail ##
function Get-WACSMSTransferFileDetail {
<#

.SYNOPSIS
Get Transfer File Detail

.DESCRIPTION
Get Transfer File Detail

.ROLE
Readers

#>

Param
(
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [Parameter(Mandatory = $true)]
    [string]$computerName,

    [Parameter(Mandatory = $false)]
    [bool]$pipeToFile,

    [Parameter(Mandatory = $false)]
    [string]$filename,

    [Parameter(Mandatory = $false)]
    [bool]$errorsOnly,

    [Parameter(Mandatory = $false)]
    [bool]$usersDownload,

    [Parameter(Mandatory = $false)]
    [bool]$groupsDownload
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('ComputerName', $computerName)

if ($usersDownload) {
    $parameters.Add('GetLocalUsersDetail', $true)
}
elseif ($groupsDownload) {
    $parameters.Add('GetLocalGroupsDetail', $true)
}
else {
    $parameters.Add('TransferFileDetail', $true)

    if ($errorsOnly) {
        $parameters.Add('ErrorsOnly', $true)
    }
}

$status = 1
$exception = $null
try {
    if (!$pipeToFile) {
        $result = Get-SmsState @parameters
    }
    else {
        # Select-Object -Skip 1 is to remove first line, as it is not CSV (powershell adds it)
        $result = Get-SmsState @parameters | Microsoft.PowerShell.Utility\ConvertTo-Csv |  Microsoft.PowerShell.Utility\Select-Object -Skip 1 | Microsoft.PowerShell.Utility\Out-File $filename
    }
}
catch {
    $exception = $_ # the exception
    $status = 0
}
@{Result = $result; Status = $status; Error = $exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSTransferFileDetail ##
function Get-WACSMSTransferSMBDetail {
<#

.SYNOPSIS
Get Transfer SMB Detail

.DESCRIPTION
Get Transfer SMB Detail

.ROLE
Readers

#>

Param
(
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [Parameter(Mandatory = $true)]
    [string]$computerName,

    [Parameter(Mandatory = $false)]
    [bool]$pipeToFile,

    [Parameter(Mandatory = $false)]
    [string]$filename
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('TransferSMBDetail', $true)
$parameters.Add('ComputerName', $computerName)


$status = 1
$exception = $null
try {
    if (!$pipeToFile) {
        $result = Get-SmsState @parameters
        if($result -eq $null) {
            $result = "null" # ability to detect null is lost in the conversion to JSON
          }
    }
    else {
      $result = Get-SmsState @parameters | Microsoft.PowerShell.Utility\ConvertTo-Csv | Out-File $filename
    }
}
catch {
    $exception = $_ # the exception
    $status = 0
}
@{Result = $result; Status = $status; Error = $exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSTransferSMBDetail ##
function Get-WACSMSTransferSummary {
<#

.SYNOPSIS
Get Transfer Summary

.DESCRIPTION
Get Transfer Summary

.ROLE
Readers

#>

Param
(
  [Parameter(Mandatory = $true)]
  [string]$jobName,

  [Parameter(Mandatory = $true)]
  [bool]$getAFSSummary
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{ }
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('TransferSummary', $true)
enum TransferState {
  Running = 2
}
$status = 1
$exception = $null
try {
  $timestamp = Get-Date
  $transferSummaries = Get-SmsState @parameters
  $transferSummaries | Microsoft.PowerShell.Utility\Add-Member -MemberType NoteProperty -Name Timestamp -Value $timestamp

  if ($getAFSSummary) {
    $sourceMachines = Get-SmsState -TransferSummary -Name $jobName
    $tieredAfsSummaries = @()
    foreach ($machine in $sourceMachines) {
      if ($machine.TransferSummary.value__ -eq [TransferState]::Running) {
        $tieredAfsSummaries += Get-SmsState -Name $machine.Job -computername $machine.SourceDevice -TransferVolumeDetail -VolumeTypes TAFSEnabled
      }
    }
    $transferSummaries | Microsoft.PowerShell.Utility\Add-Member -MemberType NoteProperty -Name TieredAFSVolumes -Value $tieredAfsSummaries
  }
  $result = $transferSummaries
}
catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result = $result; Status = $status; Error = $exception } | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSTransferSummary ##
function Get-WACSMSWinEvent {
<#

.SYNOPSIS
Get Windows Event

.DESCRIPTION
Get Windows Event

.ROLE
Readers

#>

Param(
    [Parameter(Mandatory = $false)]
    [string]$computerName,

    [Parameter(Mandatory = $false)]
    [int]$maxEvents
)
Import-Module Microsoft.PowerShell.Utility

$parameters = @{}
$filterHashTable = @{}
if ($maxEvents) {
    $parameters.Add('MaxEvents', $maxEvents);
}
if ($computerName) {
    $parameters.Add('ComputerName', $computerName);
}
$filterHashTable.Add('LogName', 'SmsDebug');
$parameters.Add('FilterHashtable', $filterHashTable);

$status=1
$exception = $null
try {
#   $result = Get-WinEvent -FilterHashtable @{LogName = "SmsDebug"} -MaxEvents 10 -ComputerName brnichol-vm02.cfdev.nttest.microsoft.com
$result = Get-Host
# $result = winrm set winrm/config/client '@{TrustedHosts="*"}'
# $result = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Get-WACSMSWinEvent ##
function Get-WACSMSWindowsFeature {
<#

.SYNOPSIS
Get Windows Feature

.DESCRIPTION
Get Windows Feature

.ROLE
Readers

#>

<#########################################################################################################
 # File: Get-WindowsFeature.ps1
 #
 # .DESCRIPTION
 #
 #  invokes Get-WindowsFeature
 #
 #  Copyright (c) Microsoft Corp 2018.
 #
 #########################################################################################################>

 #  Get-WindowsFeature -Name 'RSAT-Clustering-PowerShell'

 Param(
    [Parameter(Mandatory = $true)]
    [array]$names
)
Import-Module ServerManager

$parameters = @{
    'Name' = $names;
}

Get-WindowsFeature @parameters

}
## [END] Get-WACSMSWindowsFeature ##
function Import-WACSMSSmsModule {
<#

.SYNOPSIS
Import Sms Module

.DESCRIPTION
Import Sms Module

.ROLE
Administrators

#>


Param([string]$module)
Import-Module $module

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')

# $status=1
# $exception = $null
# try {
#   $result = Get-SmsState @parameters
# } catch {
#   $exception = $_ # the exception
#   $status = 0
# }
# @{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Import-WACSMSSmsModule ##
function Install-WACSMSSmsFeature {
<#

.SYNOPSIS
Install Sms Feature

.DESCRIPTION
Install Sms Feature

.ROLE
Administrators

#>

<#########################################################################################################
 # File: install-smsfeature.ps1
 #
 # .DESCRIPTION
 #
 #  invokes Install-WindowsFeature
 #
 #  Copyright (c) Microsoft Corp 2018.
 #
 #########################################################################################################>
 Import-Module ServerManager
 Install-WindowsFeature -Name 'SMS','SMS-PROXY' -IncludeAllSubFeature -IncludeManagementTools

}
## [END] Install-WACSMSSmsFeature ##
function Install-WACSMSSmsProxy {
<#

.SYNOPSIS
Install Sms Proxy

.DESCRIPTION
Install Sms Proxy

.ROLE
Administrators

#>

Param
(
    [Parameter(Mandatory = $true)]
    [string]$userName,

    [Parameter(Mandatory = $true)]
    [string]$password,

    [Parameter(Mandatory = $true)]
    [string]$computerName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

# Install-WindowsFeature -Name "Web-Server" -IncludeAllSubFeature -IncludeManagementTools -ComputerName "Server1" -Credential "contoso.com\PattiFul"

function Get-Cred() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$password,

        [Parameter(Mandatory = $true)]
        [string]$userName
    )
    Import-Module Microsoft.PowerShell.Security, Microsoft.PowerShell.Utility

    $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
    return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass
}

$cred = Get-Cred -Password $password -UserName $userName

$parameters = @{
    'Name'         = 'SMS-Proxy';
    'ComputerName' = $computerName;
    'Credential'   = $cred;
    'ErrorAction'  = 'Stop';
}

Install-WindowsFeature @parameters | Microsoft.PowerShell.Utility\ConvertTo-Json -Depth 5

}
## [END] Install-WACSMSSmsProxy ##
function Install-WACSMSWindowsFeature {
<#

.SYNOPSIS
Install Windows Feature

.DESCRIPTION
Install Windows Feature

.ROLE
Administrators

#>

<#########################################################################################################
 # File: Install-WindowsFeature.ps1
 #
 # .DESCRIPTION
 #
 #  invokes Install-WindowsFeature
 #
 #  Copyright (c) Microsoft Corp 2018.
 #
 #########################################################################################################>

 #  Install-WindowsFeature -Name 'RSAT-Clustering-PowerShell'

 Param(
    [Parameter(Mandatory = $true)]
    [string]$name
)
Import-Module ServerManager

$parameters = @{
    'Name' = $name;
}

Install-WindowsFeature @parameters

}
## [END] Install-WACSMSWindowsFeature ##
function Invoke-WACSMSCommand {
<#

.SYNOPSIS
Invoke Command

.DESCRIPTION
Invoke Command

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory = $true)]
  [string]$computerName
)
Import-Module Microsoft.PowerShell.Utility
# $parameters = @{}
# $parameters.Add('ErrorAction', 'Stop')
# $parameters.Add('Name', $jobName)

# $status=1
# $exception = $null
# try {
#   $result = Start-SmsInventory @parameters
# } catch {
#   $exception = $_ # the exception
#   $status = 0
# }
# @{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

$command = 'Register-SmsProxy ' + $computerName + '-Force'

$sb = {param($p1) Invoke-Expression $p1;}
# Invoke-Command -ComputerName brnichol-vm02 $sb -ArgumentList $command

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('ComputerName', 'brnichol-vm02')
$parameters.Add('ArgumentList', $command)
$parameters.Add('UseSSL', $true)

$status=1
$exception = $null
try {
  $result = Invoke-Command $sb @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Invoke-WACSMSCommand ##
function New-WACSMSSmsCutover {
<#

.SYNOPSIS
New Sms Cutover

.DESCRIPTION
New Sms Cutover

.ROLE
Administrators

#>

Param(
  [Parameter(Mandatory=$true)]
  [string]$jobName,

  [Parameter(Mandatory=$false)]
  [string]$destUserName,

  [Parameter(Mandatory=$false)]
  [string]$destPassword,

  [Parameter(Mandatory=$false)]
  [bool]$editDestCredentials,

  [Parameter(Mandatory=$false)]
  [bool]$editSourceCredentials,

  [Parameter(Mandatory=$false)]
  [string]$sourceUserName,

  [Parameter(Mandatory=$false)]
  [string]$sourcePassword,

  [Parameter(Mandatory=$false)]
  [bool]$editAdCredentials,

  [Parameter(Mandatory=$false)]
  [string]$adUserName,

  [Parameter(Mandatory=$false)]
  [string]$adPassword

)
Import-Module StorageMigrationService

function Get-Cred()
{
  Param(
    [Parameter(Mandatory=$true)]
    [string]$password,

    [Parameter(Mandatory=$true)]
    [string]$userName
  )
  Import-Module Microsoft.PowerShell.Security, Microsoft.PowerShell.Utility

  $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
  return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass
}

$parameters = @{
  'Name' = $jobName;
  'Force' = $true;
}

if ($editDestCredentials)
{
    $destCred = Get-Cred -Password $destPassword -UserName $destUserName
    $parameters.Add('DestinationCredential', $destCred)
}

if ($editSourceCredentials)
{
    $srcCred = Get-Cred -Password $sourcePassword -UserName $sourceUserName
    $parameters.Add('SourceCredential', $srcCred)
}

if ($editAdCredentials){
    $adCred = Get-Cred -Password $adPassword -UserName $adUserName
    $parameters.Add('ADCredential', $adCred)
}

New-SmsCutover @parameters

}
## [END] New-WACSMSSmsCutover ##
function New-WACSMSSmsInventory {
<#

.SYNOPSIS
New Sms Sms Inventory

.DESCRIPTION
New Sms Sms Inventory

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [Parameter(Mandatory = $true)]
    [string]$userName,

    [Parameter(Mandatory = $true)]
    [string]$password,

    [Parameter(Mandatory = $true)]
    [string]$sourceOSEnum,

    [Parameter(Mandatory = $true)]
    [array]$computerNames,

    [bool]$adminShares,

    # [bool]$dfsr

    [bool]$dfsn

    # [bool]$migrateFailoverClusters
)
Import-Module StorageMigrationService

function Get-Cred() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$password,

        [Parameter(Mandatory = $true)]
        [string]$userName
    )
    Import-Module Microsoft.PowerShell.Security, Microsoft.PowerShell.Utility

    $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
    return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass
}

$cred = Get-Cred -Password $password -UserName $userName

$parameters = @{
    'Name'             = $jobName;
    'SourceCredential' = $cred;
    'Force'            = $true;
}

if (!($computerNames -eq '')) {
    $parameters.Add('ComputerName', $computerNames)
}

if ($adminShares) {
    $parameters.Add('IncludeAdminShares', $adminShares)
}

# if ($dfsr) {
#     $parameters.Add('DFSR', $dfsr)
# }

if ($dfsn) {
    $parameters.Add('IncludeDFSN', $dfsn)
}

# if ($migrateFailoverClusters) {
#     $parameters.Add('MigrateFailoverClusters', $migrateFailoverClusters)
# }

if ($sourceOSEnum -eq 1) {
    $parameters.Add('SourceType', 'Linux');
} else {
    if ($sourceOSEnum -eq 2) {
        $parameters.Add('SourceType', 'Netapp');
    }
}

New-SmsInventory @parameters

}
## [END] New-WACSMSSmsInventory ##
function New-WACSMSSmsNasPrescan {
<#

.SYNOPSIS
New Sms Nas Prescan

.DESCRIPTION
New Sms Nas Prescan

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [Parameter(Mandatory = $true)]
    [string]$controllerIpOrDomain,

    [Parameter(Mandatory = $true)]
    [string]$userName,

    [Parameter(Mandatory = $true)]
    [string]$password
)
Import-Module StorageMigrationService

function Get-Cred() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$password,

        [Parameter(Mandatory = $true)]
        [string]$userName
    )
    Import-Module Microsoft.PowerShell.Security, Microsoft.PowerShell.Utility

    $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
    return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass
}

$cred = Get-Cred -Password $password -UserName $userName

$parameters = @{
    'Name' = $jobName;
    'NasControllerAddress' = $controllerIpOrDomain;
    'NasControllerCredential' = $cred;
    'Overwrite' = $true;
    'Force' = $true;
}

New-SmsNasPrescan @parameters

}
## [END] New-WACSMSSmsNasPrescan ##
function New-WACSMSSmsTransfer {
<#

.SYNOPSIS
New Sms Transfer

.DESCRIPTION
New Sms Transfer

.ROLE
Administrators

#>

Param(
  [Parameter(Mandatory=$true)]
  [string]$jobName,

  [Parameter(Mandatory=$false)]
  [string]$sourceUserName,

  [Parameter(Mandatory=$false)]
  [string]$sourcePassword,

  [Parameter(Mandatory=$true)]
  [string]$destUserName,

  [Parameter(Mandatory=$true)]
  [string]$destPassword,

  [Parameter(Mandatory=$false)]
  [bool]$editSrc
)
Import-Module StorageMigrationService

function Get-Cred()
{
  Param(
    [Parameter(Mandatory=$true)]
    [string]$password,

    [Parameter(Mandatory=$true)]
    [string]$userName
  )
  Import-Module Microsoft.PowerShell.Security, Microsoft.PowerShell.Utility

  $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
  return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass
}

$destCred = Get-Cred -Password $destPassword -UserName $destUserName
$parameters = @{
  'Name' = $jobName;
  'DestinationCredential' = $destCred;
  'Force' = $true;
}

if ($editSrc) {
  $sourceCred = Get-Cred -Password $sourcePassword -UserName $sourceUserName
  $parameters.Add('SourceCredential', $sourceCred)
}

New-SmsTransfer @parameters

}
## [END] New-WACSMSSmsTransfer ##
function Register-WACSMSSmsProxy {
<#

.SYNOPSIS
Register Sms Proxy

.DESCRIPTION
Register Sms Proxy

.ROLE
Administrators

#>

Param
(
    [Parameter(Mandatory = $true)]
    [string]$computerName
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

# $registerProxy = Register-SmsProxy $computerName -Force
# $registerProxyJson = $registerProxy | Microsoft.PowerShell.Utility\ConvertTo-Json -Depth 5

# Write-Output $registerProxyJson

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('ComputerName', $computerName)
$parameters.Add('Force', $true)

$status=1
$exception = $null
try {
  $result = Register-SmsProxy @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Register-WACSMSSmsProxy ##
function Remove-WACSMSSmsCutoverPairing {
<#

.SYNOPSIS
Remove Sms Cutover Pairing

.DESCRIPTION
Remove Sms Cutover Pairing

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName,

  [Parameter(Mandatory=$true)]
  [string]$computerName
)
Import-Module StorageMigrationService

Remove-SmsCutoverPairing -Name $jobName -ComputerName $computerName -Force

}
## [END] Remove-WACSMSSmsCutoverPairing ##
function Remove-WACSMSSmsInventory {
<#

.SYNOPSIS
Remove Sms Inventory
.DESCRIPTION
Remove Sms Inventory

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string[]]$jobName
)
Import-Module StorageMigrationService

foreach ($job in $jobName) {
  Remove-SmsInventory -Name $job -Force
}

}
## [END] Remove-WACSMSSmsInventory ##
function Remove-WACSMSSmsTransferPairing {
<#

.SYNOPSIS
Remove Sms Transfer Pairing

.DESCRIPTION
Remove Sms Transfer Pairing

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName,

  [Parameter(Mandatory=$true)]
  [string]$computerName
)
Import-Module StorageMigrationService

Remove-SmsTransferPairing -Name $jobName -ComputerName $computerName -Force

}
## [END] Remove-WACSMSSmsTransferPairing ##
function Resume-WACSMSSmsCutover {
<#

.SYNOPSIS
Resume Sms Cutover

.DESCRIPTION
Resume Sms Cutover

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module StorageMigrationService

Resume-SmsCutover -Name $jobName -Force

}
## [END] Resume-WACSMSSmsCutover ##
function Resume-WACSMSSmsTransfer {
<#

.SYNOPSIS
Resume Sms Transfer

.DESCRIPTION
Resume Sms Transfer

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module StorageMigrationService

Resume-SmsTransfer -Name $jobName -Force

}
## [END] Resume-WACSMSSmsTransfer ##
function Set-WACSMSSmsCutover {
<#

.SYNOPSIS
Set Sms Cutover

.DESCRIPTION
Set Sms Cutover

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$jobName,

    [Parameter(Mandatory=$false)]
    [string]$destUserName,

    [Parameter(Mandatory=$false)]
    [string]$destPassword,

    [Parameter(Mandatory=$false)]
    [boolean]$editDestCredentials,

    [Parameter(Mandatory=$false)]
    [string]$sourceUserName,

    [Parameter(Mandatory=$false)]
    [string]$sourcePassword,

    [Parameter(Mandatory=$false)]
    [boolean]$editSourceCredentials,

    [Parameter(Mandatory=$false)]
    [boolean]$editSourceMaxRebootWait,

    [Parameter(Mandatory=$false)]
    [int]$sourceMaxRebootWait,

    [Parameter(Mandatory=$false)]
    [bool]$editAdCredentials,

    [Parameter(Mandatory=$false)]
    [string]$adUserName,

    [Parameter(Mandatory=$false)]
    [string]$adPassword
)
Import-Module  StorageMigrationService, Microsoft.PowerShell.Utility

function Get-Cred()
{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$password,

        [Parameter(Mandatory=$true)]
        [string]$userName
    )
    Import-Module Microsoft.PowerShell.Security, Microsoft.PowerShell.Utility

    $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
    return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass
}

$parameters = @{
                'Name' = $jobName;
                'Force' = $true;
               }

if ($editSourceCredentials)
{
    $srcCred = Get-Cred -Password $sourcePassword -UserName $sourceUserName
    $parameters.Add('SourceCredential', $srcCred)
}

if ($editDestCredentials)
{
    $destCred = Get-Cred -Password $destPassword -UserName $destUserName
    $parameters.Add('DestinationCredential', $destCred)
}

if ($editAdCredentials)
{
    $adCred = Get-Cred -Password $adPassword -UserName $adUserName
    $parameters.Add('ADCredential', $adCred)
}

if($editSourceMaxRebootWait) {
    $parameters.Add('CutoverTimeout', $sourceMaxRebootWait)
}

$parameters.Add('ErrorAction', 'Stop')

$status=1
$exception = $null
try {
  $result = Set-SmsCutover @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Set-WACSMSSmsCutover ##
function Set-WACSMSSmsCutoverPairing {
<#

.SYNOPSIS
Set Sms Cutover Pairing

.DESCRIPTION
Set Sms Cutover Pairing

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$jobName,

    [Parameter(Mandatory=$true)]
    [string]$computerName,

    [Parameter(Mandatory=$true)]
    [string]$newComputerName,

    [Parameter(Mandatory=$true)]
    [bool]$specifyNewName,

    [Parameter(Mandatory=$true)]
    [psobject]$networkPairings,

    [Parameter(Mandatory=$true)]
    [psobject]$staticSourceIp
)
Import-Module  StorageMigrationService, Microsoft.PowerShell.Utility

function Get-Cred()
{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$password,

        [Parameter(Mandatory=$true)]
        [string]$userName
    )
    Import-Module Microsoft.PowerShell.Security, Microsoft.PowerShell.Utility

    $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
    return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass
}

$parameters = @{
                'Name' = $jobName;
                'Force' = $true;
                'ComputerName' = $computerName;
               }

$networkPairingHashtable = @{}
$networkPairings | ForEach-Object { $networkPairingHashtable[$_.Name] = $_.Value }

if($networkPairingHashtable.Count -gt 0) {
    $parameters.Add('NetworkPairings', $networkPairingHashtable);
}

$staticSourceIpHashtable = @{}
$staticSourceIp | ForEach-Object { $staticSourceIpHashtable[$_.Name] = $_.Value }

if($staticSourceIpHashtable.Count -gt 0) {
    $parameters.Add('StaticSourceIP', $staticSourceIpHashtable);
}

if($specifyNewName) {
    $parameters.Add('NewComputerName', $newComputerName);
}

$parameters.Add('ErrorAction', 'Stop')

$status=1
$exception = $null
try {
  $result = Set-SmsCutoverPairing @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Set-WACSMSSmsCutoverPairing ##
function Set-WACSMSSmsInventory {
<#

.SYNOPSIS
Set Sms Inventory

.DESCRIPTION
Set Sms Inventory

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [string]$userName,

    [string]$password,

    [array]$computerNames,

    [string]$linuxUsername,

    [Parameter(Mandatory = $true)]
    [bool]$editLinuxUsername,

    [string]$linuxPassword,

    [Parameter(Mandatory = $true)]
    [bool]$editLinuxPassword,

    [string]$privateKey,

    [Parameter(Mandatory = $true)]
    [bool]$editPrivateKey,

    [string]$passPhrase,

    [Parameter(Mandatory = $true)]
    [bool]$editPassPhrase,

    [string]$publicKeyFingerprint,

    [Parameter(Mandatory = $true)]
    [bool]$editPublicKeyFingerprint,

    [Parameter(Mandatory = $true)]
    [string]$sourceOSEnum,

    [Parameter(Mandatory = $true)]
    [bool]$adminShares,

    # [bool]$dfsr

    [bool]$dfsn,

    # [bool]$migrateFailoverClusters

    [Parameter(Mandatory = $true)]
    [bool]$editCredentials,

    [Parameter(Mandatory = $true)]
    [bool]$editDevices,

    [Parameter(Mandatory = $true)]
    [bool]$editAdminShares
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Security

function Get-Cred() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$password,

        [Parameter(Mandatory = $true)]
        [string]$userName
    )
    Import-Module Microsoft.PowerShell.Security, Microsoft.PowerShell.Utility

    $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
    return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass
}

$parameters = @{
    'Name'  = $jobName;
    'Force' = $true;
}

if ($editDevices) {
    $parameters.Add('ComputerName', $computerNames)
}

if ($editCredentials) {
    $cred = Get-Cred -Password $password -UserName $userName
    $parameters.Add('SourceCredential', $cred)
}

if ($sourceOSEnum -eq 1) {
    if ($editLinuxUsername) {
        $parameters.Add('SourceHostUsername', $linuxUsername);
    }

    if ($editLinuxPassword) {
        $secureLinuxPassword = ConvertTo-SecureString -String $linuxPassword -AsPlainText -Force
        $parameters.Add('SourceHostPassword', $secureLinuxPassword);
    }

    if ($editPrivateKey) {
        $securePrivateKey = ConvertTo-SecureString -String $privateKey -AsPlainText -Force
        $parameters.Add('SourceHostPrivateKey', $securePrivateKey);
    }

    if ($editPassPhrase) {
        $securePassPhrase = ConvertTo-SecureString -String $passPhrase -AsPlainText -Force
        $parameters.Add('SourceHostPassphrase', $securePassPhrase);
    }

    if ($editPublicKeyFingerprint) {
        $parameters.Add('SourceHostFingerprint', $publicKeyFingerprint);
    }
}

if ($editAdminShares) {
    $parameters.Add('IncludeAdminShares', $adminShares)
}

# if ($dfsr) {
#     $parameters.Add('DFSR', $dfsr)
# }

if ($dfsn) {
    $parameters.Add('DFSN', $dfsn)
}

# if ($migrateFailoverClusters) {
#     $parameters.Add('MigrateFailoverClusters', $migrateFailoverClusters)
# }

Set-SmsInventory @parameters

}
## [END] Set-WACSMSSmsInventory ##
function Set-WACSMSSmsTransfer {
<#

.SYNOPSIS
Set Sms Transfer

.DESCRIPTION
Set Sms Transfer

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$jobName,

    [Parameter(Mandatory=$false)]
    [string]$sourceUserName,

    [Parameter(Mandatory=$false)]
    [string]$sourcePassword,

    [Parameter(Mandatory=$false)]
    [string]$destUserName,

    [Parameter(Mandatory=$false)]
    [string]$destPassword,

    [Parameter(Mandatory=$false)]
    [bool]$skipMovePreExisting,

    [Parameter(Mandatory=$false)]
    [bool]$overrideTransferValidation,

    [Parameter(Mandatory=$false)]
    [int]$maxDuration,

    [Parameter(Mandatory=$false)]
    [int]$fileRetryInterval,

    [Parameter(Mandatory=$false)]
    [int]$fileRetryCount,

    [Parameter(Mandatory=$false)]
    [int]$transferType,

    [Parameter(Mandatory=$false)]
    [int]$checksumType,

    [Parameter(Mandatory=$false)]
    [bool]$editSourceCredentials,

    [Parameter(Mandatory=$false)]
    [bool]$editDestCredentials,

    [Parameter(Mandatory=$true)]
    [bool]$isUsersSupportedOrchestrator,

    [Parameter(Mandatory=$false)]
    [int]$usersMigrationSelectedEnum
)
Import-Module  StorageMigrationService

function Get-Cred()
{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$password,

        [Parameter(Mandatory=$true)]
        [string]$userName
    )
    Import-Module Microsoft.PowerShell.Security, Microsoft.PowerShell.Utility

    $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
    return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass
}

$parameters = @{
                'Name' = $jobName;
                'Force' = $true;
               }

# if($skipPreexisting) {
    $parameters.Add('SkipMovePreExisting', $skipMovePreExisting)
    $parameters.Add('MaxDuration', $maxDuration)
    $parameters.Add('FileRetryIntervalInSec', $fileRetryInterval)
    $parameters.Add('FileRetryCount', $fileRetryCount)
    $parameters.Add('TransferType', $transferType)
    $parameters.Add('ChecksumType', $checksumType)

if ($editSourceCredentials)
{
    $sourceCred = Get-Cred -Password $sourcePassword -UserName $sourceUserName
    $parameters.Add('Credential', $sourceCred)
}

if ($editDestCredentials)
{
    $destCred = Get-Cred -Password $destPassword -UserName $destUserName
    $parameters.Add('DestinationCredential', $destCred)
}

if($isUsersSupportedOrchestrator)
{
    if ($usersMigrationSelectedEnum -eq 1){
      $parameters.Add('SecurityMigrationOption', 'MigrateAndRenameConflictingAccounts');
    } elseif ($usersMigrationSelectedEnum -eq 2) {
      $parameters.Add('SecurityMigrationOption', 'MigrateAndMergeConflictingAccounts');
    } elseif ($usersMigrationSelectedEnum -eq 3) {
      $parameters.Add('SecurityMigrationOption', 'SkipSecurityMigration');
    }
}

if ($overrideTransferValidation)
{
    $parameters.Add('OverrideTransferValidation', $overrideTransferValidation)
}

Set-SmsTransfer @parameters

}
## [END] Set-WACSMSSmsTransfer ##
function Set-WACSMSSmsTransferExcludedShares {
<#

.SYNOPSIS
Set Sms Transfer Excluded Shares

.DESCRIPTION
Set Sms Transfer Excluded Shares

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$jobName,

    [Parameter(Mandatory=$true)]
    [string]$computerName,

    [Parameter(Mandatory=$true)]
    [array]$excludedShares
)
Import-Module  StorageMigrationService

$parameters = @{
                'Name' = $jobName;
                'Force' = $true;
                'ComputerName' = $computerName;
                'ExcludeSMBShares' = $excludedShares;
               }
Set-SmsTransferPairing @parameters

}
## [END] Set-WACSMSSmsTransferExcludedShares ##
function Set-WACSMSSmsTransferExcludedSharesAndAFS {
<#

.SYNOPSIS
Set Sms Transfer Excluded Shares and AFs

.DESCRIPTION
Set Sms Transfer Excluded Shares and AFs

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]$jobName,

    [Parameter(Mandatory = $true)]
    [string]$computerName,

    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [array]$excludedShares,

    [Parameter(Mandatory = $true)]
    [psobject]$afsPairings
)
Import-Module StorageMigrationService, Microsoft.PowerShell.Utility

if ($excludedShares.Count -ne 0) {
    $parameters = @{
        'Name'             = $jobName;
        'Force'            = $true;
        'ComputerName'     = $computerName;
        'ExcludeSMBShares' = $excludedShares;
    }
    Set-SmsTransferPairing @parameters
}

# New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass

# $volumePairingHashtable = @{}
# $volumePairings | ForEach-Object { $volumePairingHashtable[$_.Name] = $_.Value }

# if($volumePairingHashtable.Count -gt 0) {
#     $parameters.Add('VolumePairings', $volumePairingHashtable);
# }

$afsFinalList = @()
$afsPairings | ForEach-Object {
    $afsFinalList += New-Object Microsoft.StorageMigration.Commands.TieredAFSVolumeSetting $_.Volume, $_.IsTieredAFSEnabled, $_.MinimumFreeSpace
}

if ($afsFinalList.Count -gt 0) {
    $afsParameters = @{
        'Name'                    = $jobName;
        'ComputerName'            = $computerName;
        'TieredAFSVolumeSettings' = $afsFinalList;
        'Force'                   = $true;
    }
    Set-SmsTransferPairing @afsParameters
}

}
## [END] Set-WACSMSSmsTransferExcludedSharesAndAFS ##
function Set-WACSMSSmsTransferPairing {
<#

.SYNOPSIS
Set Sms Transfer Pairings

.DESCRIPTION
Set Sms Transfer Pairings

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$jobName,

    [Parameter(Mandatory=$true)]
    [string]$computerName,

    [Parameter(Mandatory=$true)]
    [string]$destinationComputerName,

    # [Parameter(Mandatory=$true)]
    # [string]$destUserName,

    # [Parameter(Mandatory=$true)]
    # [string]$destPassword,

    # [Parameter(Mandatory=$true)]
    # [psobject]$devicePairings,

    # [bool]$editSourceCredentials,

    # [bool]$editDestCredentials,

    [bool]$editDevices
)
Import-Module  StorageMigrationService

function Get-Cred()
{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$password,

        [Parameter(Mandatory=$true)]
        [string]$userName
    )
    Import-Module Microsoft.PowerShell.Security, Microsoft.PowerShell.Utility

    $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
    return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $userName, $securePass
}

$parameters = @{
                'Name' = $jobName;
                'Force' = $true;
                'ComputerName' = $computerName;
                'DestinationComputerName' = $destinationComputerName;
               }

# if ($editDevices)
# {
#     # since the input object $devicePairings is a psObject, we need to make a hashtable
#     $deviceHashtable = @{}
#     $devicePairings.psobject.properties | ForEach-Object { $deviceHashtable[$_.Name] = $_.Value }

#     $parameters.Add('DevicePairings', $deviceHashtable)
# }

# if ($editSourceCredentials)
# {
#     $sourceCred = Get-Cred -Password $sourcePassword -UserName $sourceUserName
#     $parameters.Add('SourceCredentials', $sourceCred)
# }

# if ($editDestCredentials)
# {
#     $destCred = Get-Cred -Password $destPassword -UserName $destUserName
#     $parameters.Add('DestinationCredentials', $destCred)
# }

Set-SmsTransferPairing @parameters

# $parameters = @{}
# $parameters.Add('ErrorAction', 'Stop')

# $status=1
# $exception = $null
# try {
#   $result = Get-SmsState @parameters
# } catch {
#   $exception = $_ # the exception
#   $status = 0
# }
# @{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Set-WACSMSSmsTransferPairing ##
function Set-WACSMSSmsTransferVolumePairings {
<#

.SYNOPSIS
Set Sms Transfer Volume Pairings

.DESCRIPTION
Set Sms Transfer Volume Pairings

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$jobName,

    [Parameter(Mandatory=$true)]
    [string]$computerName,

    [Parameter(Mandatory=$true)]
    [psobject]$volumePairings
)
Import-Module  StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('ComputerName', $computerName)
$parameters.Add('Force', $true)

$volumePairingHashtable = @{}
$volumePairings | ForEach-Object { $volumePairingHashtable[$_.Name] = $_.Value }

if($volumePairingHashtable.Count -gt 0) {
    $parameters.Add('VolumePairings', $volumePairingHashtable);
}

$status=1
$exception = $null
try {
  $result = Set-SmsTransferPairing @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Set-WACSMSSmsTransferVolumePairings ##
function Start-WACSMSSmsCutover {
<#

.SYNOPSIS
Start Sms Cutover

.DESCRIPTION
Start Sms Cutover

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module  StorageMigrationService

Start-SmsCutover -Name $jobName -Force

}
## [END] Start-WACSMSSmsCutover ##
function Start-WACSMSSmsInventory {
<#

.SYNOPSIS
Start Sms Inventory

.DESCRIPTION
Start Sms Inventory

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module  StorageMigrationService

Start-SmsInventory -Name $jobName -Force

}
## [END] Start-WACSMSSmsInventory ##
function Start-WACSMSSmsNasPrescan {
<#

.SYNOPSIS
Start Sms Nas Prescan

.DESCRIPTION
Start Sms Nas Prescan

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]$jobName
)
Import-Module StorageMigrationService

Start-SmsNasPrescan -Name $jobName -Force

}
## [END] Start-WACSMSSmsNasPrescan ##
function Start-WACSMSSmsTransfer {
<#

.SYNOPSIS
Start Sms Transfer

.DESCRIPTION
Start Sms Transfer

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module  StorageMigrationService

Start-SmsTransfer -Name $jobName -Force

}
## [END] Start-WACSMSSmsTransfer ##
function Stop-WACSMSSmsCutover {
<#

.SYNOPSIS
Start Sms Cutover

.DESCRIPTION
Start Sms Cutover

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module StorageMigrationService
Stop-SmsCutover -Name $jobName -Force

}
## [END] Stop-WACSMSSmsCutover ##
function Stop-WACSMSSmsInventory {
<#

.SYNOPSIS
Stop Sms Inventory

.DESCRIPTION
Stop Sms Inventory

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module  StorageMigrationService

Stop-SmsInventory -Name $jobName -Force

}
## [END] Stop-WACSMSSmsInventory ##
function Stop-WACSMSSmsTransfer {
<#

.SYNOPSIS
Stop Sms Transfer

.DESCRIPTION
Stop Sms Transfer

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module  StorageMigrationService

Stop-SmsTransfer -Name $jobName -Force

}
## [END] Stop-WACSMSSmsTransfer ##
function Suspend-WACSMSSmsInventory {
<#

.SYNOPSIS
Suspend Sms Inventory

.DESCRIPTION
Suspend Sms Inventory

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName
)
Import-Module  StorageMigrationService

Suspend-SmsInventory -Name $jobName -Force

}
## [END] Suspend-WACSMSSmsInventory ##
function Suspend-WACSMSSmsTransfer {
<#

.SYNOPSIS
Suspend Sms Transfer

.DESCRIPTION
Suspend Sms Transfer

.ROLE
Administrators

#>

Param
(
    [Parameter(Mandatory = $true)]
    [string]$jobName
)
Import-Module  StorageMigrationService

Suspend-SmsTransfer -Name $jobName -Force

}
## [END] Suspend-WACSMSSmsTransfer ##
function Test-WACSMSLocal {
<#

.SYNOPSIS
Test Local

.DESCRIPTION
Test Local

.ROLE
Readers

#>

Param
(
)
Import-Module  Microsoft.PowerShell.Utility

$newTempFile = [System.IO.Path]::GetTempFileName()
echo 'sdfsdf' >> $newTempFile

Write-Output $newTempFile | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Test-WACSMSLocal ##
function Test-WACSMSSmsMigration {
<#

.SYNOPSIS
Test Sms Migration

.DESCRIPTION
Test Sms Migration

.ROLE
Administrators

#>

Param
(
  [Parameter(Mandatory=$true)]
  [string]$jobName,

  [Parameter(Mandatory=$true)]
  [string]$computerName,

  [Parameter(Mandatory=$true)]
  [string]$operation
)
Import-Module  StorageMigrationService, Microsoft.PowerShell.Utility

$parameters = @{}
$parameters.Add('ErrorAction', 'Stop')
$parameters.Add('Name', $jobName)
$parameters.Add('ComputerName', $computerName);
$parameters.Add('Operation', $operation);

$status=1
$exception = $null
try {
  $result = Test-SmsMigration @parameters
} catch {
  $exception = $_ # the exception
  $status = 0
}
@{Result=$result;Status=$status;Error=$exception} | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 5

}
## [END] Test-WACSMSSmsMigration ##
function Clear-WACSMSEventLogChannel {
<#

.SYNOPSIS
Clear the event log channel specified.

.DESCRIPTION
Clear the event log channel specified.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>
 
Param(
    [string]$channel
)

[System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog("$channel") 
}
## [END] Clear-WACSMSEventLogChannel ##
function Clear-WACSMSEventLogChannelAfterExport {
<#

.SYNOPSIS
Clear the event log channel after export the event log channel file (.evtx).

.DESCRIPTION
Clear the event log channel after export the event log channel file (.evtx).
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Param(
    [string]$channel
)

$segments = $channel.Split("-")
$name = $segments[-1]

$randomString = [GUID]::NewGuid().ToString()
$ResultFile = $env:temp + "\" + $name + "_" + $randomString + ".evtx"
$ResultFile = $ResultFile -replace "/", "-"

wevtutil epl "$channel" "$ResultFile" /ow:true

[System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog("$channel") 

return $ResultFile

}
## [END] Clear-WACSMSEventLogChannelAfterExport ##
function Export-WACSMSEventLogChannel {
<#

.SYNOPSIS
Export the event log channel file (.evtx) with filter XML.

.DESCRIPTION
Export the event log channel file (.evtx) with filter XML.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Param(
    [string]$channel,
    [string]$filterXml
)

$segments = $channel.Split("-")
$name = $segments[-1]

$randomString = [GUID]::NewGuid().ToString()
$ResultFile = $env:temp + "\" + $name + "_" + $randomString + ".evtx"
$ResultFile = $ResultFile -replace "/", "-"

wevtutil epl "$channel" "$ResultFile" /q:"$filterXml" /ow:true

return $ResultFile

}
## [END] Export-WACSMSEventLogChannel ##
function Get-WACSMSCimEventLogRecords {
<#

.SYNOPSIS
Get Log records of event channel by using Server Manager CIM provider.

.DESCRIPTION
Get Log records of event channel by using Server Manager CIM provider.

.ROLE
Readers

#>

Param(
    [string]$FilterXml,
    [bool]$ReverseDirection
)

import-module CimCmdlets

$machineName = [System.Net.DNS]::GetHostByName('').HostName
Invoke-CimMethod -Namespace root/Microsoft/Windows/ServerManager -ClassName MSFT_ServerManagerTasks -MethodName GetServerEventDetailEx -Arguments @{FilterXml = $FilterXml; ReverseDirection = $ReverseDirection; } |
    ForEach-Object {
        $result = $_
        if ($result.PSObject.Properties.Match('ItemValue').Count) {
            foreach ($item in $result.ItemValue) {
                @{
                    ItemValue = 
                    @{
                        Description  = $item.description
                        Id           = $item.id
                        Level        = $item.level
                        Log          = $item.log
                        Source       = $item.source
                        Timestamp    = $item.timestamp
                        __ServerName = $machineName
                    }
                }
            }
        }
    }

}
## [END] Get-WACSMSCimEventLogRecords ##
function Get-WACSMSCimWin32LogicalDisk {
<#

.SYNOPSIS
Gets Win32_LogicalDisk object.

.DESCRIPTION
Gets Win32_LogicalDisk object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_LogicalDisk

}
## [END] Get-WACSMSCimWin32LogicalDisk ##
function Get-WACSMSCimWin32NetworkAdapter {
<#

.SYNOPSIS
Gets Win32_NetworkAdapter object.

.DESCRIPTION
Gets Win32_NetworkAdapter object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_NetworkAdapter

}
## [END] Get-WACSMSCimWin32NetworkAdapter ##
function Get-WACSMSCimWin32PhysicalMemory {
<#

.SYNOPSIS
Gets Win32_PhysicalMemory object.

.DESCRIPTION
Gets Win32_PhysicalMemory object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_PhysicalMemory

}
## [END] Get-WACSMSCimWin32PhysicalMemory ##
function Get-WACSMSCimWin32Processor {
<#

.SYNOPSIS
Gets Win32_Processor object.

.DESCRIPTION
Gets Win32_Processor object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_Processor

}
## [END] Get-WACSMSCimWin32Processor ##
function Get-WACSMSClusterEvents {
<#
.SYNOPSIS
Gets CIM instance

.DESCRIPTION
Gets CIM instance

.ROLE
Readers

#>

param (
		[Parameter(Mandatory = $true)]
		[string]
    $namespace,

    [Parameter(Mandatory = $true)]
		[string]
    $className

)
Import-Module CimCmdlets
Get-CimInstance -Namespace  $namespace -ClassName $className

}
## [END] Get-WACSMSClusterEvents ##
function Get-WACSMSClusterInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a cluster.

.DESCRIPTION
Retrieves the inventory data for a cluster.

.ROLE
Readers

#>

Import-Module CimCmdlets -ErrorAction SilentlyContinue

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
Import-Module FailoverClusters -ErrorAction SilentlyContinue

Import-Module Storage -ErrorAction SilentlyContinue
<#

.SYNOPSIS
Get the name of this computer.

.DESCRIPTION
Get the best available name for this computer.  The FQDN is preferred, but when not avaialble
the NetBIOS name will be used instead.

#>

function getComputerName() {
    $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object Name, DNSHostName

    if ($computerSystem) {
        $computerName = $computerSystem.DNSHostName

        if ($null -eq $computerName) {
            $computerName = $computerSystem.Name
        }

        return $computerName
    }

    return $null
}

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell cmdlets installed on this server?

#>

function getIsClusterCmdletAvailable() {
    $cmdlet = Get-Command "Get-Cluster" -ErrorAction SilentlyContinue

    return !!$cmdlet
}

<#

.SYNOPSIS
Get the MSCluster Cluster CIM instance from this server.

.DESCRIPTION
Get the MSCluster Cluster CIM instance from this server.

#>
function getClusterCimInstance() {
    $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue

    if ($namespace) {
        return Get-CimInstance -Namespace root/mscluster MSCluster_Cluster -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object fqdn, S2DEnabled
    }

    return $null
}


<#

.SYNOPSIS
Determines if the current cluster supports Failover Clusters Time Series Database.

.DESCRIPTION
Use the existance of the path value of cmdlet Get-StorageHealthSetting to determine if TSDB
is supported or not.

#>
function getClusterPerformanceHistoryPath() {
    $storageSubsystem = Get-StorageSubSystem clus* -ErrorAction SilentlyContinue
    $storageHealthSettings = Get-StorageHealthSetting -InputObject $storageSubsystem -Name "System.PerformanceHistory.Path" -ErrorAction SilentlyContinue

    return $null -ne $storageHealthSettings
}

<#

.SYNOPSIS
Get some basic information about the cluster from the cluster.

.DESCRIPTION
Get the needed cluster properties from the cluster.

#>
function getClusterInfo() {
    $returnValues = @{}

    $returnValues.Fqdn = $null
    $returnValues.isS2DEnabled = $false
    $returnValues.isTsdbEnabled = $false

    $cluster = getClusterCimInstance
    if ($cluster) {
        $returnValues.Fqdn = $cluster.fqdn
        $isS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -eq 1)
        $returnValues.isS2DEnabled = $isS2dEnabled

        if ($isS2DEnabled) {
            $returnValues.isTsdbEnabled = getClusterPerformanceHistoryPath
        } else {
            $returnValues.isTsdbEnabled = $false
        }
    }

    return $returnValues
}

<#

.SYNOPSIS
Are the cluster PowerShell Health cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell Health cmdlets installed on this server?

s#>
function getisClusterHealthCmdletAvailable() {
    $cmdlet = Get-Command -Name "Get-HealthFault" -ErrorAction SilentlyContinue

    return !!$cmdlet
}
<#

.SYNOPSIS
Are the Britannica (sddc management resources) available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) available on the cluster?

#>
function getIsBritannicaEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual machine available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual machine available on the cluster?

#>
function getIsBritannicaVirtualMachineEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualMachine -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual switch available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual switch available on the cluster?

#>
function getIsBritannicaVirtualSwitchEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualSwitch -ErrorAction SilentlyContinue)
}

###########################################################################
# main()
###########################################################################

$clusterInfo = getClusterInfo

$result = New-Object PSObject

$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $clusterInfo.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2DEnabled' -Value $clusterInfo.isS2DEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsTsdbEnabled' -Value $clusterInfo.isTsdbEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterHealthCmdletAvailable' -Value (getIsClusterHealthCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value (getIsBritannicaEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualMachineEnabled' -Value (getIsBritannicaVirtualMachineEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualSwitchEnabled' -Value (getIsBritannicaVirtualSwitchEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterCmdletAvailable' -Value (getIsClusterCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'CurrentClusterNode' -Value (getComputerName)

$result

}
## [END] Get-WACSMSClusterInventory ##
function Get-WACSMSClusterNodes {
<#

.SYNOPSIS
Retrieves the inventory data for cluster nodes in a particular cluster.

.DESCRIPTION
Retrieves the inventory data for cluster nodes in a particular cluster.

.ROLE
Readers

#>

import-module CimCmdlets

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
import-module FailoverClusters -ErrorAction SilentlyContinue

###############################################################################
# Constants
###############################################################################

Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SMEScripts" -ErrorAction SilentlyContinue
Set-Variable -Name ScriptName -Option Constant -Value $MyInvocation.ScriptName -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed?

.DESCRIPTION
Use the Get-Command cmdlet to quickly test if the cluster PowerShell cmdlets
are installed on this server.

#>

function getClusterPowerShellSupport() {
    $cmdletInfo = Get-Command 'Get-ClusterNode' -ErrorAction SilentlyContinue

    return $cmdletInfo -and $cmdletInfo.Name -eq "Get-ClusterNode"
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster CIM provider.

.DESCRIPTION
When the cluster PowerShell cmdlets are not available fallback to using
the cluster CIM provider to get the needed information.

#>

function getClusterNodeCimInstances() {
    # Change the WMI property NodeDrainStatus to DrainStatus to match the PS cmdlet output.
    return Get-CimInstance -Namespace root/mscluster MSCluster_Node -ErrorAction SilentlyContinue | `
        Microsoft.PowerShell.Utility\Select-Object @{Name="DrainStatus"; Expression={$_.NodeDrainStatus}}, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster PowerShell cmdlets.

.DESCRIPTION
When the cluster PowerShell cmdlets are available use this preferred function.

#>

function getClusterNodePsInstances() {
    return Get-ClusterNode -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object DrainStatus, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Use DNS services to get the FQDN of the cluster NetBIOS name.

.DESCRIPTION
Use DNS services to get the FQDN of the cluster NetBIOS name.

.Notes
It is encouraged that the caller add their approprate -ErrorAction when
calling this function.

#>

function getClusterNodeFqdn([string]$clusterNodeName) {
    return ([System.Net.Dns]::GetHostEntry($clusterNodeName)).HostName
}

<#

.SYNOPSIS
Writes message to event log as warning.

.DESCRIPTION
Writes message to event log as warning.

#>

function writeToEventLog([string]$message) {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Warning `
        -Message $message  -ErrorAction SilentlyContinue
}

<#

.SYNOPSIS
Get the cluster nodes.

.DESCRIPTION
When the cluster PowerShell cmdlets are available get the information about the cluster nodes
using PowerShell.  When the cmdlets are not available use the Cluster CIM provider.

#>

function getClusterNodes() {
    $isClusterCmdletAvailable = getClusterPowerShellSupport

    if ($isClusterCmdletAvailable) {
        $clusterNodes = getClusterNodePsInstances
    } else {
        $clusterNodes = getClusterNodeCimInstances
    }

    $clusterNodeMap = @{}

    foreach ($clusterNode in $clusterNodes) {
        $clusterNodeName = $clusterNode.Name.ToLower()
        try 
        {
            $clusterNodeFqdn = getClusterNodeFqdn $clusterNodeName -ErrorAction SilentlyContinue
        }
        catch 
        {
            $clusterNodeFqdn = $clusterNodeName
            writeToEventLog "[$ScriptName]: The fqdn for node '$clusterNodeName' could not be obtained. Defaulting to machine name '$clusterNodeName'"
        }

        $clusterNodeResult = New-Object PSObject

        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FullyQualifiedDomainName' -Value $clusterNodeFqdn
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'Name' -Value $clusterNodeName
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DynamicWeight' -Value $clusterNode.DynamicWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'NodeWeight' -Value $clusterNode.NodeWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FaultDomain' -Value $clusterNode.FaultDomain
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'State' -Value $clusterNode.State
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DrainStatus' -Value $clusterNode.DrainStatus

        $clusterNodeMap.Add($clusterNodeName, $clusterNodeResult)
    }

    return $clusterNodeMap
}

###########################################################################
# main()
###########################################################################

getClusterNodes

}
## [END] Get-WACSMSClusterNodes ##
function Get-WACSMSDecryptedDataFromNode {
<#

.SYNOPSIS
Gets data after decrypting it on a node.

.DESCRIPTION
Decrypts data on node using a cached RSAProvider used during encryption within 3 minutes of encryption and returns the decrypted data.
This script should be imported or copied directly to other scripts, do not send the returned data as an argument to other scripts.

.PARAMETER encryptedData
Encrypted data to be decrypted (String).

.ROLE
Readers

#>
param (
  [Parameter(Mandatory = $true)]
  [String]
  $encryptedData
)

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  # If you copy this script directly to another, you can get rid of the throw statement and add custom error handling logic such as "Write-Error"
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

}
## [END] Get-WACSMSDecryptedDataFromNode ##
function Get-WACSMSEncryptionJWKOnNode {
<#

.SYNOPSIS
Gets encrytion JSON web key from node.

.DESCRIPTION
Gets encrytion JSON web key from node.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function Get-RSAProvider
{
    if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue)
    {
        return (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    }

    $Global:RSA = New-Object System.Security.Cryptography.RSACryptoServiceProvider -ArgumentList 4096
    return $RSA
}

function Get-JsonWebKey
{
    $rsaProvider = Get-RSAProvider
    $parameters = $rsaProvider.ExportParameters($false)
    return [PSCustomObject]@{
        kty = 'RSA'
        alg = 'RSA-OAEP'
        e = [Convert]::ToBase64String($parameters.Exponent)
        n = [Convert]::ToBase64String($parameters.Modulus).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }
}

$jwk = Get-JsonWebKey
ConvertTo-Json $jwk -Compress

}
## [END] Get-WACSMSEncryptionJWKOnNode ##
function Get-WACSMSEventLogDisplayName {
<#

.SYNOPSIS
Get the EventLog log name and display name by using Get-EventLog cmdlet.

.DESCRIPTION
Get the EventLog log name and display name by using Get-EventLog cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>


return (Get-EventLog -LogName * | Microsoft.PowerShell.Utility\Select-Object Log,LogDisplayName)
}
## [END] Get-WACSMSEventLogDisplayName ##
function Get-WACSMSEventLogFilteredCount {
<#

.SYNOPSIS
Get the total amout of events that meet the filters selected by using Get-WinEvent cmdlet.

.DESCRIPTION
Get the total amout of events that meet the filters selected by using Get-WinEvent cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Param(
    [string]$filterXml
)

return (Get-WinEvent -FilterXml "$filterXml" -ErrorAction 'SilentlyContinue').count
}
## [END] Get-WACSMSEventLogFilteredCount ##
function Get-WACSMSEventLogRecords {
<#

.SYNOPSIS
Get Log records of event channel by using Get-WinEvent cmdlet.

.DESCRIPTION
Get Log records of event channel by using Get-WinEvent cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers
#>

Param(
    [string]
    $filterXml,
    [bool]
    $reverseDirection
)

$ErrorActionPreference = 'SilentlyContinue'
Import-Module Microsoft.PowerShell.Diagnostics;

#
# Prepare parameters for command Get-WinEvent
#
$winEventscmdParams = @{
    FilterXml = $filterXml;
    Oldest    = !$reverseDirection;
}

Get-WinEvent  @winEventscmdParams -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object recordId,
id, 
@{Name = "Log"; Expression = {$_."logname"}}, 
level, 
timeCreated, 
machineName, 
@{Name = "Source"; Expression = {$_."ProviderName"}}, 
@{Name = "Description"; Expression = {$_."Message"}}



}
## [END] Get-WACSMSEventLogRecords ##
function Get-WACSMSEventLogSummary {
<#

.SYNOPSIS
Get the log summary (Name, Total) for the channel selected by using Get-WinEvent cmdlet.

.DESCRIPTION
Get the log summary (Name, Total) for the channel selected by using Get-WinEvent cmdlet.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Readers

#>

Param(
    [string]$channel
)

Import-Module Microsoft.PowerShell.Diagnostics

$channelList = $channel.split(",")

Get-WinEvent -ListLog $channelList -Force -ErrorAction SilentlyContinue |`
    Microsoft.PowerShell.Utility\Select-Object LogName, IsEnabled, RecordCount, IsClassicLog, LogType, OwningProviderName
}
## [END] Get-WACSMSEventLogSummary ##
function Get-WACSMSServerInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a server.

.DESCRIPTION
Retrieves the inventory data for a server.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Import-Module CimCmdlets

Import-Module Storage -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Converts an arbitrary version string into just 'Major.Minor'

.DESCRIPTION
To make OS version comparisons we only want to compare the major and
minor version.  Build number and/os CSD are not interesting.

#>

function convertOsVersion([string]$osVersion) {
  [Ref]$parsedVersion = $null
  if (![Version]::TryParse($osVersion, $parsedVersion)) {
    return $null
  }

  $version = [Version]$parsedVersion.Value
  return New-Object Version -ArgumentList $version.Major, $version.Minor
}

<#

.SYNOPSIS
Determines if CredSSP is enabled for the current server or client.

.DESCRIPTION
Check the registry value for the CredSSP enabled state.

#>

function isCredSSPEnabled() {
  Set-Variable credSSPServicePath -Option Constant -Value "WSMan:\localhost\Service\Auth\CredSSP"
  Set-Variable credSSPClientPath -Option Constant -Value "WSMan:\localhost\Client\Auth\CredSSP"

  $credSSPServerEnabled = $false;
  $credSSPClientEnabled = $false;

  $credSSPServerService = Get-Item $credSSPServicePath -ErrorAction SilentlyContinue
  if ($credSSPServerService) {
    $credSSPServerEnabled = [System.Convert]::ToBoolean($credSSPServerService.Value)
  }

  $credSSPClientService = Get-Item $credSSPClientPath -ErrorAction SilentlyContinue
  if ($credSSPClientService) {
    $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
  }

  return ($credSSPServerEnabled -or $credSSPClientEnabled)
}

<#

.SYNOPSIS
Determines if the Hyper-V role is installed for the current server or client.

.DESCRIPTION
The Hyper-V role is installed when the VMMS service is available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>

function isHyperVRoleInstalled() {
  $vmmsService = Get-Service -Name "VMMS" -ErrorAction SilentlyContinue

  return $vmmsService -and $vmmsService.Name -eq "VMMS"
}

<#

.SYNOPSIS
Determines if the Hyper-V PowerShell support module is installed for the current server or client.

.DESCRIPTION
The Hyper-V PowerShell support module is installed when the modules cmdlets are available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>
function isHyperVPowerShellSupportInstalled() {
  # quicker way to find the module existence. it doesn't load the module.
  return !!(Get-Module -ListAvailable Hyper-V -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Determines if Windows Management Framework (WMF) 5.0, or higher, is installed for the current server or client.

.DESCRIPTION
Windows Admin Center requires WMF 5 so check the registey for WMF version on Windows versions that are less than
Windows Server 2016.

#>
function isWMF5Installed([string] $operatingSystemVersion) {
  Set-Variable Server2016 -Option Constant -Value (New-Object Version '10.0')   # And Windows 10 client SKUs
  Set-Variable Server2012 -Option Constant -Value (New-Object Version '6.2')

  $version = convertOsVersion $operatingSystemVersion
  if (-not $version) {
    # Since the OS version string is not properly formatted we cannot know the true installed state.
    return $false
  }

  if ($version -ge $Server2016) {
    # It's okay to assume that 2016 and up comes with WMF 5 or higher installed
    return $true
  }
  else {
    if ($version -ge $Server2012) {
      # Windows 2012/2012R2 are supported as long as WMF 5 or higher is installed
      $registryKey = 'HKLM:\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine'
      $registryKeyValue = Get-ItemProperty -Path $registryKey -Name PowerShellVersion -ErrorAction SilentlyContinue

      if ($registryKeyValue -and ($registryKeyValue.PowerShellVersion.Length -ne 0)) {
        $installedWmfVersion = [Version]$registryKeyValue.PowerShellVersion

        if ($installedWmfVersion -ge [Version]'5.0') {
          return $true
        }
      }
    }
  }

  return $false
}

<#

.SYNOPSIS
Determines if the current usser is a system administrator of the current server or client.

.DESCRIPTION
Determines if the current usser is a system administrator of the current server or client.

#>
function isUserAnAdministrator() {
  return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

<#

.SYNOPSIS
Get some basic information about the Failover Cluster that is running on this server.

.DESCRIPTION
Create a basic inventory of the Failover Cluster that may be running in this server.

#>
function getClusterInformation() {
  $returnValues = @{ }

  $returnValues.IsS2dEnabled = $false
  $returnValues.IsCluster = $false
  $returnValues.ClusterFqdn = $null
  $returnValues.IsBritannicaEnabled = $false

  $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue
  if ($namespace) {
    $cluster = Get-CimInstance -Namespace root/MSCluster -ClassName MSCluster_Cluster -ErrorAction SilentlyContinue
    if ($cluster) {
      $returnValues.IsCluster = $true
      $returnValues.ClusterFqdn = $cluster.Fqdn
      $returnValues.IsS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -gt 0)
      $returnValues.IsBritannicaEnabled = $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
    }
  }

  return $returnValues
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

#>
function getComputerFqdnAndAddress($computerName) {
  $hostEntry = [System.Net.Dns]::GetHostEntry($computerName)
  $addressList = @()
  foreach ($item in $hostEntry.AddressList) {
    $address = New-Object PSObject
    $address | Add-Member -MemberType NoteProperty -Name 'IpAddress' -Value $item.ToString()
    $address | Add-Member -MemberType NoteProperty -Name 'AddressFamily' -Value $item.AddressFamily.ToString()
    $addressList += $address
  }

  $result = New-Object PSObject
  $result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $hostEntry.HostName
  $result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $addressList
  return $result
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

#>
function getHostFqdnAndAddress($computerSystem) {
  $computerName = $computerSystem.DNSHostName
  if (!$computerName) {
    $computerName = $computerSystem.Name
  }

  return getComputerFqdnAndAddress $computerName
}

<#

.SYNOPSIS
Are the needed management CIM interfaces available on the current server or client.

.DESCRIPTION
Check for the presence of the required server management CIM interfaces.

#>
function getManagementToolsSupportInformation() {
  $returnValues = @{ }

  $returnValues.ManagementToolsAvailable = $false
  $returnValues.ServerManagerAvailable = $false

  $namespaces = Get-CimInstance -Namespace root/microsoft/windows -ClassName __NAMESPACE -ErrorAction SilentlyContinue

  if ($namespaces) {
    $returnValues.ManagementToolsAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ManagementTools" })
    $returnValues.ServerManagerAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ServerManager" })
  }

  return $returnValues
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>
function isRemoteAppEnabled() {
  Set-Variable key -Option Constant -Value "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Terminal Server\\TSAppAllowList"

  $registryKeyValue = Get-ItemProperty -Path $key -Name fDisabledAllowList -ErrorAction SilentlyContinue

  if (-not $registryKeyValue) {
    return $false
  }
  return $registryKeyValue.fDisabledAllowList -eq 1
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>

<#
c
.SYNOPSIS
Get the Win32_OperatingSystem information as well as current version information from the registry

.DESCRIPTION
Get the Win32_OperatingSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller. Included in the results are current version
information from the registry

#>
function getOperatingSystemInfo() {
  $operatingSystemInfo = Get-CimInstance Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object csName, Caption, OperatingSystemSKU, Version, ProductType, OSType, LastBootUpTime, SerialNumber
  $currentVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" | Microsoft.PowerShell.Utility\Select-Object CurrentBuild, UBR, DisplayVersion

  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name CurrentBuild -Value $currentVersion.CurrentBuild
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name UpdateBuildRevision -Value $currentVersion.UBR
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name DisplayVersion -Value $currentVersion.DisplayVersion

  return $operatingSystemInfo
}

<#

.SYNOPSIS
Get the Win32_ComputerSystem information

.DESCRIPTION
Get the Win32_ComputerSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller.

#>
function getComputerSystemInfo() {
  return Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | `
    Microsoft.PowerShell.Utility\Select-Object TotalPhysicalMemory, DomainRole, Manufacturer, Model, NumberOfLogicalProcessors, Domain, Workgroup, DNSHostName, Name, PartOfDomain, SystemFamily, SystemSKUNumber
}

<#

.SYNOPSIS
Get SMBIOS locally from the passed in machineName


.DESCRIPTION
Get SMBIOS locally from the passed in machine name

#>
function getSmbiosData($computerSystem) {
  <#
    Array of chassis types.
    The following list of ChassisTypes is copied from the latest DMTF SMBIOS specification.
    REF: https://www.dmtf.org/sites/default/files/standards/documents/DSP0134_3.1.1.pdf
  #>
  $ChassisTypes =
  @{
    1  = 'Other'
    2  = 'Unknown'
    3  = 'Desktop'
    4  = 'Low Profile Desktop'
    5  = 'Pizza Box'
    6  = 'Mini Tower'
    7  = 'Tower'
    8  = 'Portable'
    9  = 'Laptop'
    10 = 'Notebook'
    11 = 'Hand Held'
    12 = 'Docking Station'
    13 = 'All in One'
    14 = 'Sub Notebook'
    15 = 'Space-Saving'
    16 = 'Lunch Box'
    17 = 'Main System Chassis'
    18 = 'Expansion Chassis'
    19 = 'SubChassis'
    20 = 'Bus Expansion Chassis'
    21 = 'Peripheral Chassis'
    22 = 'Storage Chassis'
    23 = 'Rack Mount Chassis'
    24 = 'Sealed-Case PC'
    25 = 'Multi-system chassis'
    26 = 'Compact PCI'
    27 = 'Advanced TCA'
    28 = 'Blade'
    29 = 'Blade Enclosure'
    30 = 'Tablet'
    31 = 'Convertible'
    32 = 'Detachable'
    33 = 'IoT Gateway'
    34 = 'Embedded PC'
    35 = 'Mini PC'
    36 = 'Stick PC'
  }

  $list = New-Object System.Collections.ArrayList
  $win32_Bios = Get-CimInstance -class Win32_Bios
  $obj = New-Object -Type PSObject | Microsoft.PowerShell.Utility\Select-Object SerialNumber, Manufacturer, UUID, BaseBoardProduct, ChassisTypes, Chassis, SystemFamily, SystemSKUNumber, SMBIOSAssetTag
  $obj.SerialNumber = $win32_Bios.SerialNumber
  $obj.Manufacturer = $win32_Bios.Manufacturer
  $computerSystemProduct = Get-CimInstance Win32_ComputerSystemProduct
  if ($null -ne $computerSystemProduct) {
    $obj.UUID = $computerSystemProduct.UUID
  }
  $baseboard = Get-CimInstance Win32_BaseBoard
  if ($null -ne $baseboard) {
    $obj.BaseBoardProduct = $baseboard.Product
  }
  $systemEnclosure = Get-CimInstance Win32_SystemEnclosure
  if ($null -ne $systemEnclosure) {
    $obj.SMBIOSAssetTag = $systemEnclosure.SMBIOSAssetTag
  }
  $obj.ChassisTypes = Get-CimInstance Win32_SystemEnclosure | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty ChassisTypes
  $obj.Chassis = New-Object -TypeName 'System.Collections.ArrayList'
  $obj.ChassisTypes | ForEach-Object -Process {
    $obj.Chassis.Add($ChassisTypes[[int]$_])
  }
  $obj.SystemFamily = $computerSystem.SystemFamily
  $obj.SystemSKUNumber = $computerSystem.SystemSKUNumber
  $list.Add($obj) | Out-Null

  return $list

}

<#

.SYNOPSIS
Checks if a system lockdown policy is enforced on this computer.

.DESCRIPTION
By checking the system lockdown policy, we can infer that PowerShell is in ConstrainedLanguage mode (WDAC).
Note: $ExecutionContext.SessionState.LanguageMode should not be used within a script context as in the case of allowed scripts (by the WDAC policy),
being executed locally, the language mode will always be FullLanguage and does NOT reflect the default system lockdown policy/language mode.

#>
function isSystemLockdownPolicyEnforced() {
  return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy() -eq [System.Management.Automation.Security.SystemEnforcementMode]::Enforce
}

<#

.SYNOPSIS
Determines if the operating system is HCI.

.DESCRIPTION
Using the operating system 'Caption' (which corresponds to the 'ProductName' registry key at HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion) to determine if a server OS is HCI.

#>
function isServerOsHCI([string] $operatingSystemCaption) {
  return $operatingSystemCaption -eq "Microsoft Azure Stack HCI"
}

###########################################################################
# main()
###########################################################################

$operatingSystem = getOperatingSystemInfo
$computerSystem = getComputerSystemInfo
$isAdministrator = isUserAnAdministrator
$fqdnAndAddress = getHostFqdnAndAddress $computerSystem
$hostname = [Environment]::MachineName
$netbios = $env:ComputerName
$managementToolsInformation = getManagementToolsSupportInformation
$isWmfInstalled = isWMF5Installed $operatingSystem.Version
$clusterInformation = getClusterInformation -ErrorAction SilentlyContinue
$isHyperVPowershellInstalled = isHyperVPowerShellSupportInstalled
$isHyperVRoleInstalled = isHyperVRoleInstalled
$isCredSSPEnabled = isCredSSPEnabled
$isRemoteAppEnabled = isRemoteAppEnabled
$smbiosData = getSmbiosData $computerSystem
$isSystemLockdownPolicyEnforced = isSystemLockdownPolicyEnforced
$isHciServer = isServerOsHCI $operatingSystem.Caption

$result = New-Object PSObject
$result | Add-Member -MemberType NoteProperty -Name 'IsAdministrator' -Value $isAdministrator
$result | Add-Member -MemberType NoteProperty -Name 'OperatingSystem' -Value $operatingSystem
$result | Add-Member -MemberType NoteProperty -Name 'ComputerSystem' -Value $computerSystem
$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $fqdnAndAddress.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $fqdnAndAddress.AddressList
$result | Add-Member -MemberType NoteProperty -Name 'Hostname' -Value $hostname
$result | Add-Member -MemberType NoteProperty -Name 'NetBios' -Value $netbios
$result | Add-Member -MemberType NoteProperty -Name 'IsManagementToolsAvailable' -Value $managementToolsInformation.ManagementToolsAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsServerManagerAvailable' -Value $managementToolsInformation.ServerManagerAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsWmfInstalled' -Value $isWmfInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCluster' -Value $clusterInformation.IsCluster
$result | Add-Member -MemberType NoteProperty -Name 'ClusterFqdn' -Value $clusterInformation.ClusterFqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2dEnabled' -Value $clusterInformation.IsS2dEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value $clusterInformation.IsBritannicaEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVRoleInstalled' -Value $isHyperVRoleInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVPowershellInstalled' -Value $isHyperVPowershellInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCredSSPEnabled' -Value $isCredSSPEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsRemoteAppEnabled' -Value $isRemoteAppEnabled
$result | Add-Member -MemberType NoteProperty -Name 'SmbiosData' -Value $smbiosData
$result | Add-Member -MemberType NoteProperty -Name 'IsSystemLockdownPolicyEnforced' -Value $isSystemLockdownPolicyEnforced
$result | Add-Member -MemberType NoteProperty -Name 'IsHciServer' -Value $isHciServer

$result

}
## [END] Get-WACSMSServerInventory ##
function Set-WACSMSEventLogChannelStatus {
 <#

.SYNOPSIS
 Change the current status (Enabled/Disabled) for the selected channel.

.DESCRIPTION
Change the current status (Enabled/Disabled) for the selected channel.
The supported Operating Systems are Window Server 2012, Windows Server 2012R2, Windows Server 2016.

.ROLE
Administrators

#>

Param(
    [string]$channel,
    [boolean]$status
)

$ch = Get-WinEvent -ListLog $channel
$ch.set_IsEnabled($status)
$ch.SaveChanges()
}
## [END] Set-WACSMSEventLogChannelStatus ##

# SIG # Begin signature block
# MIIoKgYJKoZIhvcNAQcCoIIoGzCCKBcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDvlfhmqYJr/KMp
# NP3cVUqQsDQpoZgHq+vJiJkXBb4a5KCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
# esGEb+srAAAAAANOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI5WhcNMjQwMzE0MTg0MzI5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDdCKiNI6IBFWuvJUmf6WdOJqZmIwYs5G7AJD5UbcL6tsC+EBPDbr36pFGo1bsU
# p53nRyFYnncoMg8FK0d8jLlw0lgexDDr7gicf2zOBFWqfv/nSLwzJFNP5W03DF/1
# 1oZ12rSFqGlm+O46cRjTDFBpMRCZZGddZlRBjivby0eI1VgTD1TvAdfBYQe82fhm
# WQkYR/lWmAK+vW/1+bO7jHaxXTNCxLIBW07F8PBjUcwFxxyfbe2mHB4h1L4U0Ofa
# +HX/aREQ7SqYZz59sXM2ySOfvYyIjnqSO80NGBaz5DvzIG88J0+BNhOu2jl6Dfcq
# jYQs1H/PMSQIK6E7lXDXSpXzAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUnMc7Zn/ukKBsBiWkwdNfsN5pdwAw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMDUxNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAD21v9pHoLdBSNlFAjmk
# mx4XxOZAPsVxxXbDyQv1+kGDe9XpgBnT1lXnx7JDpFMKBwAyIwdInmvhK9pGBa31
# TyeL3p7R2s0L8SABPPRJHAEk4NHpBXxHjm4TKjezAbSqqbgsy10Y7KApy+9UrKa2
# kGmsuASsk95PVm5vem7OmTs42vm0BJUU+JPQLg8Y/sdj3TtSfLYYZAaJwTAIgi7d
# hzn5hatLo7Dhz+4T+MrFd+6LUa2U3zr97QwzDthx+RP9/RZnur4inzSQsG5DCVIM
# pA1l2NWEA3KAca0tI2l6hQNYsaKL1kefdfHCrPxEry8onJjyGGv9YKoLv6AOO7Oh
# JEmbQlz/xksYG2N/JSOJ+QqYpGTEuYFYVWain7He6jgb41JbpOGKDdE/b+V2q/gX
# UgFe2gdwTpCDsvh8SMRoq1/BNXcr7iTAU38Vgr83iVtPYmFhZOVM0ULp/kKTVoir
# IpP2KCxT4OekOctt8grYnhJ16QMjmMv5o53hjNFXOxigkQWYzUO+6w50g0FAeFa8
# 5ugCCB6lXEk21FFB1FdIHpjSQf+LP/W2OV/HfhC3uTPgKbRtXo83TZYEudooyZ/A
# Vu08sibZ3MkGOJORLERNwKm2G7oqdOv4Qj8Z0JrGgMzj46NFKAxkLSpE5oHQYP1H
# tPx1lPfD7iNSbJsP6LiUHXH1MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIN6MDNaGEml+9VrdRXHxqQGD
# GQ3lNFWXRyua548c8SMHMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEA1JzPjfcoByebSA0gEzfKCDpxC06lbbckzRnelSWDaoI94od0RJiV6NMn
# xr5cJxta7T+d6qauAmm69dd0gmJnz7QJSVyxx6Cl6Su5CJbxicwFEXklHd/gXUnP
# 0+Ehs54eR4/lLanUrajbEo4c5AVB8FVzE5mGW70qhNbSBoXuyOWHo7cDFUQGjf0I
# PLZjhWzuI0Ou/jL2ndO19/mH6M7YPydE5cZcvGYUsrfdyeQHycmJ9ZllJoxQZz+v
# QWNvzEOhOz/RR39JjCDUGQ5CO8VxAS/vWd5M949DDPwcPqXAfwg//kvikbrtnict
# PDLIqUYYsg0HDvtHHZBqMOOpS8g81qGCF5QwgheQBgorBgEEAYI3AwMBMYIXgDCC
# F3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCAOITfe2ma/oFZ0AbSrwAOpq8GAp5YibTDWpg3OELAs4QIGZVbId61j
# GBMyMDIzMTIwNTE3NTQ0Mi45NTJaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OEQwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wg
# ghHqMIIHIDCCBQigAwIBAgITMwAAAc1VByrnysGZHQABAAABzTANBgkqhkiG9w0B
# AQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzA1MjUxOTEy
# MDVaFw0yNDAyMDExOTEyMDVaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OEQwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDTOCLVS2jmEWOqxzygW7s6YLmm29pjvA+Ch6VL7HlT
# L8yUt3Z0KIzTa2O/Hvr/aJza1qEVklq7NPiOrpBAIz657LVxwEc4BxJiv6B68a8D
# QiF6WAFFNaK3WHi7TfxRnqLohgNz7vZPylZQX795r8MQvX56uwjj/R4hXnR7Na4L
# lu4mWsml/wp6VJqCuxZnu9jX4qaUxngcrfFT7+zvlXClwLah2n0eGKna1dOjOgyK
# 00jYq5vtzr5NZ+qVxqaw9DmEsj9vfqYkfQZry2JO5wmgXX79Ox7PLMUfqT4+8w5J
# kdSMoX32b1D6cDKWRUv5qjiYh4o/a9ehE/KAkUWlSPbbDR/aGnPJLAGPy2qA97YC
# BeeIJjRKURgdPlhE5O46kOju8nYJnIvxbuC2Qp2jxwc6rD9M6Pvc8sZIcQ10YKZV
# YKs94YPSlkhwXwttbRY+jZnQiDm2ZFjH8SPe1I6ERcfeYX1zCYjEzdwWcm+fFZml
# JA9HQW7ZJAmOECONtfK28EREEE5yzq+T3QMVPhiEfEhgcYsh0DeoWiYGsDiKEuS+
# FElMMyT456+U2ZRa2hbRQ97QcbvaAd6OVQLp3TQqNEu0es5Zq0wg2CADf+QKQR/Y
# 6+fGgk9qJNJW3Mu771KthuPlNfKss0B1zh0xa1yN4qC3zoE9Uq6T8r7G3/OtSFms
# 4wIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFKGT+aY2aZrBAJVIZh5kicokfNWaMB8G
# A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4w
# XAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2Vy
# dHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBSqG3ppKIU+i/EMwwtotoxnKfw0SX/3T16
# EPbjwsAImWOZ5nLAbatopl8zFY841gb5eiL1j81h4DiEiXt+BJgHIA2LIhKhSscd
# 79oMbr631DiEqf9X5LZR3V3KIYstU3K7f5Dk7tbobuHu+6fYM/gOx44sgRU7YQ+Y
# TYHvv8k4mMnuiahJRlU/F2vavcHU5uhXi078K4nSRAPnWyX7gVi6iVMBBUF4823o
# PFznEcHup7VNGRtGe1xvnlMd1CuyxctM8d/oqyTsxwlJAM5F/lDxnEWoSzAkad1n
# WvkaAeMV7+39IpXhuf9G3xbffKiyBnj3cQeiA4SxSwCdnx00RBlXS6r9tGDa/o9R
# S01FOABzKkP5CBDpm4wpKdIU74KtBH2sE5QYYn7liYWZr2f/U+ghTmdOEOPkXEcX
# 81H4dRJU28Tj/gUZdwL81xah8Kn+cB7vM/Hs3/J8tF13ZPP+8NtX3vu4NrchHDJY
# gjOi+1JuSf+4jpF/pEEPXp9AusizmSmkBK4iVT7NwVtRnS1ts8qAGHGPg2HPa4b2
# u9meueUoqNVtMhbumI1y+d9ZkThNXBXz2aItT2C99DM3T3qYqAUmvKUryVSpMLVp
# se4je5WN6VVlCDFKWFRH202YxEVWsZ5baN9CaqCbCS0Ea7s9OFLaEM5fNn9m5s69
# lD/ekcW2qTCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZI
# hvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# MjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAy
# MDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25Phdg
# M/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aORmsHFPPF
# dvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6
# GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5LFGc6XBp
# Dco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50Zu
# yjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9QZpGdc3E
# XzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0
# lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqEUUbi0b1q
# GFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ
# +QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435UsSFF5PA
# PBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB3TCCAdkw
# EgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxG
# NSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwGA1UdIARV
# MFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAK
# BggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
# AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvX
# zpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYI
# KwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG
# 9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0x
# M7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYunKmC
# VgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRXud2f8449
# xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVCs/wM
# nosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0DLzskYDS
# PeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2d
# Y3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFrDZ+kKNxn
# GSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+Crvs
# QWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7ntdAoGokL
# jzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+oDEzfbzL
# 6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKhggNN
# MIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEn
# MCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjhEMDAtMDVFMC1EOTQ3MSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQBo
# qfem2KKzuRZjISYifGolVOdyBKCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6RmpfTAiGA8yMDIzMTIwNTEzNTE1
# N1oYDzIwMjMxMjA2MTM1MTU3WjB0MDoGCisGAQQBhFkKBAExLDAqMAoCBQDpGal9
# AgEAMAcCAQACAjymMAcCAQACAhJkMAoCBQDpGvr9AgEAMDYGCisGAQQBhFkKBAIx
# KDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZI
# hvcNAQELBQADggEBAGmIXWZYBAP58+QJLd15s+okwjgNJVI7JG6Jm81iJ/BmJnqe
# LAeU8CgnLzVHjuJIFQ9yd4bWSj+TY7AI5UsWVb3KyJxXkuKIZAgQ+72of9UJTbc2
# b04i182qq4kGYO5z5RWEH6quIT9zK8DpI3nhF24GCYolg6Ce0ayQN9+i0iynTsQq
# 2MMcA+ELNMU2yyQKFw8db4HAw0vMDV186SwdykZt6zoxqGbgIXY97cYrOMGJMsTx
# nW4tjCg7ccx0N5eDkMHTg2YagwMcNgnRL9uzWw4HzN4/NcbJlYnxiiI3Hj7KQMYp
# IJPKTge89QQVtj48Z0lMEqIH01tOgKUi0ZOc43ExggQNMIIECQIBATCBkzB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAc1VByrnysGZHQABAAABzTAN
# BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8G
# CSqGSIb3DQEJBDEiBCCQH3Qlu0l1j1I1xReVK7a7qqP2JepBH/qACKI97Y11LTCB
# +gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIOJmpfitVr1PZGgvTEdTpStUc6GN
# h7LNroQBKwpURpkKMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAHNVQcq58rBmR0AAQAAAc0wIgQgvrOENTnwIIO68K3fkUGRLdAvrt4/
# PB1hKeV8u9NRi5cwDQYJKoZIhvcNAQELBQAEggIAHkBYsYa9jyt8x+cwgloK4d/D
# HwPVFgjxiEJkI+st300XYmZ8QNe9CIdjjVlcTbLRmlBQhW4pwG0HMI51ZoncQ84H
# FkfX7AYbYMtADE32bPAosKC9LHcphUtiJZpn15/hLSL9VLx+NEY4QY94G2o/EdSv
# KPLXjnt6/NJ7WORYd0SrDpkpdIuapx0G+TvMFlFWk48H7ySQjzu+Nb5D9xk4298R
# QcJZbZRoq/9zDVn0NsImIxdq2zNWUc+LQo8JRBmTZlXTa+7cc6e5pilOcl2h5O4I
# 6ZP26h1JytAsBwf+nmVrETG/ar/7nTz5YXXVwwj8XxYJClAbaHPQ4Jghqr5XFoJU
# qx/UmSJT6Oi3FZBsfEimvo/XZSOSa9kZkK1JSkR7cJlZk6ssQvFXOpBmZl+ThLBp
# j/jcjdZ8VE1l8UiYPkgJo4D7hrGhtz1pAQ4aEKJoNYIFEyxKYOZSKQIuIyhFHQ4f
# k10FQCDEBQhvfJaJZ3o9frEf4O1KGZmj70BwmsMxniE2wEsaR7QoMLrCSau59NRA
# QYm5z6zAhF+SX6pxPvLojWfVT1ZAqm/luleqflkDiGKWYc0oWvdEp5G3VKgWCD0N
# /LIGqFV9pbeyoTv6XOKJMiZNBfhkHh2JZdky+IWacWiKw85O8mozknfsM6rDU+uJ
# N5jPUN/g5Y9Yq3iaYLo=
# SIG # End signature block
