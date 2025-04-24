function Disconnect-WACSMHybridManagement {
<#

.SYNOPSIS
Disconnects a machine from azure hybrid agent.

.DESCRIPTION
Disconnects a machine from azure hybrid agent and uninstall the hybrid instance service.
The supported Operating Systems are Windows Server 2012 R2 and above.

.ROLE
Administrators

.PARAMETER authToken
    The authentication token for connection

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $authToken
)

Set-StrictMode -Version 5.0

<#

.SYNOPSIS
Setup script runtime environment.

.DESCRIPTION
Setup script runtime environment.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScript" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Disconnect-HybridManagement.ps1" -Scope Script
    Set-Variable -Name HybridAgentExecutable -Option ReadOnly -Value "$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe" -Scope Script
    Set-Variable -Name HybridAgentConfigFile -Option ReadOnly -Value "$env:ProgramData\AzureConnectedMachineAgent\Config\agentconfig.json" -Scope Script
    Set-Variable -Name HybridAgentPackage -Option ReadOnly -Value "Azure Connected Machine Agent" -Scope Script
}

<#

.SYNOPSIS
Cleanup script runtime environment.

.DESCRIPTION
Cleanup script runtime environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name HybridAgentExecutable -Scope Script -Force
    Remove-Variable -Name HybridAgentConfigFile -Scope Script -Force
    Remove-Variable -Name HybridAgentPackage -Scope Script -Force
}

<#

.SYNOPSIS
The main function.

.DESCRIPTION
Disconnects a machine from azure hybrid agent.

#>

function main(
    [string]$authToken
) {
    $err = $null
    $args = @{}

   # Disconnect Azure hybrid agent
   if (Test-Path $HybridAgentExecutable) {
        & $HybridAgentExecutable disconnect --access-token $authToken
   }
   else {
        throw "Could not find the Azure hybrid agent executable file."
   }


   # Uninstall Azure hybrid instance metadata service
   Uninstall-Package -Name $HybridAgentPackage -ErrorAction SilentlyContinue -ErrorVariable +err

   if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
        -Message "[$ScriptName]:Could not uninstall the package. Error: $err"  -ErrorAction SilentlyContinue

        throw $err
   }

   # Remove Azure hybrid agent config file if it exists
   if (Test-Path $HybridAgentConfigFile) {
        Remove-Item -Path $HybridAgentConfigFile -ErrorAction SilentlyContinue -ErrorVariable +err -Force

        if ($err) {
            Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Could not remove the config file. Error: $err"  -ErrorAction SilentlyContinue

            throw $err
        }
   }
}


###############################################################################
# Script execution starts here
###############################################################################
setupScriptEnv

try {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return main $authToken
} finally {
    cleanupScriptEnv
}

}
## [END] Disconnect-WACSMHybridManagement ##
function Get-WACSMAntimalwareSoftwareStatus {
<#

.SYNOPSIS
Gets the status of antimalware software on the computer.

.DESCRIPTION
Gets the status of antimalware software on the computer.

.ROLE
Readers

#>

if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue)
{
    return (Get-MpComputerStatus -ErrorAction SilentlyContinue);
}
else{
    return $Null;
}


}
## [END] Get-WACSMAntimalwareSoftwareStatus ##
function Get-WACSMAzureProtectionStatus {
<#

.SYNOPSIS
Gets the status of Azure Backup on the target.

.DESCRIPTION
Checks whether azure backup is installed on target node, and is the machine protected by azure backup.
Returns the state of azure backup.

.ROLE
Readers

#>

Function Test-RegistryValue($path, $value) {
    if (Test-Path $path) {
        $Key = Get-Item -LiteralPath $path
        if ($Key.GetValue($value, $null) -ne $null) {
            $true
        }
        else {
            $false
        }
    }
    else {
        $false
    }
}

Set-StrictMode -Version 5.0
$path = 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment'
$value = 'PSModulePath'
if ((Test-RegistryValue $path $value) -eq $false) {
    @{ Registered = $false }
} else {
    $env:PSModulePath = (Get-ItemProperty -Path $path -Name PSModulePath).PSModulePath
    $AzureBackupModuleName = 'MSOnlineBackup'
    $DpmModuleName = 'DataProtectionManager'
    $DpmModule = Get-Module -ListAvailable -Name $DpmModuleName
    $AzureBackupModule = Get-Module -ListAvailable -Name $AzureBackupModuleName
    $IsAdmin = $false;

    $CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $IsAdmin = $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (!$IsAdmin) {
        @{ Registered = $false }
    }
    elseif ($DpmModule) {
        @{ Registered = $false }
    } 
    elseif ($AzureBackupModule) {
        try {
            Import-Module $AzureBackupModuleName
            $registrationstatus = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::GetMachineRegistrationStatus(0)
            if ($registrationstatus -eq $true) {
                @{ Registered = $true }
            }
            else {
                @{ Registered = $false }
            }
        }
        catch {
            @{ Registered = $false }
        }
    }
    else {
        @{ Registered = $false }
    }
}
}
## [END] Get-WACSMAzureProtectionStatus ##
function Get-WACSMAzureVMStatus {
<#

.SYNOPSIS
Checks whether a VM is from azure or not
.DESCRIPTION
Checks whether a VM is from azure or not
.ROLE
Readers

#>

$ErrorActionPreference="SilentlyContinue"

$uri = "http://169.254.169.254/metadata/instance?api-version=2021-02-01"
$Proxy=New-object System.Net.WebProxy
$WebSession=new-object Microsoft.PowerShell.Commands.WebRequestSession
$WebSession.Proxy=$Proxy
$result = Invoke-RestMethod -Headers @{"Metadata"="true"} -Method GET -Uri $uri -WebSession $WebSession

if ( $null -eq $result){
   return $false
}
 else {
    return $true
}


}
## [END] Get-WACSMAzureVMStatus ##
function Get-WACSMBmcInfo {
<#

.SYNOPSIS
Gets current information on the baseboard management controller (BMC).

.DESCRIPTION
Gets information such as manufacturer, serial number, last known IP
address, model, and network configuration to show to user.

.ROLE
Readers

#>

Import-Module CimCmdlets
Import-Module PcsvDevice

$error.Clear()

$bmcInfo = Get-PcsvDevice -ErrorAction SilentlyContinue

$bmcAlternateInfo = Get-CimInstance Win32_Bios -ErrorAction SilentlyContinue
$serialNumber = $bmcInfo.SerialNumber

if ($bmcInfo -and $bmcAlternateInfo) {
    $serialNumber = -join($bmcInfo.SerialNumber, " / ", $bmcAlternateInfo.SerialNumber)
}

$result = New-Object -TypeName PSObject
$result | Add-Member -MemberType NoteProperty -Name "Error" $error.Count

if ($error.Count -EQ 0) {
    $result | Add-Member -MemberType NoteProperty -Name "Ip" $bmcInfo.IPv4Address
    $result | Add-Member -MemberType NoteProperty -Name "Serial" $serialNumber
}

$result

}
## [END] Get-WACSMBmcInfo ##
function Get-WACSMCimDiskRegistry {
<#

.SYNOPSIS
Get Disk Registry status by using ManagementTools CIM provider.

.DESCRIPTION
Get Disk Registry status by using ManagementTools CIM provider.

.ROLE
Readers

#>

##SkipCheck=true##

Param(
[string]$Name
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/Microsoft/Windows/ManagementTools -ClassName Msft_MTRegistryKey -Key @('Name') -Property @{Name=$Name;} -ClientOnly
Invoke-CimMethod $keyInstance -MethodName GetValues

}
## [END] Get-WACSMCimDiskRegistry ##
function Get-WACSMCimDiskSummary {
<#

.SYNOPSIS
Get Disk summary by using ManagementTools CIM provider.

.DESCRIPTION
Get Disk summary by using ManagementTools CIM provider.

.ROLE
Readers

#>

##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/Microsoft/Windows/ManagementTools -ClassName Msft_MTDisk

}
## [END] Get-WACSMCimDiskSummary ##
function Get-WACSMCimMemorySummary {
<#

.SYNOPSIS
Get Memory summary by using ManagementTools CIM provider.

.DESCRIPTION
Get Memory summary by using ManagementTools CIM provider.

.ROLE
Readers

#>

##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/Microsoft/Windows/ManagementTools -ClassName Msft_MTMemorySummary

}
## [END] Get-WACSMCimMemorySummary ##
function Get-WACSMCimNetworkAdapterSummary {
<#

.SYNOPSIS
Get Network Adapter summary by using ManagementTools CIM provider.

.DESCRIPTION
Get Network Adapter summary by using ManagementTools CIM provider.

.ROLE
Readers

#>

##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/Microsoft/Windows/ManagementTools -ClassName Msft_MTNetworkAdapter

}
## [END] Get-WACSMCimNetworkAdapterSummary ##
function Get-WACSMCimProcessorSummary {
<#

.SYNOPSIS
Get Processor summary by using ManagementTools CIM provider.

.DESCRIPTION
Get Processor summary by using ManagementTools CIM provider.

.ROLE
Readers

#>

##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/Microsoft/Windows/ManagementTools -ClassName Msft_MTProcessorSummary

}
## [END] Get-WACSMCimProcessorSummary ##
function Get-WACSMClientConnectionStatus {
<#

.SYNOPSIS
Gets status of the connection to the client computer.

.DESCRIPTION
Gets status of the connection to the client computer.

.ROLE
Readers

#>

import-module CimCmdlets
$OperatingSystem = Get-CimInstance Win32_OperatingSystem
$Caption = $OperatingSystem.Caption
$ProductType = $OperatingSystem.ProductType
$Version = $OperatingSystem.Version
$Status = @{ Label = $null; Type = 0; Details = $null; }
$Result = @{ Status = $Status; Caption = $Caption; ProductType = $ProductType; Version = $Version; }

if ($Version -and $ProductType -eq 1) {
    $V = [version]$Version
    $V10 = [version]'10.0'
    if ($V -ge $V10) {
        return $Result;
    } 
}

$Status.Label = 'unsupported-label'
$Status.Type = 3
$Status.Details = 'unsupported-details'
return $Result;

}
## [END] Get-WACSMClientConnectionStatus ##
function Get-WACSMClusterInformation {
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
## [END] Get-WACSMClusterInformation ##
function Get-WACSMComputerIdentification {
<#

.SYNOPSIS
Gets the local computer domain/workplace information.

.DESCRIPTION
Gets the local computer domain/workplace information.
Returns the computer identification information.

.ROLE
Readers

#>

import-module CimCmdlets

$ComputerSystem = Get-CimInstance -Class Win32_ComputerSystem;
$ComputerName = $ComputerSystem.DNSHostName
if ($ComputerName -eq $null) {
    $ComputerName = $ComputerSystem.Name
}

$fqdn = ([System.Net.Dns]::GetHostByName($ComputerName)).HostName

$ComputerSystem | Microsoft.PowerShell.Utility\Select-Object `
@{ Name = "ComputerName"; Expression = { $ComputerName }},
@{ Name = "Domain"; Expression = { if ($_.PartOfDomain) { $_.Domain } else { $null } }},
@{ Name = "DomainJoined"; Expression = { $_.PartOfDomain }},
@{ Name = "FullComputerName"; Expression = { $fqdn }},
@{ Name = "Workgroup"; Expression = { if ($_.PartOfDomain) { $null } else { $_.Workgroup } }}


}
## [END] Get-WACSMComputerIdentification ##
function Get-WACSMDiskSummaryDownlevel {
<#

.SYNOPSIS
Gets disk summary information by performance counter WMI object on downlevel computer.

.DESCRIPTION
Gets disk summary information by performance counter WMI object on downlevel computer.

.ROLE
Readers

#>

param
(
)

import-module CimCmdlets

function ResetDiskData($diskResults) {
    $Global:DiskResults = @{}
    $Global:DiskDelta = 0

    foreach ($item in $diskResults) {
        $diskRead = New-Object System.Collections.ArrayList
        $diskWrite = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt 60; $i++) {
            $diskRead.Insert(0, 0)
            $diskWrite.Insert(0, 0)
        }

        $Global:DiskResults.Item($item.name) = @{
            ReadTransferRate  = $diskRead
            WriteTransferRate = $diskWrite
        }
    }
}

function UpdateDiskData($diskResults) {
    $Global:DiskDelta += ($Global:DiskSampleTime - $Global:DiskLastTime).TotalMilliseconds

    foreach ($diskResult in $diskResults) {
        $localDelta = $Global:DiskDelta

        # update data for each disk
        $item = $Global:DiskResults.Item($diskResult.name)

        if ($item -ne $null) {
            while ($localDelta -gt 1000) {
                $localDelta -= 1000
                $item.ReadTransferRate.Insert(0, $diskResult.DiskReadBytesPersec)
                $item.WriteTransferRate.Insert(0, $diskResult.DiskWriteBytesPersec)
            }

            $item.ReadTransferRate = $item.ReadTransferRate.GetRange(0, 60)
            $item.WriteTransferRate = $item.WriteTransferRate.GetRange(0, 60)

            $Global:DiskResults.Item($diskResult.name) = $item
        }
    }

    $Global:DiskDelta = $localDelta
}

$counterValue = Get-CimInstance win32_perfFormattedData_PerfDisk_PhysicalDisk -Filter "name!='_Total'" | Microsoft.PowerShell.Utility\Select-Object name, DiskReadBytesPersec, DiskWriteBytesPersec
$now = get-date

# get sampling time and remember last sample time.
if (-not $Global:DiskSampleTime) {
    $Global:DiskSampleTime = $now
    $Global:DiskLastTime = $Global:DiskSampleTime
    ResetDiskData($counterValue)
}
else {
    $Global:DiskLastTime = $Global:DiskSampleTime
    $Global:DiskSampleTime = $now
    if ($Global:DiskSampleTime - $Global:DiskLastTime -gt [System.TimeSpan]::FromSeconds(30)) {
        ResetDiskData($counterValue)
    }
    else {
        UpdateDiskData($counterValue)
    }
}

$Global:DiskResults
}
## [END] Get-WACSMDiskSummaryDownlevel ##
function Get-WACSMEnvironmentVariables {
<#

.SYNOPSIS
Gets 'Machine' and 'User' environment variables.

.DESCRIPTION
Gets 'Machine' and 'User' environment variables.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

$data = @()

$system = [Environment]::GetEnvironmentVariables([EnvironmentVariableTarget]::Machine)
$user = [Environment]::GetEnvironmentVariables([EnvironmentVariableTarget]::User)

foreach ($h in $system.GetEnumerator()) {
    $obj = @{"Name" = $h.Name; "Value" = $h.Value; "Type" = "Machine"}
    $data += $obj
}

foreach ($h in $user.GetEnumerator()) {
    $obj = @{"Name" = $h.Name; "Value" = $h.Value; "Type" = "User"}
    $data += $obj
}

$data
}
## [END] Get-WACSMEnvironmentVariables ##
function Get-WACSMHybridManagementConfiguration {
<#

.SYNOPSIS
Script that return the hybrid management configurations.

.DESCRIPTION
Script that return the hybrid management configurations.

.ROLE
Administrators

#>

Set-StrictMode -Version 5.0
Import-Module Microsoft.PowerShell.Management

<#

.SYNOPSIS
Setup script runtime environment.

.DESCRIPTION
Onboards a machine for hybrid management.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScript" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Get-HybridManagementConfiguration.ps1" -Scope Script
    Set-Variable -Name HybridAgentExecutable -Option ReadOnly -Value "$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe" -Scope Script
}

<#

.SYNOPSIS
Cleanup script runtime environment.

.DESCRIPTION
Cleanup script runtime environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name HybridAgentExecutable -Scope Script -Force
}

function main() {
    $config = & $HybridAgentExecutable -j show

    if ($config) {
        $configObj = $config | ConvertFrom-Json
        @{
            machine = $configObj.resourceName;
            resourceGroup = $configObj.resourceGroup;
            subscriptionId = $configObj.subscriptionId;
            tenantId = $configObj.tenantId;
            vmId = $configObj.vmId;
            azureRegion = $configObj.location;
            agentVersion = $configObj.agentVersion;
            agentStatus = $configObj.status;
            agentLastHeartbeat = $configObj.lastHeartbeat;
        }
    } else {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
        -Message "[$ScriptName]:Could not find the Azure hybrid agent configuration."  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }
}

function getValue([string]$keyValue) {
    $splitArray = $keyValue -split ":"
    $value = $splitArray[1].trim()
    return $value
}

###############################################################################
# Script execution starts here
###############################################################################
setupScriptEnv

try {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return main

} finally {
    cleanupScriptEnv
}
}
## [END] Get-WACSMHybridManagementConfiguration ##
function Get-WACSMHybridManagementStatus {
<#

.SYNOPSIS
Script that returns if Azure Hybrid Agent is running or not.

.DESCRIPTION
Script that returns if Azure Hybrid Agent is running or not.

.ROLE
Readers

#>

Import-Module Microsoft.PowerShell.Management

$status = Get-Service -Name himds -ErrorAction SilentlyContinue
if ($null -eq $status) {
    # which means no such service is found.
    @{ Installed = $false; Running = $false }
}
elseif ($status.Status -eq "Running") {
    @{ Installed = $true; Running = $true }
}
else {
    @{ Installed = $true; Running = $false }
}

}
## [END] Get-WACSMHybridManagementStatus ##
function Get-WACSMHyperVEnhancedSessionModeSettings {
<#

.SYNOPSIS
Gets a computer's Hyper-V Host Enhanced Session Mode settings.

.DESCRIPTION
Gets a computer's Hyper-V Host Enhnaced Session Mode settings.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
Import-Module Hyper-V

Get-VMHost | Microsoft.PowerShell.Utility\Select-Object `
    EnableEnhancedSessionMode

}
## [END] Get-WACSMHyperVEnhancedSessionModeSettings ##
function Get-WACSMHyperVGeneralSettings {
<#

.SYNOPSIS
Gets a computer's Hyper-V Host General settings.

.DESCRIPTION
Gets a computer's Hyper-V Host General settings.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
Import-Module Hyper-V

Get-VMHost | Microsoft.PowerShell.Utility\Select-Object `
    VirtualHardDiskPath, `
    VirtualMachinePath

}
## [END] Get-WACSMHyperVGeneralSettings ##
function Get-WACSMHyperVHostPhysicalGpuSettings {
<#

.SYNOPSIS
Gets a computer's Hyper-V Host Physical GPU settings.

.DESCRIPTION
Gets a computer's Hyper-V Host Physical GPU settings.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
Import-Module CimCmdlets

Get-CimInstance -Namespace "root\virtualization\v2" -Class "Msvm_Physical3dGraphicsProcessor" | `
    Microsoft.PowerShell.Utility\Select-Object EnabledForVirtualization, `
    Name, `
    DriverDate, `
    DriverInstalled, `
    DriverModelVersion, `
    DriverProvider, `
    DriverVersion, `
    DirectXVersion, `
    PixelShaderVersion, `
    DedicatedVideoMemory, `
    DedicatedSystemMemory, `
    SharedSystemMemory, `
    TotalVideoMemory

}
## [END] Get-WACSMHyperVHostPhysicalGpuSettings ##
function Get-WACSMHyperVLiveMigrationSettings {
<#

.SYNOPSIS
Gets a computer's Hyper-V Host Live Migration settings.

.DESCRIPTION
Gets a computer's Hyper-V Host Live Migration settings.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
Import-Module Hyper-V

Get-VMHost | Microsoft.PowerShell.Utility\Select-Object `
    maximumVirtualMachineMigrations, `
    VirtualMachineMigrationAuthenticationType, `
    VirtualMachineMigrationEnabled, `
    VirtualMachineMigrationPerformanceOption

}
## [END] Get-WACSMHyperVLiveMigrationSettings ##
function Get-WACSMHyperVMigrationSupport {
<#

.SYNOPSIS
Gets a computer's Hyper-V migration support.

.DESCRIPTION
Gets a computer's Hyper-V  migration support.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

$migrationSettingsDatas=Microsoft.PowerShell.Management\Get-WmiObject -Namespace root\virtualization\v2 -Query "associators of {Msvm_VirtualSystemMigrationCapabilities.InstanceID=""Microsoft:MigrationCapabilities""} where resultclass = Msvm_VirtualSystemMigrationSettingData"

$live = $false;
$storage = $false;

foreach ($migrationSettingsData in $migrationSettingsDatas) {
    if ($migrationSettingsData.MigrationType -eq 32768) {
        $live = $true;
    }

    if ($migrationSettingsData.MigrationType -eq 32769) {
        $storage = $true;
    }
}

$result = New-Object -TypeName PSObject
$result | Add-Member -MemberType NoteProperty -Name "liveMigrationSupported" $live;
$result | Add-Member -MemberType NoteProperty -Name "storageMigrationSupported" $storage;
$result
}
## [END] Get-WACSMHyperVMigrationSupport ##
function Get-WACSMHyperVNumaSpanningSettings {
<#

.SYNOPSIS
Gets a computer's Hyper-V Host settings.

.DESCRIPTION
Gets a computer's Hyper-V Host settings.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
Import-Module Hyper-V

Get-VMHost | Microsoft.PowerShell.Utility\Select-Object `
    NumaSpanningEnabled

}
## [END] Get-WACSMHyperVNumaSpanningSettings ##
function Get-WACSMHyperVRoleInstalled {
<#

.SYNOPSIS
Gets a computer's Hyper-V role installation state.

.DESCRIPTION
Gets a computer's Hyper-V role installation state.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
 
$service = Microsoft.PowerShell.Management\get-service -Name "VMMS" -ErrorAction SilentlyContinue;

return ($service -and $service.Name -eq "VMMS");

}
## [END] Get-WACSMHyperVRoleInstalled ##
function Get-WACSMHyperVStorageMigrationSettings {
<#

.SYNOPSIS
Gets a computer's Hyper-V Host settings.

.DESCRIPTION
Gets a computer's Hyper-V Host settings.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
Import-Module Hyper-V

Get-VMHost | Microsoft.PowerShell.Utility\Select-Object `
    MaximumStorageMigrations

}
## [END] Get-WACSMHyperVStorageMigrationSettings ##
function Get-WACSMLicenseStatusChecks {
<#

.SYNOPSIS
Does the license checks for a server

.DESCRIPTION
Does the license checks for a server

.ROLE
Readers

#>

param (
    [Parameter(Mandatory = $true)]
    [string]
    $applicationId
)

Import-Module CimCmdlets

function Get-LicenseStatus() {
  # LicenseStatus check
  $cim = Get-CimInstance -ClassName SoftwareLicensingProduct | Where-Object { $_.ProductKeyID  -and  $_.ApplicationID -eq $applicationId }
  try {
    $licenseStatus = $cim.LicenseStatus;
  }
  catch {
    $LicenseStatus = $null;
  }

  return $LicenseStatus;
}

function Get-SoftwareLicensingService() {
  $cim = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue

  # Without the trycf it fails with the error:
  # The property 'AzureMetadataResponse' cannot be found on this object. Verify that the property exists.
  try {
    $azureMetadataResponse = $cim.AzureMetadataResponse
  }
  catch {
    $azureMetadataResponse = $null
  }

  return $azureMetadataResponse;
}


$result = New-Object PSObject
$result | Add-Member -MemberType NoteProperty -Name "LicenseStatus" -Value (Get-LicenseStatus)
$result | Add-Member -MemberType NoteProperty -Name "AzureMetadataResponse" -Value (Get-SoftwareLicensingService)

$result

}
## [END] Get-WACSMLicenseStatusChecks ##
function Get-WACSMMemorySummaryDownLevel {
<#

.SYNOPSIS
Gets memory summary information by performance counter WMI object on downlevel computer.

.DESCRIPTION
Gets memory summary information by performance counter WMI object on downlevel computer.

.ROLE
Readers

#>

import-module CimCmdlets

# reset counter reading only first one.
function Reset($counter) {
    $Global:Utilization = [System.Collections.ArrayList]@()
    for ($i = 0; $i -lt 59; $i++) {
        $Global:Utilization.Insert(0, 0)
    }

    $Global:Utilization.Insert(0, $counter)
    $Global:Delta = 0
}

$memory = Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
$now = get-date
$system = Get-CimInstance Win32_ComputerSystem
$percent = 100 * ($system.TotalPhysicalMemory - $memory.AvailableBytes) / $system.TotalPhysicalMemory
$cached = $memory.StandbyCacheCoreBytes + $memory.StandbyCacheNormalPriorityBytes + $memory.StandbyCacheReserveBytes + $memory.ModifiedPageListBytes

# get sampling time and remember last sample time.
if (-not $Global:SampleTime) {
    $Global:SampleTime = $now
    $Global:LastTime = $Global:SampleTime
    Reset($percent)
}
else {
    $Global:LastTime = $Global:SampleTime
    $Global:SampleTime = $now
    if ($Global:SampleTime - $Global:LastTime -gt [System.TimeSpan]::FromSeconds(30)) {
        Reset($percent)
    }
    else {
        $Global:Delta += ($Global:SampleTime - $Global:LastTime).TotalMilliseconds
        while ($Global:Delta -gt 1000) {
            $Global:Delta -= 1000
            $Global:Utilization.Insert(0, $percent)
        }

        $Global:Utilization = $Global:Utilization.GetRange(0, 60)
    }
}

$result = New-Object -TypeName PSObject
$result | Add-Member -MemberType NoteProperty -Name "Available" $memory.AvailableBytes
$result | Add-Member -MemberType NoteProperty -Name "Cached" $cached
$result | Add-Member -MemberType NoteProperty -Name "Total" $system.TotalPhysicalMemory
$result | Add-Member -MemberType NoteProperty -Name "InUse" ($system.TotalPhysicalMemory - $memory.AvailableBytes)
$result | Add-Member -MemberType NoteProperty -Name "Committed" $memory.CommittedBytes
$result | Add-Member -MemberType NoteProperty -Name "PagedPool" $memory.PoolPagedBytes
$result | Add-Member -MemberType NoteProperty -Name "NonPagedPool" $memory.PoolNonpagedBytes
$result | Add-Member -MemberType NoteProperty -Name "Utilization" $Global:Utilization
$result
}
## [END] Get-WACSMMemorySummaryDownLevel ##
function Get-WACSMMmaStatus {
<#

.SYNOPSIS
Script that returns if Microsoft Monitoring Agent is running or not.

.DESCRIPTION
Script that returns if Microsoft Monitoring Agent is running or not.

.ROLE
Readers

#>

Import-Module Microsoft.PowerShell.Management

$MMAStatus = Get-Service -Name HealthService -ErrorAction SilentlyContinue
if ($null -eq $MMAStatus) {
  # which means no such service is found.
  return @{ Installed = $false; Running = $false;}
}

$IsAgentRunning = $MMAStatus.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running
$ServiceMapAgentStatus = Get-Service -Name MicrosoftDependencyAgent -ErrorAction SilentlyContinue
$IsServiceMapAgentInstalled = $null -ne $ServiceMapAgentStatus -and $ServiceMapAgentStatus.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running

$AgentConfig = New-Object -ComObject 'AgentConfigManager.mgmtsvccfg'
$WorkSpaces = @($AgentConfig.GetCloudWorkspaces() | Microsoft.PowerShell.Utility\Select-Object -Property WorkspaceId, AgentId)

return @{
  Installed                     = $true;
  Running                       = $IsAgentRunning;
  IsServiceMapAgentInstalled    = $IsServiceMapAgentInstalled
  WorkSpaces                    = $WorkSpaces
}

}
## [END] Get-WACSMMmaStatus ##
function Get-WACSMNetworkSummaryDownlevel {
<#

.SYNOPSIS
Gets network adapter summary information by performance counter WMI object on downlevel computer.

.DESCRIPTION
Gets network adapter summary information by performance counter WMI object on downlevel computer.

.ROLE
Readers

#>

import-module CimCmdlets
function ResetData($adapterResults) {
    $Global:NetworkResults = @{}
    $Global:PrevAdapterData = @{}
    $Global:Delta = 0

    foreach ($key in $adapterResults.Keys) {
        $adapterResult = $adapterResults.Item($key)
        $sentBytes = New-Object System.Collections.ArrayList
        $receivedBytes = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt 60; $i++) {
            $sentBytes.Insert(0, 0)
            $receivedBytes.Insert(0, 0)
        }

        $networkResult = @{
            SentBytes = $sentBytes
            ReceivedBytes = $receivedBytes
        }
        $Global:NetworkResults.Item($key) = $networkResult
    }
}

function UpdateData($adapterResults) {
    $Global:Delta += ($Global:SampleTime - $Global:LastTime).TotalMilliseconds

    foreach ($key in $adapterResults.Keys) {
        $localDelta = $Global:Delta

        # update data for each adapter
        $adapterResult = $adapterResults.Item($key)
        $item = $Global:NetworkResults.Item($key)
        if ($item -ne $null) {
            while ($localDelta -gt 1000) {
                $localDelta -= 1000
                $item.SentBytes.Insert(0, $adapterResult.SentBytes)
                $item.ReceivedBytes.Insert(0, $adapterResult.ReceivedBytes)
            }

            $item.SentBytes = $item.SentBytes.GetRange(0, 60)
            $item.ReceivedBytes = $item.ReceivedBytes.GetRange(0, 60)

            $Global:NetworkResults.Item($key) = $item
        }
    }

    $Global:Delta = $localDelta
}

$adapters = Get-CimInstance -Namespace root/standardCimV2 MSFT_NetAdapter | Where-Object MediaConnectState -eq 1 | Microsoft.PowerShell.Utility\Select-Object Name, InterfaceIndex, InterfaceDescription
$activeAddresses = get-CimInstance -Namespace root/standardCimV2 MSFT_NetIPAddress | Microsoft.PowerShell.Utility\Select-Object interfaceIndex

$adapterResults = @{}
foreach ($adapter in $adapters) {
    foreach ($activeAddress in $activeAddresses) {
        # Find a match between the 2
        if ($adapter.InterfaceIndex -eq $activeAddress.interfaceIndex) {
            $description = $adapter | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty interfaceDescription

            if ($Global:UsePerfData -EQ $NULL) {
                $adapterData = Get-CimInstance -Namespace root/StandardCimv2 MSFT_NetAdapterStatisticsSettingData -Filter "Description='$description'" | Microsoft.PowerShell.Utility\Select-Object ReceivedBytes, SentBytes

                if ($adapterData -EQ $null) {
                    # If above doesnt return data use slower perf data below
                    $Global:UsePerfData = $true
                }
            }

            if ($Global:UsePerfData -EQ $true) {
                # Need to replace the '#' to ascii since we parse anything after # as a comment
                $sanitizedDescription = $description -replace [char]35, "_"
                $adapterData = Get-CimInstance Win32_PerfFormattedData_Tcpip_NetworkAdapter | Where-Object name -EQ $sanitizedDescription | Microsoft.PowerShell.Utility\Select-Object BytesSentPersec, BytesReceivedPersec

                $sentBytes = $adapterData.BytesSentPersec
                $receivedBytes = $adapterData.BytesReceivedPersec
            }
            else {
                # set to 0 because we dont have a baseline to subtract from
                $sentBytes = 0
                $receivedBytes = 0

                if ($Global:PrevAdapterData -ne $null) {
                    $prevData = $Global:PrevAdapterData.Item($description)
                    if ($prevData -ne $null) {
                        $sentBytes = $adapterData.SentBytes - $prevData.SentBytes
                        $receivedBytes = $adapterData.ReceivedBytes - $prevData.ReceivedBytes
                    }
                }
                else {
                    $Global:PrevAdapterData = @{}
                }

                # Now that we have data, set current data as previous data as baseline
                $Global:PrevAdapterData.Item($description) = $adapterData
            }

            $adapterResult = @{
                SentBytes = $sentBytes
                ReceivedBytes = $receivedBytes
            }
            $adapterResults.Item($description) = $adapterResult
            break;
        }
    }
}

$now = get-date

if (-not $Global:SampleTime) {
    $Global:SampleTime = $now
    $Global:LastTime = $Global:SampleTime
    ResetData($adapterResults)
}
else {
    $Global:LastTime = $Global:SampleTime
    $Global:SampleTime = $now
    if ($Global:SampleTime - $Global:LastTime -gt [System.TimeSpan]::FromSeconds(30)) {
        ResetData($adapterResults)
    }
    else {
        UpdateData($adapterResults)
    }
}

$Global:NetworkResults
}
## [END] Get-WACSMNetworkSummaryDownlevel ##
function Get-WACSMNumberOfLoggedOnUsers {
<#

.SYNOPSIS
Gets the number of logged on users.

.DESCRIPTION
Gets the number of logged on users including active and disconnected users.
Returns a count of users.

.ROLE
Readers

#>

$error.Clear()

# Use Process class to hide exe prompt when executing
$process = New-Object System.Diagnostics.Process
$process.StartInfo.FileName = "quser.exe"
$process.StartInfo.UseShellExecute = $false
$process.StartInfo.CreateNoWindow = $true
$process.StartInfo.RedirectStandardOutput = $true 
$process.StartInfo.RedirectStandardError = $true
$process.Start() | Out-Null 
$process.WaitForExit()

$result = @()
while ($line = $process.StandardOutput.ReadLine()) {
    $result += $line 
}

if ($process.StandardError.EndOfStream) {
    # quser does not return a valid ps object and includes the header.
    # subtract 1 to get actual count.
    $count = $result.count - 1
} else {
    # there is an error to get result. Set to 0 instead of -1 currently
    $count = 0
}

$process.Dispose()

@{ Count = $count }
}
## [END] Get-WACSMNumberOfLoggedOnUsers ##
function Get-WACSMPowerConfigurationPlan {
<#

.SYNOPSIS
Gets the power plans on the machine.

.DESCRIPTION
Gets the power plans on the machine.

.ROLE
Readers

#>

$GuidLength = 36
$plans = Get-CimInstance -Namespace root\cimv2\power -ClassName Win32_PowerPlan

if ($plans) {
  $result = New-Object 'System.Collections.Generic.List[System.Object]'

  foreach ($plan in $plans) {
    $currentPlan = New-Object -TypeName PSObject

    $currentPlan | Add-Member -MemberType NoteProperty -Name 'Name' -Value $plan.ElementName
    $currentPlan | Add-Member -MemberType NoteProperty -Name 'DisplayName' -Value $plan.ElementName
    $currentPlan | Add-Member -MemberType NoteProperty -Name 'IsActive' -Value $plan.IsActive
    $startBrace = $plan.InstanceID.IndexOf("{")
    $currentPlan | Add-Member -MemberType NoteProperty -Name 'Guid' -Value $plan.InstanceID.SubString($startBrace + 1, $GuidLength)

    $result.Add($currentPlan)
  }

  return $result.ToArray()
}

return $null

}
## [END] Get-WACSMPowerConfigurationPlan ##
function Get-WACSMProcessorSummaryDownlevel {
<#

.SYNOPSIS
Gets processor summary information by performance counter WMI object on downlevel computer.

.DESCRIPTION
Gets processor summary information by performance counter WMI object on downlevel computer.

.ROLE
Readers

#>

import-module CimCmdlets

# reset counter reading only first one.
function Reset($counter) {
    $Global:Utilization = [System.Collections.ArrayList]@()
    for ($i = 0; $i -lt 59; $i++) {
        $Global:Utilization.Insert(0, 0)
    }

    $Global:Utilization.Insert(0, $counter)
    $Global:Delta = 0
}

$processorCounter = Get-CimInstance Win32_PerfFormattedData_Counters_ProcessorInformation -Filter "name='_Total'"
$now = get-date
$processor = Get-CimInstance Win32_Processor
$os = Get-CimInstance Win32_OperatingSystem
$processes = Get-CimInstance Win32_Process
$percent = $processorCounter.PercentProcessorTime
$handles = 0
$threads = 0
$processes | ForEach-Object { $handles += $_.HandleCount; $threads += $_.ThreadCount }
$uptime = ($now - $os.LastBootUpTime).TotalMilliseconds * 10000

# get sampling time and remember last sample time.
if (-not $Global:SampleTime) {
    $Global:SampleTime = $now
    $Global:LastTime = $Global:SampleTime
    Reset($percent)
}
else {
    $Global:LastTime = $Global:SampleTime
    $Global:SampleTime = $now
    if ($Global:SampleTime - $Global:LastTime -gt [System.TimeSpan]::FromSeconds(30)) {
        Reset($percent)
    }
    else {
        $Global:Delta += ($Global:SampleTime - $Global:LastTime).TotalMilliseconds
        while ($Global:Delta -gt 1000) {
            $Global:Delta -= 1000
            $Global:Utilization.Insert(0, $percent)
        }

        $Global:Utilization = $Global:Utilization.GetRange(0, 60)
    }
}

$result = New-Object -TypeName PSObject
$result | Add-Member -MemberType NoteProperty -Name "Name" $processor[0].Name
$result | Add-Member -MemberType NoteProperty -Name "AverageSpeed" ($processor[0].CurrentClockSpeed / 1000)
$result | Add-Member -MemberType NoteProperty -Name "Processes" $processes.Length
$result | Add-Member -MemberType NoteProperty -Name "Uptime" $uptime
$result | Add-Member -MemberType NoteProperty -Name "Handles" $handles
$result | Add-Member -MemberType NoteProperty -Name "Threads" $threads
$result | Add-Member -MemberType NoteProperty -Name "Utilization" $Global:Utilization
$result
}
## [END] Get-WACSMProcessorSummaryDownlevel ##
function Get-WACSMRbacEnabled {
<#

.SYNOPSIS
Gets the state of the Get-PSSessionConfiguration command

.DESCRIPTION
Gets the state of the Get-PSSessionConfiguration command

.ROLE
Readers

#>

if ($null -ne (Get-Command Get-PSSessionConfiguration -ErrorAction SilentlyContinue)) {
  @{ State = 'Available' }
} else {
  @{ State = 'NotSupported' }
}

}
## [END] Get-WACSMRbacEnabled ##
function Get-WACSMRbacSessionConfiguration {
<#

.SYNOPSIS
Gets a Microsoft.Sme.PowerShell endpoint configuration.

.DESCRIPTION
Gets a Microsoft.Sme.PowerShell endpoint configuration.

.ROLE
Administrators

#>

param(
    [Parameter(Mandatory = $false)]
    [String]
    $configurationName = "Microsoft.Sme.PowerShell"
)

## check if it's full administrators
if ((Get-Command Get-PSSessionConfiguration -ErrorAction SilentlyContinue) -ne $null) {
    @{
        Administrators = $true
        Configured = (Get-PSSessionConfiguration $configurationName -ErrorAction SilentlyContinue) -ne $null
    }
} else {
    @{
        Administrators = $false
        Configured = $false
    }
}
}
## [END] Get-WACSMRbacSessionConfiguration ##
function Get-WACSMRebootPendingStatus {
<#

.SYNOPSIS
Gets information about the server pending reboot.

.DESCRIPTION
Gets information about the server pending reboot.

.ROLE
Readers

#>

import-module CimCmdlets

function Get-ComputerNameChangeStatus {
    $currentComputerName = (Get-ItemProperty "HKLM:SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName").ComputerName
    $activeComputerName = (Get-ItemProperty "HKLM:SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName").ComputerName
    return $currentComputerName -ne $activeComputerName
}

function Get-ItemPropertyValueSafe {
    param (
        [String] $Path,
        [String] $Name
    )
    # See https://github.com/PowerShell/PowerShell/issues/5906
    $value = Get-ItemProperty -Path $Path | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty $Name -ErrorAction SilentlyContinue
    if ([String]::IsNullOrWhiteSpace($value)) {
        return $null;
    }
    return $value
}

function Get-SystemNameChangeStatus {
    $nvName = Get-ItemPropertyValueSafe -Path hklm:System\CurrentControlSet\Services\Tcpip\Parameters -Name "NV Hostname"
    $name = Get-ItemPropertyValueSafe -Path hklm:System\CurrentControlSet\Services\Tcpip\Parameters -Name "Hostname"
    $nvDomain = Get-ItemPropertyValueSafe -Path hklm:System\CurrentControlSet\Services\Tcpip\Parameters -Name "NV Domain"
    $domain = Get-ItemPropertyValueSafe -Path hklm:System\CurrentControlSet\Services\Tcpip\Parameters -Name "Domain"
    return ($nvName -ne $name) -or ($nvDomain -ne $domain)
}
function Test-PendingReboot {
    $value = Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction Ignore
    if ($null -ne $value) { 
        return @{
            RebootRequired        = $true
            AdditionalInformation = 'Component Based Servicing\RebootPending'
        }
    } 
    $value = Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction Ignore
    if ($null -ne $value) { 
        return @{
            RebootRequired        = $true
            AdditionalInformation = 'WindowsUpdate\Auto Update\RebootRequired'
        } 
    }
    if (Get-ComputerNameChangeStatus) { 
        return @{
            RebootRequired        = $true
            AdditionalInformation = 'ComputerName\ActiveComputerName'
        }
    }
    if (Get-SystemNameChangeStatus) {
        return @{
            RebootRequired        = $true
            AdditionalInformation = 'Services\Tcpip\Parameters'
        }
    }
    $status = Invoke-CimMethod -Namespace root/ccm/clientsdk -ClassName CCM_ClientUtilities -MethodName DetermineIfRebootPending -ErrorAction Ignore
    if (($null -ne $status) -and $status.RebootPending) {
        return @{
            RebootRequired        = $true
            AdditionalInformation = 'CCM_ClientUtilities'
        }
    }
    return @{
        RebootRequired        = $false
        AdditionalInformation = $null
    }
}
return Test-PendingReboot

}
## [END] Get-WACSMRebootPendingStatus ##
function Get-WACSMRemoteDesktop {
<#
.SYNOPSIS
Gets the Remote Desktop settings of the system.

.DESCRIPTION
Gets the Remote Desktop settings of the system.

.ROLE
Readers
#>

Set-StrictMode -Version 5.0

Import-Module Microsoft.PowerShell.Management
Import-Module Microsoft.PowerShell.Utility
Import-Module NetSecurity -ErrorAction SilentlyContinue
Import-Module ServerManager -ErrorAction SilentlyContinue

Set-Variable -Option Constant -Name OSRegistryKey -Value "HKLM:\Software\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name OSTypePropertyName -Value "InstallationType" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name OSVersion -Value [Environment]::OSVersion.Version -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpSystemRegistryKey -Value "HKLM:\\SYSTEM\CurrentControlSet\Control\Terminal Server" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpGroupPolicyProperty -Value "fDenyTSConnections" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpNlaGroupPolicyProperty -Value "UserAuthentication" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpGroupPolicyRegistryKey -Value "HKLM:\\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpListenerRegistryKey -Value "$RdpSystemRegistryKey\WinStations" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpProtocolTypeUM -Value "{5828227c-20cf-4408-b73f-73ab70b8849f}" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpProtocolTypeKM -Value "{18b726bb-6fe6-4fb9-9276-ed57ce7c7cb2}" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpWdfSubDesktop -Value 0x00008000 -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RdpFirewallGroup -Value "@FirewallAPI.dll,-28752" -ErrorAction SilentlyContinue
Set-Variable -Option Constant -Name RemoteAppRegistryKey -Value "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\TSAppAllowList" -ErrorAction SilentlyContinue

<#
.SYNOPSIS
Gets the Remote Desktop Network Level Authentication settings of the current machine.

.DESCRIPTION
Gets the Remote Desktop Network Level Authentication settings of the system.

.ROLE
Readers
#>
function Get-RdpNlaGroupPolicySettings {
    $nlaGroupPolicySettings = @{}
    $nlaGroupPolicySettings.GroupPolicyIsSet = $false
    $nlaGroupPolicySettings.GroupPolicyIsEnabled = $false
    $registryKey = Get-ItemProperty -Path $RdpGroupPolicyRegistryKey -ErrorAction SilentlyContinue
    if (!!$registryKey) {
        if ((Get-Member -InputObject $registryKey -name $RdpNlaGroupPolicyProperty -MemberType Properties) -and ($null -ne $registryKey.$RdpNlaGroupPolicyProperty)) {
            $nlaGroupPolicySettings.GroupPolicyIsSet = $true
            $nlaGroupPolicySettings.GroupPolicyIsEnabled = $registryKey.$RdpNlaGroupPolicyProperty -eq 1
        }
    }

    return $nlaGroupPolicySettings
}

<#
.SYNOPSIS
Gets the Remote Desktop settings of the system related to Group Policy.

.DESCRIPTION
Gets the Remote Desktop settings of the system related to Group Policy.

.ROLE
Readers
#>
function Get-RdpGroupPolicySettings {
    $rdpGroupPolicySettings = @{}
    $rdpGroupPolicySettings.GroupPolicyIsSet = $false
    $rdpGroupPolicySettings.GroupPolicyIsEnabled = $false
    $registryKey = Get-ItemProperty -Path $RdpGroupPolicyRegistryKey -ErrorAction SilentlyContinue
    if (!!$registryKey) {
        if ((Get-Member -InputObject $registryKey -name $RdpGroupPolicyProperty -MemberType Properties) -and ($null -ne $registryKey.$RdpGroupPolicyProperty)) {
            $rdpGroupPolicySettings.groupPolicyIsSet = $true
            $rdpGroupPolicySettings.groupPolicyIsEnabled = $registryKey.$RdpGroupPolicyProperty -eq 0
        }
    }

    return $rdpGroupPolicySettings
}

<#
.SYNOPSIS
Gets all of the valid Remote Desktop Protocol listeners.

.DESCRIPTION
Gets all of the valid Remote Desktop Protocol listeners.

.ROLE
Readers
#>
function Get-RdpListener {
    $listeners = @()
    Get-ChildItem -Name $RdpListenerRegistryKey | Where-Object { $_.PSChildName.ToLower() -ne "console" } | ForEach-Object {
        $registryKeyValues = Get-ItemProperty -Path "$RdpListenerRegistryKey\$_" -ErrorAction SilentlyContinue
        if ($null -ne $registryKeyValues) {
            $protocol = $registryKeyValues.LoadableProtocol_Object
            $isProtocolRDP = ($null -ne $protocol) -and ($protocol -eq $RdpProtocolTypeUM -or $protocol -eq $RdpProtocolTypeKM)

            $wdFlag = $registryKeyValues.WdFlag
            $isSubDesktop = ($null -ne $wdFlag) -and ($wdFlag -band $RdpWdfSubDesktop)

            $isRDPListener = $isProtocolRDP -and !$isSubDesktop
            if ($isRDPListener) {
                $listeners += $registryKeyValues
            }
        }
    }

    return ,$listeners
}

<#
.SYNOPSIS
Gets the number of the ports that the Remote Desktop Protocol is operating over.

.DESCRIPTION
Gets the number of the ports that the Remote Desktop Protocol is operating over.

.ROLE
Readers
#>
function Get-RdpPortNumber {
    $portNumbers = @()
    Get-RdpListener | Where-Object { $null -ne $_.PortNumber } | ForEach-Object { $portNumbers += $_.PortNumber }
    return ,$portNumbers
}

<#
.SYNOPSIS
Gets the Remote Desktop settings of the system.

.DESCRIPTION
Gets the Remote Desktop settings of the system.

.ROLE
Readers
#>
function Get-RdpSettings {
    $remoteDesktopSettings = New-Object -TypeName PSObject
    $rdpEnabledSource = $null
    $rdpIsEnabled = Test-RdpEnabled
    $rdpRequiresNla = Test-RdpUserAuthentication
    $remoteAppAllowed = Test-RemoteApp
    $rdpPortNumbers = Get-RdpPortNumber
    if ($rdpIsEnabled) {
        $rdpGroupPolicySettings = Get-RdpGroupPolicySettings
        if ($rdpGroupPolicySettings.groupPolicyIsEnabled) {
            $rdpEnabledSource = "GroupPolicy"
        } else {
            $rdpEnabledSource = "System"
        }
    }
    $operatingSystemType = Get-OperatingSystemType
    $desktopFeatureAvailable = Test-DesktopFeature($operatingSystemType)
    $versionIsSupported = Test-OSVersion($operatingSystemType)

    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "IsEnabled" -Value $rdpIsEnabled
    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "RequiresNLA" -Value $rdpRequiresNla
    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "Ports" -Value $rdpPortNumbers
    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "EnabledSource" -Value $rdpEnabledSource
    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "RemoteAppAllowed" -Value $remoteAppAllowed
    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "DesktopFeatureAvailable" -Value $desktopFeatureAvailable
    $remoteDesktopSettings | Add-Member -MemberType NoteProperty -Name "VersionIsSupported" -Value $versionIsSupported

    return $remoteDesktopSettings
}

<#
.SYNOPSIS
Tests whether Remote Desktop Protocol is enabled.

.DESCRIPTION
Tests whether Remote Desktop Protocol is enabled.

.ROLE
Readers
#>
function Test-RdpEnabled {
    $rdpEnabledWithGP = $false
    $rdpEnabledLocally = $false
    $rdpGroupPolicySettings = Get-RdpGroupPolicySettings
    $rdpEnabledWithGP = $rdpGroupPolicySettings.GroupPolicyIsSet -and $rdpGroupPolicySettings.GroupPolicyIsEnabled
    $rdpEnabledLocally = !($rdpGroupPolicySettings.GroupPolicyIsSet) -and (Test-RdpSystem)

    return (Test-RdpListener) -and (Test-RdpFirewall) -and ($rdpEnabledWithGP -or $rdpEnabledLocally)
}

<#
.SYNOPSIS
Tests whether the Remote Desktop Firewall rules are enabled.

.DESCRIPTION
Tests whether the Remote Desktop Firewall rules are enabled.

.ROLE
Readers
#>
function Test-RdpFirewall {
    $firewallRulesEnabled = $true
    Get-NetFirewallRule -Group $RdpFirewallGroup | Where-Object { $_.Profile -match "Domain" } | ForEach-Object {
        if ($_.Enabled -eq "False") {
            $firewallRulesEnabled = $false
        }
    }

    return $firewallRulesEnabled
}

<#
.SYNOPSIS
Tests whether or not a Remote Desktop Protocol listener exists.

.DESCRIPTION
Tests whether or not a Remote Desktop Protocol listener exists.

.ROLE
Readers
#>
function Test-RdpListener {
    $listeners = Get-RdpListener
    return ($listeners | Microsoft.PowerShell.Utility\Measure-Object).Count -gt 0
}

<#
.SYNOPSIS
Tests whether Remote Desktop Protocol is enabled via local system settings.

.DESCRIPTION
Tests whether Remote Desktop Protocol is enabled via local system settings.

.ROLE
Readers
#>
function Test-RdpSystem {
    $registryKey = Get-ItemProperty -Path $RdpSystemRegistryKey -ErrorAction SilentlyContinue

    if ($registryKey) {
        return $registryKey.fDenyTSConnections -eq 0
    } else {
        return $false
    }
}

<#
.SYNOPSIS
Tests whether Remote Desktop connections require Network Level Authentication while enabled via local system settings.

.DESCRIPTION
Tests whether Remote Desktop connections require Network Level Authentication while enabled via local system settings.

.ROLE
Readers
#>
function Test-RdpSystemUserAuthentication {
    $listener = Get-RdpListener | Where-Object { $null -ne $_.UserAuthentication } | Microsoft.PowerShell.Utility\Select-Object -First 1

    if ($listener) {
        return $listener.UserAuthentication -eq 1
    } else {
        return $false
    }
}

<#
.SYNOPSIS
Tests whether Remote Desktop connections require Network Level Authentication.

.DESCRIPTION
Tests whether Remote Desktop connections require Network Level Authentication.

.ROLE
Readers
#>
function Test-RdpUserAuthentication {
    $nlaEnabledWithGP = $false
    $nlaEnabledLocally = $false
    $nlaGroupPolicySettings = Get-RdpNlaGroupPolicySettings
    $nlaEnabledWithGP = $nlaGroupPolicySettings.GroupPolicyIsSet -and $nlaGroupPolicySettings.GroupPolicyIsEnabled
    $nlaEnabledLocally = !($nlaGroupPolicySettings.GroupPolicyIsSet) -and (Test-RdpSystemUserAuthentication)

    return $nlaEnabledWithGP -or $nlaEnabledLocally
}

<#
.SYNOPSIS
Tests whether Remote App connections are allowed.

.DESCRIPTION
Tests whether Remote App connections are allowed.

.ROLE
Readers
#>
function Test-RemoteApp {
  $registryKey = Get-ItemProperty -Path $RemoteAppRegistryKey -Name fDisabledAllowList -ErrorAction SilentlyContinue
  if ($registryKey)
  {
      $remoteAppEnabled = $registryKey.fDisabledAllowList
      return $remoteAppEnabled -eq 1
  } else {
      return $false;
  }
}

<#
.SYNOPSIS
Gets the Windows OS installation type.

.DESCRIPTION
Gets the Windows OS installation type.

.ROLE
Readers
#>
function Get-OperatingSystemType {
    $osResult = Get-ItemProperty -Path $OSRegistryKey -Name $OSTypePropertyName -ErrorAction SilentlyContinue

    if ($osResult -and $osResult.$OSTypePropertyName) {
        return $osResult.$OSTypePropertyName
    } else {
        return $null
    }
}

<#
.SYNOPSIS
Tests the availability of desktop features based on the system's OS type.

.DESCRIPTION
Tests the availability of desktop features based on the system's OS type.

.ROLE
Readers
#>
function Test-DesktopFeature ([string] $osType) {
    $featureAvailable = $false

    switch ($osType) {
        'Client' {
            $featureAvailable = $true
        }
        'Server' {
            $DesktopFeature = Get-DesktopFeature
            if ($DesktopFeature) {
                $featureAvailable = $DesktopFeature.Installed
            }
        }
    }

    return $featureAvailable
}

<#
.SYNOPSIS
Checks for feature cmdlet availability and returns the installation state of the Desktop Experience feature.

.DESCRIPTION
Checks for feature cmdlet availability and returns the installation state of the Desktop Experience feature.

.ROLE
Readers
#>
function Get-DesktopFeature {
    $moduleAvailable = Get-Module -ListAvailable -Name ServerManager -ErrorAction SilentlyContinue
    if ($moduleAvailable) {
        return Get-WindowsFeature -Name Desktop-Experience -ErrorAction SilentlyContinue
    } else {
        return $null
    }
}

<#
.SYNOPSIS
Tests whether the current OS type/version is supported for Remote App.

.DESCRIPTION
Tests whether the current OS type/version is supported for Remote App.

.ROLE
Readers
#>
function Test-OSVersion ([string] $osType) {
    switch ($osType) {
        'Client' {
            return (Get-OSVersion) -ge (new-object 'Version' 6,2)
        }
        'Server' {
            return (Get-OSVersion) -ge (new-object 'Version' 6,3)
        }
        default {
            return $false
        }
    }
}

<#
.SYNOPSIS
Retrieves the system version information from the system's environment variables.

.DESCRIPTION
Retrieves the system version information from the system's environment variables.

.ROLE
Readers
#>
function Get-OSVersion {
    return [Environment]::OSVersion.Version
}

#########
# Main
#########

$module = Get-Module -Name NetSecurity -ErrorAction SilentlyContinue

if ($module) {
    Get-RdpSettings
}
}
## [END] Get-WACSMRemoteDesktop ##
function Get-WACSMSQLServerEndOfSupportVersion {
<#

.SYNOPSIS
Gets information about SQL Server installation on the server.

.DESCRIPTION
Gets information about SQL Server installation on the server.

.ROLE
Readers

#>

import-module CimCmdlets

$V2008 = [version]'10.0.0.0'
$V2008R2 = [version]'10.50.0.0'

Set-Variable -Name SQLRegistryRoot64Bit -Option ReadOnly -Value "HKLM:\\SOFTWARE\\Microsoft\\Microsoft SQL Server" -ErrorAction SilentlyContinue
Set-Variable -Name SQLRegistryRoot32Bit -Option ReadOnly -Value "HKLM:\\SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server" -ErrorAction SilentlyContinue
Set-Variable -Name InstanceNamesSubKey -Option ReadOnly -Value "Instance Names"-ErrorAction SilentlyContinue
Set-Variable -Name SQLSubKey -Option ReadOnly -Value "SQL" -ErrorAction SilentlyContinue
Set-Variable -Name CurrentVersionSubKey -Option ReadOnly -Value "CurrentVersion" -ErrorAction SilentlyContinue
Set-Variable -Name Running -Option ReadOnly -Value "Running" -ErrorAction SilentlyContinue

function Get-KeyPropertiesAndValues($path) {
  Get-Item $path -ErrorAction SilentlyContinue |
  Microsoft.PowerShell.Utility\Select-Object -ExpandProperty property |
  ForEach-Object {
    New-Object psobject -Property @{"Property"=$_; "Value" = (Get-ItemProperty -Path $path -Name $_ -ErrorAction SilentlyContinue).$_}
  }
}

function IsEndofSupportVersion($SQLRegistryPath) {
  $result = $false
  if (Test-Path -Path $SQLRegistryPath) {
    # construct reg key path to lead up to instances.
    $InstanceNamesKeyPath = Join-Path $SQLRegistryPath -ChildPath $InstanceNamesSubKey | Join-Path -ChildPath $SQLSubKey

    if (Test-Path -Path $InstanceNamesKeyPath) {
      # get properties and their values
      $InstanceCollection = Get-KeyPropertiesAndValues($InstanceNamesKeyPath)
      if ($InstanceCollection) {
        foreach ($Instance in $InstanceCollection) {
          if (Get-Service | Where-Object { $_.Status -eq $Running } | Where-Object { $_.Name -eq $Instance.Property }) {
            $VersionPath = Join-Path $SQLRegistryPath -ChildPath $Instance.Value | Join-Path -ChildPath $Instance.Property | Join-Path -ChildPath $CurrentVersionSubKey
            if (Test-Path -Path $VersionPath) {
              $CurrentVersion = [version] (Get-ItemPropertyValue $VersionPath $CurrentVersionSubKey -ErrorAction SilentlyContinue)
              if ($CurrentVersion -ge $V2008 -and $CurrentVersion -le $V2008R2) {
                $result = $true
                break
              }
            }
          }
        }
      }
    }
  }

  return $result
}

$Result64Bit = IsEndofSupportVersion($SQLRegistryRoot64Bit)
$Result32Bit = IsEndofSupportVersion($SQLRegistryRoot32Bit)

return $Result64Bit -OR $Result32Bit

}
## [END] Get-WACSMSQLServerEndOfSupportVersion ##
function Get-WACSMServerConnectionStatus {
<#

.SYNOPSIS
Gets status of the connection to the server.

.DESCRIPTION
Gets status of the connection to the server.

.ROLE
Readers

#>

import-module CimCmdlets

$OperatingSystem = Get-CimInstance Win32_OperatingSystem
$Caption = $OperatingSystem.Caption
$ProductType = $OperatingSystem.ProductType
$Version = $OperatingSystem.Version
$Status = @{ Label = $null; Type = 0; Details = $null; }
$Result = @{ Status = $Status; Caption = $Caption; ProductType = $ProductType; Version = $Version; }
if ($Version -and ($ProductType -eq 2 -or $ProductType -eq 3)) {
    $V = [version]$Version
    $V2016 = [version]'10.0'
    $V2012 = [version]'6.2'
    $V2008r2 = [version]'6.1'

    if ($V -ge $V2016) {
        return $Result;
    }

    if ($V -ge $V2008r2) {
        $Key = 'HKLM:\\SOFTWARE\\Microsoft\\PowerShell\\3\\PowerShellEngine'
        $WmfStatus = $false;
        $Exists = Get-ItemProperty -Path $Key -Name PowerShellVersion -ErrorAction SilentlyContinue
        if (![String]::IsNullOrEmpty($Exists)) {
            $WmfVersionInstalled = $exists.PowerShellVersion
            if ($WmfVersionInstalled.StartsWith('5.')) {
                $WmfStatus = $true;
            }
        }

        if (!$WmfStatus) {
            $status.Label = 'wmfMissing-label'
            $status.Type = 3
            $status.Details = 'wmfMissing-details'
        }

        return $result;
    }
}

$status.Label = 'unsupported-label'
$status.Type = 3
$status.Details = 'unsupported-details'
return $result;

}
## [END] Get-WACSMServerConnectionStatus ##
function Install-WACSMMonitoringDependencies {
<#

.SYNOPSIS
Script that returns if Microsoft Monitoring Agent is running or not.

.DESCRIPTION
Download and install MMAAgent & Microsoft Dependency agent

.PARAMETER WorkspaceId
  is the workspace id of the Log Analytics workspace

.PARAMETER WorkspacePrimaryKey
  is the primary key of the Log Analytics workspace

.PARAMETER IsHciCluster
 flag to indicate if the node is part of a HCI cluster

.PARAMETER AzureCloudType
  is the Azure cloud type of the Log Analytics workspace

.ROLE
Administrators

#>

[CmdletBinding()]
param (
  [Parameter()]
  [String]
  $WorkspaceId,
  [Parameter()]
  [String]
  $WorkspacePrimaryKey,
  [Parameter()]
  [bool]
  $IsHciCluster,
  [Parameter()]
  [int]
  $AzureCloudType
)

$ErrorActionPreference = "Stop"

$LogName = "Microsoft-ServerManagementExperience"
$LogSource = "SMEScript"
$ScriptName = "Install-MonitoringDependencies.ps1"

Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

<#
.SYNOPSIS
    Utility function to invoke a Windows command.
    (This command is Microsoft internal use only.)

.DESCRIPTION
    Invokes a Windows command and generates an exception if the command returns an error. Note: only for application commands.

.PARAMETER Command
    The name of the command we want to invoke.

.PARAMETER Parameters
    The parameters we want to pass to the command.
.EXAMPLE
    Invoke-WACWinCommand "netsh" "http delete sslcert ipport=0.0.0.0:9999"
#>
function Invoke-WACWinCommand {
  Param(
    [string]$Command,
    [string[]]$Parameters
  )

  try {
    Write-Verbose "$command $([System.String]::Join(" ", $Parameters))"
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $Command
    $startInfo.RedirectStandardError = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.UseShellExecute = $false
    $startInfo.Arguments = [System.String]::Join(" ", $Parameters)
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
  }
  catch {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
      -Message "[$ScriptName]: $_"  -ErrorAction SilentlyContinue
    Write-Error $_
  }

  try {
    $process.Start() | Out-Null
  }
  catch {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
      -Message "[$ScriptName]: $_"  -ErrorAction SilentlyContinue
    Write-Error $_
  }

  try {
    $process.WaitForExit() | Out-Null
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $output = $stdout + "`r`n" + $stderr
  }
  catch {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
      -Message "[$ScriptName]: $_"  -ErrorAction SilentlyContinue
    Write-Error $_
  }

  if ($process.ExitCode -ne 0) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
      -Message "[$ScriptName]: $_"  -ErrorAction SilentlyContinue
    Write-Error $_
  }

  # output all messages
  return $output
}

$MMAAgentStatus = Get-Service -Name HealthService -ErrorAction SilentlyContinue
$IsMmaRunning = $null -ne $MMAAgentStatus -and $MMAAgentStatus.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running

if (-not $IsMmaRunning) {
  # install MMA agent
  $MmaExePath = Join-Path -Path $env:temp -ChildPath 'MMASetup-AMD64.exe'
  if (Test-Path $MmaExePath) {
    Remove-Item $MmaExePath
  }
  Invoke-WebRequest -Uri https://go.microsoft.com/fwlink/?LinkId=828603 -OutFile $MmaExePath

  $ExtractFolder = Join-Path -Path $env:temp -ChildPath 'SmeMMAInstaller'
  if (Test-Path $ExtractFolder) {
    Remove-Item $ExtractFolder -Force -Recurse
  }

  &$MmaExePath /c /t:$ExtractFolder
  $SetupExePath = Join-Path -Path $ExtractFolder -ChildPath 'setup.exe'
  for ($i = 0; $i -lt 10; $i++) {
    if (-Not(Test-Path $SetupExePath)) {
      Start-Sleep -Seconds 6
    }
  }


  Invoke-WACWinCommand -Command $SetupExePath -Parameters "/qn", "NOAPM=1", "ADD_OPINSIGHTS_WORKSPACE=1", "OPINSIGHTS_WORKSPACE_AZURE_CLOUD_TYPE=$AzureCloudType", "OPINSIGHTS_WORKSPACE_ID=$WorkspaceId", "OPINSIGHTS_WORKSPACE_KEY=$WorkspacePrimaryKey", "AcceptEndUserLicenseAgreement=1"
}

$ServiceMapAgentStatus = Get-Service -Name MicrosoftDependencyAgent -ErrorAction SilentlyContinue
$IsServiceMapRunning = $null -ne $ServiceMapAgentStatus -and $ServiceMapAgentStatus.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running

if (-not $IsServiceMapRunning) {
  # Install service map/ dependency agent
  $ServiceMapExe = Join-Path -Path $env:temp -ChildPath 'InstallDependencyAgent-Windows.exe'

  if (Test-Path $ServiceMapExe) {
    Remove-Item $ServiceMapExe
  }
  Invoke-WebRequest -Uri https://aka.ms/dependencyagentwindows -OutFile $ServiceMapExe

  Invoke-WACWinCommand -Command $ServiceMapExe -Parameters "/S", "AcceptEndUserLicenseAgreement=1"
}

# Wait for agents to completely install
for ($i = 0; $i -lt 10; $i++) {
  if ($null -eq (Get-Service -Name HealthService -ErrorAction SilentlyContinue) -or $null -eq (Get-Service -Name MicrosoftDependencyAgent -ErrorAction SilentlyContinue)) {
    Start-Sleep -Seconds 6
  }
}

<#
 # .DESCRIPTION
 # Enable health settings on HCI cluster node to log faults into Microsoft-Windows-Health/Operational
 #>
if ($IsHciCluster) {
  $subsystem = Get-StorageSubsystem clus*
  $subsystem | Set-StorageHealthSetting -Name "Platform.ETW.MasTypes" -Value "Microsoft.Health.EntityType.Subsystem,Microsoft.Health.EntityType.Server,Microsoft.Health.EntityType.PhysicalDisk,Microsoft.Health.EntityType.StoragePool,Microsoft.Health.EntityType.Volume,Microsoft.Health.EntityType.Cluster"
}

}
## [END] Install-WACSMMonitoringDependencies ##
function New-WACSMEnvironmentVariable {
<#

.SYNOPSIS
Creates a new environment variable specified by name, type and data.

.DESCRIPTION
Creates a new environment variable specified by name, type and data.

.ROLE
Administrators

#>

param(
    [Parameter(Mandatory = $True)]
    [String]
    $name,

    [Parameter(Mandatory = $True)]
    [String]
    $value,

    [Parameter(Mandatory = $True)]
    [String]
    $type
)

Set-StrictMode -Version 5.0
$strings = 
ConvertFrom-StringData @'
EnvironmentErrorAlreadyExists=An environment variable of this name and type already exists.
EnvironmentErrorDoesNotExists=An environment variable of this name and type does not exist.
'@

# SIG # Begin signature block
# MIIoOQYJKoZIhvcNAQcCoIIoKjCCKCYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDmP40B4Lwr2G94
# Wl9OxoBBVvOiw6NiIq7+OROWuiUtnqCCDYUwggYDMIID66ADAgECAhMzAAADTU6R
# phoosHiPAAAAAANNMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI4WhcNMjQwMzE0MTg0MzI4WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDUKPcKGVa6cboGQU03ONbUKyl4WpH6Q2Xo9cP3RhXTOa6C6THltd2RfnjlUQG+
# Mwoy93iGmGKEMF/jyO2XdiwMP427j90C/PMY/d5vY31sx+udtbif7GCJ7jJ1vLzd
# j28zV4r0FGG6yEv+tUNelTIsFmmSb0FUiJtU4r5sfCThvg8dI/F9Hh6xMZoVti+k
# bVla+hlG8bf4s00VTw4uAZhjGTFCYFRytKJ3/mteg2qnwvHDOgV7QSdV5dWdd0+x
# zcuG0qgd3oCCAjH8ZmjmowkHUe4dUmbcZfXsgWlOfc6DG7JS+DeJak1DvabamYqH
# g1AUeZ0+skpkwrKwXTFwBRltAgMBAAGjggGCMIIBfjAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUId2Img2Sp05U6XI04jli2KohL+8w
# VAYDVR0RBE0wS6RJMEcxLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEWMBQGA1UEBRMNMjMwMDEyKzUwMDUxNzAfBgNVHSMEGDAW
# gBRIbmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDEx
# XzIwMTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIB
# ACMET8WuzLrDwexuTUZe9v2xrW8WGUPRQVmyJ1b/BzKYBZ5aU4Qvh5LzZe9jOExD
# YUlKb/Y73lqIIfUcEO/6W3b+7t1P9m9M1xPrZv5cfnSCguooPDq4rQe/iCdNDwHT
# 6XYW6yetxTJMOo4tUDbSS0YiZr7Mab2wkjgNFa0jRFheS9daTS1oJ/z5bNlGinxq
# 2v8azSP/GcH/t8eTrHQfcax3WbPELoGHIbryrSUaOCphsnCNUqUN5FbEMlat5MuY
# 94rGMJnq1IEd6S8ngK6C8E9SWpGEO3NDa0NlAViorpGfI0NYIbdynyOB846aWAjN
# fgThIcdzdWFvAl/6ktWXLETn8u/lYQyWGmul3yz+w06puIPD9p4KPiWBkCesKDHv
# XLrT3BbLZ8dKqSOV8DtzLFAfc9qAsNiG8EoathluJBsbyFbpebadKlErFidAX8KE
# usk8htHqiSkNxydamL/tKfx3V/vDAoQE59ysv4r3pE+zdyfMairvkFNNw7cPn1kH
# Gcww9dFSY2QwAxhMzmoM0G+M+YvBnBu5wjfxNrMRilRbxM6Cj9hKFh0YTwba6M7z
# ntHHpX3d+nabjFm/TnMRROOgIXJzYbzKKaO2g1kWeyG2QtvIR147zlrbQD4X10Ab
# rRg9CpwW7xYxywezj+iNAc+QmFzR94dzJkEPUSCJPsTFMIIHejCCBWKgAwIBAgIK
# YQ6Q0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEw
# OTA5WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYD
# VQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+la
# UKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc
# 6Whe0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4D
# dato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+
# lD3v++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nk
# kDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6
# A4aN91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmd
# X4jiJV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL
# 5zmhD+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zd
# sGbiwZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3
# T8HhhUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS
# 4NaIjAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRI
# bmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAL
# BgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBD
# uRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jv
# c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEF
# BQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1h
# cnljcHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkA
# YwB5AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn
# 8oalmOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7
# v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0b
# pdS1HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/
# KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvy
# CInWH8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBp
# mLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJi
# hsMdYzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYb
# BL7fQccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbS
# oqKfenoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sL
# gOppO6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtX
# cVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
# Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAANNTpGmGiiweI8AAAAA
# A00wDQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIGmE
# xt4Vo2SMLvidEoKOy0BWlnwrLtFSz/H6A6/tXaeqMEIGCisGAQQBgjcCAQwxNDAy
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20wDQYJKoZIhvcNAQEBBQAEggEAWRNT4wF78dRbCChPkMmmEMITfsHerfAA340h
# ViK1z5Yj5GYdtsaO6eMNrckGgdhY4ypdzR+wyt3WqENAfgDtlW8qnfb+e5aleTWq
# M0lKrFcmKf/rZWXG1Xbt4OPAsDMiRML8b4oaQK8a6f3xaoFt8QzRhtKmmNgmW1Um
# BAzgvmNy68PhWf+mO0zCYoArHdxkxaspc/YSAQUdQm0aQNaBc5nGvc/jiAP3v2uf
# ZeLYZAkMRZ3x3QzwnMg3+WGc4mlaiyTXnETAO7q8l2yDuGNlidJMX4ZRvIjpgEUU
# mQpGSjI1bwbjTexfSaexY2T3sgEzstp8JjDOLIBabLUWgjC85qGCF5QwgheQBgor
# BgEEAYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZI
# AWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGE
# WQoDATAxMA0GCWCGSAFlAwQCAQUABCBp9m2vxp6Pndq+mQJd4L4+LDvY8xW59oGE
# kE7K0eTYQQIGZVbJFW75GBMyMDIzMTIwNzE4MDAyNC4xMTNaMASAAgH0oIHRpIHO
# MIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQL
# ExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxk
# IFRTUyBFU046QTAwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAdB3CKrvoxfG3QAB
# AAAB0DANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MDAeFw0yMzA1MjUxOTEyMTRaFw0yNDAyMDExOTEyMTRaMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTAwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Uw
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDfMlfn35fvM0XAUSmI5qiG
# 0UxPi25HkSyBgzk3zpYO311d1OEEFz0QpAK23s1dJFrjB5gD+SMw5z6EwxC4CrXU
# 9KaQ4WNHqHrhWftpgo3MkJex9frmO9MldUfjUG56sIW6YVF6YjX+9rT1JDdCDHbo
# 5nZiasMigGKawGb2HqD7/kjRR67RvVh7Q4natAVu46Zf5MLviR0xN5cNG20xwBwg
# ttaYEk5XlULaBH5OnXz2eWoIx+SjDO7Bt5BuABWY8SvmRQfByT2cppEzTjt/fs0x
# p4B1cAHVDwlGwZuv9Rfc3nddxgFrKA8MWHbJF0+aWUUYIBR8Fy2guFVHoHeOze7I
# sbyvRrax//83gYqo8c5Z/1/u7kjLcTgipiyZ8XERsLEECJ5ox1BBLY6AjmbgAzDd
# Nl2Leej+qIbdBr/SUvKEC+Xw4xjFMOTUVWKWemt2khwndUfBNR7Nzu1z9L0Wv7TA
# Y/v+v6pNhAeohPMCFJc+ak6uMD8TKSzWFjw5aADkmD9mGuC86yvSKkII4MayzoUd
# seT0nfk8Y0fPjtdw2Wnejl6zLHuYXwcDau2O1DMuoiedNVjTF37UEmYT+oxC/OFX
# UGPDEQt9tzgbR9g8HLtUfEeWOsOED5xgb5rwyfvIss7H/cdHFcIiIczzQgYnsLyE
# GepoZDkKhSMR5eCB6Kcv/QIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFDPhAYWS0oA+
# lOtITfjJtyl0knRRMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8G
# A1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBs
# BggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUy
# MDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCXh+ckCkZaA06S
# NW+qxtS9gHQp4x7G+gdikngKItEr8otkXIrmWPYrarRWBlY91lqGiilHyIlZ3iNB
# UbaNEmaKAGMZ5YcS7IZUKPaq1jU0msyl+8og0t9C/Z26+atx3vshHrFQuSgwTHZV
# pzv7k8CYnBYoxdhI1uGhqH595mqLvtMsxEN/1so7U+b3U6LCry5uwwcz5+j8Oj0G
# UX3b+iZg+As0xTN6T0Qa8BNec/LwcyqYNEaMkW2VAKrmhvWH8OCDTcXgONnnABQH
# BfXK/fLAbHFGS1XNOtr62/iaHBGAkrCGl6Bi8Pfws6fs+w+sE9r3hX9Vg0gsRMoH
# RuMaiXsrGmGsuYnLn3AwTguMatw9R8U5vJtWSlu1CFO5P0LEvQQiMZ12sQSsQAkN
# DTs9rTjVNjjIUgoZ6XPMxlcPIDcjxw8bfeb4y4wAxM2RRoWcxpkx+6IIf2L+b7gL
# HtBxXCWJ5bMW7WwUC2LltburUwBv0SgjpDtbEqw/uDgWBerCT+Zty3Nc967iGaQj
# yYQH6H/h9Xc8smm2n6VjySRx2swnW3hr6Qx63U/xY9HL6FNhrGiFED7ZRKrnwvvX
# vMVQUIEkB7GUEeN6heY8gHLt0jLV3yzDiQA8R8p5YGgGAVt9MEwgAJNY1iHvH/8v
# zhJSZFNkH8svRztO/i3TvKrjb8ZxwjCCB3EwggVZoAMCAQICEzMAAAAVxedrngKb
# SZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmlj
# YXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIy
# NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UE
# AxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXI
# yjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjo
# YH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1y
# aa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v
# 3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pG
# ve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viS
# kR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYr
# bqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlM
# jgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSL
# W6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AF
# emzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIu
# rQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIE
# FgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWn
# G1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEW
# M2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5
# Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBi
# AEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV
# 9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3Js
# Lm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAx
# MC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2
# LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv
# 6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZn
# OlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1
# bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4
# rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU
# 6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDF
# NLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/
# HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdU
# CbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKi
# excdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTm
# dHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZq
# ELQdVTNYs6FwZvKhggNNMIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJp
# Y2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkEwMDAtMDVF
# MC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMK
# AQEwBwYFKw4DAhoDFQC8t8hT8KKUX91lU5FqRP9Cfu9MiaCBgzCBgKR+MHwxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6RxLljAi
# GA8yMDIzMTIwNzEzNDgwNloYDzIwMjMxMjA4MTM0ODA2WjB0MDoGCisGAQQBhFkK
# BAExLDAqMAoCBQDpHEuWAgEAMAcCAQACAjDRMAcCAQACAhNDMAoCBQDpHZ0WAgEA
# MDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAI
# AgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBAHvwZk3Mzaev9U7tJRsrU56Sk+6d
# RxxO0o3E3/rZ4fJFTTiqk3oc+v9+ONv7iV4C4/mBTVWxht4dgbvKez9f6vjXla3r
# Rtl80gR3tPVrVM5GtAO7vl8NA7cHmPv9rbIl1lNGRS3PuOiJC3MxomYfAbWxYvUU
# MrcytJb3eg10dQmZBsa9DL7XUPrwtE1GwQXVP9xJixgA0Ec22dP1dsKxi8tkCMxF
# 3RdJF74kyyYCd2EfzmxY5X3Mz5eQI3/QT8cQ2cffD+dxekKeLyWvE2WOn8RvubG8
# fs+p7/O4VBo+rYX0dkNPYsY1rgCS1DpIDIlJH4VMBYhRmom7XJJTbXr7mhIxggQN
# MIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdB3
# CKrvoxfG3QABAAAB0DANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
# CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCB7iw3OGYITmuR9heLeuqHpIeqO
# qIECZS4uinSRUoWrhzCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIAiVQAZf
# tNP/Md1E2Yw+fBXa9w6fjmTZ5WAerrTSPwnXMIGYMIGApH4wfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAHQdwiq76MXxt0AAQAAAdAwIgQgIjC5GU5I
# 1vumB/an/LTMlONJuNqUAMNDLVE8un+gLPAwDQYJKoZIhvcNAQELBQAEggIAh0KZ
# uoXY/X0FuusVI5L7ldhfiePqKdS0WyPaV+QJQho5I5rfxOKR7zf5tpF1zPTX+1ln
# gvDS4rjvoCNrhwJVgRtUd8ZCneWkEWvmmDJOEwr/TLqFDW2bI1cpxep/KmjJ/LZ1
# pEOXeECwyyC6OZtsiy/5RtvfkSUI3CP/f0Vjm9nZvap0zQClEWzLjzPq07BleCQH
# t0DXSzeVSx/GmT3EQW5JXPOvj6XA8f3G/yT/ilx7o/Ta3W5/nqVcQmyDS683pnem
# baMthtk2Y7QjtCJsyICfSuMagdmPuPl9sK4FvDa7+SIOrMfjoBKqqAabM47qK41d
# Va0W0KwExrTKkbIl092eembyyNVrGnjLLNn8SmwohbF6l1DALW73aYKZwnKe28Lm
# 4iyFoU3fwC//8jjofxxVD6A1RpU0crOhfb6DWJ36qiRU6VrZcgYJA98O+bbb1eve
# pMJeITHqufUFn+cx7hdXve/gtOvfRNzPYJHeFgsf9qLv4Z+9FJmzasnC1fkLtoG2
# w910HJyE4d2A7r0SeFNQfVjoHS6mZSoBdBiChsGiBMvCTjSVcx+xPAQJgqYpYwFe
# pp9oLfuhc//5kuFHw1pMxzXF/9nBaYoy57TuzG1BaMQ+bg2ANqkFh8yX3ewRD+Ry
# 0HIUZe6JGvebPR6uDmdebz3eJUP3uhjfAijBmKw=
# SIG # End signature block


If ([Environment]::GetEnvironmentVariable($name, $type) -eq $null) {
    return [Environment]::SetEnvironmentVariable($name, $value, $type)
}
Else {
    Write-Error $strings.EnvironmentErrorAlreadyExists
}
}
## [END] New-WACSMEnvironmentVariable ##
function Remove-WACSMEnvironmentVariable {
<#

.SYNOPSIS
Removes an environment variable specified by name and type.

.DESCRIPTION
Removes an environment variable specified by name and type.

.ROLE
Administrators

#>

param(
    [Parameter(Mandatory = $True)]
    [String]
    $name,

    [Parameter(Mandatory = $True)]
    [String]
    $type
)

Set-StrictMode -Version 5.0
$strings = 
ConvertFrom-StringData @'
EnvironmentErrorAlreadyExists=An environment variable of this name and type already exists.
EnvironmentErrorDoesNotExists=An environment variable of this name and type does not exist.
'@

# SIG # Begin signature block
# MIIoOQYJKoZIhvcNAQcCoIIoKjCCKCYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDmP40B4Lwr2G94
# Wl9OxoBBVvOiw6NiIq7+OROWuiUtnqCCDYUwggYDMIID66ADAgECAhMzAAADTU6R
# phoosHiPAAAAAANNMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI4WhcNMjQwMzE0MTg0MzI4WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDUKPcKGVa6cboGQU03ONbUKyl4WpH6Q2Xo9cP3RhXTOa6C6THltd2RfnjlUQG+
# Mwoy93iGmGKEMF/jyO2XdiwMP427j90C/PMY/d5vY31sx+udtbif7GCJ7jJ1vLzd
# j28zV4r0FGG6yEv+tUNelTIsFmmSb0FUiJtU4r5sfCThvg8dI/F9Hh6xMZoVti+k
# bVla+hlG8bf4s00VTw4uAZhjGTFCYFRytKJ3/mteg2qnwvHDOgV7QSdV5dWdd0+x
# zcuG0qgd3oCCAjH8ZmjmowkHUe4dUmbcZfXsgWlOfc6DG7JS+DeJak1DvabamYqH
# g1AUeZ0+skpkwrKwXTFwBRltAgMBAAGjggGCMIIBfjAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUId2Img2Sp05U6XI04jli2KohL+8w
# VAYDVR0RBE0wS6RJMEcxLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEWMBQGA1UEBRMNMjMwMDEyKzUwMDUxNzAfBgNVHSMEGDAW
# gBRIbmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDEx
# XzIwMTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIB
# ACMET8WuzLrDwexuTUZe9v2xrW8WGUPRQVmyJ1b/BzKYBZ5aU4Qvh5LzZe9jOExD
# YUlKb/Y73lqIIfUcEO/6W3b+7t1P9m9M1xPrZv5cfnSCguooPDq4rQe/iCdNDwHT
# 6XYW6yetxTJMOo4tUDbSS0YiZr7Mab2wkjgNFa0jRFheS9daTS1oJ/z5bNlGinxq
# 2v8azSP/GcH/t8eTrHQfcax3WbPELoGHIbryrSUaOCphsnCNUqUN5FbEMlat5MuY
# 94rGMJnq1IEd6S8ngK6C8E9SWpGEO3NDa0NlAViorpGfI0NYIbdynyOB846aWAjN
# fgThIcdzdWFvAl/6ktWXLETn8u/lYQyWGmul3yz+w06puIPD9p4KPiWBkCesKDHv
# XLrT3BbLZ8dKqSOV8DtzLFAfc9qAsNiG8EoathluJBsbyFbpebadKlErFidAX8KE
# usk8htHqiSkNxydamL/tKfx3V/vDAoQE59ysv4r3pE+zdyfMairvkFNNw7cPn1kH
# Gcww9dFSY2QwAxhMzmoM0G+M+YvBnBu5wjfxNrMRilRbxM6Cj9hKFh0YTwba6M7z
# ntHHpX3d+nabjFm/TnMRROOgIXJzYbzKKaO2g1kWeyG2QtvIR147zlrbQD4X10Ab
# rRg9CpwW7xYxywezj+iNAc+QmFzR94dzJkEPUSCJPsTFMIIHejCCBWKgAwIBAgIK
# YQ6Q0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEw
# OTA5WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYD
# VQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+la
# UKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc
# 6Whe0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4D
# dato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+
# lD3v++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nk
# kDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6
# A4aN91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmd
# X4jiJV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL
# 5zmhD+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zd
# sGbiwZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3
# T8HhhUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS
# 4NaIjAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRI
# bmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAL
# BgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBD
# uRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jv
# c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEF
# BQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1h
# cnljcHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkA
# YwB5AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn
# 8oalmOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7
# v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0b
# pdS1HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/
# KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvy
# CInWH8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBp
# mLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJi
# hsMdYzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYb
# BL7fQccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbS
# oqKfenoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sL
# gOppO6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtX
# cVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
# Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAANNTpGmGiiweI8AAAAA
# A00wDQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIGmE
# xt4Vo2SMLvidEoKOy0BWlnwrLtFSz/H6A6/tXaeqMEIGCisGAQQBgjcCAQwxNDAy
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20wDQYJKoZIhvcNAQEBBQAEggEAWRNT4wF78dRbCChPkMmmEMITfsHerfAA340h
# ViK1z5Yj5GYdtsaO6eMNrckGgdhY4ypdzR+wyt3WqENAfgDtlW8qnfb+e5aleTWq
# M0lKrFcmKf/rZWXG1Xbt4OPAsDMiRML8b4oaQK8a6f3xaoFt8QzRhtKmmNgmW1Um
# BAzgvmNy68PhWf+mO0zCYoArHdxkxaspc/YSAQUdQm0aQNaBc5nGvc/jiAP3v2uf
# ZeLYZAkMRZ3x3QzwnMg3+WGc4mlaiyTXnETAO7q8l2yDuGNlidJMX4ZRvIjpgEUU
# mQpGSjI1bwbjTexfSaexY2T3sgEzstp8JjDOLIBabLUWgjC85qGCF5QwgheQBgor
# BgEEAYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZI
# AWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGE
# WQoDATAxMA0GCWCGSAFlAwQCAQUABCBp9m2vxp6Pndq+mQJd4L4+LDvY8xW59oGE
# kE7K0eTYQQIGZVbJFW75GBMyMDIzMTIwNzE4MDAyNC4xMTNaMASAAgH0oIHRpIHO
# MIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQL
# ExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxk
# IFRTUyBFU046QTAwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAdB3CKrvoxfG3QAB
# AAAB0DANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MDAeFw0yMzA1MjUxOTEyMTRaFw0yNDAyMDExOTEyMTRaMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTAwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Uw
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDfMlfn35fvM0XAUSmI5qiG
# 0UxPi25HkSyBgzk3zpYO311d1OEEFz0QpAK23s1dJFrjB5gD+SMw5z6EwxC4CrXU
# 9KaQ4WNHqHrhWftpgo3MkJex9frmO9MldUfjUG56sIW6YVF6YjX+9rT1JDdCDHbo
# 5nZiasMigGKawGb2HqD7/kjRR67RvVh7Q4natAVu46Zf5MLviR0xN5cNG20xwBwg
# ttaYEk5XlULaBH5OnXz2eWoIx+SjDO7Bt5BuABWY8SvmRQfByT2cppEzTjt/fs0x
# p4B1cAHVDwlGwZuv9Rfc3nddxgFrKA8MWHbJF0+aWUUYIBR8Fy2guFVHoHeOze7I
# sbyvRrax//83gYqo8c5Z/1/u7kjLcTgipiyZ8XERsLEECJ5ox1BBLY6AjmbgAzDd
# Nl2Leej+qIbdBr/SUvKEC+Xw4xjFMOTUVWKWemt2khwndUfBNR7Nzu1z9L0Wv7TA
# Y/v+v6pNhAeohPMCFJc+ak6uMD8TKSzWFjw5aADkmD9mGuC86yvSKkII4MayzoUd
# seT0nfk8Y0fPjtdw2Wnejl6zLHuYXwcDau2O1DMuoiedNVjTF37UEmYT+oxC/OFX
# UGPDEQt9tzgbR9g8HLtUfEeWOsOED5xgb5rwyfvIss7H/cdHFcIiIczzQgYnsLyE
# GepoZDkKhSMR5eCB6Kcv/QIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFDPhAYWS0oA+
# lOtITfjJtyl0knRRMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8G
# A1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBs
# BggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUy
# MDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCXh+ckCkZaA06S
# NW+qxtS9gHQp4x7G+gdikngKItEr8otkXIrmWPYrarRWBlY91lqGiilHyIlZ3iNB
# UbaNEmaKAGMZ5YcS7IZUKPaq1jU0msyl+8og0t9C/Z26+atx3vshHrFQuSgwTHZV
# pzv7k8CYnBYoxdhI1uGhqH595mqLvtMsxEN/1so7U+b3U6LCry5uwwcz5+j8Oj0G
# UX3b+iZg+As0xTN6T0Qa8BNec/LwcyqYNEaMkW2VAKrmhvWH8OCDTcXgONnnABQH
# BfXK/fLAbHFGS1XNOtr62/iaHBGAkrCGl6Bi8Pfws6fs+w+sE9r3hX9Vg0gsRMoH
# RuMaiXsrGmGsuYnLn3AwTguMatw9R8U5vJtWSlu1CFO5P0LEvQQiMZ12sQSsQAkN
# DTs9rTjVNjjIUgoZ6XPMxlcPIDcjxw8bfeb4y4wAxM2RRoWcxpkx+6IIf2L+b7gL
# HtBxXCWJ5bMW7WwUC2LltburUwBv0SgjpDtbEqw/uDgWBerCT+Zty3Nc967iGaQj
# yYQH6H/h9Xc8smm2n6VjySRx2swnW3hr6Qx63U/xY9HL6FNhrGiFED7ZRKrnwvvX
# vMVQUIEkB7GUEeN6heY8gHLt0jLV3yzDiQA8R8p5YGgGAVt9MEwgAJNY1iHvH/8v
# zhJSZFNkH8svRztO/i3TvKrjb8ZxwjCCB3EwggVZoAMCAQICEzMAAAAVxedrngKb
# SZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmlj
# YXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIy
# NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UE
# AxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXI
# yjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjo
# YH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1y
# aa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v
# 3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pG
# ve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viS
# kR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYr
# bqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlM
# jgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSL
# W6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AF
# emzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIu
# rQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIE
# FgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWn
# G1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEW
# M2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5
# Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBi
# AEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV
# 9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3Js
# Lm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAx
# MC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2
# LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv
# 6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZn
# OlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1
# bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4
# rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU
# 6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDF
# NLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/
# HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdU
# CbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKi
# excdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTm
# dHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZq
# ELQdVTNYs6FwZvKhggNNMIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJp
# Y2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkEwMDAtMDVF
# MC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMK
# AQEwBwYFKw4DAhoDFQC8t8hT8KKUX91lU5FqRP9Cfu9MiaCBgzCBgKR+MHwxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6RxLljAi
# GA8yMDIzMTIwNzEzNDgwNloYDzIwMjMxMjA4MTM0ODA2WjB0MDoGCisGAQQBhFkK
# BAExLDAqMAoCBQDpHEuWAgEAMAcCAQACAjDRMAcCAQACAhNDMAoCBQDpHZ0WAgEA
# MDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAI
# AgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBAHvwZk3Mzaev9U7tJRsrU56Sk+6d
# RxxO0o3E3/rZ4fJFTTiqk3oc+v9+ONv7iV4C4/mBTVWxht4dgbvKez9f6vjXla3r
# Rtl80gR3tPVrVM5GtAO7vl8NA7cHmPv9rbIl1lNGRS3PuOiJC3MxomYfAbWxYvUU
# MrcytJb3eg10dQmZBsa9DL7XUPrwtE1GwQXVP9xJixgA0Ec22dP1dsKxi8tkCMxF
# 3RdJF74kyyYCd2EfzmxY5X3Mz5eQI3/QT8cQ2cffD+dxekKeLyWvE2WOn8RvubG8
# fs+p7/O4VBo+rYX0dkNPYsY1rgCS1DpIDIlJH4VMBYhRmom7XJJTbXr7mhIxggQN
# MIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdB3
# CKrvoxfG3QABAAAB0DANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
# CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCB7iw3OGYITmuR9heLeuqHpIeqO
# qIECZS4uinSRUoWrhzCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIAiVQAZf
# tNP/Md1E2Yw+fBXa9w6fjmTZ5WAerrTSPwnXMIGYMIGApH4wfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAHQdwiq76MXxt0AAQAAAdAwIgQgIjC5GU5I
# 1vumB/an/LTMlONJuNqUAMNDLVE8un+gLPAwDQYJKoZIhvcNAQELBQAEggIAh0KZ
# uoXY/X0FuusVI5L7ldhfiePqKdS0WyPaV+QJQho5I5rfxOKR7zf5tpF1zPTX+1ln
# gvDS4rjvoCNrhwJVgRtUd8ZCneWkEWvmmDJOEwr/TLqFDW2bI1cpxep/KmjJ/LZ1
# pEOXeECwyyC6OZtsiy/5RtvfkSUI3CP/f0Vjm9nZvap0zQClEWzLjzPq07BleCQH
# t0DXSzeVSx/GmT3EQW5JXPOvj6XA8f3G/yT/ilx7o/Ta3W5/nqVcQmyDS683pnem
# baMthtk2Y7QjtCJsyICfSuMagdmPuPl9sK4FvDa7+SIOrMfjoBKqqAabM47qK41d
# Va0W0KwExrTKkbIl092eembyyNVrGnjLLNn8SmwohbF6l1DALW73aYKZwnKe28Lm
# 4iyFoU3fwC//8jjofxxVD6A1RpU0crOhfb6DWJ36qiRU6VrZcgYJA98O+bbb1eve
# pMJeITHqufUFn+cx7hdXve/gtOvfRNzPYJHeFgsf9qLv4Z+9FJmzasnC1fkLtoG2
# w910HJyE4d2A7r0SeFNQfVjoHS6mZSoBdBiChsGiBMvCTjSVcx+xPAQJgqYpYwFe
# pp9oLfuhc//5kuFHw1pMxzXF/9nBaYoy57TuzG1BaMQ+bg2ANqkFh8yX3ewRD+Ry
# 0HIUZe6JGvebPR6uDmdebz3eJUP3uhjfAijBmKw=
# SIG # End signature block


If ([Environment]::GetEnvironmentVariable($name, $type) -eq $null) {
    Write-Error $strings.EnvironmentErrorDoesNotExists
}
Else {
    [Environment]::SetEnvironmentVariable($name, $null, $type)
}
}
## [END] Remove-WACSMEnvironmentVariable ##
function Restart-WACSMOperatingSystem {
<#

.SYNOPSIS
Reboot Windows Operating System by using Win32_OperatingSystem provider.

.DESCRIPTION
Reboot Windows Operating System by using Win32_OperatingSystem provider.

.ROLE
Administrators

#>
##SkipCheck=true##

Param(
)

import-module CimCmdlets

$instance = Get-CimInstance -Namespace root/cimv2 -ClassName Win32_OperatingSystem

$instance | Invoke-CimMethod -MethodName Reboot

}
## [END] Restart-WACSMOperatingSystem ##
function Set-WACSMComputerIdentification {
<#

.SYNOPSIS
Sets a computer and/or its domain/workgroup information.

.DESCRIPTION
Sets a computer and/or its domain/workgroup information.

.ROLE
Administrators

#>

param(
    [Parameter(Mandatory = $False)]
    [string]
    $ComputerName = '',

    [Parameter(Mandatory = $False)]
    [string]
    $NewComputerName = '',

    [Parameter(Mandatory = $False)]
    [string]
    $Domain = '',

    [Parameter(Mandatory = $False)]
    [string]
    $NewDomain = '',

    [Parameter(Mandatory = $False)]
    [string]
    $Workgroup = '',

    [Parameter(Mandatory = $False)]
    [string]
    $UserName = '',

    [Parameter(Mandatory = $False)]
    [string]
    $Password = '',

    [Parameter(Mandatory = $False)]
    [string]
    $UserNameNew = '',

    [Parameter(Mandatory = $False)]
    [string]
    $PasswordNew = '',

    [Parameter(Mandatory = $False)]
    [switch]
    $Restart)

function CreateDomainCred($username, $password) {
    $secureString = ConvertTo-SecureString $password -AsPlainText -Force
    $domainCreds = New-Object System.Management.Automation.PSCredential($username, $secureString)

    return $domainCreds
}

function UnjoinDomain($domain) {
    If ($domain) {
        $unjoinCreds = CreateDomainCred $UserName $Password
        Remove-Computer -UnjoinDomainCredential $unjoinCreds -PassThru -Force
    }
}

If ($NewDomain) {
    $newDomainCreds = $null
    If ($Domain) {
        UnjoinDomain $Domain
        $newDomainCreds = CreateDomainCred $UserNameNew $PasswordNew
    }
    else {
        $newDomainCreds = CreateDomainCred $UserName $Password
    }

    If ($NewComputerName) {
        Add-Computer -ComputerName $ComputerName -DomainName $NewDomain -Credential $newDomainCreds -Force -PassThru -NewName $NewComputerName -Restart:$Restart
    }
    Else {
        Add-Computer -ComputerName $ComputerName -DomainName $NewDomain -Credential $newDomainCreds -Force -PassThru -Restart:$Restart
    }
}
ElseIf ($Workgroup) {
    UnjoinDomain $Domain

    If ($NewComputerName) {
        Add-Computer -WorkGroupName $Workgroup -Force -PassThru -NewName $NewComputerName -Restart:$Restart
    }
    Else {
        Add-Computer -WorkGroupName $Workgroup -Force -PassThru -Restart:$Restart
    }
}
ElseIf ($NewComputerName) {
    If ($Domain) {
        $domainCreds = CreateDomainCred $UserName $Password
        Rename-Computer -NewName $NewComputerName -DomainCredential $domainCreds -Force -PassThru -Restart:$Restart
    }
    Else {
        Rename-Computer -NewName $NewComputerName -Force -PassThru -Restart:$Restart
    }
}
}
## [END] Set-WACSMComputerIdentification ##
function Set-WACSMEnvironmentVariable {
<#

.SYNOPSIS
Updates or renames an environment variable specified by name, type, data and previous data.

.DESCRIPTION
Updates or Renames an environment variable specified by name, type, data and previrous data.

.ROLE
Administrators

#>

param(
    [Parameter(Mandatory = $True)]
    [String]
    $oldName,

    [Parameter(Mandatory = $True)]
    [String]
    $newName,

    [Parameter(Mandatory = $True)]
    [String]
    $value,

    [Parameter(Mandatory = $True)]
    [String]
    $type
)

Set-StrictMode -Version 5.0

$nameChange = $false
if ($newName -ne $oldName) {
    $nameChange = $true
}

If (-not [Environment]::GetEnvironmentVariable($oldName, $type)) {
    @{ Status = "currentMissing" }
    return
}

If ($nameChange -and [Environment]::GetEnvironmentVariable($newName, $type)) {
    @{ Status = "targetConflict" }
    return
}

If ($nameChange) {
    [Environment]::SetEnvironmentVariable($oldName, $null, $type)
    [Environment]::SetEnvironmentVariable($newName, $value, $type)
    @{ Status = "success" }
}
Else {
    [Environment]::SetEnvironmentVariable($newName, $value, $type)
    @{ Status = "success" }
}


}
## [END] Set-WACSMEnvironmentVariable ##
function Set-WACSMHybridManagement {
<#

.SYNOPSIS
Onboards a machine for hybrid management.

.DESCRIPTION
Sets up a non-Azure machine to be used as a resource in Azure
The supported Operating Systems are Windows Server 2012 R2 and above.

.ROLE
Administrators

.PARAMETER subscriptionId
    The GUID that identifies subscription to Azure services

.PARAMETER resourceGroup
    The container that holds related resources for an Azure solution

.PARAMETER tenantId
    The GUID that identifies a tenant in AAD

.PARAMETER azureRegion
    The region in Azure where the service is to be deployed

.PARAMETER useProxyServer
    The flag to determine whether to use proxy server or not

.PARAMETER proxyServerIpAddress
    The IP address of the proxy server

.PARAMETER proxyServerIpPort
    The IP port of the proxy server

.PARAMETER authToken
    The authentication token for connection

.PARAMETER correlationId
    The correlation ID for the connection (default value is the correlation ID for WAC)

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $subscriptionId,
    [Parameter(Mandatory = $true)]
    [String]
    $resourceGroup,
    [Parameter(Mandatory = $true)]
    [String]
    $tenantId,
    [Parameter(Mandatory = $true)]
    [String]
    $azureRegion,
    [Parameter(Mandatory = $true)]
    [boolean]
    $useProxyServer,
    [Parameter(Mandatory = $false)]
    [String]
    $proxyServerIpAddress,
    [Parameter(Mandatory = $false)]
    [String]
    $proxyServerIpPort,
    [Parameter(Mandatory = $true)]
    [string]
    $authToken,
    [Parameter(Mandatory = $false)]
    [string]
    $correlationId = '88079879-ba3a-4bf7-8f43-5bc912c8cd04'
)

Set-StrictMode -Version 5.0

<#

.SYNOPSIS
Setup script runtime environment.

.DESCRIPTION
Setup script runtime environment.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScript" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Set-HybridManagement.ps1" -Scope Script
    Set-Variable -Name Machine -Option ReadOnly -Value "Machine" -Scope Script
    Set-Variable -Name HybridAgentFile -Option ReadOnly -Value "AzureConnectedMachineAgent.msi" -Scope Script
    Set-Variable -Name HybridAgentPackageLink -Option ReadOnly -Value "https://aka.ms/AzureConnectedMachineAgent" -Scope Script
    Set-Variable -Name HybridAgentExecutable -Option ReadOnly -Value "$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe" -Scope Script
    Set-Variable -Name HttpsProxy -Option ReadOnly -Value "https_proxy" -Scope Script
}

<#

.SYNOPSIS
Cleanup script runtime environment.

.DESCRIPTION
Cleanup script runtime environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name Machine -Scope Script -Force
    Remove-Variable -Name HybridAgentFile -Scope Script -Force
    Remove-Variable -Name HybridAgentPackageLink -Scope Script -Force
    Remove-Variable -Name HybridAgentExecutable -Scope Script -Force
    Remove-Variable -Name HttpsProxy -Scope Script -Force
}

<#

.SYNOPSIS
The main function.

.DESCRIPTION
Export the passed in virtual machine on this server.

#>

function main(
    [string]$subscriptionId,
    [string]$resourceGroup,
    [string]$tenantId,
    [string]$azureRegion,
    [boolean]$useProxyServer,
    [string]$proxyServerIpAddress,
    [string]$proxyServerIpPort,
    [string]$authToken,
    [string]$correlationId
) {
    $err = $null
    $args = @{}

    # Download the package
    Invoke-WebRequest -Uri $HybridAgentPackageLink -OutFile $HybridAgentFile -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Couldn't download the hybrid management package. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }

    # Install the package
    msiexec /i $HybridAgentFile /l*v installationlog.txt /qn | Out-String -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Error while installing the hybrid agent package. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }

    # Set the proxy environment variable. Note that authenticated proxies are not supported for Private Preview.
    if ($useProxyServer) {
        [System.Environment]::SetEnvironmentVariable($HttpsProxy, $proxyServerIpAddress+':'+$proxyServerIpPort, $Machine)
        $env:https_proxy = [System.Environment]::GetEnvironmentVariable($HttpsProxy, $Machine)
    }

    # Run connect command
    $ErrorActionPreference = "Stop"
    & $HybridAgentExecutable connect --resource-group $resourceGroup --tenant-id $tenantId --location $azureRegion `
        --subscription-id $subscriptionId --access-token $authToken --correlation-id $correlationId
    $ErrorActionPreference = "Continue"

    # Restart himds service
    Restart-Service -Name himds -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Couldn't restart the himds service. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return $err
    }
}


###############################################################################
# Script execution starts here
###############################################################################
setupScriptEnv

try {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return main $subscriptionId $resourceGroup $tenantId $azureRegion $useProxyServer $proxyServerIpAddress $proxyServerIpPort $authToken $correlationId

} finally {
    cleanupScriptEnv
}

}
## [END] Set-WACSMHybridManagement ##
function Set-WACSMHyperVEnhancedSessionModeSettings {
<#

.SYNOPSIS
Sets a computer's Hyper-V Host Enhanced Session Mode settings.

.DESCRIPTION
Sets a computer's Hyper-V Host Enhanced Session Mode settings.

.ROLE
Hyper-V-Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [bool]
    $enableEnhancedSessionMode
    )

Set-StrictMode -Version 5.0
Import-Module Hyper-V

# Create arguments
$args = @{'EnableEnhancedSessionMode' = $enableEnhancedSessionMode};

Set-VMHost @args

Get-VMHost | Microsoft.PowerShell.Utility\Select-Object `
    EnableEnhancedSessionMode

}
## [END] Set-WACSMHyperVEnhancedSessionModeSettings ##
function Set-WACSMHyperVHostGeneralSettings {
<#

.SYNOPSIS
Sets a computer's Hyper-V Host General settings.

.DESCRIPTION
Sets a computer's Hyper-V Host General settings.

.ROLE
Hyper-V-Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $virtualHardDiskPath,
    [Parameter(Mandatory = $true)]
    [String]
    $virtualMachinePath
    )

Set-StrictMode -Version 5.0
Import-Module Hyper-V

# Create arguments
$args = @{'VirtualHardDiskPath' = $virtualHardDiskPath};
$args += @{'VirtualMachinePath' = $virtualMachinePath};

Set-VMHost @args

Get-VMHost | Microsoft.PowerShell.Utility\Select-Object `
    VirtualHardDiskPath, `
    VirtualMachinePath

}
## [END] Set-WACSMHyperVHostGeneralSettings ##
function Set-WACSMHyperVHostLiveMigrationSettings {
<#

.SYNOPSIS
Sets a computer's Hyper-V Host Live Migration settings.

.DESCRIPTION
Sets a computer's Hyper-V Host Live Migration settings.

.ROLE
Hyper-V-Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [bool]
    $virtualMachineMigrationEnabled,
    [Parameter(Mandatory = $true)]
    [int]
    $maximumVirtualMachineMigrations,
    [Parameter(Mandatory = $true)]
    [int]
    $virtualMachineMigrationPerformanceOption,
    [Parameter(Mandatory = $true)]
    [int]
    $virtualMachineMigrationAuthenticationType
    )

Set-StrictMode -Version 5.0
Import-Module Hyper-V

if ($virtualMachineMigrationEnabled) {
    $isServer2012 = [Environment]::OSVersion.Version.Major -eq 6 -and [Environment]::OSVersion.Version.Minor -eq 2;
    
    Enable-VMMigration;

    # Create arguments
    $args = @{'MaximumVirtualMachineMigrations' = $maximumVirtualMachineMigrations};
    $args += @{'VirtualMachineMigrationAuthenticationType' = $virtualMachineMigrationAuthenticationType; };

    if (!$isServer2012) {
        $args += @{'VirtualMachineMigrationPerformanceOption' = $virtualMachineMigrationPerformanceOption; };
    }

    Set-VMHost @args;
} else {
    Disable-VMMigration;
}

Get-VMHost | Microsoft.PowerShell.Utility\Select-Object `
    maximumVirtualMachineMigrations, `
    VirtualMachineMigrationAuthenticationType, `
    VirtualMachineMigrationEnabled, `
    VirtualMachineMigrationPerformanceOption

}
## [END] Set-WACSMHyperVHostLiveMigrationSettings ##
function Set-WACSMHyperVHostNumaSpanningSettings {
<#

.SYNOPSIS
Sets a computer's Hyper-V Host settings.

.DESCRIPTION
Sets a computer's Hyper-V Host settings.

.ROLE
Hyper-V-Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [bool]
    $numaSpanningEnabled
    )

Set-StrictMode -Version 5.0
Import-Module Hyper-V

# Create arguments
$args = @{'NumaSpanningEnabled' = $numaSpanningEnabled};

Set-VMHost @args

Get-VMHost | Microsoft.PowerShell.Utility\Select-Object `
    NumaSpanningEnabled

}
## [END] Set-WACSMHyperVHostNumaSpanningSettings ##
function Set-WACSMHyperVHostStorageMigrationSettings {
<#

.SYNOPSIS
Sets a computer's Hyper-V Host Storage Migration settings.

.DESCRIPTION
Sets a computer's Hyper-V Host Storage Migrtion settings.

.ROLE
Hyper-V-Administrators

#>

param (
    [Parameter(Mandatory = $true)]
    [int]
    $maximumStorageMigrations
    )

Set-StrictMode -Version 5.0
Import-Module Hyper-V

# Create arguments
$args = @{'MaximumStorageMigrations' = $maximumStorageMigrations; };

Set-VMHost @args

Get-VMHost | Microsoft.PowerShell.Utility\Select-Object `
    MaximumStorageMigrations

}
## [END] Set-WACSMHyperVHostStorageMigrationSettings ##
function Set-WACSMPowerConfigurationPlan {
<#

.SYNOPSIS
Sets the new power plan

.DESCRIPTION
Sets the new power plan using powercfg when changes are saved by user

.ROLE
Administrators

#>

param(
	[Parameter(Mandatory = $true)]
	[String]
	$PlanGuid
)

$Error.clear()
$message = ""

# If executing an external command, then the following steps need to be done to produce correctly formatted errors:
# Use 2>&1 to store the error to the variable. FD 2 is stderr. FD 1 is stdout.
# Watch $Error.Count to determine the execution result.
# Concatenate the error message to a single string and print it out with Write-Error.
$result = & 'powercfg' /S $PlanGuid 2>&1

# $LASTEXITCODE here does not return error code, so we have to use $Error
if ($Error.Count -ne 0) {
	foreach($item in $result) {
		if ($item.Exception.Message.Length -gt 0) {
			$message += $item.Exception.Message
		}
	}
	$Error.Clear()
	Write-Error $message
}

}
## [END] Set-WACSMPowerConfigurationPlan ##
function Set-WACSMRemoteDesktop {
<#

.SYNOPSIS
Sets a computer's remote desktop settings.

.DESCRIPTION
Sets a computer's remote desktop settings.

.ROLE
Administrators

#>

param(
    [Parameter(Mandatory = $False)]
    [boolean]
    $AllowRemoteDesktop,

    [Parameter(Mandatory = $False)]
    [boolean]
    $AllowRemoteDesktopWithNLA,

    [Parameter(Mandatory=$False)]
    [boolean]
    $EnableRemoteApp)

    Import-Module NetSecurity
    Import-Module Microsoft.PowerShell.Management

function Set-DenyTSConnectionsValue {
    Set-Variable RegistryKey -Option Constant -Value 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
    Set-Variable RegistryKeyProperty -Option Constant -Value 'fDenyTSConnections'

    $KeyPropertyValue = $(if ($AllowRemoteDesktop -eq $True) { 0 } else { 1 })

    if (!(Test-Path $RegistryKey)) {
        New-Item -Path $RegistryKey -Force | Out-Null
    }

    New-ItemProperty -Path $RegistryKey -Name $RegistryKeyProperty -Value $KeyPropertyValue -PropertyType DWORD -Force | Out-Null
}

function Set-UserAuthenticationValue {
    Set-Variable RegistryKey -Option Constant -Value 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'
    Set-Variable RegistryKeyProperty -Option Constant -Value 'UserAuthentication'

    $KeyPropertyValue = $(if ($AllowRemoteDesktopWithNLA -eq $True) { 1 } else { 0 })

    if (!(Test-Path $RegistryKey)) {
        New-Item -Path $RegistryKey -Force | Out-Null
    }

    New-ItemProperty -Path $RegistryKey -Name $RegistryKeyProperty -Value $KeyPropertyValue -PropertyType DWORD -Force | Out-Null
}

function Set-RemoteAppSetting {
    Set-Variable RegistryKey -Option Constant -Value 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\TSAppAllowList'
    Set-Variable RegistryKeyProperty -Option Constant -Value 'fDisabledAllowList'

    $KeyPropertyValue = $(if ($EnableRemoteApp -eq $True) { 1 } else { 0 })

    if (!(Test-Path $RegistryKey)) {
        New-Item -Path $RegistryKey -Force | Out-Null
    }

    New-ItemProperty -Path $RegistryKey -Name $RegistryKeyProperty -Value $KeyPropertyValue -PropertyType DWORD -Force | Out-Null
}

Set-DenyTSConnectionsValue
Set-UserAuthenticationValue
Set-RemoteAppSetting

Enable-NetFirewallRule -Group "@FirewallAPI.dll,-28752" -ErrorAction SilentlyContinue

}
## [END] Set-WACSMRemoteDesktop ##
function Start-WACSMDiskPerf {
<#

.SYNOPSIS
Start Disk Performance monitoring.

.DESCRIPTION
Start Disk Performance monitoring.

.ROLE
Administrators

#>

# Update the registry key at HKLM:SYSTEM\\CurrentControlSet\\Services\\Partmgr
#   EnableCounterForIoctl = DWORD 3
& diskperf -Y

}
## [END] Start-WACSMDiskPerf ##
function Stop-WACSMCimOperatingSystem {
<#

.SYNOPSIS
Shutdown Windows Operating System by using Win32_OperatingSystem provider.

.DESCRIPTION
Shutdown Windows Operating System by using Win32_OperatingSystem provider.

.ROLE
Administrators

#>

##SkipCheck=true##

Param(
[boolean]$primary
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/cimv2 -ClassName Win32_OperatingSystem -Key @('primary') -Property @{primary=$primary;} -ClientOnly
Invoke-CimMethod $keyInstance -MethodName Shutdown

}
## [END] Stop-WACSMCimOperatingSystem ##
function Stop-WACSMDiskPerf {
<#

.SYNOPSIS
Stop Disk Performance monitoring.

.DESCRIPTION
Stop Disk Performance monitoring.

.ROLE
Administrators

#>

# Update the registry key at HKLM:SYSTEM\\CurrentControlSet\\Services\\Partmgr
#   EnableCounterForIoctl = DWORD 1
& diskperf -N


}
## [END] Stop-WACSMDiskPerf ##
function Add-WACSMAdministrators {
<#

.SYNOPSIS
Adds administrators

.DESCRIPTION
Adds administrators

.ROLE
Administrators

#>

Param (
    [Parameter(Mandatory=$true)]
    [String] $usersListString
)


$usersToAdd = ConvertFrom-Json $usersListString
$adminGroup = Get-LocalGroup | Where-Object SID -eq 'S-1-5-32-544'

Add-LocalGroupMember -Group $adminGroup -Member $usersToAdd

Register-DnsClient -Confirm:$false

}
## [END] Add-WACSMAdministrators ##
function Disconnect-WACSMAzureHybridManagement {
<#

.SYNOPSIS
Disconnects a machine from azure hybrid agent.

.DESCRIPTION
Disconnects a machine from azure hybrid agent and uninstall the hybrid instance service.
The supported Operating Systems are Windows Server 2012 R2 and above.

.ROLE
Administrators

.PARAMETER tenantId
    The GUID that identifies a tenant in AAD

.PARAMETER authToken
    The authentication token for connection

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $tenantId,
    [Parameter(Mandatory = $true)]
    [String]
    $authToken
)

Set-StrictMode -Version 5.0

<#

.SYNOPSIS
Setup script runtime environment.

.DESCRIPTION
Setup script runtime environment.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScript" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Disconnect-HybridManagement.ps1" -Scope Script
    Set-Variable -Name HybridAgentExecutable -Option ReadOnly -Value "$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe" -Scope Script
    Set-Variable -Name HybridAgentPackage -Option ReadOnly -Value "Azure Connected Machine Agent" -Scope Script
}

<#

.SYNOPSIS
Cleanup script runtime environment.

.DESCRIPTION
Cleanup script runtime environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name HybridAgentExecutable -Scope Script -Force
    Remove-Variable -Name HybridAgentPackage -Scope Script -Force
}

<#

.SYNOPSIS
The main function.

.DESCRIPTION
Disconnects a machine from azure hybrid agent.

#>

function main(
    [string]$tenantId,
    [string]$authToken
) {
    $err = $null
    $args = @{}

   # Disconnect Azure hybrid agent
   & $HybridAgentExecutable disconnect --access-token $authToken

   # Uninstall Azure hybrid instance metadata service
   Uninstall-Package -Name $HybridAgentPackage -ErrorAction SilentlyContinue -ErrorVariable +err

   if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
        -Message "[$ScriptName]:Could not uninstall the package. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        throw $err
   }

}


###############################################################################
# Script execution starts here
###############################################################################
setupScriptEnv

try {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return main $tenantId $authToken

    return @()
} finally {
    cleanupScriptEnv
}

}
## [END] Disconnect-WACSMAzureHybridManagement ##
function Get-WACSMAzureHybridManagementConfiguration {
<#

.SYNOPSIS
Script that return the hybrid management configurations.

.DESCRIPTION
Script that return the hybrid management configurations.

.ROLE
Administrators

#>

Set-StrictMode -Version 5.0
Import-Module Microsoft.PowerShell.Management

<#

.SYNOPSIS
Setup script runtime environment.

.DESCRIPTION
Onboards a machine for hybrid management.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScript" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Get-HybridManagementConfiguration.ps1" -Scope Script
    Set-Variable -Name HybridAgentExecutable -Option ReadOnly -Value "$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe" -Scope Script
}

<#

.SYNOPSIS
Cleanup script runtime environment.

.DESCRIPTION
Cleanup script runtime environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name HybridAgentExecutable -Scope Script -Force
}

function main() {
    $config = & $HybridAgentExecutable show

    if ($config -and $config.count -gt 10) {
        @{ 
            machine = getValue($config[0]);
            resourceGroup = getValue($config[1]);
            subscriptionId = getValue($config[3]);
            tenantId = getValue($config[4])
            vmId = getValue($config[5]);
            azureRegion = getValue($config[7]);
            agentVersion = getValue($config[10]);
            agentStatus = getValue($config[12]);
            agentLastHeartbeat = getValue($config[13]);
        }
    } else {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
        -Message "[$ScriptName]:Could not find the Azure hybrid agent configuration."  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }
}

function getValue([string]$keyValue) {
    $splitArray = $keyValue -split " : "
    $value = $splitArray[1].trim()
    return $value
}

###############################################################################
# Script execution starts here
###############################################################################
setupScriptEnv

try {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return main

} finally {
    cleanupScriptEnv
}
}
## [END] Get-WACSMAzureHybridManagementConfiguration ##
function Get-WACSMAzureHybridManagementOnboardState {
<#

.SYNOPSIS
Script that returns if Azure Hybrid Agent is running or not.

.DESCRIPTION
Script that returns if Azure Hybrid Agent is running or not.

.ROLE
Readers

#>

Import-Module Microsoft.PowerShell.Management

$status = Get-Service -Name himds -ErrorAction SilentlyContinue
if ($null -eq $status) {
    # which means no such service is found.
    @{ Installed = $false; Running = $false }
}
elseif ($status.Status -eq "Running") {
    @{ Installed = $true; Running = $true }
}
else {
    @{ Installed = $true; Running = $false }
}

}
## [END] Get-WACSMAzureHybridManagementOnboardState ##
function Get-WACSMCimServiceDetail {
<#

.SYNOPSIS
Gets services in details using MSFT_ServerManagerTasks class.

.DESCRIPTION
Gets services in details using MSFT_ServerManagerTasks class.

.ROLE
Readers

#>

##SkipCheck=true##

Param(
)

import-module CimCmdlets

Invoke-CimMethod -Namespace root/microsoft/windows/servermanager -ClassName MSFT_ServerManagerTasks -MethodName GetServerServiceDetail

}
## [END] Get-WACSMCimServiceDetail ##
function Get-WACSMCimSingleService {
<#

.SYNOPSIS
Gets the service instance of CIM Win32_Service class.

.DESCRIPTION
Gets the service instance of CIM Win32_Service class.

.ROLE
Readers

#>

##SkipCheck=true##

Param(
[string]$Name
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/cimv2 -ClassName Win32_Service -Key @('Name') -Property @{Name=$Name;} -ClientOnly
Get-CimInstance $keyInstance

}
## [END] Get-WACSMCimSingleService ##
function Get-WACSMCimWin32LogicalDisk {
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
## [END] Get-WACSMCimWin32LogicalDisk ##
function Get-WACSMCimWin32NetworkAdapter {
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
## [END] Get-WACSMCimWin32NetworkAdapter ##
function Get-WACSMCimWin32PhysicalMemory {
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
## [END] Get-WACSMCimWin32PhysicalMemory ##
function Get-WACSMCimWin32Processor {
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
## [END] Get-WACSMCimWin32Processor ##
function Get-WACSMClusterInventory {
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
## [END] Get-WACSMClusterInventory ##
function Get-WACSMClusterNodes {
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
## [END] Get-WACSMClusterNodes ##
function Get-WACSMDecryptedDataFromNode {
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
## [END] Get-WACSMDecryptedDataFromNode ##
function Get-WACSMEncryptionJWKOnNode {
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
## [END] Get-WACSMEncryptionJWKOnNode ##
function Get-WACSMServerInventory {
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
Get the azure arc status information

.DESCRIPTION
Get the azure arc status information

#>
function getAzureArcStatus() {

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "SMEScript"
  $ScriptName = "Get-ServerInventory.ps1 - getAzureArcStatus()"

  Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  Get-Service -Name himds -ErrorVariable Err -ErrorAction SilentlyContinue | Out-Null

  if (!!$Err) {

    $Err = "The Azure arc agent is not installed. Details: $Err"

    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
    -Message "[$ScriptName]: $Err"  -ErrorAction SilentlyContinue

    $status = "NotInstalled"
  }
  else {
    $status = (azcmagent show --json | ConvertFrom-Json -ErrorAction Stop).status
  }

  return $status
}

<#

.SYNOPSIS
Gets an EnforcementMode that describes the system lockdown policy on this computer.

.DESCRIPTION
By checking the system lockdown policy, we can infer if PowerShell is in ConstrainedLanguage mode as a result of an enforced WDAC policy.
Note: $ExecutionContext.SessionState.LanguageMode should not be used within a trusted (by the WDAC policy) script context for this purpose because
the language mode returned would potentially not reflect the system-wide lockdown policy/language mode outside of the execution context.

#>
function getSystemLockdownPolicy() {
  return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy().ToString()
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
$azureArcStatus = getAzureArcStatus
$systemLockdownPolicy = getSystemLockdownPolicy
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
$result | Add-Member -MemberType NoteProperty -Name 'AzureArcStatus' -Value $azureArcStatus
$result | Add-Member -MemberType NoteProperty -Name 'SystemLockdownPolicy' -Value $systemLockdownPolicy
$result | Add-Member -MemberType NoteProperty -Name 'IsHciServer' -Value $isHciServer

$result

}
## [END] Get-WACSMServerInventory ##
function Resolve-WACSMDNSName {
<#

.SYNOPSIS
Resolve VM Provisioning

.DESCRIPTION
Resolve VM Provisioning

.ROLE
Administrators

#>

Param
(
    [string] $computerName
)

$succeeded = $null
$count = 0;
$maxRetryTimes = 15 * 100 # 15 minutes worth of 10 second sleep times
while ($count -lt $maxRetryTimes)
{
  $resolved =  Resolve-DnsName -Name $computerName -ErrorAction SilentlyContinue

    if ($resolved)
    {
      $succeeded = $true
      break
    }

    $count += 1

    if ($count -eq $maxRetryTimes)
    {
        $succeeded = $false
    }

    Start-Sleep -Seconds 10
}

Write-Output @{ "succeeded" = $succeeded }

}
## [END] Resolve-WACSMDNSName ##
function Resume-WACSMCimService {
<#

.SYNOPSIS
Resume a service using CIM Win32_Service class.

.DESCRIPTION
Resume a service using CIM Win32_Service class.

.ROLE
Readers

#>

##SkipCheck=true##

Param(
[string]$Name
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/cimv2 -ClassName Win32_Service -Key @('Name') -Property @{Name=$Name;} -ClientOnly
Invoke-CimMethod $keyInstance -MethodName ResumeService

}
## [END] Resume-WACSMCimService ##
function Set-WACSMAzureHybridManagement {
<#

.SYNOPSIS
Onboards a machine for hybrid management.

.DESCRIPTION
Sets up a non-Azure machine to be used as a resource in Azure
The supported Operating Systems are Windows Server 2012 R2 and above.

.ROLE
Administrators

.PARAMETER subscriptionId
    The GUID that identifies subscription to Azure services

.PARAMETER resourceGroup
    The container that holds related resources for an Azure solution

.PARAMETER tenantId
    The GUID that identifies a tenant in AAD

.PARAMETER azureRegion
    The region in Azure where the service is to be deployed

.PARAMETER useProxyServer
    The flag to determine whether to use proxy server or not

.PARAMETER proxyServerIpAddress
    The IP address of the proxy server

.PARAMETER proxyServerIpPort
    The IP port of the proxy server

.PARAMETER authToken
    The authentication token for connection

.PARAMETER correlationId
    The correlation ID for the connection

#>

param (
    [Parameter(Mandatory = $true)]
    [String]
    $subscriptionId,
    [Parameter(Mandatory = $true)]
    [String]
    $resourceGroup,
    [Parameter(Mandatory = $true)]
    [String]
    $tenantId,
    [Parameter(Mandatory = $true)]
    [String]
    $azureRegion,
    [Parameter(Mandatory = $true)]
    [boolean]
    $useProxyServer,
    [Parameter(Mandatory = $false)]
    [String]
    $proxyServerIpAddress,
    [Parameter(Mandatory = $false)]
    [String]
    $proxyServerIpPort,
    [Parameter(Mandatory = $true)]
    [string]
    $authToken,
    [Parameter(Mandatory = $true)]
    [string]
    $correlationId
)

Set-StrictMode -Version 5.0

<#

.SYNOPSIS
Setup script runtime environment.

.DESCRIPTION
Setup script runtime environment.

#>

function setupScriptEnv() {
    Set-Variable -Name LogName -Option ReadOnly -Value "Microsoft-ServerManagementExperience" -Scope Script
    Set-Variable -Name LogSource -Option ReadOnly -Value "SMEScript" -Scope Script
    Set-Variable -Name ScriptName -Option ReadOnly -Value "Set-HybridManagement.ps1" -Scope Script
    Set-Variable -Name Machine -Option ReadOnly -Value "Machine" -Scope Script
    Set-Variable -Name HybridAgentFile -Option ReadOnly -Value "AzureConnectedMachineAgent.msi" -Scope Script
    Set-Variable -Name HybridAgentPackageLink -Option ReadOnly -Value "https://aka.ms/AzureConnectedMachineAgent" -Scope Script
    Set-Variable -Name HybridAgentExecutable -Option ReadOnly -Value "$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe" -Scope Script
    Set-Variable -Name HttpsProxy -Option ReadOnly -Value "https_proxy" -Scope Script
}

<#

.SYNOPSIS
Cleanup script runtime environment.

.DESCRIPTION
Cleanup script runtime environment.

#>

function cleanupScriptEnv() {
    Remove-Variable -Name LogName -Scope Script -Force
    Remove-Variable -Name LogSource -Scope Script -Force
    Remove-Variable -Name ScriptName -Scope Script -Force
    Remove-Variable -Name Machine -Scope Script -Force
    Remove-Variable -Name HybridAgentFile -Scope Script -Force
    Remove-Variable -Name HybridAgentPackageLink -Scope Script -Force
    Remove-Variable -Name HybridAgentExecutable -Scope Script -Force
    Remove-Variable -Name HttpsProxy -Scope Script -Force
}

<#

.SYNOPSIS
The main function.

.DESCRIPTION
Export the passed in virtual machine on this server.

#>

function main(
    [string]$subscriptionId,
    [string]$resourceGroup,
    [string]$tenantId,
    [string]$azureRegion,
    [boolean]$useProxyServer,
    [string]$proxyServerIpAddress,
    [string]$proxyServerIpPort,
    [string]$authToken,
    [string]$correlationId
) {
    $err = $null
    $args = @{}

    # Download the package
    Invoke-WebRequest -Uri $HybridAgentPackageLink -OutFile $HybridAgentFile -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Couldn't download the hybrid management package. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }

    # Install the package
    msiexec /i $HybridAgentFile /l*v installationlog.txt /qn | Out-String -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Error while installing the hybrid agent package. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return @()
    }

    # Set the proxy environment variable. Note that authenticated proxies are not supported for Private Preview.
    if ($useProxyServer) {
        [System.Environment]::SetEnvironmentVariable($HttpsProxy, $proxyServerIpAddress+':'+$proxyServerIpPort, $Machine)
        $env:https_proxy = [System.Environment]::GetEnvironmentVariable($HttpsProxy, $Machine)
    }

    # Run connect command
    & $HybridAgentExecutable connect --resource-group $resourceGroup --tenant-id $tenantId --location $azureRegion `
                                     --subscription-id $subscriptionId --access-token $authToken --correlation-id $correlationId

    # Restart himds service
    Restart-Service -Name himds -ErrorAction SilentlyContinue -ErrorVariable +err
    if ($err) {
        Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
            -Message "[$ScriptName]:Couldn't restart the himds service. Error: $err"  -ErrorAction SilentlyContinue

        Write-Error @($err)[0]
        return $err
    }
}


###############################################################################
# Script execution starts here
###############################################################################
setupScriptEnv

try {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    return main $subscriptionId $resourceGroup $tenantId $azureRegion $useProxyServer $proxyServerIpAddress $proxyServerIpPort $authToken $correlationId

} finally {
    cleanupScriptEnv
}

}
## [END] Set-WACSMAzureHybridManagement ##
function Set-WACSMVMPovisioning {
<#

.SYNOPSIS
Prepare VM Provisioning

.DESCRIPTION
Prepare VM Provisioning

.ROLE
Administrators

#>

Param (
    [Parameter(Mandatory = $true)]
    [array]$disks
)

$output = @{ }

$requiredDriveLetters = $disks.driveLetter
$volumeLettersInUse = (Get-Volume | Sort-Object DriveLetter).DriveLetter

$output.Set_Item('restartNeeded', $false)
$output.Set_Item('pageFileLetterChanged', $false)
$output.Set_Item('pageFileLetterNew', $null)
$output.Set_Item('pageFileLetterOld', $null)
$output.Set_Item('pageFileDiskNumber', $null)
$output.Set_Item('cdDriveLetterChanged', $false)
$output.Set_Item('cdDriveLetterNew', $null)
$output.Set_Item('cdDriveLetterOld', $null)

$cdDriveLetterNeeded = $false
$cdDrive = Get-WmiObject -Class Win32_volume -Filter 'DriveType=5' | Microsoft.PowerShell.Utility\Select-Object -First 1
if ($cdDrive -ne $null) {
    $cdDriveLetter = $cdDrive.DriveLetter.split(':')[0]
    $output.Set_Item('cdDriveLetterOld', $cdDriveLetter)

    if ($requiredDriveLetters.Contains($cdDriveLetter)) {
        $cdDriveLetterNeeded = $true
    }
}

$pageFileLetterNeeded = $false
$pageFile = Get-WmiObject Win32_PageFileusage
if ($pageFile -ne $null) {
    $pagingDriveLetter = $pageFile.Name.split(':')[0]
    $output.Set_Item('pageFileLetterOld', $pagingDriveLetter)

    if ($requiredDriveLetters.Contains($pagingDriveLetter)) {
        $pageFileLetterNeeded = $true
    }
}

if ($cdDriveLetterNeeded -or $pageFileLetterNeeded) {
    $capitalCCharNumber = 67;
    $capitalZCharNumber = 90;

    for ($index = $capitalCCharNumber; $index -le $capitalZCharNumber; $index++) {
        $tempDriveLetter = [char]$index

        $willConflict = $requiredDriveLetters.Contains([string]$tempDriveLetter)
        $inUse = $volumeLettersInUse.Contains($tempDriveLetter)
        if (!$willConflict -and !$inUse) {
            if ($cdDriveLetterNeeded) {
                $output.Set_Item('cdDriveLetterNew', $tempDriveLetter)
                $cdDrive | Set-WmiInstance -Arguments @{DriveLetter = $tempDriveLetter + ':' } > $null
                $output.Set_Item('cdDriveLetterChanged', $true)
                $cdDriveLetterNeeded = $false
            }
            elseif ($pageFileLetterNeeded) {

                $computerObject = Get-WmiObject Win32_computersystem -EnableAllPrivileges
                $computerObject.AutomaticManagedPagefile = $false
                $computerObject.Put() > $null

                $currentPageFile = Get-WmiObject Win32_PageFilesetting
                $currentPageFile.delete() > $null

                $diskNumber = (Get-Partition -DriveLetter $pagingDriveLetter).DiskNumber

                $output.Set_Item('pageFileLetterNew', $tempDriveLetter)
                $output.Set_Item('pageFileDiskNumber', $diskNumber)
                $output.Set_Item('pageFileLetterChanged', $true)
                $output.Set_Item('restartNeeded', $true)
                $pageFileLetterNeeded = $false
            }

        }
        if (!$cdDriveLetterNeeded -and !$pageFileLetterNeeded) {
            break
        }
    }
}

# case where not enough drive letters available after iterating through C-Z
if ($cdDriveLetterNeeded -or $pageFileLetterNeeded) {
    $output.Set_Item('preProvisioningSucceeded', $false)
}
else {
    $output.Set_Item('preProvisioningSucceeded', $true)
}


Write-Output $output


}
## [END] Set-WACSMVMPovisioning ##
function Start-WACSMCimService {
<#

.SYNOPSIS
Start a service using CIM Win32_Service class.

.DESCRIPTION
Start a service using CIM Win32_Service class.

.ROLE
Administrators

#>

##SkipCheck=true##

Param(
[string]$Name
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/cimv2 -ClassName Win32_Service -Key @('Name') -Property @{Name=$Name;} -ClientOnly
Invoke-CimMethod $keyInstance -MethodName StartService

}
## [END] Start-WACSMCimService ##
function Start-WACSMVMProvisioning {
<#

.SYNOPSIS
Execute VM Provisioning

.DESCRIPTION
Execute VM Provisioning

.ROLE
Administrators

#>

Param (
    [Parameter(Mandatory = $true)]
    [bool] $partitionDisks,

    [Parameter(Mandatory = $true)]
    [array]$disks,

    [Parameter(Mandatory = $true)]
    [bool]$pageFileLetterChanged,

    [Parameter(Mandatory = $false)]
    [string]$pageFileLetterNew,

    [Parameter(Mandatory = $false)]
    [int]$pageFileDiskNumber,

    [Parameter(Mandatory = $true)]
    [bool]$systemDriveModified
)

$output = @{ }

$output.Set_Item('restartNeeded', $pageFileLetterChanged)

if ($pageFileLetterChanged) {
    Get-Partition -DiskNumber $pageFileDiskNumber | Set-Partition -NewDriveLetter $pageFileLetterNew
    $newPageFile = $pageFileLetterNew + ':\pagefile.sys'
    Set-WMIInstance -Class Win32_PageFileSetting -Arguments @{name = $newPageFile; InitialSize = 0; MaximumSize = 0 } > $null
}

if ($systemDriveModified) {
    $size = Get-PartitionSupportedSize -DriveLetter C
    Resize-Partition -DriveLetter C -Size $size.SizeMax > $null
}

if ($partitionDisks -eq $true) {
    $dataDisks = Get-Disk | Where-Object PartitionStyle -eq 'RAW' | Sort-Object Number
    for ($index = 0; $index -lt $dataDisks.Length; $index++) {
        Initialize-Disk  $dataDisks[$index].DiskNumber -PartitionStyle GPT -PassThru |
        New-Partition -Size $disks[$index].volumeSizeInBytes -DriveLetter $disks[$index].driveLetter |
        Format-Volume -FileSystem $disks[$index].fileSystem -NewFileSystemLabel $disks[$index].name -Confirm:$false -Force > $null;
    }
}

Write-Output $output

}
## [END] Start-WACSMVMProvisioning ##
function Suspend-WACSMCimService {
<#

.SYNOPSIS
Suspend a service using CIM Win32_Service class.

.DESCRIPTION
Suspend a service using CIM Win32_Service class.

.ROLE
Administrators

#>

##SkipCheck=true##

Param(
[string]$Name
)

import-module CimCmdlets

$keyInstance = New-CimInstance -Namespace root/cimv2 -ClassName Win32_Service -Key @('Name') -Property @{Name=$Name;} -ClientOnly
Invoke-CimMethod $keyInstance -MethodName PauseService

}
## [END] Suspend-WACSMCimService ##

# SIG # Begin signature block
# MIIoKgYJKoZIhvcNAQcCoIIoGzCCKBcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDpQvGLiZa0r46A
# y9JkQ+ZfwA10ZBFErP941wU0q47m+6CCDXYwggX0MIID3KADAgECAhMzAAADTrU8
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIKKIKP5fhA4vrd/O6CIkXnZl
# 3C710uQghvDjBPhNgLQ6MEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAzGGKMrqju92l9sJniQNphR4b7TdN7jk+iB/W61zBOKQ4u75/5S3QL54C
# GQnZpa1yO5fUJlwezvmY0F0LbeK3YY5b2OzOFb25cjOt5YHKd5LZTAddMZCIhhxG
# MRGcN4oZCPyD7QfSNmGt7RR1iOr6Jxznp5nLfl7xVYkQhyFn2VumC4RCAnA+iEPw
# mn3exo8AfRzPyzarxFn5yrjJZfxNIM3ZjY04ANAREynJeUD0OLd+XE6GiYcG+8iw
# BDWim28o4MzXPAPbShn56DzX+68igSQdzKlqB7gP0b1yDlC0ta7+xQD0Z3gu7KGd
# Vvji7TrDvFUbQTsCsm4cDVb/zwgRLaGCF5QwgheQBgorBgEEAYI3AwMBMYIXgDCC
# F3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCBRdW7jbHLAPdP6+l6+r86CNgBatoP3LGKRy/saj15AOAIGZWita03x
# GBMyMDIzMTIwNzE4MDAyMS41MDZaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OTYwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wg
# ghHqMIIHIDCCBQigAwIBAgITMwAAAdj8SzOlHdiFFQABAAAB2DANBgkqhkiG9w0B
# AQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzA1MjUxOTEy
# NDBaFw0yNDAyMDExOTEyNDBaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OTYwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDNeOsp0fXgAz7GUF0N+/0EHcQFri6wliTbmQNmFm8D
# i0CeQ8n4bd2td5tbtzTsEk7dY2/nmWY9kqEvavbdYRbNc+Esv8Nfv6MMImH9tCr5
# Kxs254MQ0jmpRucrm3uHW421Cfva0hNQEKN1NS0rad1U/ZOme+V/QeSdWKofCThx
# f/fsTeR41WbqUNAJN/ml3sbOH8aLhXyTHG7sVt/WUSLpT0fLlNXYGRXzavJ1qUOe
# Pzyj86hiKyzQJLTjKr7GpTGFySiIcMW/nyK6NK7Rjfy1ofLdRvvtHIdJvpmPSze3
# CH/PYFU21TqhIhZ1+AS7RlDo18MSDGPHpTCWwo7lgtY1pY6RvPIguF3rbdtvhoyj
# n5mPbs5pgjGO83odBNP7IlKAj4BbHUXeHit3Da2g7A4jicKrLMjo6sGeetJoeKoo
# j5iNTXbDwLKM9HlUdXZSz62ftCZVuK9FBgkAO9MRN2pqBnptBGfllm+21FLk6E3v
# VXMGHB5eOgFfAy84XlIieycQArIDsEm92KHIFOGOgZlWxe69leXvMHjYJlpo2VVM
# tLwXLd3tjS/173ouGMRaiLInLm4oIgqDtjUIqvwYQUh3RN6wwdF75nOmrpr8wRw1
# n/BKWQ5mhQxaMBqqvkbuu1sLeSMPv2PMZIddXPbiOvAxadqPkBcMPUBmrySYoLTx
# wwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFPbTj0x8PZBLYn0MZBI6nGh5qIlWMB8G
# A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4w
# XAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2Vy
# dHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQCunA6aSP48oJ1VD+SMF1/7SFiTGD6zyLC3
# Ju9HtLjqYYq1FJWUx10I5XqU0alcXTUFUoUIUPSvfeX/dX0MgofUG+cOXdokaHHS
# lo6PZIDXnUClpkRix9xCN37yFBpcwGLzEZlDKJb2gDq/FBGC8snTlBSEOBjV0eE8
# ICVUkOJzIAttExaeQWJ5SerUr63nq6X7PmQvk1OLFl3FJoW4+5zKqriY/PKGssOa
# A5ZjBZEyU+o7+P3icL/wZ0G3ymlT+Ea4h9f3q5aVdGVBdshYa/SehGmnUvGMA8j5
# Ct24inx+bVOuF/E/2LjIp+mEary5mOTrANVKLym2kW3eQxF/I9cj87xndiYH55Xf
# rWMk9bsRToxOpRb9EpbCB5cSyKNvxQ8D00qd2TndVEJFpgyBHQJS/XEK5poeJZ5q
# gmCFAj4VUPB/dPXHdTm1QXJI3cO7DRyPUZAYMwQ3KhPlM2hP2OfBJIr/VsDsh3sz
# LL2ZJuerjshhxYGVboMud9aNoRjlz1Mcn4iEota4tam24FxDyHrqFm6EUQu/pDYE
# DquuvQFGb5glIck4rKqBnRlrRoiRj0qdhO3nootVg/1SP0zTLC1RrxjuTEVe3PKr
# ETbtvcODoGh912Xrtf4wbMwpra8jYszzr3pf0905zzL8b8n8kuMBChBYfFds916K
# Tjc4TGNU9TCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZI
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
# MCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjk2MDAtMDVFMC1EOTQ3MSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQBI
# p++xUJ+f85VrnbzdkRMSpBmvL6CBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6RxmOjAiGA8yMDIzMTIwNzE1NDE0
# NloYDzIwMjMxMjA4MTU0MTQ2WjB0MDoGCisGAQQBhFkKBAExLDAqMAoCBQDpHGY6
# AgEAMAcCAQACAhPpMAcCAQACAhO4MAoCBQDpHbe6AgEAMDYGCisGAQQBhFkKBAIx
# KDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZI
# hvcNAQELBQADggEBAAAz2HRk7+/xRlHR6KQ8qjWfR0wRYQqKP/wFpK4xAbATtRF9
# f5PgEl4pFgdP9vRTAc7M9vlBD7xPQcEpzMJemNXLl3kcON8xvJE45OL4u5z5thao
# POO+XTnZvSVe87uMz9ymx1apyceEXn1RsOkB7cso9miwWay8Ep+vpicp32d/Nlzp
# 8o0A0JtAI3Wj0Ig8+7KVWSlrk5meSpMJVQUJBfECn/O/9VwRxpOS6Rzxo3a5PB1A
# TrSBoQS67ktIehEbyTDFIdMOCEiXoMNARWiw7xhAhVjRK2/6Wy/gvXZmIAl/bc6D
# MnvdVqDgDx3On3IfHy4UVHflrMNovQGyx06914QxggQNMIIECQIBATCBkzB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdj8SzOlHdiFFQABAAAB2DAN
# BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8G
# CSqGSIb3DQEJBDEiBCDH+8LEVoQt+3cnqjB3P6M/8A0HlBwsQkkQWofWP8chEjCB
# +gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIDrjIX/8CZN3RTABMNt5u73Mi3o3
# fmvq2j8Sik+2s75UMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAHY/EszpR3YhRUAAQAAAdgwIgQg20gM8bcnlNFHQ9on8zio2r15GX62
# WbfJRGUgew+rJPowDQYJKoZIhvcNAQELBQAEggIAdaWXYBAVDV2I8w8pr5wjm3CR
# CJ8LN0MoV1OTTz3+gjPSbi+xDKuvVdUvw8z33D5PFs3kI+ETSDz+Zr7iBC2bxmbm
# yXtxRKPgDfcnEr0Um5GWnni49Jr7VBPQmYbKGCG1xb8dJhsZod75UMq2ADP4ypwr
# wak1v8GyD/Ee4XS67XvEZL9XcAxMTN39rO8Qdzs8vrkDNN+aI9atKxvjHSVGDIsu
# NgoPvdQAhFbWqnWg/ikUKXM+il69RL1LDbGIxbS2ysKo4JWohsEe+U7JTjXq8/em
# +D/DZU/cb8UV3zySlek1wisdkQIQ2QHn9YlVVe+p4TGx9VuoQyBJLEDGJ5q308O3
# kJl2+vRyi85RcGbd/bcMJzVpWDyEWMZjk4YfA0xF43JIm5SVdIiDi/kCd4I5SsfU
# hg1IrwqWz1bdVkm5Lz5gQI2Utp8qaZ4qHqwSU4b40ehb4FsU9Sco1vtZ1jC+ADId
# Qy9gPI+NY2MHMRsvPE79AeJZSkJEPkn5UDzn3hmLqPWTuG4/g7FGyf7wklYt1Rq7
# fTP5ebDzzVLdgbz3jpb0pUc+QB0UjN3tc6A2FfMZ3ek02fPggHofiUVuct3cPvGf
# rMFuRaymUMmTdpgYelw75O+IVLrM9AdJxS0Y/kHo5iYrkQ3UReoc/ZO1fJhp5zpX
# YhOlXLAGxl+UGd23Wgs=
# SIG # End signature block
