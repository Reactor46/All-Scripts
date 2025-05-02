function Add-WACSEWdacPolicy {
<#
.SYNOPSIS
    Add a Windows Defender Application Control (WDAC) supplemental policy
.DESCRIPTION
    Add a Windows Defender Application Control (WDAC) supplemental policy
.ROLE
    Administrators
#>

param (
    [Parameter(Mandatory = $true)]
    [String]$filePath
)
    
Set-StrictMode -Version 5.0;

$Script:eventId = 0
function Write-WdacEventLog {
    param (
        [Parameter(Mandatory = $false)]
        [String]$EntryType,
        [Parameter(Mandatory = $true)]
        [String]$Message
    )

    if (!$entryType) {
        $entryType = 'Error'
    }

    $LogName = "Microsoft-ServerManagementExperience"
    $LogSource = "msft.sme.security"
    $ScriptName = "Add-WdacPolicy.ps1"

    # Create the event log if it does not exists
    New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    # EntryType: Error, Information, FailureAudit, SuccessAudit, Warning
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId $Script:eventId -Category 0 -EntryType $EntryType `
        -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue

    $Script:eventId += 1
}

function Add-WdacPolicy {
    param (
        [Parameter(Mandatory = $true)]
        [String]$filePath
    )

    $policyPath = [IO.Path]::GetFullPath( $filePath )
    Add-ASLocalWDACSupplementalPolicy -Path $policyPath -ErrorAction SilentlyContinue -ErrorVariable err

    # See https://github.com/PowerShell/PowerShell/pull/10840
    if (!!$err -and $err[0] -isnot [System.Management.Automation.FlowControlException]) {
        $errorMessage = "There was an error adding the supplemental WDAC policy.  Error: $err. File: $policyPath"
        Write-WdacEventLog -message $errorMessage
        throw $err
    }

    [xml]$xmlFile = Get-Content -Path $policyPath -ErrorAction SilentlyContinue
    $policyId = $xmlFile.SiPolicy.PolicyID
    Copy-Item -Path "$env:InfraCSVRootFolderPath\CloudMedia\Security\WDAC\Stage\$policyId.cip" -Destination "$env:InfraCSVRootFolderPath\CloudMedia\Security\WDAC\Active" -Force -ErrorAction SilentlyContinue -ErrorVariable err

    if (!!$err) {
        $errorMessage = "There was an error applying the supplemental WDAC policy. Error: $err."
        Write-WdacEventLog -message $errorMessage
        throw $err
    }
}

###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
    $errorMessage = "Couldn't add supplemental WDAC policy. Module 'Microsoft.AS.Infra.Security.WDAC' does not exist on this server. Ensure that you are running the latest version of Azure Stack HCI."

    Import-Module -Name Microsoft.AS.Infra.Security.WDAC -ErrorAction SilentlyContinue -ErrorVariable err

    if (!!$err) {
        Write-WdacEventLog -Message $errorMessage
        Write-Error -Message $errorMessage
    }
    else {
        Add-WdacPolicy -FilePath $filePath
    }
}

}
## [END] Add-WACSEWdacPolicy ##
function Get-WACSEAszProperties {
<#

.SYNOPSIS
Check-ASZ

.DESCRIPTION
Checks if server has ASZ deployed

.ROLE
Readers

#>

$bitlockerModuleExists = $null -ne (Get-Command -Module AzureStackBitlockerAgent)
$osConfigAgentExists = $null -ne (Get-Command -Module AzureStackOSConfigAgent)
$wdacModuleExists = $null -ne (Get-Command -Module Microsoft.AS.Infra.Security.WDAC)
$serverInfo = Get-CimInstance Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object operatingSystemSKU, buildNumber

$aszProperties = @{}
$aszProperties.Add("bitlockerModuleExists", $bitlockerModuleExists)
$aszProperties.Add("osConfigAgentExists", $osConfigAgentExists)
$aszProperties.Add("wdacModuleExists", $wdacModuleExists)
$aszProperties.Add("serverInfo", $serverInfo)

$aszProperties

}
## [END] Get-WACSEAszProperties ##
function Get-WACSECluster {
<#
.SYNOPSIS
Gets cluster object

.DESCRIPTION
Gets cluster object

.ROLE
Readers

#>

Import-Module FailoverClusters
FailoverClusters\Get-Cluster | Microsoft.PowerShell.Utility\Select-Object securityLevel, securityLevelForStorage

}
## [END] Get-WACSECluster ##
function Get-WACSEClusterSecuritySettings {
<#
.SYNOPSIS
Get array of Azure Stack HCI Security Settings

.DESCRIPTION
Get array of Azure Stack HCI Security Settings
    [
        Boot volume bitlocker encyption
        Side Channel Mitigation
        Credential Guard
        SMB signing
        Drift Control
    ]

.ROLE
Administrators

#>

function Get-SecurityBaseline {
  $securityBaselineDocument = Get-OsConfigurationDocumentResult -Id 64329a05-92b9-450e-a0b3-b2f9185100c1 | ConvertFrom-Json
  $status = ($securityBaselineDocument.OsConfiguration.Scenario).status

  foreach ($setting in $status) {
    if ($setting.state -ne "completed") {
      return $false
    }
  }

  return $true
}

function Get-AszSecuritySettings {
  # Get Boot volume bitlocker encyption setting
  $bootVolumeBitlockerEncryption = Get-ASLocalBitlockerEnforced

  # Get Side Channel Mitigation setting
  $sideChannelMitigation = Get-ASSecurity -FeatureName SideChannelMitigation -Local

  # Get Credential Guard setting
  $credentialGuard = Get-ASSecurity -FeatureName CredentialGuard -Local

  # Get SMB signing setting
  $smbSigning = Get-ASSecurity -FeatureName SMBSigning -Local

  # Get Drift Control setting
  $driftControl = Get-ASSecurity -FeatureName DriftControl -Local

  # Get Security Baseline
  $securityBaseline = Get-SecurityBaseline


  $settingsStatus = @{}
  $settingsStatus.Add("bootVolumeBitlockerEncryption", $bootVolumeBitlockerEncryption)
  $settingsStatus.Add("sideChannelMitigation", $sideChannelMitigation)
  $settingsStatus.Add("credentialGuard", $credentialGuard)
  $settingsStatus.Add("smbSigning", $smbSigning)
  $settingsStatus.Add("driftControl", $driftControl)
  $settingsStatus.Add("securityBaseline", $securityBaseline)

  $settingsStatus
}

$Script:eventId = 0
function Write-GetAszSettingsEventLog {
  param (
    [Parameter(Mandatory = $false)]
    [String]$entryType,
    [Parameter(Mandatory = $true)]
    [String]$message
  )

  if (!$entryType) {
    $entryType = 'Error'
  }

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "msft.sme.security"
  $ScriptName = "Get-ClusterSecuritySettings.ps1"

  # Create the event log if it does not exists
  New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  # EntryType: Error, Information, FailureAudit, SuccessAudit, Warning
  Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId $Script:eventId -Category 0 -EntryType $EntryType `
    -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue

  $Script:eventId += 1
}

###############################################################################
# Script execution starts here
###############################################################################
$bitlockerModuleExists = $null -ne (Get-Command -Module AzureStackBitlockerAgent)
$osConfigAgentExists = $null -ne (Get-Command -Module AzureStackOSConfigAgent)
if ($bitlockerModuleExists -and $osConfigAgentExists) {
  Get-AszSecuritySettings
}
else {
  Write-GetAszSettingsEventLog -Message "Couldn't query ASZ Security Settings. Modules 'AzureStackBitlockerAgent' and 'AzureStackOSConfigAgent' do not exist on this server. Ensure that you are running the latest version of Azure Stack HCI."
}

}
## [END] Get-WACSEClusterSecuritySettings ##
function Get-WACSEClusterSpecificSettings {
<#
.SYNOPSIS
Get array of cluster-specific Azure Stack HCI Security Settings

.DESCRIPTION
Get array of cluster-specific Azure Stack HCI Security Settings
    [
        Data Volume Bitlocker
        SMB Cluster encryption
    ]

.ROLE
Readers

#>

Set-StrictMode -Version 5.0;

# Get SMB Cluster encryption setting
$smbClusterEncryption = Get-ASSecurity -FeatureName SMBEncryption -Local


# Get Data volumes bitlocker setting
# (Only a cluster cmdlet - no local support for data volume encryption)
# TODO 03-03-2023: Update when ECE removal is done. We cannot use ECE with WAC since ECE requires CredSSP
# $dataVolumeEncryption = Get-ASBitlockerDataVolumeEncryptionStatus

$settingsStatus = @{}
# $settingsStatus.Add("dataVolumeEncryption", $dataVolumeEncryption)
$settingsStatus.Add("smbClusterEncryption", $smbClusterEncryption)

$settingsStatus

}
## [END] Get-WACSEClusterSpecificSettings ##
function Get-WACSEComputerInfo {
<#

.SYNOPSIS
Get-ComputerInfo

.DESCRIPTION
Gets OSDisplayVersion, OsOperatingSystemSKU, and OsBuildNumber information

.ROLE
Readers

#>

$computerInfo = Get-ComputerInfo | Microsoft.PowerShell.Utility\Select-Object OSDisplayVersion, OsOperatingSystemSKU
$computerInfo | Add-Member -Name 'OsBuildNumber' -Type NoteProperty -Value ([System.Environment]::OSVersion.Version.Build)

$computerInfo

}
## [END] Get-WACSEComputerInfo ##
function Get-WACSENodeDataVolumesBitlockerStatus {
<#
.SYNOPSIS
Get array of cluster shared data volumes bitlocker status

.DESCRIPTION
Get array of cluster shared data volumes bitlocker status

.ROLE
Readers

#>

Set-StrictMode -Version 5.0;

$Script:eventId = 0
function Write-WdacEventLog {
    param (
        [Parameter(Mandatory = $false)]
        [String]$EntryType,
        [Parameter(Mandatory = $true)]
        [String]$Message
    )

    if (!$entryType) {
        $entryType = 'Error'
    }

    $LogName = "Microsoft-ServerManagementExperience"
    $LogSource = "msft.sme.security"
    $ScriptName = "Get-NodeDataVolumesBitlockerStatus.ps1"

    # Create the event log if it does not exists
    New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    # EntryType: Error, Information, FailureAudit, SuccessAudit, Warning
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId $Script:eventId -Category 0 -EntryType $EntryType `
        -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue

    $Script:eventId += 1
}

function Get-BitlockerStatus {
    $result = @()

    $bitLockerVolumes = Get-AsBitlocker -volumeType ClusterSharedVolume -Local -ErrorAction SilentlyContinue -ErrorVariable err
    if (!!$err) {
        $errorMessage = "There was an error getting the bitlocker status. Error: $err."
        Write-WdacEventLog -message $errorMessage
        throw $err
    }

    foreach ($volume in $bitLockerVolumes) {
        $volumeInfo = Get-Volume -FilePath $volume.mountPoint
        $volumeName = $volumeInfo.FileSystemLabel
        $path = $volumeInfo.Path
        $volumeId = (Split-Path -Path $path -Leaf).Split('{}')[1]

        $newObject = $volume | Microsoft.PowerShell.Utility\Select-Object *, @{Name = 'VolumeId'; Expression = { $volumeId } }, @{Name = 'VolumeName'; Expression = { $volumeName } }
        $result += $newObject 
    }

    return $result
}

###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
    $errorMessage = "Couldn't get shared volumes bitlocker status. Module 'Microsoft.AS.Infra.Security.WDAC' does not exist on this server. Ensure that you are running the latest version of Azure Stack HCI."

    Import-Module -Name Microsoft.AS.Infra.Security.WDAC -ErrorAction SilentlyContinue -ErrorVariable err

    if (!!$err) {
        Write-WdacEventLog -Message $errorMessage
        Write-Error -Message $errorMessage
    }
    else {
        Get-BitlockerStatus
    }
}

}
## [END] Get-WACSENodeDataVolumesBitlockerStatus ##
function Get-WACSEPreferenceActions {
<#

.SYNOPSIS
Get Actions for Threats.

.DESCRIPTION
Get Custom SetActions Which Are More Prefferable.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0;

$Preference = Get-MpPreference

$preferenceDefaultActions = $Preference.ThreatIDDefaultAction_Actions
$preferenceDefaultActionsIDs = $Preference.ThreatIDDefaultAction_Ids

$actionsHash = @{}

$actionsHash.Add('Actions', $preferenceDefaultActions)
$actionsHash.Add('Ids', $preferenceDefaultActionsIDs)

return $actionsHash
}
## [END] Get-WACSEPreferenceActions ##
function Get-WACSERealTimeMonitoringState {
<#

.SYNOPSIS
Get Real Time Monitoring State.

.DESCRIPTION
Get Real Time Monitoring State.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0;

$Preference = Get-MpPreference

return $Preference.DisableRealtimeMonitoring
}
## [END] Get-WACSERealTimeMonitoringState ##
function Get-WACSESecuredCoreFeatures {
<#
.SYNOPSIS
Get Secured-Core Features

.DESCRIPTION
Get array of secured-core features
    [
        TPM 2.0,
        Secure Boot,
        VBS,
        HVCI,
        Boot DMA Protection,
        SystemGuard
    ]

.ROLE
Readers

#>

Set-StrictMode -Version 5.0;

# tpm version check
function CheckTpmVersion {
  $TpmObj = Get-CimInstance -classname Win32_Tpm -namespace root\cimv2\Security\MicrosoftTpm

  if ($null -ne $TpmObj) {
    return $TpmObj.SpecVersion[0] -eq "2"
  }

  return $false
}

<#
Check whether VBS is enabled and running
0.	VBS is not enabled.
1.	VBS is enabled but not running.
2.	VBS is enabled and running.
#>
function CheckVBS {
  return (Get-CimInstance -classname Win32_DeviceGuard -namespace root\Microsoft\Windows\DeviceGuard).VirtualizationBasedSecurityStatus
}

<#
# device guard checked used for hcvi and system guard
0.	No services running.
1.	If present, Windows Defender Credential Guard is running.
2.	If present, HVCI is running.
3.	If present, System Guard Secure Launch is running.
4.	If present, SMM Firmware Measurement is running.
#>
function CheckDGSecurityServicesRunning($_val) {
  $DGObj = Get-CimInstance -classname Win32_DeviceGuard -namespace root\Microsoft\Windows\DeviceGuard

  # loop to avoid out of index out of bounds errors
  for ($i = 0; $i -lt $DGObj.SecurityServicesRunning.length; $i++) {
    if ($DGObj.SecurityServicesRunning[$i] -eq $_val) {
      return $true
    }
  }

  return $false
}

<#
Indicates whether the Windows Defender Credential Guard or HVCI service has been configured.
0.	No services configured.
1.	If present, Windows Defender Credential Guard is configured.
2.	If present, HVCI is configured.
3.	If present, System Guard Secure Launch is configured.
4.	If present, SMM Firmware Measurement is configured.
#>
function CheckDGSecurityServicesConfigured($_val) {
  $DGObj = Get-CimInstance -classname Win32_DeviceGuard -namespace root\Microsoft\Windows\DeviceGuard
  if ($_val -in $DGObj.SecurityServicesConfigured) {
    return $true
  }

  return $false
}

# bootDMAProtection check
$bootDMAProtectionCheck =
@"
  namespace SystemInfo
    {
      using System;
      using System.Runtime.InteropServices;

      public static class NativeMethods
      {
        internal enum SYSTEM_DMA_GUARD_POLICY_INFORMATION : int
        {
            /// </summary>
            SystemDmaGuardPolicyInformation = 202
        }

        [DllImport("ntdll.dll")]
        internal static extern Int32 NtQuerySystemInformation(
          SYSTEM_DMA_GUARD_POLICY_INFORMATION SystemDmaGuardPolicyInformation,
          IntPtr SystemInformation,
          Int32 SystemInformationLength,
          out Int32 ReturnLength);

        public static byte BootDmaCheck() {
          Int32 result;
          Int32 SystemInformationLength = 1;
          IntPtr SystemInformation = Marshal.AllocHGlobal(SystemInformationLength);
          Int32 ReturnLength;

          result = NativeMethods.NtQuerySystemInformation(
                    NativeMethods.SYSTEM_DMA_GUARD_POLICY_INFORMATION.SystemDmaGuardPolicyInformation,
                    SystemInformation,
                    SystemInformationLength,
                    out ReturnLength);

          if (result == 0) {
            byte info = Marshal.ReadByte(SystemInformation, 0);
            return info;
          }

          return 0;
        }
      }
    }
"@
Add-Type -TypeDefinition $bootDMAProtectionCheck

function checkSecureBoot {
  if ((Get-Command Confirm-SecureBootUEFI -ErrorAction  SilentlyContinue) -ne $null) {
    <#
    For devices that Standard hardware security is not supported, this means that the device does not meet
    at least one of the requirements of standard hardware security.
    This causes the Confirm-SecureBootUEFI command to fail with the error:
      Cmdlet not supported on this platform: 0xC0000002
   #>
    try {
      return Confirm-SecureBootUEFI
    }
    catch {
      return $false
    }
  }
  return $false
}


###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
  $securedCoreFeatures = @{}

  # Status: Security is running
  # Configured: Security service is enabled/configured
  $TPM20Obj = @{"Status" = CheckTpmVersion; "Configured" = $null }
  $secureBoot = @{"Status" = checkSecureBoot; "Configured" = $null }
  $bootDMAProtection = @{"Status" = ([SystemInfo.NativeMethods]::BootDmaCheck()) -ne 0; "Configured" = $null }

  $vbsStatus = [int](CheckVBS)
  $vbsRunning = if ($vbsStatus -eq 2) { $true } else { $false }
  $vbsConfigured = if ($vbsStatus -gt 0) { $true } else { $false }
  $VBS = @{"Status" = $vbsRunning; "Configured" = $vbsConfigured }
  $HVCI = @{"Status" = CheckDGSecurityServicesRunning(2); "Configured" = CheckDGSecurityServicesConfigured(2) }
  $systemGuard = @{"Status" = CheckDGSecurityServicesRunning(3); "Configured" = CheckDGSecurityServicesConfigured(3) }

  $securedCoreFeatures.Add("tpm20", $TPM20Obj)
  $securedCoreFeatures.Add("secureBoot", $secureBoot)
  $securedCoreFeatures.Add("bootDMAProtection", $bootDMAProtection)
  $securedCoreFeatures.Add("vbs", $VBS)
  $securedCoreFeatures.Add("hvci", $HVCI)
  $securedCoreFeatures.Add("systemGuard", $systemGuard)

  $securedCoreFeatures
}

}
## [END] Get-WACSESecuredCoreFeatures ##
function Get-WACSESecuredCoreOsConfigFeatures {
<#
.SYNOPSIS
Get Secured-Core Features using OsConfiguration module

.DESCRIPTION
Get array of secured-core features
    [
        TPM 2.0,
        Secure Boot,
        VBS,
        HVCI,
        Boot DMA Protection,
        SystemGuard
    ]

.ROLE
Administrators

#>

Set-StrictMode -Version 5.0;

# Set OsConfiguration document and get the result, or $null on failure
function OsConfigurationSetDocumentGetResult {

  [CmdletBinding()]
  Param (
    [Parameter(Mandatory)]
    [String] $Id,

    [Parameter(Mandatory)]
    [String] $Content
  )

  # Set the document to get securedcore settings
  #Set-OsConfigurationDocument -Content $Content -Wait -TimeoutInSeconds 300
  Set-OsConfigurationDocument -Content $Content -Wait

  $result = Get-OsConfigurationDocumentResult -Id $Id | ConvertFrom-Json

  return $result.OsConfiguration.Scenario[0]
}

# Use OsConfiguration to check SecuredCore states
$jsonDocumentToGetSecuredCoreSettingStates =
@"
{
  "OsConfiguration":{
      "Document":{
        "schemaversion":"1.0",
        "id":"10088660-1861-4131-96e8-f32e85011100",
        "version":"10056C2C71F6A41F9AB4A601AD00C8B5BC7531576233010B13A221A9FE1BE100",
        "context":"device",
        "scenario":"SecuredCoreState"
      },
      "Scenario":[
        {
            "name":"SecuredCoreState",
            "schemaversion":"1.0",
            "action":"get",
            "SecuredCoreState":{
              "VirtualizationBasedSecurityStatus": "0",
              "HypervisorEnforcedCodeIntegrityStatus": "0",
              "SystemGuardStatus": "0",
              "SecureBootState": "0",
              "TPMVersion": "",
              "BootDMAProtection": "0"
            }
        }
      ]
  }
}
"@

function GetSecuredCoreSettingStates {

  # Set the document to get securedcore settings
  $result = OsConfigurationSetDocumentGetResult -Id "10088660-1861-4131-96e8-f32e85011100" -Content $jsonDocumentToGetSecuredCoreSettingStates
  $Script:securedCoreStatus = $result.status

  return $result.SecuredCoreState
}

$jsonDocumentToGetSecuredCoreSettingConfigurations =
@"
{
  "OsConfiguration":{
      "Document":{
        "schemaversion":"1.0",
        "id":"47e88660-1861-4131-96e8-f32e85011e55",
        "version":"3C356C2C71F6A41F9AB4A601AD00C8B5BC7531576233010B13A221A9FE1BE7A0",
        "context":"device",
        "scenario":"SecuredCore"
      },
      "Scenario":[
        {
            "name":"SecuredCore",
            "schemaversion":"1.0",
            "action":"get",
            "SecuredCore":{
              "EnableVirtualizationBasedSecurity": "0",
              "HypervisorEnforcedCodeIntegrity": "0",
              "ConfigureSystemGuardLaunch": "0"
            }
        }
      ]
  }
}
"@

function GetSecuredCoreSettingConfigurations {

  # Set the document to get securedcore settings
  $result = OsConfigurationSetDocumentGetResult -Id "47e88660-1861-4131-96e8-f32e85011e55" -Content $jsonDocumentToGetSecuredCoreSettingConfigurations

  return $result.SecuredCore
}

function CheckTpmVersion {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory)]
    $status
  )
  if ([bool]($SecuredCoreStates.PSobject.Properties.name -match "TPMVersion") -and $status -ne "failed") {
    return $null -ne $SecuredCoreStates.TPMVersion -and $SecuredCoreStates.TPMVersion[0] -eq "2"
  }
  return $false
}

function getFeatureStatus {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory)]
    [String] $securedCoreState,

    [Parameter(Mandatory)]
    $expectedStateValue
  )

  foreach($status in $Script:securedCoreStatus) {
    # handle tpm separately to make sure it exists
    if ($securedCoreState -eq "TPMVersion") {
      return CheckTpmVersion $status.state
    }
    # check for failed status
    if ($status.name -eq $securedCoreState -and $status.state -eq "failed") {
      return $false
    }
    return ($SecuredCoreStates | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty $securedCoreState) -eq $expectedStateValue
  }
}

$SecuredCoreStates = GetSecuredCoreSettingStates
$SecuredCoreConfigurations = GetSecuredCoreSettingConfigurations

# String with a list of TPM version such as "2.0, 0, 1.16"
$TPM20Obj = @{"Status" = getFeatureStatus "TPMVersion" "2"; "Configured" = $null }
# Indicates whether secure boot is enabled. The value is one of the following: 0 - Not supported, 1 - Enabled, 2 - Disabled
$secureBoot = @{"Status" = getFeatureStatus "SecureBootState" 1; "Configured" = $null }
# Boot DMA Protection status. 1 - Enabled, 2 - Disabled
$bootDMAProtection = @{"Status" = getFeatureStatus "BootDMAProtection" 1; "Configured" = $null }
# Virtualization-based security status. Value is one of the following: 0 - Running, 1 - Reboot required, 2 - 64 bit architecture required, 3 - not licensed, 4 - not configured, 5 - System doesn't meet hardware requirements, 42 - Other. Event logs in Microsoft-Windows-DeviceGuard have more details
$VBS = @{"Status" = getFeatureStatus "VirtualizationBasedSecurityStatus" 0; "Configured" = $SecuredCoreConfigurations.EnableVirtualizationBasedSecurity -eq 1 }
# Hypervisor Enforced Code Integrity (HVCI) status. 0 - Running, 1 - Reboot required, 2 - Not configured, 3 - VBS not running
$HVCI = @{"Status" = getFeatureStatus "HypervisorEnforcedCodeIntegrityStatus" 0; "Configured" = $SecuredCoreConfigurations.HypervisorEnforcedCodeIntegrity -eq 2 }
# System Guard status. 0 - Running, 1 - Reboot required, 2 - Not configured, 3 - System doesn't meet hardware requirements
$systemGuard = @{"Status" = getFeatureStatus "SystemGuardStatus" 0; "Configured" = $SecuredCoreConfigurations.ConfigureSystemGuardLaunch -eq 1 }

$securedCoreFeatures = @{}

$securedCoreFeatures.Add("tpm20", $TPM20Obj)
$securedCoreFeatures.Add("secureBoot", $secureBoot)
$securedCoreFeatures.Add("bootDMAProtection", $bootDMAProtection)
$securedCoreFeatures.Add("vbs", $VBS)
$securedCoreFeatures.Add("hvci", $HVCI)
$securedCoreFeatures.Add("systemGuard", $systemGuard)

$securedCoreFeatures

}
## [END] Get-WACSESecuredCoreOsConfigFeatures ##
function Get-WACSEStatusSummary {
<#

.SYNOPSIS
Get Summary Object.

.DESCRIPTION
Get Summary:
    - Latest Detected Threat
    - Latest Scan Date and Type
    - Scheduled Scan Date and Type
    - Threat Definition Version
    - Version Creation Date.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0;

$ComputerStatus = Get-MpComputerStatus;
$ThreatDetection = Get-MpThreatDetection;
$MpPreference = Get-MpPreference;

$summaryHash = @{}

# Latest Threat Detected
if ($ThreatDetection) {
    $threatsTime = $ThreatDetection.InitialDetectionTime
    if ($threatsTime) {
        if ($threatsTime -is [system.array]) {
            $summaryHash.Add('latestThreatTime', $threatsTime[$threatsTime.Length-1]);
        } else {
            $summaryHash.Add('latestThreatTime', $threatsTime);
        }
    }
} else {
    $summaryHash.Add('latestThreatDate', $null);
}

# Latest Scan
$latestScanTime = $null
$latestScanType = $null
if ($null -eq $ComputerStatus.QuickScanStartTime ) {
    if ($null -eq $ComputerStatus.FullScanStartTime) {
        $summaryHash.Add('latestScanTime', $null);
        $summaryHash.Add('latestScanType', $null);
    }
    else {
        $latestScanTime = $ComputerStatus.FullScanStartTime;
        $latestScanType = 2;
    }
}
else {
    if ($null -eq $ComputerStatus.FullScanStartTime) {
        $latestScanTime = $ComputerStatus.QuickScanStartTime
        $latestScanType = 1;
    }
    else {
        if ($ComputerStatus.QuickScanStartTime -gt $ComputerStatus.FullScanStartTime) {
            $latestScanTime = $ComputerStatus.QuickScanStartTime
            $latestScanType = 1;
        }
        else {
            $latestScanTime = $ComputerStatus.FullScanStartTime
            $latestScanType = 2;
        }
    }
}

if ($summaryHash.ContainsKey('latestScanTime')) {
  $summaryHash.latestScanTime = $latestScanTime
} else {
  $summaryHash.Add('latestScanTime', $latestScanTime)
}

if ($summaryHash.ContainsKey('latestScanType')) {
  $summaryHash.latestScanType = $latestScanType
} else {
  $summaryHash.Add('latestScanType', $latestScanType)
}

# Next Scheduled Scan
$scanDay = [int]$MpPreference.ScanScheduleDay
$scanMinutes = [int]$MpPreference.ScanScheduleTime.TotalMinutes
$scanTimeSpan = New-Timespan -minutes $scanMinutes
$scanDateTime = [DateTime]($scanTimeSpan.Ticks);
$scanTime = $scanDateTime.ToString('t')
$scanType = [int]$MpPreference.ScanParameters
$summaryHash.Add('scanDay', $scanDay)
$summaryHash.Add('scanTime', $scanTime)
$summaryHash.Add('scanType', $scanType)

return $summaryHash

}
## [END] Get-WACSEStatusSummary ##
function Get-WACSEThreatDetections {
<#

.SYNOPSIS
Get Array of Detected Threats.

.DESCRIPTION
Get Array of Detected Threat Object
    [
        ThreatID,
        Detected Threat Name,
        Detected File,
        Time and Date,
        Threat Alert Level,
        Threat Status,
        Threat Category,
        Threat Default Action
    ].

.ROLE
Readers

#>

Set-StrictMode -Version 5.0;

$ThreatDetection = Get-MpThreatDetection

$tableEntryObjectsArray = @()

foreach ($detectedThreat in $ThreatDetection) {
    $tableEntryHash = @{}

    $threatID = $detectedThreat.ThreatID
    $tableEntryHash.Add('ThreatID', [string]$threatID)

    # Detected Threat
    $threat = Get-MpThreat -ThreatID $threatID
    $threatName = $threat.ThreatName
    $tableEntryHash.Add('DetectedThreat', $threatName)

    # Item
    $threatFileLongName = $threat.Resources[0]
    $threatFileSplitted = $threatFileLongName -split "file:_"
    $tableEntryHash.Add('Item', $threatFileSplitted[1])

    $dateTime = $detectedThreat.InitialDetectionTime
    $tableEntryHash.Add('DateTime', $dateTime.ToString('g'))

    $tableEntryHash.Add('AlertLevel', $threat.SeverityID)

    $tableEntryHash.Add('Status', $detectedThreat.ThreatStatusID)

    $tableEntryHash.Add('Category', $threat.CategoryID)

    $tableEntryHash.Add('DefaultAction', $detectedThreat.CleaningActionID)

    $tableEntryObject = New-Object -TypeName psobject -Property $tableEntryHash
    $tableEntryObjectsArray += $tableEntryObject
}

return $tableEntryObjectsArray

}
## [END] Get-WACSEThreatDetections ##
function Get-WACSEWdacPolicyInfo {
<#
.SYNOPSIS
Get Windows Defender Application Control (WDAC) Policy Info

.DESCRIPTION
Get Windows Defender Application Control (WDAC) Policy Info

.ROLE
Administrators

#>

Set-StrictMode -Version 5.0;

$Script:eventId = 0
function Write-GetWdacEventLog {
  param (
    [Parameter(Mandatory = $false)]
    [String]$entryType,
    [Parameter(Mandatory = $true)]
    [String]$message
  )

  if (!$entryType) {
    $entryType = 'Error'
  }

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "msft.sme.security"
  $ScriptName = "Set-WdacPolicyMode.ps1"

  # Create the event log if it does not exists
  New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  # EntryType: Error, Information, FailureAudit, SuccessAudit, Warning
  Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId $Script:eventId -Category 0 -EntryType $EntryType `
    -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue

  $Script:eventId += 1
}

###############################################################################
# Script execution starts here
###############################################################################
$wdacModuleExists = $null -ne (Get-Command -Module Microsoft.AS.Infra.Security.WDAC)
if ($wdacModuleExists) {
  # Refresh policy when we have some event flooding issue
  Invoke-WDACRefreshPolicyTool | Out-Null

  Get-ASLocalWDACPolicyInfo
} else {
  Write-GetWdacEventLog -Message "Couldn't query WDAC policy. Module 'Microsoft.AS.Infra.Security.WDAC' does not exist on this server. Ensure that you are running the latest version of Azure Stack HCI."
}

}
## [END] Get-WACSEWdacPolicyInfo ##
function Get-WACSEWdacPolicyMode {
<#
.SYNOPSIS
Get Windows Defender Application Control (WDAC) Policy setting

.DESCRIPTION
Get Windows Defender Application Control (WDAC) Policy setting
Returned values
    0: Not deployed
    1: Audit
    2: Enforcement

.ROLE
Readers

#>

Set-StrictMode -Version 5.0;

Get-ASLocalWDACPolicyMode

}
## [END] Get-WACSEWdacPolicyMode ##
function Remove-WACSEWdacPolicy {
<#
.SYNOPSIS
    Remove a Windows Defender Application Control (WDAC) supplemental policy
.DESCRIPTION
    Remove a Windows Defender Application Control (WDAC) supplemental policy
.ROLE
    Administrators
#>

param (
    [Parameter(Mandatory = $true)]
    [String]$policyGuid
)
    
Set-StrictMode -Version 5.0;

$Script:eventId = 0
function Write-WdacEventLog {
    param (
        [Parameter(Mandatory = $false)]
        [String]$EntryType,
        [Parameter(Mandatory = $true)]
        [String]$Message
    )

    if (!$entryType) {
        $entryType = 'Error'
    }

    $LogName = "Microsoft-ServerManagementExperience"
    $LogSource = "msft.sme.security"
    $ScriptName = "Add-WdacPolicy.ps1"

    # Create the event log if it does not exists
    New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    # EntryType: Error, Information, FailureAudit, SuccessAudit, Warning
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId $Script:eventId -Category 0 -EntryType $EntryType `
        -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue

    $Script:eventId += 1
}

function Remove-WdacPolicy {
    param (
        [Parameter(Mandatory = $true)]
        [String]$policyGuid
    )

    Remove-ASLocalWDACSupplementalPolicy -PolicyGuid $policyGuid -ErrorAction SilentlyContinue -ErrorVariable err

    # See https://github.com/PowerShell/PowerShell/pull/10840
    if (!!$err -and $err[0] -isnot [System.Management.Automation.FlowControlException]) {
        $errorMessage = "There was an error removing the supplemental WDAC policy.  Error: $err. Policy: $policyGuid"
        Write-WdacEventLog -message $errorMessage
        throw $err
    }

    $policyFilePath = "$env:InfraCSVRootFolderPath\CloudMedia\Security\WDAC\Active\$policyGuid.cip"

    if (Test-Path -Path $policyFilePath -PathType Leaf) {
        Remove-Item -Path $policyFilePath -Force -ErrorAction SilentlyContinue -ErrorVariable err

        if (!!$err) {
            $errorMessage = "There was an error removing the supplemental WDAC policy. Error: $err. File: $policyGuid"
            Write-WdacEventLog -message $errorMessage
            throw $err
        }
    }
}

###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
    $errorMessage = "Couldn't remove supplemental WDAC policy. Module 'Microsoft.AS.Infra.Security.WDAC' does not exist on this server. Ensure that you are running the latest version of Azure Stack HCI."

    Import-Module -Name Microsoft.AS.Infra.Security.WDAC -ErrorAction SilentlyContinue -ErrorVariable err

    if (!!$err) {
        Write-WdacEventLog -Message $errorMessage
        Write-Error -Message $errorMessage
    }
    else {
        Remove-WdacPolicy -PolicyGuid $policyGuid
    }
}

}
## [END] Remove-WACSEWdacPolicy ##
function Set-WACSEClusterSecuritySettings {
<#
.SYNOPSIS
Script that enables and disables Azure Stack HCI Security Settings

.DESCRIPTION
Script that enables and disables Azure Stack HCI Security Settings

.Parameter SettingName
Name of cluster security setting whose status needs to be toggled. Either one of the following:
 bootVolumeBitlockerEncryption, sideChannelMitigation, credentialGuard,
 smbClusterEncryption, smbSigning, driftControl

.Parameter action
Action to perform, either enable (1) or disable (0) a cluster security setting
Allowed value:
  0: Disable
  1: Enable

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [String]$settingName,
  [Parameter(Mandatory = $true)]
  [String]$action,
  [Parameter()]
  [array]$mountPoints
)

# Enable Action
function EnableClusterSecuritySetting {
  [CmdletBinding()]
  param (
    [Parameter()]
    [string]$SettingName,
    [Parameter()]
    [array]$mountPoints
  )

  switch ($SettingName) {
    "bootVolumeBitlockerEncryption" {
      Enable-ASHostLocalVolumeEncryption
    }
    "sideChannelMitigation" {
      # Enabling these settings requires system reboot for the settings to take effect
      Enable-ASSecurity -FeatureName SideChannelMitigation -Local
    }
    "credentialGuard" {
      # Enabling these settings requires system reboot for the settings to take effect
      Enable-ASSecurity -FeatureName CredentialGuard -Local
    }
    "smbClusterEncryption" {
      Enable-ASSecurity -FeatureName SMBEncryption -Local
    }
    "smbSigning" {
      Enable-ASSecurity -FeatureName SMBSigning -Local
    }
    "driftControl" {
      Enable-ASSecurity -FeatureName DriftControl -Local
    }
    "dataVolumeBitLocker" {
      $mountPoints | ForEach-Object {
        Enable-ASBitlocker -VolumeType ClusterSharedVolume -Local -MountPoint $_
      }
    }
  }
}


# Disable Action
function DisableClusterSecuritySetting {
  [CmdletBinding()]
  param (
    [Parameter()]
    [string]$SettingName,
    [Parameter()]
    [array]$mountPoints
  )

  switch ($SettingName) {
    "bootVolumeBitlockerEncryption" {
      Disable-ASHostLocalVolumeEncryption
    }
    "sideChannelMitigation" {
      # Disabling these settings requires system reboot for the settings to take effect
      # Applying these settings will put your system at risk to silicon-based microarchitectural
      # and speculative execution side-channel vulnerabilities
      # To run this and bypass the confirmation prompt, add -Confirm:$false
      Disable-ASSecurity -FeatureName SideChannelMitigation -Local -Confirm:$false
    }
    "credentialGuard" {
      # Disabling these settings requires system reboot for the settings to take effect
      Disable-ASSecurity -FeatureName CredentialGuard -Local
    }
    "smbClusterEncryption" {
      Disable-ASSecurity -FeatureName SMBEncryption -Local
    }
    "smbSigning" {
      Disable-ASSecurity -FeatureName SMBSigning -Local
    }
    "driftControl" {
      # By disabling OSConfig drift control the system will no longer be able to auto-correct
      # any out-of-band security related changes
      # To run this and cmdlet by pass the confirmation prompt use -Confirm:$false
      Disable-ASSecurity -FeatureName DriftControl -Local -Confirm:$false
    }
    "dataVolumeBitLocker" {
      $mountPoints | ForEach-Object {
        Disable-ASBitlocker -VolumeType ClusterSharedVolume -Local -MountPoint $_
      }
    }
  }
}


Add-Type -TypeDefinition @"
   public enum ActionType {
        Disable,
        Enable
    }
"@

function ToggleSecurityFeature {
  [CmdletBinding()]
  param (
    [Parameter()]
    [string]$action,
    [Parameter()]
    [String]$settingName,
    [Parameter()]
    [array]$mountPoints
  )
  if ([ActionType]$action -eq [ActionType]::Enable) {
    EnableClusterSecuritySetting -SettingName $settingName -MountPoints $mountPoints
  }
  elseif ([ActionType]$action -eq [ActionType]::Disable) {
    DisableClusterSecuritySetting -SettingName $settingName -MountPoints $mountPoints
  }
  else {
    $LogName = "Microsoft-ServerManagementExperience"
    $LogSource = "msft.sme.security"
    $ScriptName = "Set-ClusterSecuritySettings.ps1"
    $Message = "Invalid toggle action passed: $action"
    $EntryType = 'Error'

    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource `
      -EventId 0 -Category 0 -EntryType $EntryType `
      -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue
  }
}

$Script:eventId = 0
function Write-SetAszSettingsEventLog {
  param (
    [Parameter(Mandatory = $false)]
    [String]$entryType,
    [Parameter(Mandatory = $true)]
    [String]$message
  )

  if (!$entryType) {
    $entryType = 'Error'
  }

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "msft.sme.security"
  $ScriptName = "Set-ClusterSecuritySettings.ps1"

  # Create the event log if it does not exists
  New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  # EntryType: Error, Information, FailureAudit, SuccessAudit, Warning
  Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId $Script:eventId -Category 0 -EntryType $EntryType `
    -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue

  $Script:eventId += 1
}

###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
  $bitlockerModuleExists = $null -ne (Get-Command -Module AzureStackBitlockerAgent)
  $osConfigAgentExists = $null -ne (Get-Command -Module AzureStackOSConfigAgent)
  if ($bitlockerModuleExists -and $osConfigAgentExists) {
    ToggleSecurityFeature -Action $action -SettingName $settingName -MountPoints $mountPoints
  } else {
    Write-SetAszSettingsEventLog -Message "Couldn't toggle ASZ Security Settings. Modules 'AzureStackBitlockerAgent' and 'AzureStackOSConfigAgent' do not exist on this server. Ensure that you are running the latest version of Azure Stack HCI."
  }
}

}
## [END] Set-WACSEClusterSecuritySettings ##
function Set-WACSEClusterTrafficEncryption {
<#

.SYNOPSIS
Sets cluster traffic encryption

.DESCRIPTION
Sets cluster traffic encryption

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [uint32]
  $encryptCoreTraffic,
  [Parameter(Mandatory = $false)]
  [uint32]
  $encryptStorageTraffic
)
Import-Module FailoverClusters
$cluster = Get-Cluster;

$cluster.SecurityLevel = $encryptCoreTraffic

if ($encryptStorageTraffic)
{
  $cluster.SecurityLevelForStorage = $encryptStorageTraffic
}

}
## [END] Set-WACSEClusterTrafficEncryption ##
function Set-WACSERealTimeMonitoringState {
<#

.SYNOPSIS
Set Real Time Monitoring State.

.DESCRIPTION
Set Real Time Monitoring State to On or Off.

.ROLE
Administrators

#>

Param(
    [Int32]$DisableRealtimeMonitoring
)

Set-StrictMode -Version 5.0;

Set-MpPreference -DisableRealtimeMonitoring $DisableRealtimeMonitoring
}
## [END] Set-WACSERealTimeMonitoringState ##
function Set-WACSEScheduledScan {
<#

.SYNOPSIS
Set Date, Time and Type for Recurrent Scan.

.DESCRIPTION
Set Date, Time and Type for Recurrent Scan.

.Parameter ScanParameters
Specifies the scan type to use during a scheduled scan. The acceptable values for this parameter are:
  1: Quick scan
  2: Full scan

.Parameter ScanScheduleDay
Specifies the day of the week on which to perform a scheduled scan.
Alternatively, specify everyday for a scheduled scan or never. The acceptable values for this parameter are:
  0: Everyday
  1: Sunday
  2: Monday
  3: Tuesday
  4: Wednesday
  5: Thursday
  6: Friday
  7: Saturday
  8: Never

.ROLE
Administrators

#>

Param(
    [Int32]$ScanParameters,
    [Int32]$ScanScheduleDay
)

switch ($ScanParameters) {
  1 { $ScanType = 'QuickScan' }
  2 { $ScanType = 'FullScan' }
}

Set-StrictMode -Version 5.0;
Set-MpPreference -ScanParameters $ScanType -ScanScheduleDay $ScanScheduleDay

}
## [END] Set-WACSEScheduledScan ##
function Set-WACSESecuredCoreFeatures {
<#
.SYNOPSIS
Script that enables and disables Secured Core Features

.DESCRIPTION
Script that enables and disables Secured Core Features
  1. You CAN enable configurable code integrity without either HVCI or Cred Guard.
  2. You CAN enable HVCI without either configurable code integrity or Cred Guard.
  3. You CAN enable Cred Guard without either configurable code integrity or HVCI.
  4. You CANNOT enable either Cred Guard or HVCI without Virtualization Based Security.

.Parameter selectedFeatures
All selected features to toggle on/off.

.Parameter action
Value to set to either enable (1) or disable (0) feature

.Parameter secureBoot
Set RequirePlatformSecurityFeatures to 1 (Secure Boot only) or 3 (Secure Boot and DMA protection)

.Parameter featureDetails
An list of object containing the details (e.g. status) of the secured core features

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [PSCustomObject[]]$selectedFeatures,
  [Parameter(Mandatory = $true)]
  [String]$action,
  [Parameter(Mandatory = $false)]
  [String]$secureBoot,
  [Parameter(Mandatory = $false)]
  [PSCustomObject[]]$featureDetails
)


$Script:action = [int]$action
if (-not($Script:action -in @(0, 1))) {
  Throw "Invalid value for parameter $Script:action. Use 0 to disable or 1 to enable a secured core feature"
}
$Script:actionWord = if ($Script:action -eq 0 ) { 'disable' } else { 'enable' };


$Script:secureBoot = [int]$secureBoot
if (-not($Script:secureBoot -in @(0, 1, 3))) {
  Throw "Invalid value for parameter $Script:secureBoot. Use 0 (Disable RequirePlatformSecurityFeatures), 1 (Secure Boot only) or 3 (Secure Boot and DMA protection)"
}

$Script:selectedFeatures = $selectedFeatures
$Script:featureDetails = $featureDetails

<##
  https://social.technet.microsoft.com/wiki/contents/articles/36183.windows-server-2016-device-guard-faq.aspx#:~:text=Credential%20Guard%20%3D%20a%20credential%20protection%20feature%20that,user%20credentials%20from%20being%20accessible%20from%20the%20OS
  1. You CAN enable configurable code integrity without either HVCI or Cred Guard.
  2. You CAN enable HVCI without either configurable code integrity or Cred Guard.
  3. You CAN enable Cred Guard without either configurable code integrity or HVCI.
  4. You CANNOT enable either Cred Guard or HVCI without Virtualization Based Security.
#>


<#
  Check if the script is being run remotely. This check is necessary because
  when UEFI lock is set, it should not be possible to toggle secured core features
 #>
function checkRemote {
  if ($PSSenderInfo -or (Get-Host).Name -eq 'ServerRemoteHost') {
    return $true
  }
  return $false
}
$Script:isRemoteConnection = checkRemote

function getFeatureStatuses($featureName) {
  foreach ($feature in $Script:featureDetails) {
    if ($feature.securityFeature -like $featureName) {
      return ($feature.configured -or $feature.status)
    }
  }
  return
}

function ExecuteCommandAndLog($_cmd) {
  try {
    Invoke-Expression $_cmd | Out-String
  }
  catch {
    Write-Host "Exception while exectuing $_cmd"
    Throw $_.Exception.Message
  }
}

# Check if is Virtual Machine
function checkIsVM {
  $model = (Get-WmiObject win32_computersystem).model
  $model = $model.ToString().ToUpperInvariant()
  return $model.Contains("VM") -or $model.Contains("VIRTUAL")
}
$Script:isVirtualMachine = checkIsVM

# TODO: Do we need to check for prerequisites

<# NOTE: System Guard is NOT supported for virtual machines #>
function toggleSystemGuard() {
  if ($Script:isVirtualMachine) {
    Write-Error "System Guard is not supported for Virtual Machines"
    return;
  }

  $systemGuardStatus = getFeatureStatuses('systemGuard')
  if ($systemGuardStatus -ne $null -and
    (
      ($Script:action -eq 0 -and !$systemGuardStatus) -or
      ($Script:action -eq 1 -and $systemGuardStatus)
    )) {
    return
  }

  $path = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\SystemGuard"
  if (-Not(Test-Path $path)) {
    New-Item $path -Force
  }
  Set-ItemProperty -path $path -name "Enabled" -value $Script:action -Type "DWORD" -Force
  return $path
}

function toggleVBS() {
  $vbsStatus = getFeatureStatuses('vbs')
  if (
    $vbsStatus -ne $null -and
    (
      ($Script:action -eq 0 -and !$vbsStatus) -or
      ($Script:action -eq 1 -and $vbsStatus)
    )) {
    return
  }

  $path = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\DeviceGuard"
  if (-Not(Test-Path $path)) {
    New-Item $path -Force
  }

  # For Windows 10 version 1607 and later
  $uefiLock = Get-ItemProperty -Path $path | `
    Microsoft.PowerShell.Utility\Select-Object "Locked" -ExpandProperty "Locked" -ErrorAction SilentlyContinue

  # Check UEFI lock for Windows 10 version 1511 and earlier
  $uefiUnlock = Get-ItemProperty -Path $path | `
    Microsoft.PowerShell.Utility\Select-Object "Unocked" -ExpandProperty "Unocked" -ErrorAction SilentlyContinue

  if (($uefiLock -eq 1 -or $uefiUnlock -eq 0) -and $Script:isRemoteConnection) {
    Throw "UEFI lock enabled. Cannot $Script:actionWord VBS remotely."
  }

  $currentSecureBootValue = Get-ItemProperty -Path $path | `
    Microsoft.PowerShell.Utility\Select-Object "RequirePlatformSecurityFeatures" -ExpandProperty "RequirePlatformSecurityFeatures" -ErrorAction SilentlyContinue

  if ($Script:action -eq 0) {
    # Note: all other VBS features (HVCI, Cred Guard) need to be disabled as well, or VBS will automatically turn on
    toggleHVCI($Script:action)
    toggleCredentialGuard($Script:action)

    Set-ItemProperty -path $path -Name "EnableVirtualizationBasedSecurity" -Value 0 -Type "DWORD" -Force

    # Disable secure boot
    if ($currentSecureBootValue) {
      Remove-ItemProperty -Path $path -Name "RequirePlatformSecurityFeatures" -ErrorAction SilentlyContinue
    }
  }

  if ($Script:action -eq 1) {
    Set-ItemProperty -path $path -Name "EnableVirtualizationBasedSecurity" -Value 1 -Type "DWORD" -Force

    # Enable Secure Boot
    if ($Script:secureBoot -in @(1, 3)) {
      <#
        - 1: Secure Boot only
        - 3: Secure Boot and DMA protection.
      #>
      Set-ItemProperty -path $path -name "RequirePlatformSecurityFeatures" -value $Script:secureBoot -Type "DWORD" -Force
    }
  }

  return $path
}


function toggleHVCI () {
  # https://docs.microsoft.com/en-us/windows/security/threat-protection/device-guard/enable-virtualization-based-protection-of-code-integrity#how-to-turn-off-hvci

  $hvciStatus = getFeatureStatuses('hvci')
  if (
    $hvciStatus -ne $null -and
    (
      ($Script:action -eq 0 -and !$hvciStatus) -or
      ($Script:action -eq 1 -and $hvciStatus)
    )) {
    return
  }

  $path = "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity"
  if (-Not(Test-Path $path)) {
    New-Item $path -Force
  }

  $uefiLock = Get-ItemProperty -Path $path | `
    Microsoft.PowerShell.Utility\Select-Object "Locked" -ExpandProperty "Locked" -ErrorAction SilentlyContinue

  if ($uefiLock -eq 1 -and $Script:isRemoteConnection) {
    Throw "UEFI lock enabled. Cannot $Script:actionWord HVCI remotely."
  }

  if ($Script:action -eq 0) {
    Remove-ItemProperty -Path $path -Name "WasEnabledBy" -ErrorAction SilentlyContinue
  }

  if ($Script:action -eq 1) {
    # Note: VBS will automatically turn on if you enable a VBS feature (HVCI, Cred Guard)
    toggleVBS(1)
    Set-ItemProperty -path $path -name "WasEnabledBy" -value 0 -Type "DWORD" -Force
  }

  # Toggle HVCI
  Set-ItemProperty -path $path -name "Enabled" -value $Script:action -Type "DWORD" -Force

  return $path
}

function toggleCredentialGuard() {
  <## TODO: To be used for credential guard Phase 3
  $credentialGuardStatus = getFeatureStatuses('credentialGuard')
  if (
    $credentialGuardStatus -ne $null -and
    (
      ($Script:action -eq 0 -and !$credentialGuardStatus) -or ($Script:action -eq 1 -and $credentialGuardStatus)
    )) {
    return
  } #>

  # HACK: Phase 2. Credential guard is not returned in the features object
  $securityServicesConfigured = (Get-CimInstance -classname Win32_DeviceGuard -namespace root\Microsoft\Windows\DeviceGuard).SecurityServicesConfigured
  if ( $securityServicesConfigured.Length -gt 0 -and
    ( $Script:action -eq 0 -and (-not(1 -in $securityServicesConfigured))) -or
    ($Script:action -eq 1 -and (1 -in $securityServicesConfigured))) {
    return;
  }

  $path = "Registry::HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\LSA"
  if (-Not(Test-Path $path)) {
    New-Item $path -Force
  }

  <#
    0: (Disabled) Turns off Windows Defender Credential Guard remotely if configured previously without UEFI Lock
    1: (Enabled with UEFI lock) Turns on Windows Defender Credential Guard with UEFI lock
    2: (Enabled without lock) Turns on Windows Defender Credential Guard without UEFI lock
  #>
  $credGuardValue = Get-ItemProperty -Path $path | `
    Microsoft.PowerShell.Utility\Select-Object "LsaCfgFlags" -ExpandProperty "LsaCfgFlags" -ErrorAction SilentlyContinue

  if ($credGuardValue -eq 1 -and $Script:isRemoteConnection) {
    Throw "UEFI lock enabled. Cannot $Script:actionWord Credential Guard remotely."
  }

  if ($Script:action -eq 0) {
    Remove-ItemProperty -path $path -name "LsaCfgFlags" -ErrorAction SilentlyContinue -Force

    # This setting is persisted in EFI (firmware) variables so we need to delete it
    $settingPath = "Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\DeviceGuard
    "
    Remove-ItemProperty -Path $settingPath -Name "LsaCfgFlags" -ErrorAction SilentlyContinue

    # Set of commands to run SecConfig.efi to delete UEFI variables if were set in pre OS
    $FreeDrive = Get-ChildItem function:[s-z]: -Name | Where-Object { !(Test-Path $_) } | Get-random
    ExecuteCommandAndLog 'mountvol $FreeDrive /s'
    Copy-Item "$env:windir\System32\SecConfig.efi" $FreeDrive\EFI\Microsoft\Boot\SecConfig.efi -Force | Out-String
    ExecuteCommandAndLog 'bcdedit /create "{0cb3b571-2f2e-4343-a879-d86a476d7215}" /d DGOptOut /application osloader'
    ExecuteCommandAndLog 'bcdedit /set "{0cb3b571-2f2e-4343-a879-d86a476d7215}" path \EFI\Microsoft\Boot\SecConfig.efi'
    ExecuteCommandAndLog 'bcdedit /set "{bootmgr}" bootsequence "{0cb3b571-2f2e-4343-a879-d86a476d7215}"'
    ExecuteCommandAndLog 'bcdedit /set "{0cb3b571-2f2e-4343-a879-d86a476d7215}" loadoptions DISABLE-LSA-ISO,DISABLE-VBS'
    ExecuteCommandAndLog 'bcdedit /set "{0cb3b571-2f2e-4343-a879-d86a476d7215}" device partition=$FreeDrive'
    ExecuteCommandAndLog 'mountvol $FreeDrive /d'
  }
  else {
    toggleVBS(1)
    Set-ItemProperty -path $path -name "LsaCfgFlags" -value 1 -Type "DWORD" -Force
  }

  return $path
}


###############################################################################
# Script execution starts here
###############################################################################

foreach ($selectedFeat in $Script:selectedFeatures) {
  if ($selectedFeat.securityFeature -eq 'vbs') {
    toggleVBS
  }
  elseif ($selectedFeat.securityFeature -eq 'hvci') {
    toggleHVCI
  }
  elseif ($selectedFeat.securityFeature -eq 'credentialGuard') {
    toggleCredentialGuard
  }
  elseif ($selectedFeat.securityFeature -eq 'systemGuard') {
    toggleSystemGuard
  }
  else {
    $errorMessage = 'Allowed feature values (case-sensitive): vbs, hvci, systemGuard, and credentialGuard'
    Throw $errorMessage
  }
}

}
## [END] Set-WACSESecuredCoreFeatures ##
function Set-WACSESecuredCoreOsConfigFeatures {
<#
.SYNOPSIS
Script that enables and disables Secured Core Features

.DESCRIPTION
Script that enables and disables Secured Core Features
  1. You CAN enable configurable code integrity without either HVCI or Cred Guard.
  2. You CAN enable HVCI without either configurable code integrity or Cred Guard.
  3. You CAN enable Cred Guard without either configurable code integrity or HVCI.
  4. You CANNOT enable either Cred Guard or HVCI without Virtualization Based Security.

.Parameter selectedFeatures
All selected features to toggle on/off.

.Parameter action
Value to set to either enable (1) or disable (0) feature

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [PSCustomObject[]]$selectedFeatures,
  [Parameter(Mandatory = $true)]
  [String]$action
)

$Script:selectedFeatures = $selectedFeatures

$Script:failedState = @{failureState = @{}}
# Check status of selected features after set calls and capture error on failure
function CheckStatus {
  Param([Parameter(Mandatory = $true)] $Status)

  foreach ($featStatus in $Status) {
    if ($Script:SecuredCoreConfigurations.ContainsKey($featStatus.name) -and $featStatus.state -eq "failed") {
      $errState = $featStatus.state
      $errCode = $featStatus.ErrorCode
      $errorMsg = "State: $errState, Error: $errCode"
      $Script:failedState.failureState.Add($featStatus.name, $errorMsg)
    }
  }

  return $Script:failedState
}


# Set OsConfiguration document to get current configuration values
function OsConfigurationSetDocumentGetResult {

  [CmdletBinding()]
  Param (
    [Parameter(Mandatory)]
    [String] $Id,

    [Parameter(Mandatory)]
    [String] $Content
  )

  # Set the document to get securedcore settings
  #Set-OsConfigurationDocument -Content $Content -Wait -TimeoutInSeconds 300
  Set-OsConfigurationDocument -Content $Content -Wait

  $result = Get-OsConfigurationDocumentResult -Id $Id | ConvertFrom-Json

  return $result.OsConfiguration.Scenario[0]
}

$jsonDocumentToGetSecuredCoreSettingConfigurations =
@"
{
  "OsConfiguration":{
      "Document":{
        "schemaversion":"1.0",
        "id":"47e88660-1861-4131-96e8-f32e85011e55",
        "version":"3C356C2C71F6A41F9AB4A601AD00C8B5BC7531576233010B13A221A9FE1BE7A0",
        "context":"device",
        "scenario":"SecuredCore"
      },
      "Scenario":[
        {
            "name":"SecuredCore",
            "schemaversion":"1.0",
            "action":"get",
            "SecuredCore":{
              "EnableVirtualizationBasedSecurity": "0",
              "HypervisorEnforcedCodeIntegrity": "0",
              "ConfigureSystemGuardLaunch": "0"
            }
        }
      ]
  }
}
"@

function GetSecuredCoreSettingConfigurations {

  # Set the document to get securedcore settings
  $result = OsConfigurationSetDocumentGetResult -Id "47e88660-1861-4131-96e8-f32e85011e55" -Content $jsonDocumentToGetSecuredCoreSettingConfigurations

  return $result.SecuredCore
}

$jsonDocumentToSetSecuredCoreSettingConfigurationsTemplate =
@"
{
  "OsConfiguration":{
      "Document":{
        "schemaversion":"1.0",
        "id":"74e88660-1861-4131-96e8-f32e85011e55",
        "version":"C8B5BC7531576233010B13A221A9FE1BE7A03C356C2C71F6A41F9AB4A601AD00",
        "context":"device",
        "scenario":"SecuredCore"
      },
      "Scenario":[
        {
            "name":"SecuredCore",
            "schemaversion":"1.0",
            "action":"set",
            "SecuredCore":{
              "EnableVirtualizationBasedSecurity": "1",
              "HypervisorEnforcedCodeIntegrity": "2",
              "ConfigureSystemGuardLaunch": "1"
            }
        }
      ]
  }
}
"@

# Set the configurations based on $Script:SecuredCoreConfigurations
function SetSecuredCoreSettingsUsingOsConfiguration() {

  # Get current configuration values
  $SecuredCoreConfigurations = GetSecuredCoreSettingConfigurations

  # Toggle the settings
  $jsonDocumentObject = $jsonDocumentToSetSecuredCoreSettingConfigurationsTemplate | ConvertFrom-Json

  if ($Script:SecuredCoreConfigurations.ContainsKey("EnableVirtualizationBasedSecurity"))
  {
    $jsonDocumentObject.OsConfiguration.Scenario[0].SecuredCore.EnableVirtualizationBasedSecurity = $Script:SecuredCoreConfigurations.EnableVirtualizationBasedSecurity
  }
  else
  {
    $jsonDocumentObject.OsConfiguration.Scenario[0].SecuredCore.EnableVirtualizationBasedSecurity = $SecuredCoreConfigurations.EnableVirtualizationBasedSecurity
  }

  if ($Script:SecuredCoreConfigurations.ContainsKey("HypervisorEnforcedCodeIntegrity"))
  {
    $jsonDocumentObject.OsConfiguration.Scenario[0].SecuredCore.HypervisorEnforcedCodeIntegrity = $Script:SecuredCoreConfigurations.HypervisorEnforcedCodeIntegrity
  }
  else
  {
    $jsonDocumentObject.OsConfiguration.Scenario[0].SecuredCore.HypervisorEnforcedCodeIntegrity = $SecuredCoreConfigurations.HypervisorEnforcedCodeIntegrity
  }

  if ($Script:SecuredCoreConfigurations.ContainsKey("ConfigureSystemGuardLaunch"))
  {
    $jsonDocumentObject.OsConfiguration.Scenario[0].SecuredCore.ConfigureSystemGuardLaunch = $Script:SecuredCoreConfigurations.ConfigureSystemGuardLaunch
  }
  else
  {
    $jsonDocumentObject.OsConfiguration.Scenario[0].SecuredCore.ConfigureSystemGuardLaunch = $SecuredCoreConfigurations.ConfigureSystemGuardLaunch
  }

  $jsonDocumentToSetSecuredCoreSettings = ConvertTo-Json -InputObject $jsonDocumentObject -Depth 5

  # Set the document to get securedcore settings
  #Set-OsConfigurationDocument -Content $jsonDocumentToSetSecuredCoreSettings -Wait -TimeoutInSeconds 300
  Set-OsConfigurationDocument -Content $jsonDocumentToSetSecuredCoreSettings -Wait

  # Return false on timeout.
  $documentState = Get-OsConfigurationDocument -Id "74e88660-1861-4131-96e8-f32e85011e55" | Microsoft.PowerShell.Utility\Select-Object "State"
  if ("DocumentStateCompleted" -ne $documentState.state) {
    return $null
  }

  $result = Get-OsConfigurationDocumentResult -Id "74e88660-1861-4131-96e8-f32e85011e55" | ConvertFrom-Json

  return CheckStatus $result.OsConfiguration.Scenario[0].Status
}

$Script:SecuredCoreConfigurations = @{}

function ToggleSecuredCoreSettingConfiguration() {
  foreach ($selectedFeat in $Script:selectedFeatures) {
    if ($selectedFeat.securityFeature -eq 'vbs') {
      $Script:SecuredCoreConfigurations.Add("EnableVirtualizationBasedSecurity", $action)
    }
    elseif ($selectedFeat.securityFeature -eq 'hvci') {
      if ($action -eq "1") {
        $Script:SecuredCoreConfigurations.Add("HypervisorEnforcedCodeIntegrity", "2")
      }
      else {
        $Script:SecuredCoreConfigurations.Add("HypervisorEnforcedCodeIntegrity", "0")
      }
    }
    elseif ($selectedFeat.securityFeature -eq 'credentialGuard') {
      $Script:SecuredCoreConfigurations.Add("ConfigureCredentialGuard", $action)
    }
    elseif ($selectedFeat.securityFeature -eq 'systemGuard') {
      if ($action -eq "1") {
        $Script:SecuredCoreConfigurations.Add("ConfigureSystemGuardLaunch", "1")
      }
      else {
        $Script:SecuredCoreConfigurations.Add("ConfigureSystemGuardLaunch", "2")
      }
    }
    else {
      $errorMessage = 'Allowed feature values (case-sensitive): vbs, hvci, systemGuard, and credentialGuard'
      Throw $errorMessage
    }
  }
  $Script:SecuredCoreConfigurations
}

$Script:SecuredCoreConfigurations
ToggleSecuredCoreSettingConfiguration
$Script:SecuredCoreConfigurations
SetSecuredCoreSettingsUsingOsConfiguration

}
## [END] Set-WACSESecuredCoreOsConfigFeatures ##
function Set-WACSEThreatAction {
<#

.SYNOPSIS
Set Given Threat Default Action to Given Threat.

.DESCRIPTION
Set Given Threat Default Action to Given Threat.

.ROLE
Administrators

#>

Param(
    [string]$chosenAction,
    [string]$threatID
)

Set-StrictMode -Version 5.0;

$threatID = [int64]$threatID

Set-MpPreference -ThreatIDDefaultAction_Ids $threatID -ThreatIDDefaultAction_Actions $chosenAction
}
## [END] Set-WACSEThreatAction ##
function Set-WACSEWdacPolicyMode {
<#
.SYNOPSIS
Set Windows Defender Application Control (WDAC) Policy setting

.DESCRIPTION
Set Windows Defender Application Control (WDAC) Policy mode to Audit(1) or Enforced(2)

.Parameter mode
Policy mode to set to either: Audit (1) or Enforcement (2)

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [String]$mode
)

Add-Type -TypeDefinition @"
   public enum PolicyMode {
        Audit = 1,
        Enforced = 2
    }
"@

function ToggleWdacPolicyMode {
  [CmdletBinding()]
  param (
    [Parameter()]
    [string]$mode
  )

  if ([PolicyMode]$mode -eq [PolicyMode]::Audit) {
    Enable-ASLocalWDACPolicy -Mode Audit
  }
  elseif ([PolicyMode]$mode -eq [PolicyMode]::Enforced) {
    Enable-ASLocalWDACPolicy -Mode Enforced
  }
  else {
    $LogName = "Microsoft-ServerManagementExperience"
    $LogSource = "msft.sme.security"
    $ScriptName = "Set-WDACPolicyMode.ps1"
    $Message = "Invalid WDAC policy mode passed: $mode"
    $EntryType = 'Error'

    # EntryType: Error, Information, FailureAudit, SuccessAudit, Warning
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource `
      -EventId 0 -Category 0 -EntryType $EntryType `
      -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue
  }
}

$Script:eventId = 0
function Write-SetWdacEventLog {
  param (
    [Parameter(Mandatory = $false)]
    [String]$entryType,
    [Parameter(Mandatory = $true)]
    [String]$message
  )

  if (!$entryType) {
    $entryType = 'Error'
  }

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "msft.sme.security"
  $ScriptName = "Set-WdacPolicyMode.ps1"

  # Create the event log if it does not exists
  New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  # EntryType: Error, Information, FailureAudit, SuccessAudit, Warning
  Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId $Script:eventId -Category 0 -EntryType $EntryType `
    -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue

  $Script:eventId += 1
}


###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
  $wdacModuleExists = $null -ne (Get-Command -Module Microsoft.AS.Infra.Security.WDAC)
  if ($wdacModuleExists) {
    ToggleWdacPolicyMode -Mode $mode
  } else {
    Write-SetWdacEventLog -Message "Couldn't toggle WDAC policy mode. Module 'Microsoft.AS.Infra.Security.WDAC' does not exist on this server. Ensure that you are running the latest version of Azure Stack HCI."
  }
}

}
## [END] Set-WACSEWdacPolicyMode ##
function Start-WACSEMpScan {
<#

.SYNOPSIS
Start Scan.

.DESCRIPTION
Start Scan.

.Parameter ScanType
Specifies the scan type to use during a scheduled scan. The acceptable values for this parameter are:
  FullScan
  QuickScan

.ROLE
Readers

#>

param (
  [Parameter(Mandatory = $true)]
  [string]$ScanType
)

Set-StrictMode -Version 5.0;

switch ($ScanType) {
  1 { $ScanTypeValue = 'QuickScan' }
  2 { $ScanTypeValue = 'FullScan' }
}

Start-MpScan -ScanType $ScanTypeValue

}
## [END] Start-WACSEMpScan ##
function Test-WACSEOsConfigModule {
<#

.SYNOPSIS
Test-OSConfigModule

.DESCRIPTION
Checks if OSConfiguration Module is present

.ROLE
Readers

#>

$Script:eventId = 0
function Write-OsConfigEventToEventLog {
  param (
    [Parameter(Mandatory = $false)]
    [String]$entryType,
    [Parameter(Mandatory = $true)]
    [String]$message
  )

  if (!$entryType) {
    $entryType = 'Warning'
  }

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "msft.sme.security"
  $ScriptName = "Test-OsConfigModule.ps1"

  # Create the event log if it does not exists
  New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  # EntryType: Error, Information, FailureAudit, SuccessAudit, Warning
  Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId $Script:eventId -Category 0 -EntryType $EntryType `
    -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue

  $Script:eventId += 1
}

$jsonDocumentToGetSecuredCoreSettingStates =
@"
{
  "OsConfiguration":{
      "Document":{
        "schemaversion":"1.0",
        "id":"10088660-1861-4131-96e8-f32e85011100",
        "version":"10056C2C71F6A41F9AB4A601AD00C8B5BC7531576233010B13A221A9FE1BE100",
        "context":"device",
        "scenario":"SecuredCoreState"
      },
      "Scenario":[
        {
            "name":"SecuredCoreState",
            "schemaversion":"1.0",
            "action":"get",
            "SecuredCoreState":{
              "VirtualizationBasedSecurityStatus": "0",
              "HypervisorEnforcedCodeIntegrityStatus": "0",
              "SystemGuardStatus": "0",
              "SecureBootState": "0",
              "TPMVersion": "",
              "BootDMAProtection": "0"
            }
        }
      ]
  }
}
"@

function Test-OsConfigFeatureEnabled {
  ########## 13-Spetember-2022 ##########
  ### 1. Should WAC be enabling OSConfig if it's not enabled so that these security
  # settings can be read through WAC? - OsConfig is designed to be released only to
  # ASZ HCI (for now), but because WSD does not support velocity based composition,
  # its binaries are released to every server edition as part of 7C/8C update/KB. The
  # feature itself is guarded/blocked by WSD EKB. EKB is designed and can only be
  # enabled on ASZ HCI. It leaves the "weird" situation, OsConfig is released/present
  # in server edition (FE server 2022) but not enabled (cannot be enabled for server
  # edition except ASZ HCI.
  ### 2. Then how can we improve validation such that an OS update doesn't break
  # OSConfig using WAC? This last update broke every single WAC customer out there
  # using WAC to manage their HCI cluster - We should validate that these OS updates
  # don't break existing WAC functionality. However, it is one exception (which is also surprised to us)
  # we have to handle, we do not expect to have another one in future.
  Import-Module -Name OsConfiguration -ErrorAction SilentlyContinue -ErrorVariable err

  if (!!$err) {
    Write-OsConfigEventToEventLog -Message "There was an error importing the OsConfiguration module. Error: $err"
    return $false
  }

  try {
    Set-OsConfigurationDocument -Content $jsonDocumentToGetSecuredCoreSettingStates -Wait
  }
  catch {
    Write-OsConfigEventToEventLog -Message "There was an error setting the OS configuration document. Error: $err"
    return $false
  }

  return $true
}

###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
  Test-OsConfigFeatureEnabled
}

}
## [END] Test-WACSEOsConfigModule ##
function Test-WACSEWdacPolicyFilePath {
<#
.SYNOPSIS
    Test if a filepath belongs to a cluster shared volume
.DESCRIPTION
    Test if a filepath belongs to a cluster shared volume
.ROLE
    Administrators
#>

param (
    [Parameter(Mandatory = $true)]
    [String]$filePath
)

enum ValidationErrorType {
    FileDoesNotExist
    NoClusterVolume
    FileNotOnClusterVolume
}

$ErrorActionPreference = "Stop"

$Script:eventId = 0
function Write-WdacEventLog {
    param (
        [Parameter(Mandatory = $false)]
        [String]$EntryType,
        [Parameter(Mandatory = $true)]
        [String]$Message
    )

    if (!$entryType) {
        $entryType = 'Error'
    }

    $LogName = "Microsoft-ServerManagementExperience"
    $LogSource = "msft.sme.security"
    $ScriptName = "Test-WdacPolicyFilePath.ps1"

    # Create the event log if it does not exists
    New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

    # EntryType: Error, Information, FailureAudit, SuccessAudit, Warning
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId $Script:eventId -Category 0 -EntryType $EntryType `
        -Message "[$ScriptName]: $Message" -ErrorAction SilentlyContinue

    $Script:eventId += 1
}

function Test-WdacPolicyFilePath {
    param (
        [Parameter(Mandatory = $true)]
        [String]$filePath
    )

    $policyPath = [IO.Path]::GetFullPath($filePath)
    if (-not (Test-Path -Path $policyPath -PathType Leaf)) {
        $errorMessage = "The policy file path does not exist."
        Write-WdacEventLog -message $errorMessage
        return @{ result = $false; error = [ValidationErrorType]::FileDoesNotExist }
    }

    $clusterName = (Get-Cluster).name
    $clusterSharedVolumes = Get-ClusterSharedVolume -Cluster $clusterName

    if ($clusterSharedVolumes.Count -eq 0) {
        $errorMessage = "No cluster shared volumes were found."
        Write-WdacEventLog -message $errorMessage
        return @{ result = $false; error = [ValidationErrorType]::NoClusterVolume }
    }

    $matchFound = $false
    foreach ($volume in $clusterSharedVolumes) {
        $volumePath = [IO.Path]::GetFullPath($volume.SharedVolumeInfo.FriendlyVolumeName)
        if ($policyPath.StartsWith( $volumePath, [StringComparison]::OrdinalIgnoreCase )) {
            $matchFound = $true
            break
        }
    }

    if (-not $matchFound) {
        $errorMessage = "The policy file must be on a cluster shared volume."
        Write-WdacEventLog -message $errorMessage
        return @{ result = $false; error = [ValidationErrorType]::FileNotOnClusterVolume }
    }

    return @{ result = $true; error = $null }
}

###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
    Test-WdacPolicyFilePath -filePath $filePath
}

}
## [END] Test-WACSEWdacPolicyFilePath ##
function Get-WACSECimWin32LogicalDisk {
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
## [END] Get-WACSECimWin32LogicalDisk ##
function Get-WACSECimWin32NetworkAdapter {
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
## [END] Get-WACSECimWin32NetworkAdapter ##
function Get-WACSECimWin32PhysicalMemory {
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
## [END] Get-WACSECimWin32PhysicalMemory ##
function Get-WACSECimWin32Processor {
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
## [END] Get-WACSECimWin32Processor ##
function Get-WACSEClusterInventory {
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
## [END] Get-WACSEClusterInventory ##
function Get-WACSEClusterNodes {
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
## [END] Get-WACSEClusterNodes ##
function Get-WACSEDecryptedDataFromNode {
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
## [END] Get-WACSEDecryptedDataFromNode ##
function Get-WACSEEncryptionJWKOnNode {
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
## [END] Get-WACSEEncryptionJWKOnNode ##
function Get-WACSEServerInventory {
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
## [END] Get-WACSEServerInventory ##

# SIG # Begin signature block
# MIIoLQYJKoZIhvcNAQcCoIIoHjCCKBoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBAE5mSwzStrVbJ
# 8UGZssk/xkoiijj7rPhJFVRiQXcfUaCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGg0wghoJAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIHULcIvBnZWCpCJwi3o1OevA
# z421uzPu6qJ7a7V46djtMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAM4mrBsvy8zE+/ErSqSr7XMj6AL3kLL2a/vhU30uo3If+9i88GYysgfti
# G5WjoLdZshBl/aKwRul++Xtdz/duafDCRwn6CnSdtMRrGu/awsUVktO5qKU0SOEB
# pmtsN0LO7voJ4CJHJJD3pOI678z9K82XepFhw7Nt1cWbih8e9Q4/4jftI5dEYMxI
# +LtQI7c9rL2inXyDXal05C1PKJDkTwdTIz6dimTmFeVmBNbOFe3cmuxGp5cznyw5
# 1x/gP622pveuDZPWAoFyE5AHVCjzYs30efnVrfZs4zREuCvxWVcm5qX6ds6zJUCW
# 8coc1Qp/kHJ6sUt0nWI6wi9FnZFcKaGCF5cwgheTBgorBgEEAYI3AwMBMYIXgzCC
# F38GCSqGSIb3DQEHAqCCF3AwghdsAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCAevjBLLUVZNI4c2GO7+ksR22oPDbNnXROvjvguLGb9qQIGZWitmjYc
# GBMyMDIzMTIwNzAzNDQ0Ny42MjNaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046N0YwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wg
# ghHtMIIHIDCCBQigAwIBAgITMwAAAdWpAs/Fp8npWgABAAAB1TANBgkqhkiG9w0B
# AQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzA1MjUxOTEy
# MzBaFw0yNDAyMDExOTEyMzBaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046N0YwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDFfak57Oph9vuxtloABiLc6enT+yKH619b+OhGdkyh
# gNzkX80KUGI/jEqOVMV4Sqt/UPFFidx2t7v2SETj2tAzuVKtDfq2HBpu80vZ0vyQ
# DydVt4MDL4tJSKqgYofCxDIBrWzJJjgBolKdOJx1ut2TyOc+UOm7e92tVPHpjdg+
# Omf31TLUf/oouyAOJ/Inn2ih3ASP0QYm+AFQjhYDNDu8uzMdwHF5QdwsscNa9PVS
# GedLdDLo9jL6DoPF4NYo06lvvEQuSJ9ImwZfBGLy/8hpE7RD4ewvJKmM1+t6eQuE
# sTXjrGM2WjkW18SgUZ8n+VpL2uk6AhDkCa355I531p0Jkqpoon7dHuLUdZSQO40q
# mVIQ6qQCanvImTqmNgE/rPJ0rgr0hMPI/uR1T/iaL0mEq4bqak+3sa8I+FAYOI/P
# C7V+zEek+sdyWtaX+ndbGlv/RJb5mQaGn8NunbkfvHD1Qt5D0rmtMOekYMq7QjYq
# E3FEP/wAY4TDuJxstjsa2HXi2yUDEg4MJL6/JvsQXToOZ+IxR6KT5t5fB5FpZYBp
# VLMma3pm5z6VXvkXrYs33NXJqVWLwiswa7NUFV87Es2sou9Idw3yAZmHIYWgOQ+D
# IY1nY3aG5DODiwN1rJyEb+mbWDagrdVxcncr6UKKO49eoNTXEW+scUf6GwXG0KEy
# mQIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFK/QXKNO35bBMOz3R5giX7Ala2OaMB8G
# A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4w
# XAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2Vy
# dHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBmRddqvQuyjRpx0HGxvOqffFrbgFAg0j82
# v0v7R+/8a70S2V4t7yKYKSsQGI6pvt1A8JGmuZyjmIXmw23AkI5bZkxvSgws8rrB
# tJw9vakEckcWFQb7JG6b618x0s9Q3DL0dRq46QZRnm7U6234lecvjstAow30dP0T
# nIacPWKpPc3QgB+WDnglN2fdT1ruQ6WIVBenmpjpG9ypRANKUx5NRcpdJAQW2FqE
# HTS3Ntb+0tCqIkNHJ5aFsF6ehRovWZp0MYIz9bpJHix0VrjdLVMOpe7wv62t90E3
# UrE2KmVwpQ5wsMD6YUscoCsSRQZrA5AbwTOCZJpeG2z3vDo/huvPK8TeTJ2Ltu/I
# tXgxIlIOQp/tbHAiN8Xptw/JmIZg9edQ/FiDaIIwG5YHsfm2u7TwOFyd6OqLw18Z
# 5j/IvDPzlkwWJxk6RHJF5dS4s3fnyLw3DHBe5Dav6KYB4n8x/cEmD/R44/8gS5Pf
# uG1srjLdyyGtyh0KiRDSmjw+fa7i1VPoemidDWNZ7ksNadMad4ZoDvgkqOV4A6a+
# N8HIc/P6g0irrezLWUgbKXSN8iH9RP+WJFx5fBHE4AFxrbAUQ2Zn5jDmHAI3wYcQ
# DnnEYP51A75WFwPsvBrfrb1+6a1fuTEH1AYdOOMy8fX8xKo0E0Ys+7bxIvFPsUpS
# zfFjBolmhzCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZI
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
# 6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKhggNQ
# MIICOAIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEn
# MCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjdGMDAtMDVFMC1EOTQ3MSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQBO
# Ei+S/ZVFe6w1Id31m6Kge26lNKCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6Ru9bDAiGA8yMDIzMTIwNzAzNDEz
# MloYDzIwMjMxMjA4MDM0MTMyWjB3MD0GCisGAQQBhFkKBAExLzAtMAoCBQDpG71s
# AgEAMAoCAQACAgxiAgH/MAcCAQACAhOsMAoCBQDpHQ7sAgEAMDYGCisGAQQBhFkK
# BAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJ
# KoZIhvcNAQELBQADggEBAGMr/V9zs81P6aKN9JUhAQNyu1qTl0W9rO2JUX5IqTzk
# S3nJhXsRqKoXPFV47x+C6wc4Q2cEzP+4S6jPX/60qMYuwm8+Gmk+VqVvmRIdeEbp
# hx6y+G110G5oZtRJz7ZslT9EjAzNTGjZ5UUb1sTHsQcTeebrNWI/kETAKYBE5Iny
# mxKxWxU0WEtBGhoDzMvEwriFadT4yABOWcr+1on+mnUEcMz0vKDfo3Wgbeqnw14R
# 1T33+75YwJoYrLAGlZkH2KbVPIVKAjuEic21aZBdSPGUG/0Jxxu2Y1bumtFaoffP
# X5ldCvJBehnKFHNOf+7jLLNo4TFQr3nS/EekNIjLnfcxggQNMIIECQIBATCBkzB8
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1N
# aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdWpAs/Fp8npWgABAAAB
# 1TANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEE
# MC8GCSqGSIb3DQEJBDEiBCAM1dsegLfXG0uNzD0RJwohe4dPccHYzOWq+vSVY3fr
# uTCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EINm/I4YM166JMM7EKIcYvlcb
# r2CHjKC0LUOmpZIbBsH/MIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
# IDIwMTACEzMAAAHVqQLPxafJ6VoAAQAAAdUwIgQgcTw1OKOE3d5h1x08YVTlvg63
# biSHFBb8VzsmVArfbuQwDQYJKoZIhvcNAQELBQAEggIAeR3aDzu013KjbOXq6UVe
# rRHQis9Sp1v06SjSxVCltpkn1vGdNknMTpWIphh+Rq9Bth4tghTgVfEMJIWBfEzr
# 5CWy3i7MN4nRSYu76JKM4ONeZHyChymhpMLDlzwLMlFuXahdsYcLPDC6dn/EQEBk
# YrWFp955beabSLPkcdcidrOB3BEH7QdxuTFW3qaWVRYQY9xUo03xZwjIOWqLGmax
# 3UZtEq/FVemKfqXD99HqRamZ+t0RwB7Xtl0/dN3+5qgo8Oi7U4jChT7qXqsrrG1e
# PWs/qaja5G4hOID9Bj2JPcjJkIazJczpwFBsqDboOty03G/N7lqTFYXcH4/gR6t7
# X7sFOMeFpbm4G1JfKplqC1XcMBnqCv/Idml7dTxPvYgVkex89DzzOhKodSrrG3Mr
# RG5sIg9+u+6u5dqZ6b7wGDpcA1epQEeSHuxs36RSCcXaTMDCJaAhp6ROVVWQEQVa
# NW28pTKssKhVuaXhDvkRn/lTjfeAFPtIKmKqoScUl8/GG8BMPnAGla0GgW3KVEWh
# XAWID3bdsp2dHNNzEZothnoAQmJ8c2kHSahYCuAnTLL2G7s+pYux6LnRynXChux2
# x/m3MRgSMNHI9TYVgEKIB0Z/1/uTFBTyRN0y8YUl7rCQSMcBxaMWLi2mc9iL9vtV
# iZsxQE7xsyivJNytxqx4eQM=
# SIG # End signature block
