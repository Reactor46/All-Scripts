function Disable-WACSIInsightsCapability {
<#

.SYNOPSIS
Deactivates a capability, which stops data collection for that capability and prevents the capability from being invoked.

.DESCRIPTION
Deactivates a capability, which stops data collection for that capability and prevents the capability from being invoked.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.PARAMETER name
Specifies a capability using a capability name

.ROLE
Administrators

#>
param (
  [Parameter(Mandatory = $true)]
  [string]
  $name
)
Import-Module SystemInsights

Disable-InsightsCapability -Name $name -confirm:$false

}
## [END] Disable-WACSIInsightsCapability ##
function Disable-WACSIInsightsCapabilitySchedule {
<#

.SYNOPSIS
Disables periodic predictions for the specified capabilities.

.DESCRIPTION
Disables periodic predictions for the specified capabilities.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.PARAMETER name
Specifies a capability using a capability name

.ROLE
Administrators

#>
param (
  [Parameter(Mandatory = $true)]
  [string]
  $name
 )
Import-Module SystemInsights

Disable-InsightsCapabilitySchedule -Name $name -Confirm:$false

}
## [END] Disable-WACSIInsightsCapabilitySchedule ##
function Enable-WACSIInsightsCapability {
<#

.SYNOPSIS
Activates a capability, which starts all data collection for that capability, allows the capability to be invoked, and enables users to set custom configuration information.

.DESCRIPTION
Activates a capability, which starts all data collection for that capability, allows the capability to be invoked, and enables users to set custom configuration information.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.PARAMETER name
Specifies a capability using a capability name

.ROLE
Administrators

#>
param (
  [Parameter(Mandatory = $true)]
  [string]
  $name
)
Import-Module SystemInsights

Enable-InsightsCapability -Name $name

}
## [END] Enable-WACSIInsightsCapability ##
function Enable-WACSIInsightsCapabilitySchedule {
<#

.SYNOPSIS
Enables periodic predictions for the specified capabilities.

.DESCRIPTION
Enables periodic predictions for the specified capabilities.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.PARAMETER name
Specifies a capability using a capability name

.ROLE
Administrators

#>
param (
  [Parameter(Mandatory = $true)]
  [string]
  $name
 )
Import-Module SystemInsights

Enable-InsightsCapabilitySchedule -Name $name

}
## [END] Enable-WACSIInsightsCapabilitySchedule ##
function Get-WACSIClusterSettings {
<#

.SYNOPSIS
Gets item property values for clustered storage settings.

.DESCRIPTION
Gets item property values for clustered storage settings.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.ROLE
Readers

#>
Import-Module Microsoft.PowerShell.Management, Microsoft.PowerShell.Utility

$ClusterNode       = $false
$CollectionEnabled = $null
$TotalStorageState = $null
$VolumeState       = $null

try
{
    Import-Module FailoverClusters -ErrorAction Stop
    # store result in a variable so the script will only output one item in the array
    $cluster = Get-Cluster -Name $env:COMPUTERNAME -WarningAction SilentlyContinue -ErrorAction Stop
    if ($cluster -ne $null) {
      $ClusterNode = $true;
    }
}
catch
{
    $ClusterNode = $false
}

if ($ClusterNode -eq $true)
{
    $systemDataArchiver = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\SystemDataArchiver -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if (($systemDataArchiver -ne $null) -and ($systemDataArchiver.ClusterVolumesAndDisks -ne $null))
    {
        $CollectionEnabled = $systemDataArchiver.ClusterVolumesAndDisks -ne 0
    }

    $totalStorage = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\SystemInsights\Capabilities\Total storage consumption forecasting' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if (($totalStorage -ne $null) -and ($totalStorage.ClusterVolumesAndDisks -ne $null))
    {
        $TotalStorageState = $totalStorage.ClusterVolumesAndDisks
    }

    $volume = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\SystemInsights\Capabilities\Volume consumption forecasting' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if (($volume -ne $null) -and ($volume.ClusterVolumes -ne $null))
    {
        $VolumeState = $volume.ClusterVolumes
    }
}

@{
    ClusterNode       = $ClusterNode;
    CollectionEnabled = $CollectionEnabled;
    TotalStorageState = $TotalStorageState;
    VolumeState       = $VolumeState
} | Write-Output

}
## [END] Get-WACSIClusterSettings ##
function Get-WACSIInsightsCapability {
<#

.SYNOPSIS
Invokes Get-InsightsCapability

.DESCRIPTION
Invokes Get-InsightsCapability
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.ROLE
Readers

#>
Import-Module SystemInsights
Get-InsightsCapability

}
## [END] Get-WACSIInsightsCapability ##
function Get-WACSIInsightsCapabilityAction {
<#

.SYNOPSIS
Invokes Get-InsightsCapabilityAction.

.DESCRIPTION
Invokes Get-InsightsCapabilityAction.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.PARAMETER name
Specifies a capability using a capability name

.ROLE
Readers

#>
param (
  [Parameter(Mandatory = $true)]
  [string]
  $name
)
Import-Module SystemInsights

Get-InsightsCapabilityAction -Name $name

}
## [END] Get-WACSIInsightsCapabilityAction ##
function Get-WACSIInsightsCapabilityResultCombo {
<#

.SYNOPSIS
Invokes Get-InsightsCapabilityResultCombo.

.DESCRIPTION
Invokes Get-InsightsCapabilityResultCombo.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.PARAMETER name

.PARAMETER usesAnomalyFilter

.PARAMETER minutesToSubtract

.PARAMETER hoursToSubtract

.PARAMETER daysToSubtract

.PARAMETER monthsToSubtract

.PARAMETER diskIdentifier

.ROLE
Readers

#>
param (
  [Parameter(Mandatory = $true)]
  [string]$name,

  [Parameter(Mandatory = $true)]
  [boolean]$usesAnomalyFilter,
  
  [Parameter(Mandatory = $true)]
  [int]$minutesToSubtract,
  
  [Parameter(Mandatory = $true)]
  [int]$hoursToSubtract,
  
  [Parameter(Mandatory = $true)]
  [int]$daysToSubtract,

  [Parameter(Mandatory = $true)]
  [int]$monthsToSubtract,

  [string]$diskIdentifier

)
Import-Module SystemInsights, Microsoft.PowerShell.Management, Microsoft.PowerShell.Utility
$resultWithoutHistory = Get-InsightsCapabilityResult -Name $name
$resultWithHistory = Get-InsightsCapabilityResult -Name $name -History

$outputData = $null
$diskList = @()

if ($resultWithHistory) {
  $path = $resultWithoutHistory.output
  $outputData = Get-Content -Path $path -Encoding UTF8 | Microsoft.PowerShell.Utility\ConvertFrom-Json

  if ($usesAnomalyFilter) {
    if ($outputData -and $outputData.AnomalyDetectionResults) {
      # Use += to avoid single disks not returning an array
      $diskList += $outputData.AnomalyDetectionResults | ForEach-Object { $_.Identifier } | Sort-Object | Get-Unique
      if ($diskList.length -gt 0) {
        if (!$diskIdentifier -or !$diskList.contains($diskIdentifier)) {
          $diskIdentifier = ''
          $diskIdentifier = $diskList | Microsoft.PowerShell.Utility\Select-Object -First 1
        }
        foreach ($set in $outputData.AnomalyDetectionResults) {
          if ($set.Identifier) {
            if ($set.Identifier -eq $diskIdentifier) {
              if ($set.Series -and ($set.Series.length -gt 0)) {
                $set.Series = $set.Series | Microsoft.PowerShell.Utility\Sort-Object -Property DateTime
                # We use the last value in the series as the day to filter back from
                $lastDate = $set.Series.DateTime | Microsoft.PowerShell.Utility\Select-Object -Last 1
                if ($lastDate) {
                  $filterDate = $lastDate.AddMinutes(-$minutesToSubtract).AddHours(-$hoursToSubtract).AddDays(-$daysToSubtract).AddMonths(-$monthsToSubtract)
                  $set.Series = $set.Series | Where-Object { $_.DateTime -gt $filterDate }
                }
              }
              else {
                $set.Series = $()
              }
            }
            else {
              $set.Series = $()
            }
          }
        }
      }
    }
  }
  $outputData | Microsoft.PowerShell.Utility\Add-Member -NotePropertyName 'DiskList' -NotePropertyValue $diskList
  $outputData = $outputData | Microsoft.PowerShell.Utility\ConvertTo-Json -Compress -Depth 5
}

$scheduleData = Get-InsightsCapabilitySchedule -name $name
$capabilityList = Get-InsightsCapability

$combinedResults = @{ }
$combinedResults += @{"historyResult" = $resultWithHistory }
$combinedResults += @{"capabilityResult" = $resultWithoutHistory }
$combinedResults += @{"outputResult" = $outputData }
$combinedResults += @{"scheduleResult" = $scheduleData }
$combinedResults += @{"capabilityListResult" = $capabilityList }
Write-Output $combinedResults

}
## [END] Get-WACSIInsightsCapabilityResultCombo ##
function Get-WACSIInsightsCapabilitySchedule {
<#

.SYNOPSIS
Invokes Get-InsightsCapabilitySchedule

.DESCRIPTION
Invokes Get-InsightsCapabilitySchedule
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.PARAMETER name
Specifies a capability using a capability name

.ROLE
Readers

#>
param (
  [Parameter(Mandatory = $true)]
  [string]
  $name
)
Import-Module SystemInsights

Get-InsightsCapabilitySchedule -Name $name

}
## [END] Get-WACSIInsightsCapabilitySchedule ##
function Get-WACSIInsightsFeature {
<#

.SYNOPSIS
Invokes Get-WindowsFeature

.DESCRIPTION
Invokes Get-WindowsFeature
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.ROLE
Readers

#>
Import-Module ServerManager

Get-WindowsFeature -Name 'System-Insights','RSAT-System-Insights', 'System-DataArchiver'

}
## [END] Get-WACSIInsightsFeature ##
function Get-WACSISystemDriveLetter {
<#

.SYNOPSIS
Get system drive

.DESCRIPTION
Get system drive

.ROLE
Readers

#>
Import-Module Microsoft.PowerShell.Management, Microsoft.PowerShell.Utility

$systemDriveLetter = (Get-WmiObject Win32_OperatingSystem).SystemDrive

Write-Output $systemDriveLetter

}
## [END] Get-WACSISystemDriveLetter ##
function Get-WACSITempPath {
<#

.SYNOPSIS
Returns the path of the current user's temporary folder.

.DESCRIPTION
Returns the path of the current user's temporary folder.

.ROLE
Readers

#>
Import-Module Microsoft.PowerShell.Utility

$newTempFile = [System.IO.Path]::GetTempPath() | Microsoft.PowerShell.Utility\ConvertTo-Json -Depth 5

Write-Output $newTempFile

}
## [END] Get-WACSITempPath ##
function Install-WACSIInsightsFeature {
<#

.SYNOPSIS
Invokes Install-WindowsFeature.

.DESCRIPTION
Invokes Install-WindowsFeature.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.ROLE
Administrators

#> 
Import-Module ServerManager

Install-WindowsFeature -Name 'System-Insights','RSAT-System-Insights', 'System-DataArchiver'

}
## [END] Install-WACSIInsightsFeature ##
function Install-WACSINugetFromTemp {
<#

.SYNOPSIS

.DESCRIPTION

.PARAMETER path

.PARAMETER title

.PARAMETER id

.PARAMETER version

.PARAMETER dllName

.ROLE
Administrators

#>
Param
(
    [Parameter(Mandatory = $true)]
    [string]$path,

    [Parameter(Mandatory = $true)]
    [string]$title,

    [Parameter(Mandatory = $true)]
    [string]$id,

    [Parameter(Mandatory = $true)]
    [string]$version,

    [Parameter(Mandatory = $true)]
    [string]$dllName
)
Import-Module SystemInsights, Microsoft.PowerShell.Management, Microsoft.PowerShell.Utility

$sourceDirectoryPath = $path + $id + "." + $version

$destinationDirectoryPath = $env:SystemDrive + "\ProgramData\Microsoft\Windows\SystemInsights\InstalledCapabilities\" + $id

try {
    if (Test-Path $destinationDirectoryPath) { Remove-Item $destinationDirectoryPath -Force -Recurse; }

    Copy-Item -Path $sourceDirectoryPath -Destination $destinationDirectoryPath -Recurse -Force

    $dllPath = $destinationDirectoryPath + "\" + $dllName

    Add-InsightsCapability -Name $title -Library $dllPath -Confirm:$false

    Restart-Service DPS

    $policiesPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Insights\Parameters"
    if (Test-Path $policiesPath) {
        $currentSetting = Get-ItemProperty -Path $policiesPath -Name MaxSerializedLengthInMB -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        if ($currentSetting -eq $null) {
            Set-ItemProperty -Path $policiesPath -Name MaxSerializedLengthInMB -Value 100 -Force
            Restart-Service Insights
        }
        else {
            if ($currentSetting.MaxSerializedLengthInMB -lt 100) {
                Set-ItemProperty -Path $policiesPath -Name MaxSerializedLengthInMB -Value 100 -Force
            }
            Restart-Service Insights
        }
    }
    else {
        New-Item -Path $policiesPath -Force
        New-ItemProperty -Path $policiesPath -Name MaxSerializedLengthInMB -Value 100 -PropertyType DWORD -Force
        Restart-Service Insights
    }

    if (Test-Path $path) { Remove-Item $path -Force -Recurse; }

}
catch {
    @{
        errorDetail = ("$($_.Exception.Message) $($_.CategoryInfo.GetMessage())");
        hResult     = $_.Exception.hResult;
    } | Microsoft.PowerShell.Utility\Write-Output
    throw $_
}
}
## [END] Install-WACSINugetFromTemp ##
function Invoke-WACSIInsightsCapability {
<#

.SYNOPSIS
Invokes Invoke-InsightsCapability.

.DESCRIPTION
Invokes Invoke-InsightsCapability.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $true)]
    [string]
    $name
)
Import-Module SystemInsights

Invoke-InsightsCapability -Name $name -Confirm:$false

}
## [END] Invoke-WACSIInsightsCapability ##
function Remove-WACSIInsightsCapability {
<#

.SYNOPSIS
Remove insights capability and restart service.

.DESCRIPTION
Remove insights capability and restart service.

.ROLE
Administrators

#>
Param
(
    [Parameter(Mandatory = $true)]
    [string]$title
)
Import-Module SystemInsights, Microsoft.PowerShell.Management, Microsoft.PowerShell.Utility

try {
    Remove-InsightsCapability -Name $title -Confirm:$false -ErrorAction Stop

    Restart-Service DPS -ErrorAction Stop
} catch {
    @{
        errorDetail = ("$($_.Exception.Message) $($_.CategoryInfo.GetMessage())");
        hResult = $_.Exception.hResult;
    } | Microsoft.PowerShell.Utility\Write-Output
    throw $_
}
}
## [END] Remove-WACSIInsightsCapability ##
function Set-WACSIClusterSettings {
<#

.SYNOPSIS
Sets item property values for clustered storage settings.

.DESCRIPTION
Sets item property values for clustered storage settings.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.PARAMETER systemDataArchiverValue

.PARAMETER totleStorageValue

.PARAMETER volumeValue

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [uint32]
  $systemDataArchiverValue,
  [Parameter(Mandatory = $true)]
  [uint32]
  $totalStorageValue,
  [Parameter(Mandatory = $true)]
  [uint32]
  $volumeValue
)
Import-Module Microsoft.PowerShell.Management

$systemDataArchiverPath = 'HKLM:\SOFTWARE\Microsoft\Windows\SystemDataArchiver'
$totalStoragePath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\SystemInsights\Capabilities\Total storage consumption forecasting'
$volumePath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\SystemInsights\Capabilities\Volume consumption forecasting'

# system data archiver
Stop-Service -Name DPS

if (Test-Path $systemDataArchiverPath)
{
  Set-ItemProperty -Path $systemDataArchiverPath -Name ClusterVolumesAndDisks -Value $systemDataArchiverValue
}
else
{
  New-Item -Path $systemDataArchiverPath -Force
  New-ItemProperty -Path $systemDataArchiverPath -Name ClusterVolumesAndDisks -Value $systemDataArchiverValue -PropertyType DWORD -Force
}

Start-Service -Name DPS

#  total storage
if (Test-Path $totalStoragePath)
{
  Set-ItemProperty -Path $totalStoragePath -Name ClusterVolumesAndDisks -Value $totalStorageValue
}
else
{
  New-Item -Path $totalStoragePath -Force
  New-ItemProperty -Path $totalStoragePath -Name ClusterVolumesAndDisks -Value $totalStorageValue -PropertyType DWORD -Force
}

#  volume
if (Test-Path $volumePath)
{
  Set-ItemProperty -Path $volumePath -Name ClusterVolumes -Value $volumeValue
}
else
{
  New-Item -Path $volumePath -Force
  New-ItemProperty -Path $volumePath -Name ClusterVolumes -Value $volumeValue -PropertyType DWORD -Force
}

}
## [END] Set-WACSIClusterSettings ##
function Set-WACSIInsightsCapabilityActionCombo {
<#

.SYNOPSIS

.DESCRIPTION
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.PARAMETER name
.PARAMETER okPath
.PARAMETER warningPath
.PARAMETER criticalPath
.PARAMETER errorPath
.PARAMETER nonePath
.PARAMETER okUsername
.PARAMETER okPassword
.PARAMETER warningUsername
.PARAMETER warningPassword
.PARAMETER criticalUsername
.PARAMETER criticalPassword
.PARAMETER errorUsername
.PARAMETER errorPassword
.PARAMETER noneUsername
.PARAMETER nonePassword
.PARAMETER okIncluded
.PARAMETER warningIncluded
.PARAMETER criticalIncluded
.PARAMETER errorIncluded
.PARAMETER noneIncluded
.PARAMETER okDelete
.PARAMETER warningDelete
.PARAMETER criticalDelete
.PARAMETER errorDelete
.PARAMETER noneDelete
.PARAMETER commonCredentialsUsed
.PARAMETER commonUsername
.PARAMETER commonPassword

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $true)]
    [string] $name,
    [string] $okPath,
    [string] $warningPath,
    [string] $criticalPath,
    [string] $errorPath,
    [string] $nonePath,
    [string] $okUsername,
    [string] $okPassword,
    [string] $warningUsername,
    [string] $warningPassword,
    [string] $criticalUsername,
    [string] $criticalPassword,
    [string] $errorUsername,
    [string] $errorPassword,
    [string] $noneUsername,
    [string] $nonePassword,
    [bool] $okIncluded,
    [bool] $warningIncluded,
    [bool] $criticalIncluded,
    [bool] $errorIncluded,
    [bool] $noneIncluded,
    [bool] $okDelete,
    [bool] $warningDelete,
    [bool] $criticalDelete,
    [bool] $errorDelete,
    [bool] $noneDelete,
    [bool] $commonCredentialsUsed,
    [string] $commonUsername,
    [string] $commonPassword

)
Import-Module SystemInsights

function Get-Cred() {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$password,

        [Parameter(Mandatory = $true)]
        [string]$username
    )
    Import-Module Microsoft.PowerShell.Utility, Microsoft.PowerShell.Security

    $securePass = ConvertTo-SecureString -String $password -AsPlainText -Force
    return New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $username, $securePass
}

$okException = $null
$warningException = $null
$criticalException = $null
$errorException = $null
$noneException = $null

if ($okIncluded) {
    try {
        if ($okDelete) {
            Remove-InsightsCapabilityAction -Name $name -Type "Ok" -Confirm:$false
        }
        else {
            $cred = $null;
            if ($commonCredentialsUsed) {
                $cred = Get-Cred -UserName $commonUsername -Password $commonPassword;
            }
            else {
                $cred = Get-Cred -UserName $okUsername -Password $okPassword;
            }
            Set-InsightsCapabilityAction -Name $name -Type "Ok" -Action $okPath -ErrorAction "stop" -Confirm:$false -ActionCredential $cred;
        }
    }
    catch {
        $okException = $_
    }
}

if ($warningIncluded) {
    try {
        if ($warningDelete) {
            Remove-InsightsCapabilityAction -Name $name -Type "Warning" -Confirm:$false
        }
        else {
            $cred = $null;
            if ($commonCredentialsUsed) {
                $cred = Get-Cred -UserName $commonUsername -Password $commonPassword;
            }
            else {
                $cred = Get-Cred -UserName $warningUsername -Password $warningPassword;
            }
            Set-InsightsCapabilityAction -Name $name -Type "Warning" -Action $warningPath -ErrorAction "stop" -Confirm:$false -ActionCredential $cred;
        }
    }
    catch {
        $okException = $_
    }
}

if ($criticalIncluded) {
    try {
        if ($criticalDelete) {
            Remove-InsightsCapabilityAction -Name $name -Type "Critical" -Confirm:$false
        }
        else {
            $cred = $null;
            if ($commonCredentialsUsed) {
                $cred = Get-Cred -UserName $commonUsername -Password $commonPassword;
            }
            else {
                $cred = Get-Cred -UserName $criticalUsername -Password $criticalPassword;
            }
            Set-InsightsCapabilityAction -Name $name -Type "Critical" -Action $criticalPath -ErrorAction "stop" -Confirm:$false -ActionCredential $cred;
        }
    }
    catch {
        $okException = $_
    }
}

if ($errorIncluded) {
    try {
        if ($errorDelete) {
            Remove-InsightsCapabilityAction -Name $name -Type "Error" -Confirm:$false
        }
        else {
            $cred = $null;
            if ($commonCredentialsUsed) {
                $cred = Get-Cred -UserName $commonUsername -Password $commonPassword;
            }
            else {
                $cred = Get-Cred -UserName $errorUsername -Password $errorPassword;
            }
            Set-InsightsCapabilityAction -Name $name -Type "Error" -Action $errorPath -ErrorAction "stop" -Confirm:$false -ActionCredential $cred;
        }
    }
    catch {
        $okException = $_
    }
}

if ($noneIncluded) {
    try {
        if ($noneDelete) {
            Remove-InsightsCapabilityAction -Name $name -Type "None" -Confirm:$false
        }
        else {
            $cred = $null;
            if ($commonCredentialsUsed) {
                $cred = Get-Cred -UserName $commonUsername -Password $commonPassword;
            }
            else {
                $cred = Get-Cred -UserName $noneUsername -Password $nonePassword;
            }
            Set-InsightsCapabilityAction -Name $name -Type "None" -Action $nonePath -ErrorAction "stop" -Confirm:$false -ActionCredential $cred;
        }
    }
    catch {
        $okException = $_
    }
}

# return obj with specific errs to give to notifications pane
@{okException = $okException; warningException = $warningException; criticalException = $criticalException; errorException = $errorException; noneException = $noneException}

}
## [END] Set-WACSIInsightsCapabilityActionCombo ##
function Set-WACSIInsightsCapabilitySchedule {
<#

.SYNOPSIS
Invokes Set-InsightsCapabilitySchedule.

.DESCRIPTION
Invokes Set-InsightsCapabilitySchedule.
The supported Operating Systems are Windows Server 2019.
Copyright (c) Microsoft Corp 2018.

.PARAMETER name
.PARAMETER daily
.PARAMETER hourly
.PARAMETER minute
.PARAMETER monday
.PARAMETER tuesday
.PARAMETER wednesday
.PARAMETER thursday
.PARAMETER friday
.PARAMETER saturday
.PARAMETER sunday
.PARAMETER at
.PARAMETER minutesInterval
.PARAMETER hoursInterval
.PARAMETER daysInterval

.ROLE
Administrators

#>
param (
  [Parameter(Mandatory = $true)] [string] $name,
  [Parameter(Mandatory = $false)] [bool] $daily,
  [Parameter(Mandatory = $false)] [bool] $hourly,
  [Parameter(Mandatory = $false)] [bool] $minute,
  [Parameter(Mandatory = $false)] [bool] $monday,
  [Parameter(Mandatory = $false)] [bool] $tuesday,
  [Parameter(Mandatory = $false)] [bool] $wednesday,
  [Parameter(Mandatory = $false)] [bool] $thursday,
  [Parameter(Mandatory = $false)] [bool] $friday,
  [Parameter(Mandatory = $false)] [bool] $saturday,
  [Parameter(Mandatory = $false)] [bool] $sunday,
  [Parameter(Mandatory = $false)] [datetime] $at,
  [Parameter(Mandatory = $false)] [uint16] $minutesInterval,
  [Parameter(Mandatory = $false)] [uint16] $hoursInterval,
  [Parameter(Mandatory = $false)] [uint16] $daysInterval
)
Import-Module SystemInsights

$arguments = @{}
$arguments += @{"name" = $name}

$daysOfWeek = @()

if ($monday) {
  $daysOfWeek += "Monday"
}

if ($tuesday) {
  $daysOfWeek += "Tuesday"
}

if ($wednesday) {
  $daysOfWeek += "Wednesday"
}

if ($thursday) {
  $daysOfWeek += "Thursday"
}

if ($friday) {
  $daysOfWeek += "Friday"
}

if ($saturday) {
  $daysOfWeek += "Saturday"
}

if ($sunday) {
  $daysOfWeek += "Sunday"
}


if ($daily) {
  $arguments += @{"Daily" = $true}
  $arguments += @{"at" = $at}
  if ($daysInterval)
  {
    $arguments += @{"daysInterval" = $daysInterval}
  }
  else
  {
    $arguments += @{"daysOfWeek" = $daysOfWeek}
  }
}

if ($hourly) {
  $arguments += @{"Hourly" = $true}
  $arguments += @{"hoursInterval" = $hoursInterval}
  $arguments += @{"daysOfWeek" = $daysOfWeek}
}


if ($minute) {
  $arguments += @{"Minute" = $true}
  $arguments += @{"minutesInterval" = $minutesInterval}
  $arguments += @{"daysOfWeek" = $daysOfWeek}
}

Set-InsightsCapabilitySchedule @arguments

}
## [END] Set-WACSIInsightsCapabilitySchedule ##
function Update-WACSINugetFromTemp {
<#

.SYNOPSIS

.DESCRIPTION

.PARAMETER path

.PARAMETER title

.PARAMETER id

.PARAMETER version

.PARAMETER dllName

.ROLE
Administrators

#>
Param
(
    [Parameter(Mandatory = $true)]
    [string]$path,

    [Parameter(Mandatory = $true)]
    [string]$title,

    [Parameter(Mandatory = $true)]
    [string]$id,

    [Parameter(Mandatory = $true)]
    [string]$version,

    [Parameter(Mandatory = $true)]
    [string]$dllName
)
Import-Module SystemInsights, Microsoft.PowerShell.Utility, Microsoft.PowerShell.Security

$sourceDirectoryPath = $path + $id + "." + $version

$destinationDirectoryPath = $env:SystemDrive + "\ProgramData\Microsoft\Windows\SystemInsights\InstalledCapabilities\" + $id

try {
    if (Test-Path $destinationDirectoryPath) { Remove-Item $destinationDirectoryPath -Force -Recurse; }

    Copy-Item -Path $sourceDirectoryPath -Destination $destinationDirectoryPath -Recurse -Force

    $dllPath = $destinationDirectoryPath + "\" + $dllName

    Update-InsightsCapability -Name $title -Library $dllPath -Confirm:$false

    Restart-Service DPS

    $policiesPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Insights\Parameters"
    if (Test-Path $policiesPath) {
        $currentSetting = Get-ItemProperty -Path $policiesPath -Name MaxSerializedLengthInMB -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        if ($currentSetting -eq $null) {
            Set-ItemProperty -Path $policiesPath -Name MaxSerializedLengthInMB -Value 100 -Force
            Restart-Service Insights
        }
        else {
            if ($currentSetting.MaxSerializedLengthInMB -lt 100) {
                Set-ItemProperty -Path $policiesPath -Name MaxSerializedLengthInMB -Value 100 -Force
            }
            Restart-Service Insights
        }
    }
    else {
        New-Item -Path $policiesPath -Force
        New-ItemProperty -Path $policiesPath -Name MaxSerializedLengthInMB -Value 100 -PropertyType DWORD -Force
        Restart-Service Insights
    }

    if (Test-Path $path) { Remove-Item $path -Force -Recurse; }

}
catch {
    @{
        errorDetail = ("$($_.Exception.Message) $($_.CategoryInfo.GetMessage())");
        hResult     = $_.Exception.hResult;
    } | Microsoft.PowerShell.Utility\Write-Output
    throw $_
}
}
## [END] Update-WACSINugetFromTemp ##
function Get-WACSICimWin32LogicalDisk {
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
## [END] Get-WACSICimWin32LogicalDisk ##
function Get-WACSICimWin32NetworkAdapter {
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
## [END] Get-WACSICimWin32NetworkAdapter ##
function Get-WACSICimWin32PhysicalMemory {
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
## [END] Get-WACSICimWin32PhysicalMemory ##
function Get-WACSICimWin32Processor {
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
## [END] Get-WACSICimWin32Processor ##
function Get-WACSIClusterInventory {
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
    return $null -ne (Get-StorageSubSystem clus* | Get-StorageHealthSetting -Name "System.PerformanceHistory.Path")
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
## [END] Get-WACSIClusterInventory ##
function Get-WACSIClusterNodes {
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
## [END] Get-WACSIClusterNodes ##
function Get-WACSIDecryptedDataFromNode {
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
## [END] Get-WACSIDecryptedDataFromNode ##
function Get-WACSIEncryptionJWKOnNode {
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
## [END] Get-WACSIEncryptionJWKOnNode ##
function Get-WACSIServerInventory {
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
  $operatingSystemInfo = Get-CimInstance Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object csName, Caption, OperatingSystemSKU, Version, ProductType, OSType, LastBootUpTime
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
script to query SMBIOS locally from the passed in machineName


.DESCRIPTION
script to query SMBIOS locally from the passed in machine name
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

$result

}
## [END] Get-WACSIServerInventory ##

# SIG # Begin signature block
# MIIoKgYJKoZIhvcNAQcCoIIoGzCCKBcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAaL6kfVSTAhZqD
# GHd/I0up0EeFGJHtkiLvojXuztTGlaCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIBRimzq5E/BsKUcIopF68GER
# YUrezIhBph4xCfYkDHc2MEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAvjb3sqPb25nnMmz/1dokR26WBnaBguvv45wZ6zCJqzoThPdm7xQJr8Mt
# kTCk7c12aaycURrD+k2gLX0twZfS46SrgBBccOT3D+Y5EW3SxZZ5IJwaemEO30vf
# mJV9ZXv53OY/AeWPts0zqbbCOUhJ/gT6AZvokcj7ZLskTw/bT6wBnq2OzTHwbZqC
# yNWJ2CXkM5NYw6R3X+PDxUuIqPDEeJi73E1iI+ObKfzeLeZDJloawzWzsAjItMGH
# unqZMnlu24leQ8R0pbFNECpb8ouaAR2WbsPVckHmeCmFCcdwO/eWE9RlJDSzLcJ0
# 5949cxyfaLRdCb5TlFVaPY8AXYADe6GCF5QwgheQBgorBgEEAYI3AwMBMYIXgDCC
# F3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCAa+Wpz7PQPLs+7bJwVKNmXQa6YC3bseRpp5Mo39NByCAIGZWitW3Yg
# GBMyMDIzMTIwNTE3NDkwNi4xNjNaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
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
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6RnDOjAiGA8yMDIzMTIwNTE1NDE0
# NloYDzIwMjMxMjA2MTU0MTQ2WjB0MDoGCisGAQQBhFkKBAExLDAqMAoCBQDpGcM6
# AgEAMAcCAQACAlTYMAcCAQACAhpcMAoCBQDpGxS6AgEAMDYGCisGAQQBhFkKBAIx
# KDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZI
# hvcNAQELBQADggEBAA6FJBrmGKpshwLO5FnzlbvtN4GiTOokRbJLvXzdhKZqyz1m
# YwHCZmHk/TGDdbfBnUif5zLOfEmYnR4NSRdNHCQl2t0cJS9jjAodGwH0SdXpZWTL
# ZFrzgej78m0MvLbWevJBbu4ulPPRqPPiiWFQPA7CY0J702U1Z5b42Snoj8XU7WKp
# QA8Ux+7VAWpsDSF0So40hZDAicY0wYmxcOCp1XUN/cloO0jim34XjKVxpnfAd4/Y
# PhcpYvDcvoVs1bb8pE2+X/zWq+VJSeuB6xsdaRl4TYBZoZnTPjstbgcdzxtRcBQw
# qQgZjNEROkbzack+NJtLuBCPhkBKOi3WPMVgh3gxggQNMIIECQIBATCBkzB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdj8SzOlHdiFFQABAAAB2DAN
# BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8G
# CSqGSIb3DQEJBDEiBCDhOFu1qC5Xdb6q9Gs6cNCcqt3zz+Lv9j5OUYJeTcHOszCB
# +gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIDrjIX/8CZN3RTABMNt5u73Mi3o3
# fmvq2j8Sik+2s75UMIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAHY/EszpR3YhRUAAQAAAdgwIgQgW1gJkPqemSJqTCUVQrnGASg4rsmD
# PBXbjmDRjddUlU0wDQYJKoZIhvcNAQELBQAEggIAwS8XyA/PqT1fBo7vm+Ek5pv/
# d05IsLFMUUCGm7HEqfU8UcT/TDpTxnYd6vMtvvxZHthnMKvSWqh+QxcpKymxKX/Q
# UoJctOJNNKU3n1gSF/WTLPhyCFnoTO40f9CtuCt63POtWQOY7ga95O8Teg3gHTXi
# wyCd5WaaDp5pcETXZjkIi+ERRzZjMGFeaq6KiNlqtCdIXtdciG4a2i3RclhiToEi
# UUNJ5QsCHnwUBjzdvxQTtYpJLaWMeu/2imDWK9vJGjYc6eCVkAj5p2/IzVWIKzST
# fvPDvDU8VSaXKD+W6XOp9jrLrghoXdeFvct3E6/NYpCEWOBzAzyVHrfp+V/VzysA
# 4RKLOjHpcs8nGAxsDEn2TlEJzHAqtG//HyTzfz9iBAHP5XpYMMb8CslR4Z2JoeVB
# OsT4YiuWtRI87o8phrgz2mAPpVn08QbciLMCm24mIeDdHJQRkxlh+LM1YBkCskmg
# K53y2XrN1kD8evfT+dIjrGMI5S3cs3c2uxhhNWXVbIbtAs5D6H56s2HURxVO5vex
# 8Anu/X73IMn86m6PqpFF2aIay1kM5v03EtqOA38nq0P0dxHu1dwAgHa8jn+xqiTZ
# iAIC7FfRw65wlXY5D6C9ch9MrSa6Zjkk12YZJGRVaWC8PVJJFRm26O9llN9MVIOK
# JHsENbstQUibfMQB+Rk=
# SIG # End signature block
