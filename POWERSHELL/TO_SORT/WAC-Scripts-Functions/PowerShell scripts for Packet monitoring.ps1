function Confirm-WACPMPrerequisitesInstalled {
<#
.SYNOPSIS
Script that checks if prerequisites are installed

.DESCRIPTION
This script checks if Wireshark Dissector and Payload Parser are installed.
These applications are required to parse ETL and ETW data captured by Packet Monitor

.ROLE
Readers

#>

function checkIfWiresharkDissectorIsInstalled {
  $regPath = @(
    "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    "Registry::HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
  )
  if (Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object {
    [bool]($_.PSobject.Properties.name -match "DisplayName") -and  $_.DisplayName -like "*wireshark*" }) {
    return $true
  }
  return $false
}

function checkIfPayloadParserIsInstalled {
  if (Get-Service -Name "PayloadParser" -ErrorAction SilentlyContinue) {
    return $true
  }
  return $false
}

###############################################################################
# Script execution starts here...
###############################################################################
$applications = @{}
$wiresharkInstalled = checkIfWiresharkDissectorIsInstalled
$payloadParserInstalled = checkIfPayloadParserIsInstalled

$applications.Add('wireshark', $wiresharkInstalled)
$applications.Add('payloadParser', $payloadParserInstalled)

$applications

}
## [END] Confirm-WACPMPrerequisitesInstalled ##
function ConvertFrom-WACPMCapturedData {
<#
.SYNOPSIS
Parse captured data

.DESCRIPTION
Parse captured data by running it through the PayloadParser

.ROLE
Readers

#>

Param(
  [Parameter(Mandatory = $true)]
  [String] $filePath
)

if ([String]::IsNullOrWhiteSpace($filePath) -or -not(Test-Path($filePath))) {
  return;
}

function Delete-File($Path) {
  if (Test-Path $Path) {
    Remove-Item -Path $Path -Force
  }
}

Push-Location

$parserDir = (Get-ChildItem $filePath).DirectoryName.TrimEnd('\')
$fileName = (Get-ChildItem $filePath).Name.TrimEnd('\')

# Set environment variable for PayloadParser to get generated ETL files
$PARSER_FILES_PATH = $parserDir + '\'
[Environment]::SetEnvironmentVariable("PARSER_FILES_PATH", $PARSER_FILES_PATH, "Machine")

Set-Location $parserDir

# PayloadParser only accepts an ETL file with the name 'PktMon.etl'.
# So, if the file the user passed has a different name, we rename it
$wasRenamed = $false;
$pktMonFileName = "PktMon.etl"
$pktmonETLPath = Join-Path $parserDir $pktMonFileName
if ($fileName.ToLower() -ne "pktmon.etl") {
  Delete-File -Path $pktmonETLPath
  Rename-Item -Path $filePath -NewName $pktmonETLPath -Force
  $wasRenamed = $true;
}

$logfilePath = $pktmonETLPath.Replace('etl', 'txt')

# Delete the existing PktMOn.txt file since the PayloadParser creates a new file
Delete-File -Path $logfilePath

# Parse data using payload parser. Generates a file PktMon.txt
Start-Service -Name PayloadParser

# We sleep to give the Payload Parser time to complete
Start-Sleep -Seconds 5
Stop-Service -Name PayloadParser -Force

if ($wasRenamed) {
  Rename-Item -Path $pktmonETLPath -NewName $filePath -Force
}

Pop-Location

if (Test-Path($logfilePath)) {
  return $logfilePath
}

}
## [END] ConvertFrom-WACPMCapturedData ##
function Copy-WACPMFileToServer {
<#

.SYNOPSIS
Upload file from localhost to remote server

.DESCRIPTION
Upload file from localhost to remote server

.Parameter source
Source Path

.Parameter destination
Destination Path

.Parameter destinationServer
Server to upload file to

.Parameter username
Server to upload file to

.Parameter password
User password

.ROLE
Readers

#>

param (
  [Parameter(Mandatory = $true)]
  [ValidateScript( { Test-Path $_ -PathType 'Leaf' })]
  [String]$source,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [String]$destination,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [String]$destinationServer,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [String]$username,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [String]$encryptedPassword
)

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode($encryptedData) {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

function Get-UserCredentials {
    param (
      [Parameter(Mandatory = $true)]
      [ValidateNotNullOrEmpty()]
      [String]$username,
      [Parameter(Mandatory = $true)]
      [ValidateNotNullOrEmpty()]
      [String]$encryptedPassword
    )

    $password = DecryptDataWithJWKOnNode $encryptedPassword
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securePassword
    return $credential
}

$Script:credential = Get-UserCredentials $username $encryptedPassword
$Script:serverName = $destinationServer

function Copy-FileToDestinationServer() {
    param (
      [Parameter(Mandatory = $true)]
      [ValidateNotNullOrEmpty()]
      [String]$source,
      [Parameter(Mandatory = $true)]
      [ValidateScript( { Test-Path $_ -isValid })]
      [String]$destination
    )

    # Delete the remote file first if it exists
    Invoke-Command -ComputerName $Script:serverName -Credential $Script:credential -ScriptBlock {
      param($destination)
      if (Test-Path $destination) {
        Remove-Item -Path $destination -Force -ErrorAction SilentlyContinue
      }
    } -ArgumentList $destination

    # Upload the file
    $session = New-PSSession -ComputerName $Script:serverName -Credential $Script:credential
    Copy-Item -Path $source -ToSession $session -Destination $Destination
  }

  function Get-AdminSmbShare {
    $adminSharedFolder = Invoke-Command -ComputerName $Script:serverName -Credential $Script:credential -ScriptBlock {
      return (Get-SmbShare -Name "ADMIN$").Name;
    } -ArgumentList $destination

    return $adminSharedFolder
  }

  function Get-FreeDisk {
    $disk = Invoke-Command -ComputerName $Script:serverName -Credential $Script:credential -ScriptBlock {
      Get-ChildItem function:[d-z]: -n | Where-Object { !(test-path $_) } | Microsoft.PowerShell.Utility\Select-Object  -Last 1
    }

    return $disk
  }

  function Transfer-FileToServer() {
    param (
      [Parameter(Mandatory = $true)]
      [ValidateNotNullOrEmpty()]
      [String]$username,
      [Parameter(Mandatory = $true)]
      [ValidateNotNullOrEmpty()]
      [String]$password,
      [Parameter(Mandatory = $true)]
      [ValidateScript( { Test-Path $_ -isValid })]
      [String]$source
    )
    $serverName = $Script:serverName
    $smbShare = Get-AdminSmbShare
    $shareName = Get-FreeDisk

    # TODO: This is under test to see which between Start-BitsTransfer and Copy-Item is faster
    net use $shareName \\$serverName\$smbShare $password /USER:$username
    Start-BitsTransfer -Source $source -Destination "S:\Pktmon" -TransferType Upload
    net use $shareName /delete
  }


  ###############################################################################
  # Script execution starts here
  ###############################################################################
  if (-not ($env:pester)) {

    $sourcePath = $ExecutionContext.InvokeCommand.ExpandString($source)
    Copy-FileToDestinationServer -Source $sourcePath -Destination $destination

    $fileExists = Invoke-Command -ComputerName $Script:serverName -Credential $Script:credential -ScriptBlock {
      param($destination)
      return (Test-Path $destination);
    } -ArgumentList $destination

    if ($fileExists) {
      return New-Object PSObject -Property @{success = $true; }
    }
    else {
      return New-Object PSObject -Property @{success = $false; }
    }
  }

}
## [END] Copy-WACPMFileToServer ##
function Get-WACPMComponents {
<#

.SYNOPSIS
Get a list of all pktmon components as json and parse to custom result object

.DESCRIPTION
Get a list of all pktmon components as json and parse to custom result object

.ROLE
Readers

#>

# Method is used to convert components' properties format to a more workable one
# from Properties: {@{p1=v1}, @{p1=v2}, @{p2=v3}} (array of hashtables)
# to Properties: {p1,p2} where p1 = {v1,v2} and p2=v3 (hashtable of values)
function convertProperties($componentList) {
  $componentList.components | ForEach-Object {

    $convertedProperties = @{ }
    $_.Properties | ForEach-Object {
      $propName = $_.Name
      if ($propName -eq "Nic ifIndex" -or $propName -eq "EtherType") {
        $convertedProperties[$propName] += , $_.Value
      }
      else {
        $convertedProperties.Add($propName, $_.Value)
      }
    }

    $_.Properties = $convertedProperties
  }
}

function getVmSwitches($componentList) {
    return $componentList | Where-Object { $_.Type -eq "VMS Protocol Nic" }
}

# For the purpose of building correct associations
# consider adapters to be the components that are not filter, protocol or a virtual switch
function getAdapters($componentList, $nameMap) {
    $adapterList = $componentList | Where-Object { $_.Type -ne "Protocol" -and $_.Type -ne "Filter" -and $_.Type -ne "VMS Protocol Nic" }
    $adapterList | ForEach-Object {
      # Handle the adapter duplicates
      handleAdapterDuplicate $_ $adapterList $nameMap
    }

    return $adapterList
}

function getFilters($componentList) {
  return $componentList | Where-Object {$_.Type -eq "Filter"}
}

function getProtocols($componentList) {
  return $componentList | Where-Object {$_.Type -eq "Protocol"}
}

# Method finds virtual adapters associated with a given virtual switch in a given list of adapters
# It also updates the names of the duplicate items in the name map
function getVirtualAdaptersPerVSwitch($vmSwitchComponent, $adapterComponents, $nameMap) {
    $vSwitchExtIfIndex = $vmSwitchComponent.Properties."Ext ifIndex"

    $vadapters = $adapterComponents | Where-Object { $_.Properties."Ext ifIndex" -eq $vSwitchExtIfIndex }
    $vadapters | ForEach-Object {
        $currentAdapter = $_
        $currentAdapter.Grouped = $true
        # Each adapter and its duplicate belong to its group
        $currentAdapter.ComponentGroup = $currentAdapter.Name
        $nameMap["$($currentAdapter.Id)"].ComponentGroup = $currentAdapter.ComponentGroup

        # Find the filters for each virtual adapter
        processComponentFilters $currentAdapter $filters $nameMap

        # Find the protocols for each virtual adapter
        processComponentProtocols $currentAdapter $protocols $nameMap
    }

    return $vadapters
}

# Method finds the virtual network adapters in a given list of virtual adapters
function getVirtualNetworkAdapters($virtualAdapters) {
    return $virtualAdapters | Where-Object { $_.Type -eq "Host vNic" }
}

# Method finds the vm network adapters in a given list of virtual adapters
function getVirtualMachineNetworkAdapters($virtualAdapters) {
    return $virtualAdapters | Where-Object { $_.Type -eq "VM Nic" }
}

# Method finds the physical adapters associated with a virtual switch in a given list of adapters
function getPhysicalAdaptersPerVSwitch($vmSwitchComponent, $adapterComponents, $nameMap) {
  $physicalComponents = @()

  # Get all Nic ifIndex values for the vm switch. One VM switch can have multiple Nic ifIndex values. Each Nic ifIndex maps to one physical adapter.
  $nicIfIndices = $vmSwitchComponent.Properties."Nic ifIndex"

  $nicIfIndices | ForEach-Object {
    $nicIdx = $_
    $adapterComponents | Where-Object { $_.Properties."ifIndex" -eq $nicIdx } | ForEach-Object {
        $_.Grouped = $true
        $_.ComponentGroup = $_.Name
        $nameMap["$($_.Id)"].ComponentGroup = $_.ComponentGroup

        # Get the filters for each physical adapter
        processComponentFilters $_ $filters $nameMap

        # Get the protocols for each physical adapter
        processComponentProtocols $_ $protocols $nameMap

        $physicalComponents += $_
    }
  }

  return $physicalComponents
}

# Finds the duplicate adapter and updates its name in the map and updates the component's edges
function handleAdapterDuplicate($adapter, $adapterComponents, $nameMap) {
  # All adapter duplicates are of type Miniport. If the adapter we're trying to process is Miniport, then ignore
  if ($adapter.Type -eq "Miniport") {
    return
  }

  $adapter | Add-Member -NotePropertyName DuplicateIds -NotePropertyValue @()

  $duplicate = $adapterComponents | Where-Object { $_.Id -ne $adapter.Id -and $_.Properties.ifIndex -eq $adapter.Properties.ifIndex -and $_.Properties."MAC Address" -eq $adapter.Properties."MAC Address" }
  if ($duplicate) {
      $duplicate = $duplicate[0]
      $duplicate.Grouped = $true
      $duplicate.ComponentGroup = $adapter.ComponentGroup

      $nameMap["$($duplicate.Id)"].Name = $adapter.Name
      $nameMap["$($duplicate.Id)"].ComponentGroup = $duplicate.ComponentGroup

      # Only duplicate components carry the info about the edges, so make sure to add it to the actual component
      if($duplicate.Type -eq "Miniport") {
        $adapter.Edges = $duplicate.Edges
      }
      $adapter.DuplicateIds += $duplicate.Id
  }
}

# Process data for current filter and return the next one
function getNextFilter($component, $currentFilter, $filters, $nameMap) {
  $ifIndex = $currentFilter.Properties["ifIndex"]

  # Each filter belongs to the group of its adapter component
  $nextFilter = $filters | Where-Object {$_.Properties["Lower ifIndex"] -eq $ifIndex}

  if ($nextFilter) {
    $nextFilter.Grouped = $true
    $nextFilter.ComponentGroup = $component.ComponentGroup
    $nameMap["$($nextFilter.Id)"].ComponentGroup = $nextFilter.ComponentGroup
  }

  return $nextFilter
}

# Method finds all filters build on top of a component in order and
# adds them to the component as a property
function processComponentFilters($component, $filters, $nameMap) {
  $ifIndex = $component.Properties["ifIndex"]

  $componentFilters = $filters | Where-Object {$_.Properties["Miniport ifIndex"] -eq $ifIndex}

  if ($componentFilters) {
    # Array will contain the component's filters in the order they are applied
    $orderedFilters = @()

    # Handle 1st filter separately - 1st filter doesn't have Lower ifIndex in its properties
    $firstFilter = $componentFilters | Where-Object {-not $_.Properties["Lower ifIndex"]}
    $firstFilter.Grouped = $true
    $firstFilter.ComponentGroup = $component.ComponentGroup

    $nameMap["$($firstFilter.Id)"].ComponentGroup = $firstFilter.ComponentGroup
    $orderedFilters += $firstFilter

    # The rest of the filtes in the sequence are chained one after the other
    $currentFilter = $firstFilter
    while ($currentFilter) {
      $nextFilter = getNextFilter $component $currentFilter $componentFilters $nameMap
      $orderedFilters += $nextFilter
      $currentFilter = $nextFilter
    }

    $component | Add-Member -NotePropertyName Filters -NotePropertyValue $orderedFilters
  }

}

function processComponentProtocols($component, $protocols, $nameMap) {
  $ifIndex = $component.Properties["ifIndex"]

  $componentProtocols = $protocols | Where-Object {$_.Properties["Miniport ifIndex"] -eq $ifIndex}

  $componentProtocols | ForEach-Object {
    # Each protocol belongs to the group of its adapter component
    $_.Grouped = $true
    $_.ComponentGroup = $component.ComponentGroup
    $nameMap["$($_.Id)"].ComponentGroup = $_.ComponentGroup
  }

  if ($componentProtocols) {
    if ($componentProtocols.GetType().name -eq 'PSCustomObject') {
      $componentProtocols = @($componentProtocols)
    }
    $component | Add-Member -NotePropertyName Protocols -NotePropertyValue $componentProtocols
  }
}

# Method builds the adapter associations for a given virtual switch from a given list of adapters.
# It adds 3 properties to the virtual switch component:
# virtualNetworkAdapters - the list of virtual network adapters for this switch out of all adapters
# virtualMachineNetworkAdapters - the list of vm network adapters for this switch
# physicalNetworkAdapters - the list of physical adapters associated with this switch
# Filters - list of filters applied on top of the switch in order
# Protocols - list of protocols applied on top of the switch
function processVmSwitchComponent($vmSwitchComponent, $adapterComponents, $filters, $protocols, $nameMap) {
    $addedProperties = @{ }

    # 1. Populate the switch name (in case it's missing) and component group
    if (-not $vmSwitchComponent.Name) {
      $name = $vmSwitchComponent.Properties."Switch Name"
      $vmSwitchComponent.Name = $name

      $nameMap["$($vmSwitchComponent.Id)"].Name = $name
    }

    $vmSwitchComponent.Grouped = $true
    $vmSwitchComponent.ComponentGroup = $vmSwitchComponent.Name
    $nameMap["$($vmSwitchComponent.Id)"].ComponentGroup = $vmSwitchComponent.ComponentGroup

    # 2. Handle the vswitch duplicates - virtual switches have 1 original and (at least) 2 duplicates
    $vmSwitchComponent | Add-Member -NotePropertyName DuplicateIds -NotePropertyValue @()

    $duplicates = $adapterComponents | Where-Object {$_.Properties.ifIndex -eq $vmSwitchComponent.Properties."Ext ifIndex"}
    $duplicates | ForEach-Object {
      $_.Grouped = $true
      $_.ComponentGroup = $vmSwitchComponent.ComponentGroup
      $vmSwitchComponent.DuplicateIds += $_.Id

      $nameMap["$($_.Id)"].Name = $vmSwitchComponent.Name
      $nameMap["$($_.Id)"].ComponentGroup = $vmSwitchComponent.ComponentGroup

      # Only the Miniport duplicate has the vswitch Edges
      # Also grab the ifIndex, we need it for finding the filters and protocols
      if($_.Type -eq "Miniport") {
        $vmSwitchComponent.Edges = $_.Edges
        $vmSwitchComponent.Properties["ifIndex"] = $_.Properties["ifIndex"]
      }
    }

    # 3. Find all virtual adapters associated with a given virtual switch
    $virtualAdapters = getVirtualAdaptersPerVSwitch $vmSwitchComponent $adapterComponents $nameMap

    # 4 Group the virtual adapters into categories
    # 4.1 Find the Virtual Network Adapters from the virtual components:
    $virtualNetworkAdapters = @()
    $vnas = getVirtualNetworkAdapters $virtualAdapters
    $vnas | ForEach-Object {
      $virtualNetworkAdapters += $_
    }
    $addedProperties += @{"virtualNetworkAdapters" = $virtualNetworkAdapters }

    # 4.2 Find the Virtual Machine Network Adapters from the virtual components:
    $virtualMachineNetworkAdapters = @()
    $vmnas = getVirtualMachineNetworkAdapters $virtualAdapters
    $vmnas | ForEach-Object {
      $virtualMachineNetworkAdapters += $_
    }
    $addedProperties += @{"virtualMachineNetworkAdapters" = $virtualMachineNetworkAdapters }

    # 5. Find all physical adapters associated with a given virtual switch
    $physicalAdapters = @()
    $pas = getPhysicalAdaptersPerVSwitch $vmSwitchComponent $adapterComponents $nameMap
    $pas | ForEach-Object {
      $physicalAdapters += $_
    }
    $addedProperties += @{"physicalNetworkAdapters" = $physicalAdapters }

    # 6. Get the filters for the switch
    processComponentFilters $vmSwitchComponent $filters $nameMap

    # 7. Get the protocols for the switch
    processComponentProtocols $vmSwitchComponent $protocols $nameMap

    $vmSwitchComponent | Add-Member -NotePropertyMembers $addedProperties
    return $vmSwitchComponent
}


###############################################################################
# Script execution starts here...
###############################################################################

$components = pktmon list -i --json | ConvertFrom-Json

# (1) Convert Properties and (2) Name map between component id and component name
$nameMap = @{ }
$components | ForEach-Object {
  $componentGroup = $_.Group

  $_.Components | ForEach-Object {
    $componentName = $_.Name

    $_ | Add-Member -NotePropertyName Grouped -NotePropertyValue $false
    $_ | Add-Member -NotePropertyName ComponentGroup -NotePropertyValue $componentGroup

    $convertedProperties = @{ }

    if (-not($nameMap.ContainsKey("$($_.Id)"))) {
        $nameMap.Add("$($_.Id)", @{"Name" = $componentName; "ComponentGroup" = $componentGroup})
    }

    $_.Properties | ForEach-Object {

      $propName = $_.Name
      if ($propName -eq "Nic ifIndex" -or $propName -eq "EtherType") {
        $convertedProperties[$propName] += , $_.Value
      }
      else {
        $convertedProperties.Add($propName, $_.Value)
      }
    }

    $_.Properties = $convertedProperties
  }
}


$componentList = $components.Components


$vmSwitchComponents = getVmSwitches $componentList
$adapters = getAdapters $componentList $nameMap
$filters = getFilters $componentList
$protocols = getProtocols $componentList

$vmSwitches = @( )

# Construct the associated objects for each vm switch on the system
$vmSwitchComponents | ForEach-Object {
    $vmSwitches += processVmSwitchComponent $_ $adapters $filters $protocols $nameMap
}

$floatingGroup = @{"Name" = "Standalone Adapters"; "Type" = "Unbound" }

$floatingAdapters = @()
$adapters | Where-Object { !$_.Grouped -and $_.Type -ne "HTTP" } | ForEach-Object {
  $_.Grouped = $true

  processComponentFilters $_ $filters $nameMap
  processComponentProtocols $_ $protocols $nameMap

  $floatingAdapters += $_
}

$floatingGroup += @{"Adapters" = $floatingAdapters }

$floatingGroup = [PSCustomObject]$floatingGroup
$vmSwitches += $floatingGroup

# THE END RESULT
# Tree - list of VM switch association trees
# NameMap - map between component id and component name
$result = [PSCustomObject]@{"treeList" = $vmSwitches; "nameMap" = $nameMap }
$result

}
## [END] Get-WACPMComponents ##
function Get-WACPMCounters {
<#

.SYNOPSIS
Get current pktmon counters

.DESCRIPTION
Get current pktmon counters

.ROLE
Readers

#>
pktmon counters

}
## [END] Get-WACPMCounters ##
function Get-WACPMLogPath {
<#

.SYNOPSIS
Get path for pktmon log file

.DESCRIPTION
Get path for pktmon log file

.PARAMETER logType
File extension. Recognised file extensions are etl, txt, and pcapng.

.ROLE
Administrators

#>

Param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string] $logType
)

if (-not(@('ETL', 'TXT', 'PCAPNG').Contains($logType.ToUpper()))) {
  Throw "Invalid file extensions: $logType. Recognised file extensions are etl, txt, and pcapng.";
}

Push-Location

$Script:logFileDir = [Environment]::GetEnvironmentVariable("PARSER_FILES_PATH", "Machine").TrimEnd('\')

Set-Location $Script:logFileDir

function getExistingPathToType($extension) {
  $fileName = switch($extension) {
    ETL { "PktMon.etl" }
    TXT { "PktMonText.txt" }
    PCAPNG { "PktMon.pcapng"}
  }

  $logPath = $Script:logFileDir + "\$fileName"
  if (Test-Path $logPath) {
    return $logPath
  }

  return $null
}

function getNewPathToType($type) {
  $etlLogPath = getExistingPathToType "ETL"
  if (!$etlLogPath) {
    # We need the ETL file to convert to txt or pcapng format
    return $null
  }

  if ($type -eq "ETL") {
    return $etlLogPath
  }

  # If we have a previously stored log with given extension, we want to make sure to clean it up first
  $existingLogPath = getExistingPathToType $type
  if ($existingLogPath) {
    Remove-Item $existingLogPath
  }

  switch($type) {
    TXT {
      # Convert log file to text format.
      # pktmon etl2txt PktMon.etl --out PktMonText.txt | Out-Null
      pktmon format PktMon.etl --out PktMonText.txt | Out-Null
    }

    PCAPNG {
      # Convert log file to pcapng format. Dropped packets are not included by default.
      # pktmon etl2pcap PktMon.etl --out PktMon.pcapng | Out-Null
      pktmon pcapng PktMon.etl --out PktMon.pcapng | Out-Null
    }
  }

  return getExistingPathToType $type
}

###############################################################################
# Script execution starts here...
###############################################################################
$logType = $logType.ToUpper()
getNewPathToType $logType

}
## [END] Get-WACPMLogPath ##
function Get-WACPMPacketMonitorLogFile {
<#
.SYNOPSIS
Gets the packet monitoring log file

.DESCRIPTION
Return path to captured data file

.ROLE
Readers

#>

Param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [string] $action,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullorEmpty()]
    [string] $logFilesDir
)

function Get-PktmonLogFilePath($PktmonResult, $logFilesDir) {
    if ($null -eq $PktmonResult) {
      return
    }


    $logFilesDir = $ExecutionContext.InvokeCommand.ExpandString($logFilesDir)
    $pathToEtl = Join-Path $logFilesDir "PktMon.etl"
    if (-not(Test-Path $pathToEtl)) {
        return
    }

    return $pathToEtl
  }

function Get-PktmonStatus() {
    $pktmonHelp = pktmon help
    if (-not ($pktmonHelp -match "status")) {
        return $null
    }

    <##
    There are time when you stop pktmon and it shows All Counters zero message
    and there is no packet event data file. This checks that pktmon is running and
    data is being save to a file. This file we are checking for is later passed to
    the PayloadParser where it is converted from ETL to json format.
  #>
    $pktmonStatus = pktmon status

    # if packetmon is not running, the size of the array will be 2.
    if ($pktmonStatus.length -eq 2) {
        return $null
    }

    return $pktmonStatus
}


###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
    $action = $action.ToLower()

    if ($action -eq "stop") {
        $pktmonResult = pktmon stop
    }
    elseif ($action -eq "status") {
        $pktmonResult = Get-PktmonStatus
    }
    else {
        return;
    }

    return Get-PktmonLogFilePath -PktmonResult $pktmonResult -logFilesDir $logFilesDir
}

}
## [END] Get-WACPMPacketMonitorLogFile ##
function Get-WACPMPartitionedData {
<#
.SYNOPSIS
Partition parsed packetMon data

.DESCRIPTION
Partition parsed packetMon data into chunks of 100MB or less

.PARAMETER sourceFile
Path to log file

.ROLE
Readers

#>

Param(
  [Parameter(Mandatory = $true)]
  [string] $sourceFile
)

$generatedFile = [System.Collections.ArrayList]@()

$upperBound = 100MB

$fileSize = (Get-Item $sourceFile).length / 1MB

$parentFolder = Split-Path -Parent $sourceFile

# Delete existing files
Get-Item -Path $parentFolder\PktMonChunk*.txt | Remove-Item -Force -ErrorAction SilentlyContinue


$reader = [io.file]::OpenRead($sourceFile)

$buffer = New-Object byte[] $upperBound

$count = $idx = 1

try {
  # "Splitting $sourceFile using $upperBound bytes per file."
  do {
    $count = $reader.Read($buffer, 0, $buffer.Length)
    if ($count -gt 0) {
      $destinationFile = (Join-Path $parentFolder "PktMonChunk{0}.txt") -f ($idx)
      $writer = [io.file]::OpenWrite($destinationFile)
      try {
        # "Writing to $destinationFile"
        $writer.Write($buffer, 0, $count)
      }
      finally {
        [Void]$generatedFile.Add($destinationFile)
        $writer.Close()
      }
    }
    $idx ++
  } while ($count -gt 0)
}
finally {
  $reader.Close()
}

return $generatedFile

}
## [END] Get-WACPMPartitionedData ##
function Get-WACPMPktMonInstallStatus {
<#

.SYNOPSIS
Check if pktmon is installed or not

.DESCRIPTION
Check if pktmon is installed or not

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
$pktmonHelp = pktmon help
if ($pktmonHelp -match "unload") {
    @{ state = "Available" }
}
else {
    @{ state = "NotSupported" }
}

}
## [END] Get-WACPMPktMonInstallStatus ##
function Get-WACPMSavedLogs {
<#

.SYNOPSIS
Get path for pktmon log file

.DESCRIPTION
Get path for pktmon log file

.ROLE
Readers

#>
Param(
  [Parameter(Mandatory = $true)]
  [string] $logFolder
)


function Get-PktMonSavedLogs($logFolder) {

  $result = [System.Collections.ArrayList]@()

  $savesLocation = Join-Path $env:SystemDrive $logFolder

  if (-not(Test-Path($savesLocation))) {
    New-Item -Path $savesLocation -ItemType "directory" -Force | Out-Null
  }

  # We only open ETL files because we need to run it through the PayloadParser.
  Get-ChildItem $savesLocation | Where-Object { $_.Name -match "ETL" } | Sort-Object -Property LastWriteTime -Descending | ForEach-Object {
    $log = @{"Name" = $_.Name; "Path" = Join-Path $savesLocation $_.Name; "LastWriteTime" = $_.LastWriteTime }
    # $result += $log
    $result.Add($log) | Out-Null
  }

  return ,$result
}

Get-PktMonSavedLogs -LogFolder $logFolder

}
## [END] Get-WACPMSavedLogs ##
function Import-WACPMCapture {
<#
.SYNOPSIS
Get pktmon capture results

.DESCRIPTION
Get packet monitor logs. These are the results from the PayloadParser

.PARAMETER pathToLog
Path to log file

.ROLE
Readers

#>


Param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullorEmpty()]
  [string] $pathToLog
)

$contents = Get-Content $pathToLog -Raw -Encoding UTF8
$contents

}
## [END] Import-WACPMCapture ##
function Import-WACPMPrerequisite {
<#
.SYNOPSIS
Script downloads Packet Monintor prerequsites

.DESCRIPTION
This script downloads prerequsites needed to parse Packet Monitor ETL data

.Parameter uri
URI for the prerequisite

.Parameter destinationFolder
Folder to save file to

.Parameter fileName
Name of the prerequisite to install. wireshark.exe or payloadparser.zip

.ROLE
Readers
#>


param (
  [Parameter(Mandatory = $true)]
  [String]$uri,
  [Parameter(Mandatory = $true)]
  [String]$destinationFolder,
  [Parameter(Mandatory = $true)]
  [String]$fileName
)

function Compress-WiresharkDissectorFile {
  param (
    [Parameter(Mandatory = $true)]
    [String]$sourcePath,
    [Parameter(Mandatory = $true)]
    [String]$destinationPath
  )

  $prerequisiteName = (Get-Item $sourcePath).BaseName.ToLower()
  if ($prerequisiteName -eq "wireshark") {
    Compress-Archive -Path $sourcePath -DestinationPath $destinationPath -Update
  }

  return destinationPath;
}

$destinationFolder = $ExecutionContext.InvokeCommand.ExpandString($destinationFolder)
if (-not(Test-Path $destinationFolder)) {
  New-Item -Path $destinationFolder -ItemType "Directory" -Force | Out-Null
}

$downloadLocation = Join-Path -Path $destinationFolder -ChildPath $fileName

# Remove the file if it exists. This is because the file could be corrupted or a more recent version is available
if (Test-Path $downloadLocation) {
  Remove-Item -Path $downloadLocation -Recurse -Force
}

Invoke-WebRequest -Uri $uri -OutFile $downloadLocation

if (Test-Path $downloadLocation) {
  $fileName = (Get-Item -Path $downloadLocation).Name
  return New-Object PSObject -Property @{executablePath = $downloadLocation; fileName = $fileName }
}

}
## [END] Import-WACPMPrerequisite ##
function Install-WACPMPayloadParser {
<#
.SYNOPSIS
Script installs PayloadParser

.DESCRIPTION
This script installs PayloadParser

.Parameter executablePath
Path to PayloadParser

.ROLE
Readers
#>


param (
  [Parameter(Mandatory = $true)]
  [String]$path
)

$path = $ExecutionContext.InvokeCommand.ExpandString($path)

$parentDir = Split-path -Parent $path
Expand-Archive $path -DestinationPath $parentDir -ErrorAction SilentlyContinue

$destinationDir = Join-Path -Path $parentDir -ChildPath "PayloadParser"
$installerFilePath = "$destinationDir\PayloadParserSetup.msi"
Start-Process -FilePath "$env:systemroot\system32\msiexec.exe" -ArgumentList "/i `"$installerFilePath`" /qn /passive" -Wait | Out-Null

# Remove the executable downloaded
Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path $destinationDir -Recurse -Force -ErrorAction SilentlyContinue

if (Get-Service "PayloadParser") {
  return New-Object PSObject -Property @{success = $true; }
}
else {
  return New-Object PSObject -Property @{success = $false; }
}

}
## [END] Install-WACPMPayloadParser ##
function Install-WACPMWiresharkDissector {
<#
.SYNOPSIS
Script installs Wireshark Dissector

.DESCRIPTION
This script installs Wireshark Dissector

.Parameter path
Path to install Wireshark Dissector

.ROLE
Readers
#>


param (
  [Parameter(Mandatory = $true)]
  [String]$path
)


function checkIfWiresharkDissectorIsInstalled {
  $regPath = @(
    "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    "Registry::HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
  )
  if (Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object {
    [bool]($_.PSobject.Properties.name -match "DisplayName") -and  $_.DisplayName -like "*wireshark*" }) {
    return $true
  }
  return $false
}

$path = $ExecutionContext.InvokeCommand.ExpandString($path)

Start-Process -FilePath $path -ArgumentList "/S", "/v", "/qn" -PassThru | Out-Null

# Stop the process on completion
$wiresharkProcess = Get-Process | Where-Object { $_.Product -eq "Wireshark Dissect" -or $_.ProcessName -eq "Un_A" }
if ($wiresharkProcess) {

  $count = 0
  while ($true) {

    # Somwtimes, the Wireshark installer process does not stop and we need to force it
    # to stop if Wireshark is installed successfully. We continue to poll until it is installed
    $wiresharkInstalled = checkIfWiresharkDissectorIsInstalled
    if ($wiresharkInstalled) {
      $wiresharkProcess | Stop-Process -Force | Out-Null
      return New-Object PSObject -Property @{success = $true; }
    }

    $count += 1

    # This buffer time ensures the installation and post-installation clean-up is done before we stop the process
    Start-Sleep -Seconds 5

    # If the installer is not done in 10seconds, we might have a problem. We force stop the installer and throw a timeOut error
    if ($count -gt 2) {
      $wiresharkProcess | Stop-Process -Force | Out-Null
      Throw (new-object System.TimeoutException)
    }
  }
}

# Remove the executable downloaded
Remove-Item -Path $path -Force -ErrorAction SilentlyContinue

if ($wiresharkInstalled) {
  return New-Object PSObject -Property @{success = $true; }
} else {
  return New-Object PSObject -Property @{success = $false; }
}

}
## [END] Install-WACPMWiresharkDissector ##
function Resolve-WACPMDestinationFilePath {
<#
.SYNOPSIS
Resolves a string to a valid path that includes the destination server

.DESCRIPTION
Resolves a string to a valid path that includes the destination server

.Parameter path
String of path to resolve

.Parameter server
Destination server

.ROLE
Readers
#>

param (
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [String]$path,
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [String]$server
)


function Resolve-DestinationPath() {
  param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String]$path,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [String]$server
  )

  $path = $ExecutionContext.InvokeCommand.ExpandString($path)

  # Create the destination folder if it does not exist
  $parentFolder = Split-Path -Parent $path

  if (-not(Test-Path $parentFolder)) {
    New-Item -Path $parentFolder -ItemType Directory -Force | Out-Null
  }

  $rootDrive = (Get-Location).Drive.Root
  $newPath = $path.Replace($rootDrive, "").Trim("\")

  $server = "\\" + $server.Trim("\")
  $resolvedPath = Join-Path -Path $server -ChildPath $newPath
  return $resolvedPath
}

$resolvedPath = Resolve-DestinationPath -Path $path -Server $server
return $resolvedPath

}
## [END] Resolve-WACPMDestinationFilePath ##
function Resolve-WACPMFilePath {
<#
.SYNOPSIS
Resolves a string to a valid path

.DESCRIPTION
Resolves a string to a valid path

.Parameter path
String of path to resolve

.ROLE
Readers
#>

param (
  [Parameter(Mandatory = $false)]
  [String]$path,
  [Parameter(Mandatory = $false)]
  [String]$smbShare
)

# We do not know the drives that are available in the destination node we need to upload
# the file to. So we need to use $ENV:Temp. This is a string that we need to resolve
# and get a valid name for. This string will be the resoved path we pass to the
# upload and install functions
if (-not([String]::IsNullOrWhiteSpace($path))) {
  $resolvedPath = $ExecutionContext.InvokeCommand.ExpandString($Path)

  $parentDir = Split-path -Parent $resolvedPath
  if (-not(Test-Path $parentDir)) {
    New-Item -Path $parentDir -ItemType "Directory" -Force | Out-Null
  }

  return $resolvedPath
}
elseif (-not([String]::IsNullOrWhiteSpace($smbShare))) {
  $destinationFolder = (Get-SmbShare -Name $smbShare).Path
  $resolvedPath = Join-Path -Path $destinationFolder -ChildPath $path
  return $resolvedPath
}

}
## [END] Resolve-WACPMFilePath ##
function Save-WACPMLog {
<#

.SYNOPSIS
Get path for pktmon log file

.DESCRIPTION
Get path for pktmon log file

.ROLE
Readers

#>
Param(
  [Parameter(Mandatory = $true)]
  [string] $srcLogPath,
  [Parameter(Mandatory = $true)]
  [string] $destLogFolder,
  [Parameter(Mandatory = $true)]
  [string] $logName,
  [Parameter(Mandatory = $true)]
  [boolean] $newCapture
)

$Script:srcLogPath = $srcLogPath
$Script:destLogFolder = $destLogFolder
$Script:logName = $logName


function Remove-FilesByExtension {
  Param(
    [Parameter(Mandatory = $true)]
    [string] $extension,
    [Parameter(Mandatory = $true)]
    [string] $location
  )

  $logsSorted = Get-ChildItem $location -File | Where-Object { $_.Name -match $extension } | Sort-Object -Property LastWriteTime -Descending

  # If destination folder exists, check if we need to clear some of its contents
  $savedLogsCount = $logsSorted.Count

  # Limit number of saved logs - only keep the 5 most recent logs
  $maxSaveCount = 5

  # If we have more logs than our limit clean the oldest ones
  if ($maxSaveCount -le $savedLogsCount) {
    $logsToDelete = $logsSorted | Microsoft.PowerShell.Utility\Select-Object -Last ($savedLogsCount - $maxSaveCount + 1)

    $logsToDelete | ForEach-Object {
      Remove-Item -Path "$($location)\$($_.Name)"
    }
  }
}

function Get-FileExtension($fileName) {
  if ($fileName -match "etl") { return "etl" }
  elseif ($fileName -match "txt") { return "txt" }
  elseif ($fileName -match "pcapng") { return "pcapng" }
}

function Save-CapturedLog($isNewCapture) {
  if ($isNewCapture) {
    # If no source - nothing to copy
    if (-not (Test-Path $Script:srcLogPath)) {
      return $null
    }

    $savesLocation = Join-Path $env:SystemDrive $Script:destLogFolder
    if (Test-Path $savesLocation) {
      # If destination folder exists, check if we need to clear some of its contents
      $extension = Get-FileExtension $Script:logName
      Remove-FilesByExtension -Extension $extension -Location $savesLocation
    }
    else {
      New-Item $savesLocation -ItemType Directory -Force | Out-Null
    }

    # Finally, copy the file
    $destination = Join-Path $savesLocation $Script:logName
    Copy-Item -Path $Script:srcLogPath -Destination $destination
  }
  else {
    $savedLog = Get-Item $Script:logName
    if ($savedLog) {
      $savedLog.LastWriteTime = (Get-Date)
    }
  }
}

###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
  Save-CapturedLog -isNewCapture $newCapture
}

}
## [END] Save-WACPMLog ##
function Set-WACPMFilters {

<#

.SYNOPSIS
add pktmon filters

.DESCRIPTION
add pktmon filters

.ROLE
Readers

#>
Param(
    [Parameter(Mandatory = $true)]
    [string[]] $filters
)

pktmon unload

foreach ($filter in $filters) {
  $command = 'pktmon filter add' + $filter
  Invoke-Expression $command
}

}
## [END] Set-WACPMFilters ##
function Start-WACPMCapture {
<#

.SYNOPSIS
start pktmon capture

.DESCRIPTION
start pktmon capture

.PARAMETER startArgs
A string of flags (and their values) to pass to the `pktmon start` command.
Example "--components nics --etw".
For more details on usage, see: pktmon start help

.PARAMETER filters
An array of string of flags (and their values) to pass to the `pktmon filter add` command.
Example "--ip 192.168.20.1 192.168.20.100".
For more details on usage, see: pktmon filter add help

.ROLE
Readers
#>

Param(
    [Parameter(Mandatory = $false)]
    [PSCustomObject] $startArgs,
    [Parameter(Mandatory = $false)]
    [System.Array] $filters,
    [Parameter(Mandatory = $false)]
    [string] $logFilesDir
)

function isEmpty ($object) {
    if ($null -eq $object) {
        return $true
    }

    return $object.count -eq 0
}

function Add-Filters($filters) {
    if ( isEmpty($filters) ) { return; }

    foreach ($filter in $filters) {

        # Add capture filter
        if (-not([string]::IsNullOrWhitespace($filter))) {
            $filterCommand = "pktmon filter add " + $filter
            Invoke-Expression -Command $filterCommand
        }
    }
}

function Parse-StartArguments() {
    param (
        [Parameter(Mandatory = $false)]
        [PSCustomObject]$arguments
    )

    if (isEmpty($arguments)) {
        return '';
    }

    $result = ''
    if (-not(isEmpty($arguments.component))) {
        $components = $arguments.component -join " "
        $result += " --comp $($components)"
    }

    if ($arguments.dropped) {
        $result += " --type drop"
    }

    return $result.Trim();
}

function Start-PktMon($parsedStartArguments, $logFilesDir) {
    $logFilesDir = $ExecutionContext.InvokeCommand.ExpandString($logFilesDir)
    if (-not(Test-Path $logFilesDir)) {
        New-Item -Path $logFilesDir -ItemType Directory | Out-Null
    }

    # TODO: Test with smaller files.
    # File size in MB
    # NOTE: (22 April 2021) File size limited to 100MB since the Payload Parser
    # as of this date hangs if you try to parse larger files
    $fileSize = 40

    $pathToEtl = Join-Path $logFilesDir "PktMon.etl"
    $startCommand = "pktmon start --etw --file-size $($fileSize) --file-name '$($pathToEtl)'"

    # Add parse start arguments
    $startCommand = $startCommand + ' ' + $parsedStartArguments

    Invoke-Expression -Command $startCommand.Trim()
}

###############################################################################
# Script execution starts here
###############################################################################
if (-not ($env:pester)) {
    # Reset any previous state of the tool
    Invoke-Expression -Command "pktmon unload"

    Add-Filters -Filters $filters

    $parsedStartArguments = Parse-StartArguments -Arguments $startArgs
    Start-PktMon $parsedStartArguments $logFilesDir
}

}
## [END] Start-WACPMCapture ##
function Get-WACPMCimWin32LogicalDisk {
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
## [END] Get-WACPMCimWin32LogicalDisk ##
function Get-WACPMCimWin32NetworkAdapter {
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
## [END] Get-WACPMCimWin32NetworkAdapter ##
function Get-WACPMCimWin32PhysicalMemory {
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
## [END] Get-WACPMCimWin32PhysicalMemory ##
function Get-WACPMCimWin32Processor {
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
## [END] Get-WACPMCimWin32Processor ##
function Get-WACPMClusterInventory {
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
## [END] Get-WACPMClusterInventory ##
function Get-WACPMClusterNodes {
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
## [END] Get-WACPMClusterNodes ##
function Get-WACPMDecryptedDataFromNode {
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
## [END] Get-WACPMDecryptedDataFromNode ##
function Get-WACPMEncryptionJWKOnNode {
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
## [END] Get-WACPMEncryptionJWKOnNode ##
function Get-WACPMServerInventory {
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
## [END] Get-WACPMServerInventory ##

# SIG # Begin signature block
# MIInvwYJKoZIhvcNAQcCoIInsDCCJ6wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC6qk+In3vap2du
# mZJUldhXCMPuhIFH4CpUmKIeRgk2EKCCDXYwggX0MIID3KADAgECAhMzAAADTrU8
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGZ8wghmbAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIOQkriWWzTjAE72V8cthRsKW
# wYIaNrfxjMaNpCs3lITWMEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAE9Ujtmoqv8Ff0YwM6Fmdq0aBmg0udFjToH9fqQGk5AtC5zV6ics1PXiW
# 8lVgcb0fyaI2gCvbldOFRgs+/F/POzFNFK5N6oeVQ6eVPV+nKyzVXx3B4NOukwFD
# sHdesQ0t+pEy/rVqIwKMVe38leAo6O/l0ZqPLNGMIpM7s2Iv+1dp6i634Y6xBbMM
# banS7PPh7Ev1sNUhcj5WqH2VOoWPKUF79n4tiLeCiOlHawOsEIbm4KyNnK+uCMKt
# zVoJYQvoy0czq/v6XPgiLznGQyjgwl6KiVtZnU0D0HrU/KF+XYGv7U4wkPLcZ9s2
# xR8/guGPBPyjV3mTC2g89gkUCqoTxaGCFykwghclBgorBgEEAYI3AwMBMYIXFTCC
# FxEGCSqGSIb3DQEHAqCCFwIwghb+AgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFZBgsq
# hkiG9w0BCRABBKCCAUgEggFEMIIBQAIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCBQLSfjLqx/iOdUfCwWQ/n7WzA/z+dLCndaQDOX/eiw2gIGZV4e871a
# GBMyMDIzMTIwNjIzNDkzMy43MzhaMASAAgH0oIHYpIHVMIHSMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJl
# bGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNO
# OjE3OUUtNEJCMC04MjQ2MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloIIReDCCBycwggUPoAMCAQICEzMAAAHg1PwfExUffl0AAQAAAeAwDQYJ
# KoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjMx
# MDEyMTkwNzE5WhcNMjUwMTEwMTkwNzE5WjCB0jELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3Bl
# cmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjoxNzlFLTRC
# QjAtODI0NjElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKyHnPOhxbvRATnGjb/6fuBh
# h3ZLzotAxAgdLaZ/zkRFUdeSKzyNt3tqorMK7GDvcXdKs+qIMUbvenlH+w53ssPa
# 6rYP760ZuFrABrfserf0kFayNXVzwT7jarJOEjnFMBp+yi+uwQ2TnJuxczceG5FD
# HrII6sF6F879lP6ydY0BBZkZ9t39e/svNRieA5gUnv/YcM/bIMY/QYmd9F0B+ebF
# Yi+PH4AkXahNkFgK85OIaRrDGvhnxOa/5zGL7Oiii7+J9/QHkdJGlfnRfbQ3QXM/
# 5/umBOKG4JoFY1niZ5RVH5PT0+uCjwcqhTbnvUtfK+N+yB2b9rEZvp2Tv4ZwYzEd
# 9A9VsYMuZiCSbaFMk77LwVbklpnw4aHWJXJkEYmJvxRbcThE8FQyOoVkSuKc5OWZ
# 2+WM/j50oblA0tCU53AauvUOZRoQBh89nHK+m5pOXKXdYMJ+ceuLYF8h5y/cXLQM
# OmqLJz5l7MLqGwU0zHV+MEO8L1Fo2zEEQ4iL4BX8YknKXonHGQacSCaLZot2kyJV
# RsFSxn0PlPvHVp0YdsCMzdeiw9jAZ7K9s1WxsZGEBrK/obipX6uxjEpyUA9mbVPl
# jlb3R4MWI0E2xI/NM6F4Ac8Ceax3YWLT+aWCZeqiIMLxyyWZg+i1KY8ZEzMeNTKC
# EI5wF1wxqr6T1/MQo+8tAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUcF4XP26dV+8S
# usoA1XXQ2TDSmdIwHwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYD
# VR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# cmwvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwG
# CCsGAQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIw
# MjAxMCgxKS5jcnQwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcD
# CDAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAMATzg6R/A0ldO7M
# qGxD1VJji5yVA1hHb0Hc0Yjtv7WkxQ8iwfflulX5Us64tD3+3NT1JkphWzaAWf2w
# KdAw35RxtQG1iON3HEZ0X23nde4Kg/Wfbx5rEHkZ9bzKnR/2N5A16+w/1pbwJzdf
# RcnJT3cLyawr/kYjMWd63OP0Glq70ua4WUE/Po5pU7rQRbWEoQozY24hAqOcwuRc
# m6Cb0JBeTOCeRBntEKgjKep4pRaQt7b9vusT97WeJcfaVosmmPtsZsawgnpIjbBa
# 55tHfuk0vDkZtbIXjU4mr5dns9dnanBdBS2PY3N3hIfCPEOszquwHLkfkFZ/9bxw
# 8/eRJldtoukHo16afE/AqP/smmGJh5ZR0pmgW6QcX+61rdi5kDJTzCFaoMyYzUS0
# SEbyrDZ/p2KOuKAYNngljiOlllct0uJVz2agfczGjjsKi2AS1WaXvOhgZNmGw42S
# FB1qaloa8Kaux9Q2HHLE8gee/5rgOnx9zSbfVUc7IcRNodq6R7v+Rz+P6XKtOgyC
# qW/+rhPmp/n7Fq2BGTRkcy//hmS32p6qyglr2K4OoJDJXxFs6lwc8D86qlUeGjUy
# o7hVy5VvyA+y0mGnEAuA85tsOcUPlzwWF5sv+B5fz35OW3X4Spk5SiNulnLFRPM5
# XCsSHqvcbC8R3qwj2w1evPhZxDuNMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJ
# mQAAAAAAFTANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNh
# dGUgQXV0aG9yaXR5IDIwMTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMwMTgzMjI1
# WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBAOThpkzntHIhC3miy9ckeb0O1YLT/e6cBwfSqWxOdcjK
# NVf2AX9sSuDivbk+F2Az/1xPx2b3lVNxWuJ+Slr+uDZnhUYjDLWNE893MsAQGOhg
# fWpSg0S3po5GawcU88V29YZQ3MFEyHFcUTE3oAo4bo3t1w/YJlN8OWECesSq/XJp
# rx2rrPY2vjUmZNqYO7oaezOtgFt+jBAcnVL+tuhiJdxqD89d9P6OU8/W7IVWTe/d
# vI2k45GPsjksUZzpcGkNyjYtcI4xyDUoveO0hyTD4MmPfrVUj9z6BVWYbWg7mka9
# 7aSueik3rMvrg0XnRm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75xqRdbZ2De+JKR
# Hh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9fvzZnkXftnIv231fgLrbqn427DZM9itu
# qBJR6L8FA6PRc6ZNN3SUHDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1DTsEzOUyO
# ArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7XKHYC4jMYctenIPDC+hIK12NvDMk2ZItb
# oKaDIV1fMHSRlJTYuVD5C4lh8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqvUAV6
# bMURHXLvjflSxIUXk8A8FdsaN8cIFRg/eKtFtvUeh17aj54WcmnGrnu3tz5q4i6t
# AgMBAAGjggHdMIIB2TASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsGAQQBgjcVAgQW
# BBQqp1L+ZMSavoKRPEY1Kc8Q/y8E7jAdBgNVHQ4EFgQUn6cVXQBeYl2D9OXSZacb
# UzUZ6XIwXAYDVR0gBFUwUzBRBgwrBgEEAYI3TIN9AQEwQTA/BggrBgEFBQcCARYz
# aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9zaXRvcnku
# aHRtMBMGA1UdJQQMMAoGCCsGAQUFBwMIMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIA
# QwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFNX2
# VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwu
# bWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEw
# LTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYt
# MjMuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCdVX38Kq3hLB9nATEkW+Geckv8qW/q
# XBS2Pk5HZHixBpOXPTEztTnXwnE2P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7pZmc6
# U03dmLq2HnjYNi6cqYJWAAOwBb6J6Gngugnue99qb74py27YP0h1AdkY3m2CDPVt
# I1TkeFN1JFe53Z/zjj3G82jfZfakVqr3lbYoVSfQJL1AoL8ZthISEV09J+BAljis
# 9/kpicO8F7BUhUKz/AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNhcy4sa3tuPywJeBTp
# kbKpW99Jo3QMvOyRgNI95ko+ZjtPu4b6MhrZlvSP9pEB9s7GdP32THJvEKt1MMU0
# sHrYUP4KWN1APMdUbZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMBV0lUZNlz138e
# W0QBjloZkWsNn6Qo3GcZKCS6OEuabvshVGtqRRFHqfG3rsjoiV5PndLQTHa1V1QJ
# sWkBRH58oWFsc/4Ku+xBZj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8vwLBgqJ7
# Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrpNPgkNWcr4A245oyZ1uEi6vAnQj0llOZ0
# dFtq0Z4+7X6gMTN9vMvpe784cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFxBmoQ
# tB1VM1izoXBm8qGCAtQwggI9AgEBMIIBAKGB2KSB1TCB0jELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxh
# bmQgT3BlcmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjox
# NzlFLTRCQjAtODI0NjElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vy
# dmljZaIjCgEBMAcGBSsOAwIaAxUAbfPR1fBX6HxYfyPx8zYzJU5fIQyggYMwgYCk
# fjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIF
# AOkbEhgwIhgPMjAyMzEyMDYyMzMwMzJaGA8yMDIzMTIwNzIzMzAzMlowdDA6Bgor
# BgEEAYRZCgQBMSwwKjAKAgUA6RsSGAIBADAHAgEAAgIXfDAHAgEAAgIRQjAKAgUA
# 6RxjmAIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAID
# B6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBABTnTo5eNtnpzJC+W7Ch
# xS6L4UyimX7WrZXW2eEDEgKTZpTXFjLS47RIstWph87RenrwyKOwKSUOyb5xkSJv
# 2Jc3oQFsy/BMQX4/YNXzZgesCMBSP/gmK7WuMH/mXoqBel+M6EVQ8q6DZNNbrQlT
# s5HkfB0o6pIDd7LYHNRr32xQMYIEDTCCBAkCAQEwgZMwfDELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBIDIwMTACEzMAAAHg1PwfExUffl0AAQAAAeAwDQYJYIZIAWUDBAIB
# BQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQx
# IgQga8zQNBxBTVvb77DuTFU/c9NuNabf/Zio0Aig0//wcXswgfoGCyqGSIb3DQEJ
# EAIvMYHqMIHnMIHkMIG9BCDj7lK/8jnlbTjPvc77DCCSb4TZApY9nJm5whsK/2kK
# wTCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB4NT8
# HxMVH35dAAEAAAHgMCIEIHr0oHOg0FxtjunDtngABWsbU9jTti/pxC1hTOeMf4dt
# MA0GCSqGSIb3DQEBCwUABIICAERcR0cyRnhyz4JLg05H9S9UaGqFy/c8G6CPI5jb
# JaMDXIlT5D/GBNyOBlw+hEkilzQ/FfjfTKS6oct9ESIBdWj/t7z2nPeL0Kni44L8
# CffN854USqZDQ7rzI7GgxKe6gN1NTq1x5esXQCaYHz7Sbv2niIrMGWPeLfp/7S7U
# gWUB7dodHqocNKMfQ6PdBPZfraMystUZTFLKbGW8U6tIjv1J79Aed8K3ltHigHDS
# ZMupaAnMg7/1PI0namjVDiQobuqRAnybQR7yqV1nAejVWW7f+Duq2qfOHKB7ZoGd
# LhyZhhrmgpcln3YHklWYA4ASzmEyHJUZrdskIcm7u3pardCzA+u4PQXP9l8Or3DG
# //ZE2f38VxmwIbvEdNYM36cPCs7f78jT8ZIMteN8CHlZDoF6/1cMHsZgvgd+W55K
# cDghYE7Z9yHx6XCAF+LbDe9KiYMHHvorB4UhYi+asqISrokh7cO7ppS+tNp7WpKr
# u926zvpolTo1vS3dsCRqxal+gh12Iq/xIRow5gX1CE7PMoWWnndHfGEzJevuJ6K6
# ZbYvT3fnDqVUIzlYHyuvZ8KetbHqLpQVWI1lQBkJWwq15EgOKa7ZlK1/sCMtEPiv
# BQeHKnBde64PV5S7oPDlud8DkISQKMnUhu0i22m+5vAMwxwQndHnZMObpn5Es6gQ
# hyCl
# SIG # End signature block
