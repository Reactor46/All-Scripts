function Add-WACSRVolumes {
<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER sourceAddVolumePartnership

.PARAMETER destinationAddVolumePartnership

.PARAMETER sourceRGName

.PARAMETER sourceComputerName

.PARAMETER destinationRGName

.PARAMETER destinationComputerName

.ROLE
Administrators

#>
Param
(
  [Parameter(Mandatory=$true)]
  [array] $sourceAddVolumePartnership,
  [Parameter(Mandatory=$true)]
  [array] $destinationAddVolumePartnership,
  [Parameter(Mandatory=$true)]
  [String] $sourceRGName,
  [Parameter(Mandatory=$true)]
  [String] $sourceComputerName,
  [Parameter(Mandatory=$true)]
  [String] $destinationRGName,
  [Parameter(Mandatory=$true)]
  [String] $destinationComputerName
)
Import-Module CimCmdlets
Invoke-CimMethod -Namespace 'root/Microsoft/Windows/StorageReplica' -ClassName 'MSFT_WvrAdminTasks' -MethodName 'SetPartnershipAddVolumes' -Arguments @{'SourceComputerName'= $sourceComputerName; 'SourceRGName'= $sourceRGName;  'DestinationComputerName'= $destinationComputerName; 'DestinationRGName'= $destinationRGName;  'SourceAddVolumePartnership'=$sourceAddVolumePartnership; 'DestinationAddVolumePartnership' = $destinationAddVolumePartnership}

}
## [END] Add-WACSRVolumes ##
function Dismount-WACSRSRDestination {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
Param
(
    [Parameter(Mandatory=$true)]
    [String] $resourceGroupName
)

Import-Module CimCmdlets, Microsoft.PowerShell.Utility
Dismount-SRDestination -Name $resourceGroupName -Force
}
## [END] Dismount-WACSRSRDestination ##
function Edit-WACSRSRParnership {

<#

.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
Param
(
  [Parameter(Mandatory=$true)]
  [Uint32] $asyncRPO,
  [Parameter(Mandatory=$true)]
  [bool] $encryption,
  [Parameter(Mandatory=$true)]
  [Uint32] $replicationMode,
  [Parameter(Mandatory=$true)]
  [Uint64] $logSizeInBytes,
  [Parameter(Mandatory=$true)]
  [String] $destinationComputerName,
  [Parameter(Mandatory=$true)]
  [String] $destinationRGName,
  [Parameter(Mandatory=$true)]
  [String] $sourceComputerName,
  [Parameter(Mandatory=$true)]
  [String] $sourceRGName,
  [Parameter(Mandatory = $true)]
  [bool] $compression,
  [Parameter(Mandatory = $true)]
  [bool] $osVersion23H2OrLater
)
Import-Module CimCmdlets
$customArgs =  @{
  'SourceComputerName'= $sourceComputerName;
  'SourceRGName' = $sourceRGName;
  'DestinationComputerName'= $destinationComputerName;
  'DestinationRGName' = $destinationRGName;
  'Encryption' = $encryption;
  'LogSizeInBytes' = $logSizeInBytes;
  'ReplicationMode' = $replicationMode;
  # no asyncRPO property unless we need it
};

if ($asyncRPO -gt 0)
{
    $customArgs.AsyncRPO = $asyncRPO
}

if ($osVersion23H2OrLater) {
    $customArgs.Compression = $compression
} 

Invoke-CimMethod -Namespace 'root/Microsoft/Windows/StorageReplica' -ClassName 'MSFT_WvrAdminTasks' -MethodName 'SetPartnershipModifyPartnership' -Arguments $customArgs | Out-Null
}
## [END] Edit-WACSRSRParnership ##
function Get-WACSRCounterData {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Readers

#>
 Param
 (
     [Parameter(Mandatory=$true)]
     [uint16] $counterType,
     [Parameter(Mandatory=$true)]
     [String] $partitionId
 )
Import-Module CimCmdlets, Microsoft.PowerShell.Utility

$result = Invoke-CimMethod -Namespace 'root/Microsoft/Windows/StorageReplica' -ClassName 'MSFT_WvrAdminTasks' -MethodName 'QueryCounterData' -Arguments @{'partitionId'= $partitionId; 'counterType'= $counterType}

Write-Output $result.itemValue

}
## [END] Get-WACSRCounterData ##
function Get-WACSRDisk {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Readers

#>
Import-Module Storage
Get-Disk

}
## [END] Get-WACSRDisk ##
function Get-WACSRFileSystemRoot {
<#
.SYNOPSIS
    Name: Get-FileSystemRoot
    Description: Gets the local file system root entities of the machine.
.DESCRIPTION

.ROLE
Administrators

.Returns
    The local file system root entities.
#>
param(
    [Parameter(Mandatory = $true)]
    [bool]
    $osVersion23H2OrLater,

    [Parameter(Mandatory = $false)]
    [bool]
    $getVolumesOnlyFromAvailableStorage
)
function Get-FileSystemRoot
{
    Import-Module  Storage, Microsoft.PowerShell.Utility

    $localVolumes = Get-LocalVolumes;

    return $localVolumes | % {
        $disk = $_ | Get-Partition | Get-Disk

        $caption = $null;
        $displayName = $null;

        if ([string]::IsNullOrWhiteSpace($_.DriveLetter))
        {
          $caption = $_.Path
        } else
        {
          $caption = $_.DriveLetter + ':\'
        }

        if ([string]::IsNullOrWhiteSpace($_.FileSystemLabel))
        {
          $displayName = $caption
        }
        else
        {
          $displayName =  $_.FileSystemLabel + ' (' + $caption + ')' # e.g. MyVolumeLabel (F:\) or MyVolumeLabel (\\?\VOLUME{EB824AA2-6E0A-4D29-BEB8-56112CAD3B5C}\)
        }

        $_ | Microsoft.PowerShell.Utility\Select-Object @{ Name = 'Caption'; Expression = { $caption } },
            @{ Name = 'DisplayName'; Expression = { $displayName } },
            @{ Name = 'Size'; Expression = { $_.Size} },
            @{ Name = 'SizeRemaining'; Expression = { $_.SizeRemaining } },
            @{ Name = 'DiskLogicalSectorSize'; Expression = { $disk.LogicalSectorSize } },
            @{ Name = 'DiskPhysicalSectorSize'; Expression = { $disk.PhysicalSectorSize } },
            @{ Name = 'DriveLetter'; Expression = { $_.DriveLetter } },
            @{ Name = 'FileSystem'; Expression = { $_.FileSystemType  } },
            @{ Name = 'isCSV'; Expression = { $false } },
            @{ Name = 'clusterFQDN'; Expression = { $null } },
            @{ Name = 'FileSystemNumber'; Expression = { $_.psBase.CimInstanceProperties["FileSystemType"].Value } }
        }
}


<#
.Synopsis
    Name: Get-LocalVolumes
    Description: Gets the local volumes of the machine.

.Returns
    The local volumes.
#>
function Get-LocalVolumes
{
    Remove-Module Storage -ErrorAction Ignore; # Remove the Storage module to prevent it from automatically localizing

    Import-Module Storage, StorageReplica

    $partitionsInReplication = @{}
    $localVolumes = @()
    $subsystem = Get-StorageSubSystem -FriendlyName Win*

    #
    # Lets find all partitions that are participating in replication
    # either as log / data
    #
    $allGroups = Get-SRGroup

    foreach ($group in $allGroups)
    {
        #
        # We will assume that SRGroup.LogVolume is either
        # a) drive letter in the 'C:\' format
        # b) OR a volume guid
        #
        $logVolume = $null;

        #  first try getting volume by drive letter
        $logVolumeDriveLetter = $group.LogVolume.Substring(0,1)
        $logVolume = Get-Volume -DriveLetter $logVolumeDriveLetter -ErrorAction SilentlyContinue

        # fall back to path if cannot find by drive letter
        if ($logVolume -eq $null)
        {
            $logVolume = Get-Volume -Path $group.LogVolume -ErrorAction SilentlyContinue
        }

        if ($logVolume -ne $null)
        {
            $partition = $logVolume | Get-Partition

            if ($partition -ne $null)
            {
                $partitionsInReplication.Add($partition.Guid, $true)
            }
        }

        #
        # Now just add all data partitions
        #
        foreach ($dp in $group.Partitions)
        {
            $partitionsInReplication.Add(('{' + $dp + '}'), $true)
        }
    }

    #
    # Now lets get the rest of the volumes
    #

    $disks = $subsystem | Get-Disk | Where-Object { ($_.IsSystem -eq $false) -and ($_.PartitionStyle -eq 'GPT') }
    foreach ($disk in $disks)
    {
        $partitions = $disk | Get-Partition

        foreach ($part in $partitions)
        {
            # Skip partitions that are not in replication already
            if ($part.Guid -ne $null -and -not $partitionsInReplication.ContainsKey($part.Guid))
            {
                $currentVolume = $part | Get-Volume

                if ($currentVolume -ne $null)
                {
                    $localVolumes += $currentVolume
                }
            }
        }
    }

    return Get-StandAloneServerVolumes -localVolumes $localVolumes -osVersion23H2OrLater $osVersion23H2OrLater
}

<#
.Synopsis
    Name: Get-StandAloneServerVolumes
    Description: Gets the required volumes to return for a standalone server.

.Returns
    The local volumes.
#>
function Get-StandAloneServerVolumes {
    Param (
        [Parameter(Mandatory=$true)]
        [System.Array] $localVolumes,

        [Parameter(Mandatory=$true)]
        [bool] $osVersion23H2OrLater
    )

    # Seperate the volumes to ones with size > 0
    # If there volumes with size 0 then check for type, keep ones wih "Basic" type
    $filteredLocalVolumes = @()

    foreach($volume in $localVolumes) {
        if ($volume.Size -gt 0) {
            $filteredLocalVolumes += $volume
        } elseif ($volume | Get-Partition | Where-Object { $_.type -eq "Basic" }) {
            $filteredLocalVolumes += $volume
        } 
    }

    if ($osVersion23H2OrLater) {
        return $filteredLocalVolumes
    } else {
        # we now have logic to allow for drive paths if no drive letter so let's show them all as long as they have some size
        return $localVolumes | Where-Object { $_.Size -gt 0 }
    }
}

<#
.Synopsis
    Name: Get-ClusterVolumes
    Description: Gets the volumes on the cluster primary or secondary.

.Returns
    The cluster volumes.
#>
function Get-ClusterVolumes {
    Param
    (
        [Parameter(Mandatory=$true)]
        [System.Object] $cluster
    )

    Remove-Module Storage -ErrorAction Ignore; # Remove the Storage module to prevent it from automatically localizing

    Import-Module FailoverClusters, Storage, Microsoft.PowerShell.Utility

    $clusterVolumes            = @();
    $s2dEnabled         = (($cluster).S2DEnabled -eq 1)
    $node               = $env:COMPUTERNAME
    $available_storage  = Get-ClusterGroup | Where-Object { $_.GroupType -eq 'AvailableStorage' }

    $result = Move-ClusterGroup -InputObject $available_storage -Node $node

    if ($result -ne $null) {
        $availableResources = @()

        #
        # Find all groups, that are online on this node
        #
        $groups = Get-ClusterGroup | Where-Object { $_.OwnerNode -eq $env:COMPUTERNAME -and ( $_.State -eq 'Online' -or $_.State -eq 'PartialOnline' ) }

        #
        # Filter them
        # a) remove resources which are not online
        # b) remove groups which contain SR resources
        #
        if ($getVolumesOnlyFromAvailableStorage) {
                
            $availableResources += $available_storage | Get-ClusterResource | Where-Object { $_.ResourceType -eq 'Physical Disk' -and $_.State -eq 'Online' }
            
        } else {
            foreach ($group in $groups) {
                $srResources = $group | Get-ClusterResource | ? ResourceType -eq 'Storage Replica'
                $isReplicationParticipant = (($srResources | Microsoft.PowerShell.Utility\Measure-Object).Count -ne 0)

                if ($isReplicationParticipant -eq $false) {

                    $diskResources = $group | Get-ClusterResource | Where-Object { $_.ResourceType -eq 'Physical Disk' -and $_.State -eq 'Online' }

                    foreach ($disk in $diskResources) {
                        $availableResources += $disk
                    }
                }
            }
        }

        # Find disks which are physically connected to this node
        $physicallyConnectedDisks = Get-PhysicalDiskSNV | ? IsPhysicallyConnected -eq $true

        # Find all CSVs
        $csvs = Get-ClusterSharedVolume

        #
        # filter them
        # a) remove groups which contain SR resources
        # b) remove CSVs that are not physically connected to this node
        #
        foreach ($csv in $csvs) {

            $csvGroup                 = $csv | Get-ClusterGroup
            $srResources              = $csvGroup | Get-ClusterResource | ? ResourceType -eq 'Storage Replica'
            $isReplicationParticipant = (($srResources | Microsoft.PowerShell.Utility\Measure-Object).Count -ne 0)

            if ($isReplicationParticipant -eq $false) {

                $diskIdGuid = ($csv | Get-ClusterParameter -Name 'DiskIdGuid').Value
                $msftDisk   = Get-Disk | ? Guid -eq $diskIdGuid

                if ($msftDisk.PartitionStyle -eq 'GPT') {
                    if ($s2dEnabled -eq $true) {
                        $availableResources += $csv
                    }
                    else {
                        $pd = Get-PhysicalDisk | ? UniqueId -eq $msftDisk.UniqueId
                        if ($pd -ne $null) {
                            if (($physicallyConnectedDisks).PhysicalDisk | ? ObjectId -eq $pd.ObjectId) {
                                $availableResources += $csv
                            }
                        }
                    }
                }
            }
        }


        $csvType = [Microsoft.FailoverClusters.PowerShell.ClusterSharedVolume].FullName

        foreach ($resource in $availableResources) {
            $paramGuid = $resource | Get-ClusterParameter -Name DiskIdGuid
            $disk = Get-Disk | Where-Object { $_.Guid -eq $paramGuid.Value }

            if ($disk.PartitionStyle -eq 'GPT') {
                $volume = $disk | Get-Partition | Get-Volume

                $caption = $null;
                $displayName = $null;
                $isCSV = $false;

                if ([string]::IsNullOrWhiteSpace($volume.DriveLetter))
                {
                  $caption = $volume.Path
                } else
                {
                  $caption = $volume.DriveLetter + ':\'
                }

                if ([string]::IsNullOrWhiteSpace($volume.FileSystemLabel))
                {
                  $displayName = $caption
                }
                else
                {
                  $displayName = $volume.FileSystemLabel + ' (' + $caption + ')' # e.g. MyVolumeLabel (F:\) or MyVolumeLabel (\\?\VOLUME{EB824AA2-6E0A-4D29-BEB8-56112CAD3B5C}\)
                }

                if ($resource.GetType().FullName -eq $csvType) {
                    $caption = $resource.SharedVolumeInfo.FriendlyVolumeName
                    $displayName = $resource.SharedVolumeInfo.FriendlyVolumeName
                    $isCSV = $true;
                }

                $volumeObject = $volume | Microsoft.PowerShell.Utility\Select-Object @{ Name = 'Caption'; Expression = { $caption } },
                    @{ Name = 'DisplayName'; Expression = { $displayName } },
                    @{ Name = 'Size'; Expression = { $_.Size } },
                    @{ Name = 'SizeRemaining'; Expression = { $_.SizeRemaining } },
                    @{ Name = 'DiskLogicalSectorSize'; Expression = { $disk.LogicalSectorSize } },
                    @{ Name = 'DiskPhysicalSectorSize'; Expression = { $disk.PhysicalSectorSize } },
                    @{ Name = 'DriveLetter'; Expression = { $_.DriveLetter } },
                    @{ Name = 'FileSystem'; Expression = { $_.FileSystemType  } },
                    @{ Name = 'isCSV'; Expression = { $isCSV } },
                    @{ Name = 'clusterFQDN'; Expression = { $cluster.Name + '.' + $cluster.Domain } },
                    @{ Name = 'FileSystemNumber'; Expression = { $_.psBase.CimInstanceProperties["FileSystemType"].Value } }

                $clusterVolumes += $volumeObject
            }
        }
    }

    if ($clusterVolumes.Count -gt 0) {
        return Get-FilteredClusterSharedVolumes -clusterVolumes $clusterVolumes -osVersion23H2OrLater $osVersion23H2OrLater
    } else {
        return $clusterVolumes
    }
}


<#
.Synopsis
    Name: Get-FilteredClusterSharedVolumes
    Description: Gets the required volumes to return for a CSV.

.Returns
    The cluster volumes.
#>
function Get-FilteredClusterSharedVolumes {
    Param (
        [Parameter(Mandatory=$true)]
        [System.Array] $clusterVolumes,

        [Parameter(Mandatory=$true)]
        [bool] $osVersion23H2OrLater
    )

    if ($osVersion23H2OrLater) {
        return $clusterVolumes
    } 

    $filteredClusterVolumes = @()

    foreach($volumeObject in $clusterVolumes) {
        if ($volumeObject.FileSystemNumber -ne 0) {
            $filteredClusterVolumes += $volumeObject
        }
    }

    return $filteredClusterVolumes;
}

$cluster = $null
try {
  Import-Module FailoverClusters -ErrorAction SilentlyContinue
  $cluster = Get-Cluster -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
}
catch {
  # swallow the err if Get-Cluster cmdlet is not recognized because FailoverClustering is not installed
}

if ($cluster -ne $null) {
  Get-ClusterVolumes -Cluster $cluster
} else {
  Get-FileSystemRoot
}

}
## [END] Get-WACSRFileSystemRoot ##
function Get-WACSRNodeClusteredState {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Readers

#>
$result = @{
    isClustered = $false;
    isStretch = $false;
}

try {
    Import-Module Microsoft.PowerShell.Utility
    Import-Module FailoverClusters -ErrorAction Stop
    $cluster = Get-Cluster -ErrorAction SilentlyContinue
    if ($cluster)
    {
        $result.isClustered = $true
        try
        {
            $sites = Get-ClusterFaultDomain -Type Site -ErrorAction SilentlyContinue

            # a more thorough check would be to check each child and each child's child of each site and determine if the site contains nodes
            if (($sites | Microsoft.PowerShell.Utility\Measure-Object).count -gt 1)
            {
                $result.isStretch = $true
            }
        }
        catch
        {
            # swallow the error - no sites
            Write-Output $_ | Out-Null
        }
    }

}
catch {
    # swallow the error - it is not a cluster because get-cluster failed
    Write-Output $_ | Out-Null
}

Write-Output $result

}
## [END] Get-WACSRNodeClusteredState ##
function Get-WACSRNodeFqdnsAndState {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Readers

#>
Import-Module -Name FailoverClusters -ErrorAction SilentlyContinue
Import-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue

$nodes = @()


Get-Clusternode | ForEach-Object {
  $fqdn = $null
  $err = $null

  try {
      #  there could be a DNS lookup issue
      $fqdn = [System.Net.Dns]::GetHostEntry($_.Name).HostName;
  }
  catch {
      $err = $_
  }

    $nodes += @{
        name = $_.Name
        fqdn = $fqdn
        state = $_.State.value__;
        error = $err
    }

}

$nodes

}
## [END] Get-WACSRNodeFqdnsAndState ##
function Get-WACSROSBuild {
<#
.SYNOPSIS
Gets OS Build and UBR

.DESCRIPTION
Gets OS Build and UBR

.ROLE
Readers

#>
Import-Module  Microsoft.PowerShell.Management
$item = Get-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion"

Write-Output @{
"buildNumber" = $item.CurrentBuildNumber;
"ubr" = $item.UBR;
}

}
## [END] Get-WACSROSBuild ##
function Get-WACSRPartitionInfo {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Readers

#>
Param
(
    [Parameter(Mandatory=$true)]
    [String] $PartitionId,
    [Parameter(Mandatory=$true)]
    [String] $ReplicationGroupName
)
Import-Module CimCmdlets, Storage, Microsoft.PowerShell.Utility
$result = Invoke-CimMethod -Namespace 'root/Microsoft/Windows/StorageReplica' -ClassName 'MSFT_WvrAdminTasks' -MethodName 'QueryPartitionInfo' -Arguments @{'partitionId'= $PartitionId; 'replicationGroupName'= $ReplicationGroupName}
$partitionIdArgument = '{' + $PartitionId + '}'
$partitionResult = Get-Partition | ? Guid -eq $partitionIdArgument | Get-Volume
if ($partitionResult -ne $null)
{
    $result | Add-Member -NotePropertyName PartitionFreeSpaceInBytes -NotePropertyValue $partitionResult.SizeRemaining
}
Write-Output $result

}
## [END] Get-WACSRPartitionInfo ##
function Get-WACSRSRGroup {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Readers

#>
Param
(
    [Parameter(Mandatory=$true)]
    [String] $name
)

# function Get-PartitionInfo {
#     Param (
#       [Parameter(Mandatory=$true)]
#       [String] $PartitionId,
#       [Parameter(Mandatory=$true)]
#       [String] $ReplicationGroupName,
#       [Parameter(Mandatory=$true)]
#       [String] $ComputerName
#     )

#     $result = Invoke-CimMethod -Namespace 'root/Microsoft/Windows/StorageReplica' -ClassName 'MSFT_WvrAdminTasks' -MethodName 'QueryPartitionInfo' -Arguments @{'partitionId'= $PartitionId; 'replicationGroupName'= $ReplicationGroupName} -ComputerName $ComputerName
#     $partitionIdArgument = '{' + $PartitionId + '}'
#     $partitionResult = Get-Partition | ? Guid -eq $partitionIdArgument | Get-Volume

#     if ($partitionResult -ne $null)
#     {
#         $result | Add-Member -NotePropertyName PartitionFreeSpaceInBytes -NotePropertyValue $partitionResult.SizeRemaining
#     }

#     Write-Output $result
# }


Import-Module StorageReplica, Microsoft.PowerShell.Utility
$group = Get-SRGroup -Name $name
$result = @{
  "group" = $group;
  "ownerNode" = $null;
}

# $computerName = $group.computerName;

# used to set group.computerName to be the owner node's name instead of cluster's name
if ($group.isCluster)
{
    Import-Module FailoverClusters -ErrorAction SilentlyContinue
    $replicationIds = Get-ClusterResource | Where-Object { $_.resourcetype -eq "storage replica" }  | Get-ClusterParameter  -Name "replicationGroupId"
    $resource = $replicationIds | Where-Object { $_.value -eq  ('{' + $group.id + '}') }

    if (($resource -ne $null) -and ($resource.ClusterObject -ne $null -and $resource.ClusterObject.OwnerNode -ne $null))
    {
      $result.ownerNode = $resource.ClusterObject.OwnerNode.Name
      # $computerName =  $resource.ClusterObject.OwnerNode.Name
    }
}

# foreach ($replica in $group.replicas)
# {
#     $partitionInfo = get-PartitionInfo -PartitionId $replica.partitionId -ReplicationGroupName $group.Name -ComputerName $computerName
# }
# #todo need to get this info merged with the group data somehow otherwise need seperate call as before....



Write-Output $result


}
## [END] Get-WACSRSRGroup ##
function Get-WACSRSRPartnership {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Readers

#>
Import-Module StorageReplica
Get-SRPartnership

}
## [END] Get-WACSRSRPartnership ##
function Get-WACSRSRServerFeature {
<#################################################################################################################################################
 # File: Get-SRFeature.ps1
 #
 # .DESCRIPTION
 #
 # Gets the Windows Feature Storage Replica and returns if it is installed.
 #
 #  The supported Operating Systems are Windows Server 2016.
 #
 #  Copyright (c) Microsoft Corp 2016.
 #
 #################################################################################################################################################>
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Readers

#>
Import-Module ServerManager
Get-WindowsFeature -Name Storage-Replica, RSAT-Storage-Replica
}
## [END] Get-WACSRSRServerFeature ##
function Install-WACSRSRFeature {
<#################################################################################################################################################
 # File: Install-SRFeature.ps1
 #
 # .DESCRIPTION
 #
 # Installs Storage Replica Feature and all management tools. 
 #
 #  The supported Operating Systems are Windows Server 2016.
 #
 #  Copyright (c) Microsoft Corp 2016.
 #
 #################################################################################################################################################>
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
Param
(
    [Boolean] $RestartMachine
)
Import-Module ServerManager
Add-WindowsFeature -Name Storage-Replica -IncludeManagementTools -Restart:$RestartMachine
}
## [END] Install-WACSRSRFeature ##
function Mount-WACSRSRDestination {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
Param
(
    [Parameter(Mandatory=$true)]
    [String] $resourceGroupName,
    [Parameter(Mandatory=$true)]
    [String] $temporaryPath
)

Import-Module CimCmdlets, Microsoft.PowerShell.Utility
Mount-SRDestination -Name $resourceGroupName -TemporaryPath $temporaryPath -Force

}
## [END] Mount-WACSRSRDestination ##
function New-WACSRSrPartnership {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
param (
  [Parameter(Mandatory = $true)]
  [string]
  $sourceComputerName,

  [Parameter(Mandatory = $true)]
  [string]
  $destinationComputerName,

  [Parameter(Mandatory = $true)]
  [string]
  $sourceRGName,

  [Parameter(Mandatory = $true)]
  [string]
  $destinationRGName,

  [Parameter(Mandatory = $true)]
  [array]
  $sourceVolumeName,

  [Parameter(Mandatory = $true)]
  [string]
  $sourceLogVolumeName,

  [Parameter(Mandatory = $true)]
  [array]
  $destinationVolumeName,

  [Parameter(Mandatory = $true)]
  [string]
  $destinationLogVolumeName,

  [Parameter(Mandatory=$true)]
  [bool]
  $enableEncryption,

  [Parameter(Mandatory=$true)]
  [bool]
  $enableConsistencyGroups,

  [Parameter(Mandatory=$true)]
  [Uint64]
  $logSizeInBytes,

  [Parameter(Mandatory = $true)]
  [bool]
  $seeded,

  [Parameter(Mandatory = $true)]
  [Uint32]
  $replicationMode,

  [Parameter(Mandatory = $true)]
  [uint32]
  $asyncRPO,

  [Parameter(Mandatory = $true)]
  [bool]
  $enableCompression,

  [Parameter(Mandatory = $true)]
  [int]
  $logType,

  [Parameter(Mandatory = $true)]
  [bool]
  $osVersion23H2OrLater
)
Import-Module StorageReplica
$customArgs = @{
    "SourceComputerName" =  $sourceComputerName;
    "DestinationComputerName" =  $destinationComputerName;
    "SourceRGName" =  $sourceRGName;
    "DestinationRGName" =  $destinationRGName;
    "ReplicationMode" =  $replicationMode;
    "SourceVolumeName" = $sourceVolumeName;
    "SourceLogVolumeName" = $sourceLogVolumeName;
    "DestinationVolumeName" = $destinationVolumeName;
    "DestinationLogVolumeName" = $destinationLogVolumeName;
    "LogSizeInBytes" = $logSizeInBytes;
}

if ($asyncRPO -gt 0)
{
    $customArgs.AsyncRPO = $asyncRPO
}

# Older OS versions do not support compression, so even if it's false we don't want to pass the flag in
if ($osVersion23H2OrLater) {
  New-SRPartnership @customArgs -Seeded:$seeded -EnableConsistencyGroups:$enableConsistencyGroups -EnableEncryption:$enableEncryption -LogType $logType -EnableCompression:$enableCompression -Force
} else {
  New-SRPartnership @customArgs -Seeded:$seeded -EnableConsistencyGroups:$enableConsistencyGroups -EnableEncryption:$enableEncryption -Force
}


}
## [END] New-WACSRSrPartnership ##
function Remove-WACSRSRGroup {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
Param
(
    [Parameter(Mandatory=$true)]
    [String] $computerName,
    [String] $rgName
)
Import-Module CimCmdlets

Invoke-CimMethod -Namespace 'root/Microsoft/Windows/StorageReplica' -ClassName 'MSFT_WvrAdminTasks' -MethodName 'RemoveSrGroup' -Arguments @{'ComputerName'= $computerName; 'Name'= $rgName;}

}
## [END] Remove-WACSRSRGroup ##
function Remove-WACSRSRPartnership {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
Param
(
    [Parameter(Mandatory=$true)]
    [String] $sourceComputerName,
    [Parameter(Mandatory=$true)]
    [String] $sourceRGName,
    [Parameter(Mandatory=$true)]
    [String] $destinationComputerName,
    [Parameter(Mandatory=$true)]
    [String] $destinationRGName
)
Import-Module CimCmdlets

Invoke-CimMethod -Namespace 'root/Microsoft/Windows/StorageReplica' -ClassName 'MSFT_WvrAdminTasks' -MethodName 'RemovePartnership' -Arguments @{'SourceComputerName'= $sourceComputerName; 'SourceRGName'= $sourceRGName; "DestinationComputerName"=$destinationComputerName; 'DestinationRGName' = $destinationRGName; IgnoreRemovalFailure = $true;}

}
## [END] Remove-WACSRSRPartnership ##
function Remove-WACSRVolumes {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
Param
(
  [Parameter(Mandatory=$true)]
  [array] $removeVolumeNames,
  [Parameter(Mandatory=$true)]
  [String] $sourceRGName,
  [Parameter(Mandatory=$true)]
  [String] $sourceComputerName
)

Import-Module CimCmdlets

Invoke-CimMethod -Namespace 'root/Microsoft/Windows/StorageReplica' -ClassName 'MSFT_WvrAdminTasks' -MethodName 'SetGroupRemoveVolumes' -Arguments @{'ComputerName'= $sourceComputerName; 'Name'= $sourceRGName; "RemoveVolumeName"=$removeVolumeNames;}

}
## [END] Remove-WACSRVolumes ##
function Set-WACSRSRPartnershipRoles {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
Param
(
    [Parameter(Mandatory=$true)]
    [String] $destinationComputerName,
    [Parameter(Mandatory=$true)]
    [String] $destinationRGName,
    [Parameter(Mandatory=$true)]
    [String] $newSourceComputerName,
    [Parameter(Mandatory=$true)]
    [String] $sourceRGName
)
Import-Module StorageReplica

Set-SRPartnership -NewSourceComputerName $newSourceComputerName -SourceRGName $sourceRGName -DestinationComputerName $destinationComputerName -DestinationRGName $destinationRGName -Force

}
## [END] Set-WACSRSRPartnershipRoles ##
function Suspend-WACSRSRGroup {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
Param
(
    [Parameter(Mandatory=$true)]
    [String] $name
)
Import-Module StorageReplica

Suspend-SRGroup -Name $name -Force

}
## [END] Suspend-WACSRSRGroup ##
function Sync-WACSRSRGroup {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Administrators

#>
Param
(
    [Parameter(Mandatory=$true)]
    [String] $name
)
Import-Module StorageReplica

Sync-SRGroup -Name $name -Force

}
## [END] Sync-WACSRSRGroup ##
function Test-WACSRConnection {
<#
.SYNOPSIS

.DESCRIPTION

.ROLE
Readers

#>
param (
		[Parameter(Mandatory = $true)]
		[String]
    $nodeName
)
Import-Module Microsoft.PowerShell.Management

Test-Connection -ComputerName $nodeName -Quiet

}
## [END] Test-WACSRConnection ##
function Clear-WACSREventLogChannel {
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
## [END] Clear-WACSREventLogChannel ##
function Clear-WACSREventLogChannelAfterExport {
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
## [END] Clear-WACSREventLogChannelAfterExport ##
function Export-WACSREventLogChannel {
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
## [END] Export-WACSREventLogChannel ##
function Get-WACSRCimEventLogRecords {
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
## [END] Get-WACSRCimEventLogRecords ##
function Get-WACSRCimWin32LogicalDisk {
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
## [END] Get-WACSRCimWin32LogicalDisk ##
function Get-WACSRCimWin32NetworkAdapter {
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
## [END] Get-WACSRCimWin32NetworkAdapter ##
function Get-WACSRCimWin32PhysicalMemory {
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
## [END] Get-WACSRCimWin32PhysicalMemory ##
function Get-WACSRCimWin32Processor {
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
## [END] Get-WACSRCimWin32Processor ##
function Get-WACSRClusterEvents {
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
## [END] Get-WACSRClusterEvents ##
function Get-WACSRClusterInventory {
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
## [END] Get-WACSRClusterInventory ##
function Get-WACSRClusterNodes {
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
## [END] Get-WACSRClusterNodes ##
function Get-WACSRDecryptedDataFromNode {
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
## [END] Get-WACSRDecryptedDataFromNode ##
function Get-WACSREncryptionJWKOnNode {
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
## [END] Get-WACSREncryptionJWKOnNode ##
function Get-WACSREventLogDisplayName {
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
## [END] Get-WACSREventLogDisplayName ##
function Get-WACSREventLogFilteredCount {
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
## [END] Get-WACSREventLogFilteredCount ##
function Get-WACSREventLogRecords {
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
## [END] Get-WACSREventLogRecords ##
function Get-WACSREventLogSummary {
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
## [END] Get-WACSREventLogSummary ##
function Get-WACSRServerInventory {
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
Get the Win32_OperatingSystem information

.DESCRIPTION
Get the Win32_OperatingSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller.

#>
function getOperatingSystemInfo() {
  return Get-CimInstance Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object csName, Caption, OperatingSystemSKU, Version, ProductType, OSType, LastBootUpTime
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
## [END] Get-WACSRServerInventory ##
function Set-WACSREventLogChannelStatus {
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
## [END] Set-WACSREventLogChannelStatus ##

# SIG # Begin signature block
# MIIoOQYJKoZIhvcNAQcCoIIoKjCCKCYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBq5SvpxFgthlwJ
# OVPauFjVoJI1HT/DZVLTmu8t8OyZU6CCDYUwggYDMIID66ADAgECAhMzAAADTU6R
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
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIO34
# YENL1lRp+OPuYwpFOUUKRpQnz07t9tZ5nn0X+hLDMEIGCisGAQQBgjcCAQwxNDAy
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20wDQYJKoZIhvcNAQEBBQAEggEAFkhlT2AaxRU0CTnYZDBn0KQaY5+/uGS9w6vk
# duEAG3sGfsqsmURavoqaNxG7kQx0w9NKfbuc+cb/73FuHgrjGDt+KU8NHiKHRqTe
# EtkBhFtFPw9nPnSx50ZH1ZDajovdDt8IWK2EEPeAtBtNebvDzI6sgZeTcPsZ1Ad/
# E8oNNnXWwq9IfbDbrDfEftBmZ9eOY9Lqo54EcdpnyZP7/3ZJ8WYcUexnikJfN4hp
# K8x9uVE55R4CsRiqJtI1sFXiFcFwbyCfCt6dPuP+qlJxQy9xzXgx0DTNEQYFK0o7
# LYI0sK/LQU0ISNSnjYAiAiwZYaG4w/qQZjqqsyx4VMtQMMmFjKGCF5QwgheQBgor
# BgEEAYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZI
# AWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGE
# WQoDATAxMA0GCWCGSAFlAwQCAQUABCCWI3ts2SKGYqP7PljJfDSmUKKtaAfA0M28
# HaBlpeaXdQIGZVbCsg72GBMyMDIzMTIwNTE4NDIzNS44NTlaMASAAgH0oIHRpIHO
# MIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQL
# ExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxk
# IFRTUyBFU046QTQwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAdYnaf9yLVbIrgAB
# AAAB1jANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MDAeFw0yMzA1MjUxOTEyMzRaFw0yNDAyMDExOTEyMzRaMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTQwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Uw
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDPLM2Om8r5u3fcbDEOXydJ
# tbkW5U34KFaftC+8QyNqplMIzSTC1ToE0zcweQCvPIfpYtyPB3jt6fPRprvhwCks
# Uw9p0OfmZzWPDvkt40BUStu813QlrloRdplLz2xpk29jIOZ4+rBbKaZkBPZ4R4LX
# QhkkHne0Y/Yh85ZqMMRaMWkBM6nUwV5aDiwXqdE9Jyl0i1aWYbCqzwBRdN7CTlAJ
# xrJ47ov3uf/lFg9hnVQcqQYgRrRFpDNFMOP0gwz5Nj6a24GZncFEGRmKwImL+5KW
# PnVgvadJSRp6ZqrYV3FmbBmZtsF0hSlVjLQO8nxelGV7TvqIISIsv2bQMgUBVEz8
# wHFyU3863gHj8BCbEpJzm75fLJsL3P66lJUNRN7CRsfNEbHdX/d6jopVOFwF7omm
# TQjpU37A/7YR0wJDTt6ZsXU+j/wYlo9b22t1qUthqjRs32oGf2TRTCoQWLhJe3cA
# IYRlla/gEKlbuDDsG3926y4EMHFxTjsjrcZEbDWwjB3wrp11Dyg1QKcDyLUs2anB
# olvQwJTN0mMOuXO8tBz20ng/+Xw+4w+W9PMkvW1faYi435VjKRZsHfxIPjIzZ0wf
# 4FibmVPJHZ+aTxGsVJPxydChvvGCf4fe8XfYY9P5lbn9ScKc4adTd44GCrBlJ/JO
# soA4OvNHY6W+XcKVcIIGWwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFGGaVDY7TQBi
# MCKg2+j/zRTcYsZOMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8G
# A1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBs
# BggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUy
# MDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQDUv+RjNidwJxSb
# Mk1IvS8LfxNM8VaVhpxR1SkW+FRY6AKkn2s3On29nGEVlatblIv1qVTKkrUxLYMZ
# 0z+RA6mmfXue2Y7/YBbzM5kUeUgU2y1Mmbin6xadT9DzECeE7E4+3k2DmZxuV+GL
# FYQsqkDbe8oy7+3BSiU29qyZAYT9vRDALPUC5ZwyoPkNfKbqjl3VgFTqIubEQr56
# M0YdMWlqCqq0yVln9mPbhHHzXHOjaQsurohHCf7VT8ct79po34Fd8XcsqmyhdKBy
# 1jdyknrik+F3vEl/90qaon5N8KTZoGtOFlaJFPnZ2DqQtb2WWkfuAoGWrGSA43My
# l7+PYbUsri/NrMvAd9Z+J9FlqsMwXQFxAB7ujJi4hP8BH8j6qkmy4uulU5SSQa6X
# kElcaKQYSpJcSjkjyTDIOpf6LZBTaFx6eeoqDZ0lURhiRqO+1yo8uXO89e6kgBeC
# 8t1WN5ITqXnjocYgDvyFpptsUDgnRUiI1M/Ql/O299VktMkIL72i6Qd4BBsrj3Z+
# iLEnKP9epUwosP1m3N2v9yhXQ1HiusJl63IfXIyfBJaWvQDgU3Jk4eIZSr/2KIj4
# ptXt496CRiHTi011kcwDpdjQLAQiCvoj1puyhfwVf2G5ZwBptIXivNRba34KkD5o
# qmEoF1yRFQ84iDsf/giyn/XIT7YY/zCCB3EwggVZoAMCAQICEzMAAAAVxedrngKb
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
# Y2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkE0MDAtMDVF
# MC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMK
# AQEwBwYFKw4DAhoDFQD5r3DVRpAGQo9sTLUHeBC87NpK+qCBgzCBgKR+MHwxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6RmjDTAi
# GA8yMDIzMTIwNTEzMjQyOVoYDzIwMjMxMjA2MTMyNDI5WjB0MDoGCisGAQQBhFkK
# BAExLDAqMAoCBQDpGaMNAgEAMAcCAQACAhmlMAcCAQACAhPKMAoCBQDpGvSNAgEA
# MDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAI
# AgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBAJ0aZc6G0YI6UPwGrq3z5PCmU/nF
# zY/5gDuYRjXRLCOyAUWR7B5/HzrarF+S2kEta2tZ+X1jWK6KiEg7M7XUisfKkSwp
# 5pZbD4KevWJokBR6dNj6ArAWGDgVtWe6M0clAO14aLKZCGMHszamBdRrt+44lgcY
# Hs8H+YK9X1JQAvLumdd8IVMSNqi2bmsPIZPd+V9q+EPHXKYy+2OKlmpMr3a/N7mj
# EqdlvCX6JKGoZ8mImSGoQSZ7Bii9qDTKqUIsiYM7UFoKFSPVAa/WRcyyHHKOrXFq
# 7oePvw5Hcj3Cnjzdh6hh0Mj6brv1tHqhP+vb8AP5w2YTmEnCrxC5ZS/w1C0xggQN
# MIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdYn
# af9yLVbIrgABAAAB1jANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
# CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCCl/0xJmfflnajvdsadRDQIUf/h
# vXVysURosEI18lVbJjCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EINbLTQ1X
# eNM+EBinOEJMjZd0jMNDur+AK+O8P12j5ST8MIGYMIGApH4wfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAHWJ2n/ci1WyK4AAQAAAdYwIgQgEpNLgKfY
# Yeh4xCS1NXPUVZ1aNKZVS32XZloGSfvv1zswDQYJKoZIhvcNAQELBQAEggIALZWs
# sTX3pzssd9Hn6wrD9SyOdRKRPPi23HpfmtIS/9haEwRVVq0L8rsFbszmjP28CGaf
# b3KYdtuqqI+RJw4SIf4/XKbNVo72uGuZn4bt2TtGhUAieyCz9O8Oj3ohxDDL3KiQ
# EtWg9ppif3/VZFTIfj6eg0p8LXfOHCIO4PqPaF2ZzXaNGZEHJTmFBPyG0I/oYRY4
# w2TkxhLgko+HkhMr7UBiosiw8lC20ikPhPoSrtUTiW3efUFingfKZgJDUjirKh2r
# zOgELtNsOx7jU+BrG7hiwqSX6sNu8vTSl5bSU/A7jJ6Jo1NGVhInBy62L0ZhGRqN
# iYdvd2rJlEkE9IYfagft4D7Zs0oTgsOz/Fkm8D448dFB/xUQIZIwnNB5V9MeMPs/
# l0AT4skzahJMirWYAAuRkiL5aBqhq/5lmD90kvzloBKTgYe3TxJyc2U9E7TOD1wQ
# jFRVAOCYeDxy0On+F/SUz8/TlMVa7N+uKOQVuviohh6WVOBCU8jZpR2CUAsSvPWm
# uvH5sp4sjuwfa+oRvhwPXJ7Jba7M5i9an9TBIyr8ihEdG38SD8yCw4sgpuvh5Abn
# ZrVax4eCsKcYdYLXZWTTC7cyPVwe2Wvczg4WEavv9L3jdCZg8hZGUoFDwsT3ZpiA
# 9KZ/r9+ivZo6CeDaz5ElOCGWmXMrMwJp20iIgMk=
# SIG # End signature block
