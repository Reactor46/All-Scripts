function Add-WACAKAzCliExtensions {
<#

.SYNOPSIS
Adds a list of Az Cli extensions

.DESCRIPTION
Adds a list of Az Cli extensions by name

.ROLE
Administrators

#>
Param(
    [parameter(Mandatory = $true)] [object[]] $extensions
)

$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine')

# If no account is available will send back null -> will then parse to UI that needs az login
$value = az account show

# Az account is logged in and valid
if ($null -ne $value) {
    foreach ($extension in $extensions) {
        if ($null -ne $extension) {
            az extension add --name $extension.Name --upgrade --only-show-errors
        }
    }
}
}
## [END] Add-WACAKAzCliExtensions ##
function Add-WACAKAzCliExtensionsExternal {
<#

.SYNOPSIS
Adds a list of Az Cli extensions

.DESCRIPTION
Adds a list of Az Cli extensions from an external source

.ROLE
Administrators

#>
Param(
    [parameter(Mandatory = $true)] [string[]] $extensions
)

$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine')

# If no account is available will send back null -> will then parse to UI that needs az login
$value = az account show

# Az account is logged in and valid
if ($null -ne $value) {
    foreach ($extension in $extensions) {
        az extension add --source $extension --yes --upgrade --only-show-errors
    }
}
}
## [END] Add-WACAKAzCliExtensionsExternal ##
function Enable-WACAKAzureArc {
<#

.SYNOPSIS
Onboard to Azure Arc.

.DESCRIPTION
Onboard to Azure Arc.

.EXAMPLE
./Enable-AzureArc.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $SubscriptionId,
    [Parameter(Mandatory = $true)]
    [string]
    $ResourceGroup,
    [Parameter(Mandatory = $true)]
    [string]
    $Region,
    [Parameter(Mandatory = $true)]
    [string]
    $ConnectedCluster,
    [Parameter(Mandatory = $true)]
    [string]
    $TenantId,
    [Parameter(Mandatory = $true)]
    [string]
    $ClientId,
    [Parameter(Mandatory = $true)]
    [string]
    $ClientSecret
)

Import-Module AksHci

$secure_clientSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential($ClientId, $secure_clientSecret)
$EnvironmentName = 'AzureCloud'

Connect-AzAccount -Environment $EnvironmentName -TenantId $TenantId -SubscriptionId $SubscriptionId -Credential $credentials -ServicePrincipal | Out-Null
Enable-AksHciArcConnection -Name $ConnectedCluster -tenantId $TenantId -subscriptionId $SubscriptionId -resourceGroup $ResourceGroup -credential $credentials -location $Region -WarningAction SilentlyContinue
}
## [END] Enable-WACAKAzureArc ##
function Get-WACAKAdAuthPodsExist {
<#

.SYNOPSIS
Checks if the AD Auth pods exist.

.DESCRIPTION
Checks if the Ad Auth webhook pod exists.

.EXAMPLE
./Get-AdAuthPodsExist.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Readers

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $kubectl,
    [Parameter(Mandatory = $true)]
    [string]
    $clusterName
)

Import-Module AksHci

$kubeconfigLocation = $(New-Item -Path "$env:TEMP\workload-kubeconfig" -ItemType File -Force).FullName
$kubectl = $ExecutionContext.InvokeCommand.ExpandString($kubectl)

AksHci\Get-AksHciCredential -Name $clusterName -configPath $kubeconfigLocation -Confirm:$false

$result = & $kubectl --kubeconfig=$kubeconfigLocation get pods --all-namespaces | ConvertTo-Json

Remove-Item -Path $kubeconfigLocation -Force

return $result.contains("ad-auth-webhook")
}
## [END] Get-WACAKAdAuthPodsExist ##
function Get-WACAKAksClusters {
<#
.SYNOPSIS
Retrieves the list of cluster in an AKS HCI setup
.DESCRIPTION
Returns an array of kubernetes workload clusters
.EXAMPLE
./Get-AksClusters.ps1
Return a inventory object.
.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.
.ROLE
Readers
#>

Param(
    [Parameter(Mandatory = $false)]
    [bool]
    $IsWacGateway = $false
)

function main {
    Import-Module AksHci

    # TODO:  Get rid of this function when data is available from AksHci PowerShell module, funtions to remove: Get-TempKubeconfigFilePath, Invoke-CommandLine, Invoke-Kubectl and  Get-WorkloadCusterKubeConfig
    $BinariesFilePath = Join-Path -Path $Env:ProgramFiles -ChildPath 'AksHci'
    $KubeCtlFullPath = Join-Path -Path $BinariesFilePath -ChildPath 'kubectl.exe'

    function Get-TempKubeconfigFilePath {
        param (
            [string] $ClusterName
        )

        $KubeConfigFileName = "$ClusterName-kubeconfig"
        $TempDirectory = $(New-Item -Path "$env:TEMP\AksHciKubeConfigs" -ItemType Directory -Force).FullName
        Join-Path -Path $TempDirectory -ChildPath $KubeConfigFileName
    }

    function Invoke-CommandLine {
        <#
        .DESCRIPTION
            Executes a command and optionally ignores errors.
        .PARAMETER command
            Comamnd to execute.
        .PARAMETER arguments
            Arguments to pass to the command.
        .PARAMETER ignoreError
            Optionally, ignore errors from the command (don't throw).
        .PARAMETER showOutput
            Optionally, show live output from the executing command.
        #>
        param (
            [String]$command,
            [String]$arguments,
            [Switch]$ignoreError,
            [Switch]$showOutput
        )
        try {
            if ($showOutput.IsPresent) {
                $result = (& $command ([Text.RegularExpressions.Regex]::Split( $Arguments, " (?=(?:[^']|'[^']*')*$)" ) -Replace "'", "") | Microsoft.PowerShell.Core\Out-Default)
            }
            else {
                $result = (& $command ([Text.RegularExpressions.Regex]::Split( $Arguments, " (?=(?:[^']|'[^']*')*$)" ) -Replace "'", "") 2>&1)
            }
        }
        catch {
            if ($ignoreError.IsPresent) {
                return
            }
            throw
        }
        $out = $result | Where-Object { $_.gettype().Name -ine "ErrorRecord" }  # On a non-zero exit code, this may contain the error
        #$outString = ($out | Out-String).ToLowerInvariant()
        if ($LASTEXITCODE) {
            if ($ignoreError.IsPresent) {
                return
            }
            $err = $result | Where-Object { $_.gettype().Name -eq "ErrorRecord" }
            throw "$command $arguments failed to execute [$err]"
        }
        return $out
    }

    function Invoke-Kubectl {
        <#
        .DESCRIPTION
            Executes a kubectl command.
        .PARAMETER kubeconfig
            The kubeconfig file to use. Defaults to the management kubeconfig.
        .PARAMETER arguments
            Arguments to pass to the command.
        .PARAMETER ignoreError
            Optionally, ignore errors from the command (don't throw).
        .PARAMETER showOutput
            Optionally, show live output from the executing command.
        #>
        param (
            [string] $Kubeconfig,
            [string] $Arguments,
            [switch] $IgnoreError,
            [switch] $ShowOutput
        )
        return Invoke-CommandLine -command $global:KubeCtlFullPath -arguments $("--kubeconfig='$kubeconfig' $Arguments") -showOutput:$showOutput.IsPresent -ignoreError:$ignoreError.IsPresent
    }

    # TODO: Refactor or remove this when data is avaliable from PowershellModule
    function Get-WorkloadCusterKubeConfig {
        <#
        .DESCRIPTION
            Retrieves workload cluster's kubeconfig

        .PARAMETER ClusterName
            The name of the cluster who's kubeconfig is being retrieved

        .PARAMETER BinariesPath
            Installtion directory for AKS HCI binaries
        #>
        param (
            [Parameter(Mandatory = $true)]
            [string]
            $ClusterName,
            [Parameter(Mandatory = $false)]
            [string]
            $KubeconfigFilePath
        )

        if ([string]::IsNullOrEmpty($KubeconfigFilePath)) {
            $KubeconfigFilePath = Get-TempKubeconfigFilePath -ClusterName $ClusterName
        }

        try {
            Get-AksHciCredential -name $ClusterName -configPath $KubeconfigFilePath -adAuth -Confirm:$false 2>$null
        }
        catch {
            Get-AksHciCredential -name $ClusterName -configPath $KubeconfigFilePath -Confirm:$false 2>$null
        }

        $MaxRetries = 5
        $Retries = 0
        While (!(Test-Path $KubeconfigFilePath) -and $Retries -lt $MaxRetries) {
            $Retries ++
            Start-Sleep 5
        }

    }

    function Get-AzureArcDetails {
        Param(
            [Parameter(Mandatory = $true)]
            [string]
            $ClusterName
        )

        try {
            $KubeconfigFilePath = Get-TempKubeconfigFilePath -ClusterName $ClusterName
            Get-WorkloadCusterKubeConfig -ClusterName $ClusterName -KubeconfigFilePath $KubeconfigFilePath | Out-Null

            $Data = Invoke-Kubectl -Kubeconfig $KubeconfigFilePath "get cm azure-clusterconfig -n azure-arc -o json" | ConvertFrom-Json | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty Data

            if ($null -ne $Data) {
                return @{
                    ResourceName   = $Data.AZURE_RESOURCE_NAME
                    ResourceGroup  = $Data.AZURE_RESOURCE_GROUP
                    SubscriptionId = $Data.AZURE_SUBSCRIPTION_ID
                }
            }
        }
        catch {
            # Do nothing, we'll just ignore errors we encounter trying to get cluster list, If we can retrieve the list we won't let this block the list from showing.
        }

        return $null
    }

    function Get-AKShybridClusters {
        try {
            [array]$Clusters = @()
            $account = az account show --only-show-errors | ConvertFrom-Json
            if ($null -ne $account.id) {
                $AKShybridClusters = az hybridaks list --subscription $account.id --only-show-errors | ConvertFrom-Json
                if ($null -ne $AKShybridClusters) {
                    foreach ($Cluster in $AKShybridClusters) {
                        Add-Member -InputObject $Cluster -NotePropertyName ProvisioningState -NotePropertyValue $Cluster.properties.provisioningState
                        Add-Member -InputObject $Cluster -NotePropertyName KubernetesVersion -NotePropertyValue $Cluster.properties.kubernetesVersion
                        $AzureArcDetails = @{
                            ResourceName = $Cluster.name
                            ResourceGroup = $Cluster.resourceGroup
                            SubscriptionId = $account.id
                        }
                        Add-Member -InputObject $Cluster -NotePropertyName AzureArcDetails -NotePropertyValue $AzureArcDetails
                        Add-Member -InputObject $Cluster -NotePropertyName NodePoolDetails -NotePropertyValue $Cluster.properties.agentPoolProfiles
                        Add-Member -InputObject $Cluster -NotePropertyName ClusterType -NotePropertyValue 'Hybrid'
                        $Clusters += $Cluster
                    }
                }
            }
            return $Clusters
        }
        catch {
            return $null
        }
    }

    [array]$AksHciClusters = Get-AksHciCluster
    [array]$Clusters = @()
    if ($null -ne $AksHciClusters) {
        foreach ($Cluster in $AksHciClusters) {
            # TODO: Refactor or get rid of this when data is available from AksHci PowerShell module
            $AzureArcDetails = Get-AzureArcDetails -ClusterName $Cluster.Name
            Add-Member -InputObject $Cluster -NotePropertyName AzureArcDetails -NotePropertyValue $AzureArcDetails
            [array]$NodePoolDetails = Get-AksHciNodePool -ClusterName $Cluster.Name
            Add-Member -InputObject $Cluster -NotePropertyName NodePoolDetails -NotePropertyValue $NodePoolDetails
            $ClusterType = 'Local'
            if ($null -ne $AzureArcDetails) {
                $ClusterType = 'LocalArc'
            }
            Add-Member -InputObject $Cluster -NotePropertyName ClusterType -NotePropertyValue $ClusterType
            $Clusters += $Cluster
        }
    }
    $errorActionPreference = 'silentlycontinue'
    [array]$AksHybridClusters = Get-AKShybridClusters
    if ($null -ne $AksHybridClusters) {
        foreach($Cluster in $AksHybridClusters) {
            $Clusters += $Cluster
        }
    }

    $Clusters
}

if ($IsWacGateway) {
    PowerShell -NoProfile -Windowstyle Hidden  ${function:main}
}
else {
    main
}

}
## [END] Get-WACAKAksClusters ##
function Get-WACAKAksHciBillingStatus {
<#

.SYNOPSIS
Get AKS HCI billing status information

.DESCRIPTION
Get AKS HCI billing status information

.EXAMPLE
./Get-AksHciBillingStatus.ps1

.NOTES
Returns a list of available updates for the current version

.ROLE
Readers

#>

PowerShell -NoProfile -Windowstyle Hidden {
    Import-Module AksHci

    AksHci\Get-AksHciBillingStatus
}

}
## [END] Get-WACAKAksHciBillingStatus ##
function Get-WACAKAksHciCluster {
<#

.SYNOPSIS
Retrieves a list of all Aks Hci clusters running on the Aks host.

.DESCRIPTION
Retrieves a list of all Aks Hci clusters running on the Aks host.

.EXAMPLE
./Get-AksHciCluster.ps1

.NOTES
The supported operating systems are Azure Stack HCI, Windows Server 2019, and Windows Server 2016.

.ROLE
Readers

#>
Param(
    [Parameter(Mandatory=$false)]
    [string]
    $ClusterName
)

Import-Module AksHci

if ($ClusterName -eq [String]::empty) {
    AksHci\Get-AksHciCluster
} else {
    AksHci\Get-AksHciCluster -Name $ClusterName
}

}
## [END] Get-WACAKAksHciCluster ##
function Get-WACAKAksHciClusterNetwork {
<#
.SYNOPSIS
Retrieves the list of virtual networks from the AKS-HCI host.

.DESCRIPTION
Retrieves the list of virtual networks from the AKS-HCI host.

.EXAMPLE
./Get-AksHciClusterNetwork.ps1
Returns an array of virtual networks.

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Readers
#>

Param(
    [AllowNull()]
    [AllowEmptyString()]
    [Parameter(Mandatory = $false)]
    [string]
    $VnetName
)

Import-Module AksHci

if (($null -eq $VnetName) -or (-$VnetName -eq [String]::empty)) {
    Get-AksHciClusterNetwork
} else {
    Get-AksHciClusterNetwork -name $VnetName
}

}
## [END] Get-WACAKAksHciClusterNetwork ##
function Get-WACAKAksHciClusterUpdates {
<#

.SYNOPSIS
Gets target cluster updates

.DESCRIPTION
Gets target cluster updates

.EXAMPLE
./Get-AksHciClusterUpdates.ps1

.NOTES
Returns a list of available updates for the current version

.ROLE
Readers

#>
Param(
    [Parameter(Mandatory = $true)]
    [string]
    $Name
)

Import-Module AksHci

AksHci\Get-AksHciClusterUpdates -Name $Name

}
## [END] Get-WACAKAksHciClusterUpdates ##
function Get-WACAKAksHciCredential {
<#

.SYNOPSIS
Gets the kubeconfig for the cluster.

.DESCRIPTION
Gets the kubeconfig for the cluster.

.EXAMPLE
./Get-AksHciCredential.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string[]]
    $clusterNames
)

Import-Module AksHci

$kubeconfigLocations = @()
foreach($clusterName in $clusterNames) {
    $fileName = "$clusterName-kubeconfig"
    $kubeconfigLocation = $(New-Item -Path "$env:TEMP\$fileName" -ItemType File -Force).FullName
    try {
        Get-AksHciCredential -Name $clusterName -configPath $kubeconfigLocation -adAuth -Confirm:$false
    }
    catch {
        Get-AksHciCredential -Name $clusterName -configPath $kubeconfigLocation -Confirm:$false
    }
    $MaxRetries = 5
    $Retries = 0
    While (!(Test-Path $kubeconfigLocation) -and $Retries -lt $MaxRetries) {
        $Retries ++
        Start-Sleep 5
    }
    $kubeconfigLocations += $kubeconfigLocation
}
return $kubeconfigLocations
}
## [END] Get-WACAKAksHciCredential ##
function Get-WACAKAksHciEventLogs {
<#

.SYNOPSIS
Get Aks Hci Event logs.

.DESCRIPTION
Get Aks Hci Event logs.

.EXAMPLE
./Get-AksHciEventLogs.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $LogName,
    [Parameter(Mandatory = $true)]
    [string]
    $MatchString,
    [Parameter(Mandatory = $true)]
    [int]
    $TimeElapsedInSec
)

$requiredLogGenerationTime = (get-date).AddSeconds(-$TimeElapsedInSec);

Get-EventLog -LogName $LogName |
Where-Object { $_.Message -Match $MatchString  -and $_.TimeGenerated -ge $requiredLogGenerationTime } |
Microsoft.PowerShell.Utility\Select-Object -First 1 -ErrorAction SilentlyContinue;

}
## [END] Get-WACAKAksHciEventLogs ##
function Get-WACAKAksHciKubernetesVersion {
<#
.SYNOPSIS
Retrieves the of supported Kubernetes versions from the Aks Hci host.
.DESCRIPTION
Retrieves the of supported Kubernetes versions from the Aks Hci host.
.EXAMPLE
./Get-AksHciKubernetesVersion.ps1
Returns an object with mapping of OS to a list of Kubernetes versions.
.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.
.ROLE
Readers
#>

Import-Module AksHci

Get-AksHciKubernetesVersion
}
## [END] Get-WACAKAksHciKubernetesVersion ##
function Get-WACAKAksHciLogs {
<#

.SYNOPSIS
Collects the Aks-HCI logs.

.DESCRIPTION
Collects the Aks-HCI logs.

.EXAMPLE
./Get-AksHciLogs.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Import-Module AksHci

Get-AksHciLogs
}
## [END] Get-WACAKAksHciLogs ##
function Get-WACAKAksHciNodePool {
<#

.SYNOPSIS
Retrieves a list of all node pools associated with the cluster for AKS on Azure Stack HCI.

.DESCRIPTION
Retrieves a list of all node pools associated with the cluster for AKS on Azure Stack HCI.

.EXAMPLE
./Get-AksHciNodePool.ps1

.NOTES
The supported operating systems are Azure Stack HCI, Windows Server 2019, and Windows Server 2016.

.ROLE
Readers

#>
Param(
    [Parameter(Mandatory=$false)]
    [string]
    $ClusterName
)

Import-Module AksHci

AksHci\Get-AksHciNodePool -clusterName $ClusterName

}
## [END] Get-WACAKAksHciNodePool ##
function Get-WACAKAksHciVmSize {

<#
.SYNOPSIS
Gets AksHci VM Sizes
.DESCRIPTION
Gets AksHci VM Sizes
.EXAMPLE
./Get-AksHciVmSize.ps1
.NOTES
Return information on VM sizes, each entry in the returned array has the following fields: VMSize, CPU and memoryGB
.ROLE
Readers
#>

Import-Module AksHci

AksHci\Get-AksHciVmSize | Microsoft.PowerShell.Utility\Select-Object @{Name='Value'; Expression={$_.VmSize.ToString()}}, CPU, MemoryGB

}
## [END] Get-WACAKAksHciVmSize ##
function Get-WACAKAksHostInventory {
<#
.SYNOPSIS
Gets information about the Azure Kubernetes platform that is already setup on a server / cluster.
.DESCRIPTION
Retrieves inventory from the Azure Kubernetes platform.
.EXAMPLE
./Get-AksHostInventory.ps1
Return a inventory object.
.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.
.ROLE
Readers
#>
Param(
    [Parameter(Mandatory = $false)]
    [bool]
    $IsFailoverCluster = $false,
    [Parameter(Mandatory = $false)]
    [bool]
    $IsWacGateway = $false
)

function main {
    Param(
        $IsFailoverCluster = $false
    )
    $ModuleName = 'AksHci'

    function Get-SetupConfig {
        $AksHciModule = Get-Module -Name $ModuleName -ErrorAction SilentlyContinue
        if ($null -ne $AksHciModule) {
            try {
                return Get-AksHciConfig;
            }
            catch {
                if ($_.Exception.Message -like '*This machine does not appear to be configured for deployment*') {
                    return $null
                }
                throw
            }
        }
        else {
            return $null
        }
    }

    function Get-ResourceBridgeConfig {
        try {
            return Get-MocConfig;
        }
        catch {
            if (($_.Exception.Message -like '*This machine does not appear to be configured for deployment*') -or
            ($_.Exception.Message -like '*The term ''Get-MocConfig'' is not recognized as the name of a cmdlet*')) {
                return $null
            }
            throw
        }
        else {
            return $null
        }
    }

    function Get-DiskInfo {
        param (
            [Parameter(Mandatory = $true)]
            [bool]
            $IsFailoverCluster
        )
        $DiskInfo = @{
            TotalDiskSpace = 0;
            FreeDiskSpace  = 0;
        }

        if ($IsFailoverCluster) {
            $ClusterDiskInfo = $DiskInfo
            Import-Module -Name FailoverClusters
            $csvs = Get-ClusterSharedVolume
            foreach ($csv in $csvs) {
                $ClusterDiskInfo.FreeDiskSpace += $csv.SharedVolumeInfo.Partition.FreeSpace;

                $ClusterDiskInfo.TotalDiskSpace += $csv.SharedVolumeInfo.Partition.Size;
            }
            return $ClusterDiskInfo;
        }

        # DriveType 3 = local disk
        # https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-logicaldisk
        $DiskInfoSum = Get-CimInstance -Class CIM_LogicalDisk `
        | Microsoft.PowerShell.Utility\Select-Object @{Name = "TotalDiskSpace"; Expression = { $_.size } }, @{Name = "FreeDiskSpace"; Expression = { $_.freespace } }, DriveType `
        | Where-Object DriveType -EQ '3' `
        | Microsoft.PowerShell.Utility\Measure-Object -Property TotalDiskSpace, FreeDiskSpace -Sum | Sort-Object Sum

        $DiskInfo.FreeDiskSpace = $DiskInfoSum[0].Sum
        $DiskInfo.TotalDiskSpace = $DiskInfoSum[1].Sum

        $DiskInfo
    }

    function Get-MemoryInfo {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [bool]
            $IsFailoverCluster
        )

        if ($IsFailoverCluster) {
            $ClusterMemoryInfo = @{
                TotalMemory = 0;
                FreeMemory  = 0;
            }

            $clusterData = Get-WmiObject Win32_OperatingSystem `
            | Microsoft.PowerShell.Utility\Select-Object @{l = 'FreeMemory'; e = { $_.FreePhysicalMemory + $_.FreeVirtualMemory } }, @{l = 'TotalMemory'; e = { $_.TotalVisibleMemorySize + $_.TotalVirtualMemorySize } }
            $ClusterMemoryInfo.FreeMemory += $clusterData.freememory * 1024;
            $ClusterMemoryInfo.TotalMemory += $clusterData.totalmemory * 1024;

            return $ClusterMemoryInfo;
        }

        $OsWmiObject = Get-CimInstance -Class Win32_OperatingSystem
        $TotalMemory = Get-CimInstance -Class Win32_PhysicalMemory | Microsoft.PowerShell.Utility\Measure-Object -Property capacity -Sum | ForEach-Object { $_.Sum }

        $MemoryInfo = @{
            TotalMemory = $TotalMemory;
            FreeMemory  = $OsWmiObject.FreePhysicalMemory * 1024;
        }

        $MemoryInfo
    }

    Import-Module $ModuleName -ErrorAction SilentlyContinue -ErrorVariable ImportErr
    if (($null -ne $ImportErr) -and ($null -ne $ImportErr[0])) {
        if ($ImportErr.FullyQualifiedErrorId -ne "Modules_ModuleNotFound,Microsoft.PowerShell.Commands.ImportModuleCommand") {
            throw $ImportErr[0]
        }
    }

    $Inventory = @{
        InstallationState = 'NotInstalled';
        StoragePath       = $null;
        TotalDiskSpace    = $null;
        FreeDiskSpace     = $null;
        TotalMemory       = $null;
        FreeMemory        = $null;
        KubernetesVersion = $null;
        AksHciVersion     = $null;
        Networking        = $null;
        AksHciClusters    = $null;
        OperatingSystem   = $null;
        ModuleVersion     = $null;
        Catalog           = $null;
        Ring              = $null;
        RBInstalledFirst  = $null;
    }

    $AksHciConfigs = Get-SetupConfig

    if ($null -ne $AksHciConfigs.AksHci.InstallState) {
        $ValidInstallStates = @(
            [InstallState]::Updating,
            [InstallState]::UpdateFailed,
            [InstallState]::Installed
        )

        $CurrentInstallState = $AksHciConfigs.AksHci.InstallState

        if ($CurrentInstallState -in $ValidInstallStates) {
            $Inventory.InstallationState = $AksHciConfigs.AksHci.installState.toString()
            $Inventory.Catalog = $AksHciConfigs.AksHci.catalog.toString()
            $Inventory.ring = $AksHciConfigs.AksHci.ring.toString()
            $Inventory.StoragePath = $AksHciConfigs.Kva.imageDir

            $InstalledVersions = Get-Module -Name $ModuleName | Sort-Object -Property Version -Descending
            $Inventory.ModuleVersion = $InstalledVersions[0].version.ToString()
            $Inventory.AksHciVersion = $AksHciConfigs.Kva.version
            $Inventory.KubernetesVersion = $AksHciConfigs.Kva.kvaK8sVersion

            $DiskSpaceInfo = Get-DiskInfo -IsFailoverCluster $IsFailoverCluster
            $Inventory.TotalDiskSpace = $DiskSpaceInfo.TotalDiskSpace
            $Inventory.FreeDiskSpace = $DiskSpaceInfo.FreeDiskSPace

            $MemoryInfo = Get-MemoryInfo -IsFailoverCluster $IsFailoverCluster
            $Inventory.TotalMemory = $MemoryInfo.TotalMemory
            $Inventory.FreeMemory = $MemoryInfo.FreeMemory
        }
    } else {
        $ResourceBridgeConfig = Get-ResourceBridgeConfig
        if ($null -ne $ResourceBridgeConfig) {
            $Inventory.RBInstalledFirst = $true
        } else {
            $Inventory.RBInstalledFirst = $false
        }
    }

    $Inventory.OperatingSystem = (Get-CimInstance -Class Win32_OperatingSystem -ErrorAction SilentlyContinue).caption

    $Inventory
}

if ($IsWacGateway) {
    PowerShell -NoProfile  ${function:main} -args $IsFailoverCluster
}
else {
    main -IsFailoverCluster $IsFailoverCluster
}
}
## [END] Get-WACAKAksHostInventory ##
function Get-WACAKAllEventLogs {
<#

.SYNOPSIS
Get all Event logs.

.DESCRIPTION
Get all Event logs.

.EXAMPLE
./Get-AllEventLogs.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $LogName,
    [Parameter(Mandatory = $true)]
    [string]
    $MatchString,
    [Parameter(Mandatory = $true)]
    [int]
    $TimeElapsedInSec
)
$ErrorActionPreference = 'SilentlyContinue'
$requiredLogGenerationTime = (get-date).AddSeconds(-$TimeElapsedInSec);

Get-EventLog -LogName $LogName |
Where-Object { $_.Message -Match $MatchString -and $_.TimeGenerated -ge $requiredLogGenerationTime } |
Sort-Object -Property TimeGenerated

}
## [END] Get-WACAKAllEventLogs ##
function Get-WACAKArcPods {
<#

.SYNOPSIS
Gets the Azure Arc pods.

.DESCRIPTION
Gets the Azure Arc pods.

.EXAMPLE
./Get-ArcPods.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $kubectl,
    [Parameter(Mandatory = $true)]
    [string]
    $clusterName
)

Import-Module AksHci

$kubeconfigLocation = $(New-Item -Path "$env:TEMP\workload-kubeconfig" -ItemType File -Force).FullName
$kubectl = $ExecutionContext.InvokeCommand.ExpandString($kubectl)

AksHci\Get-AksHciCredential -Name $clusterName -configPath $kubeconfigLocation -Confirm:$false

$namespace = 'azure-arc-onboarding'
$result = & $kubectl --kubeconfig=$kubeconfigLocation get pods -n $namespace 2>$null

Remove-Item -Path $kubeconfigLocation -Force

return $result
}
## [END] Get-WACAKArcPods ##
function Get-WACAKAzureStackHCIStatus {
<#

.SYNOPSIS
Get the Azure Stack HCI status.

.DESCRIPTION
Get the Azure Stack HCI registration and connection status.

.EXAMPLE
Get-AzureStackHCIStatus

.ROLE
Readers

#>

Import-Module AzureStackHCI

$azureStackHCIStatus = @{
    RegistrationStatus = $null;
    ConnectionStatus   = $null;
}

$azureStackHCI = Get-AzureStackHCI

if ($null -ne $azureStackHCI.RegistrationStatus) {
    $azureStackHCIStatus.RegistrationStatus = $azureStackHCI.RegistrationStatus.toString()
}
if ($null -ne $azureStackHCI.ConnectionStatus) {
    $azureStackHCIStatus.ConnectionStatus = $azureStackHCI.ConnectionStatus.toString()
}

$azureStackHCIStatus
}
## [END] Get-WACAKAzureStackHCIStatus ##
function Get-WACAKCimInstanceIWin32OperatingSystem {
<#

.SYNOPSIS
Gets information about a server for host configuration settings.

.DESCRIPTION
Retrieves inventory based on the requirements of the server.

.EXAMPLE
./Get-IsVirtualMachine.ps1
Return a server inventory object with data related to its configuration.

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Readers

#>


Get-CimInstance -Namespace root\CIMv2 -className Win32_OperatingSystem

}
## [END] Get-WACAKCimInstanceIWin32OperatingSystem ##
function Get-WACAKClusterSharedVolume {
<#

.SYNOPSIS
Gets a list of Cluster Shared Volumes.

.DESCRIPTION
Gets information about Cluster Shared Volumes in a failover cluster.

.EXAMPLE
Get-ClusterSharedVolume -Cluster cluster1

.NOTES
Takes in parameters -Cluster, -InputObject, and -Name.

.ROLE
Readers

#>

# Filter out Infrastructure volumes to prevent users from selecting these because these volumes are reserved for the infra services
# of the ASZ worker stamp. Future scenarios like update and optional feature enablement would be blocked if there is not enough space on these volumes.
$infrastructureVolumePrefix = 'Cluster Virtual Disk (Infrastructure_'
$Volumes = FailoverClusters\Get-ClusterSharedVolume | Where-Object { !($_.Name.StartsWith($infrastructureVolumePrefix,'CurrentCultureIgnoreCase')) }

function Get-ClusterSharedVolumeFreeSpace {
  Param(
    [Parameter(Mandatory = $true)]
    [string]$CsvName
  )

  $FreeSpace = (FailoverClusters\Get-ClusterSharedVolume -Name $CsvName `
    | Microsoft.PowerShell.Utility\Select-Object -Expand SharedVolumeInfo `
    | Microsoft.PowerShell.Utility\Select-Object @{n = 'FreeSpace'; e = { ($_.Partition.Size - $_.Partition.UsedSpace) / 1GB } } `
    | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty 'FreeSpace'
  )
  return $FreeSpace
}

function Get-FQDN {
  Param(
    [Parameter(Mandatory = $true)]
    [object]$Volume
  )
  return ("$($Volume.OwnerNode.Name).$env:userdnsdomain").ToLower();
}

function Get-DriveLetter {
  Param(
    [Parameter(Mandatory = $true)]
    [object]$Volume
  )

  return (Get-Item $Volume.sharedVolumeInfo.FriendlyVolumeName).PSDrive.Name;
}

function Get-AdminShare {
  Param(
    [Parameter(Mandatory = $true)]
    [object]$Volume
  )

  return @{
    Name = "$(Get-DriveLetter $Volume)$"
    UncPath = "\\$(Get-FQDN $Volume)\$(Get-DriveLetter $Volume)$"
  }
}

$Results = @()
foreach ($Volume in $Volumes) {
  $isOnline = $Volume.State -eq 'Online'
  $Result = @{
    Name   = $Volume.Name;
    Online = $isOnline;
    AdminShare = $null;
    Path = $null;
    FreeSpace = $null;
  }
  if ($isOnline) {
    $Result.AdminShare = (Get-AdminShare $Volume);
    $Result.Path  = $Volume.sharedVolumeInfo.FriendlyVolumeName;
    $Result.FreeSpace = (Get-ClusterSharedVolumeFreeSpace -CsvName $Volume.Name)
  }
  $Results += $Result
}

$Results

}
## [END] Get-WACAKClusterSharedVolume ##
function Get-WACAKFileContent {
<#

.SYNOPSIS
Gets the content of a file.

.DESCRIPTION
Gets the content of a file.

.EXAMPLE
./Get-FileContent.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Readers

#>

Param(
    [Parameter(Mandatory=$true)]
    [string]
    $Path
)

$Path = $ExecutionContext.InvokeCommand.ExpandString($Path);
$Content = Get-Content -Path $Path

return $Content
}
## [END] Get-WACAKFileContent ##
function Get-WACAKInstallEventsProgress {
<#

.SYNOPSIS
Gets event logs logged by an install script

.DESCRIPTION
Gets event logs logged by an install script

.EXAMPLE
./Get-InstallEventsProgress.ps1

.NOTES
Outputs all event logs by an install script using a unique logsource

.ROLE
Administrators

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $LogSource
)
$LogName = "Microsoft-ServerManagementExperience";
New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
Microsoft.PowerShell.Management\Get-EventLog  -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
}
## [END] Get-WACAKInstallEventsProgress ##
function Get-WACAKPendingRebootStatus {
<#

.SYNOPSIS
Determines if a server needed to be rebooted.

.DESCRIPTION
Determines if a server needed to be rebooted based on this criteria: 
- A role or feature requires a restart of a server.

.EXAMPLE
./Get-PendingRebootStatus.ps1
Return true or false depending on if the server needs to be restarted.

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Readers

#>

# Determines if a reboot is needed for Component Based Servicing (Windows roles & features)
$subKeyNames = (get-item "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing").GetSubKeyNames()
$pendingReboot = $subKeyNames -contains "RebootPending"		
return $pendingReboot
}
## [END] Get-WACAKPendingRebootStatus ##
function Get-WACAKPlatformUpdateInfo {
<#

.SYNOPSIS
Gets update info for the host

.DESCRIPTION
Get current version and available updates for the host

.EXAMPLE
./Get-PlatformUpdateInfo.ps1

.NOTES
Returns a list of available updates and the current version

.ROLE
Readers

#>

Param(
    [Parameter(Mandatory = $false)]
    [bool]
    $IsWacGateway = $false
)

function main {
    Import-Module AksHci

    $AllUpdates = Get-AksHciUpdates

    $UpdateInfo = @{
        CurrentVersion = Get-AksHciVersion
        AllUpdates    = $AllUpdates
    }

    $UpdateInfo
}

if ($IsWacGateway) {
    PowerShell -NoProfile -Windowstyle Hidden  ${function:main}
}
else {
    main
}

}
## [END] Get-WACAKPlatformUpdateInfo ##
function Get-WACAKResourceBridgeInventory {
<#
.SYNOPSIS
Gets information about the Resource Bridge platform that is setup on a cluster.
.DESCRIPTION
Retrieves inventory from the Resource Bridge
.EXAMPLE
./Get-ResourceBridgeInventory.ps1
Return a inventory object.
.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.
.ROLE
Readers
#>
Param(
    [Parameter(Mandatory = $false)]
    [bool]
    $IsFailoverCluster = $false
)

Set-StrictMode -Version 5.0
$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine')

$MocModuleName = 'Moc'

function Get-MocSetupConfig {
    $MocModule = Get-Module -Name $MocModuleName -ErrorAction SilentlyContinue
    if ($null -ne $MocModule) {
        try {
            return Get-MocConfig;
        }
        catch {
            if ($_.Exception.Message -like '*This machine does not appear to be configured for deployment*') {
                return $null
            }
            throw
        }
    }
    else {
        return $null
    }
}

function Get-DiskInfo {
    param (
        [Parameter(Mandatory = $true)]
        [bool]
        $IsFailoverCluster
    )
    $DiskInfo = @{
        TotalDiskSpace = 0;
        FreeDiskSpace  = 0;
    }

    if ($IsFailoverCluster) {
        $ClusterDiskInfo = $DiskInfo
        Import-Module -Name FailoverClusters
        $csvs = FailoverClusters\Get-ClusterSharedVolume
        foreach ($csv in $csvs) {
            $ClusterDiskInfo.FreeDiskSpace += $csv.SharedVolumeInfo.Partition.FreeSpace;

            $ClusterDiskInfo.TotalDiskSpace += $csv.SharedVolumeInfo.Partition.Size;
        }
        return $ClusterDiskInfo;
    }

    # DriveType 3 = local disk
    # https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-logicaldisk
    $DiskInfoSum = Get-CimInstance -Class CIM_LogicalDisk `
    | Microsoft.PowerShell.Utility\Select-Object @{Name = "TotalDiskSpace"; Expression = { $_.size } }, @{Name = "FreeDiskSpace"; Expression = { $_.freespace } }, DriveType `
    | Where-Object DriveType -EQ '3' `
    | Microsoft.PowerShell.Utility\Measure-Object -Property TotalDiskSpace, FreeDiskSpace -Sum | Sort-Object Sum

    $DiskInfo.FreeDiskSpace = $DiskInfoSum[0].Sum
    $DiskInfo.TotalDiskSpace = $DiskInfoSum[1].Sum

    $DiskInfo
}

function Get-MemoryInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [bool]
        $IsFailoverCluster
    )

    if ($IsFailoverCluster) {
        $ClusterMemoryInfo = @{
            TotalMemory = 0;
            FreeMemory  = 0;
        }

        $clusterData = Get-WmiObject Win32_OperatingSystem `
        | Microsoft.PowerShell.Utility\Select-Object @{l = 'FreeMemory'; e = { $_.FreePhysicalMemory + $_.FreeVirtualMemory } }, @{l = 'TotalMemory'; e = { $_.TotalVisibleMemorySize + $_.TotalVirtualMemorySize } }
        $ClusterMemoryInfo.FreeMemory += $clusterData.freememory * 1024;
        $ClusterMemoryInfo.TotalMemory += $clusterData.totalmemory * 1024;

        return $ClusterMemoryInfo;
    }

    $OsWmiObject = Get-CimInstance -Class Win32_OperatingSystem
    $TotalMemory = Get-CimInstance -Class Win32_PhysicalMemory | Microsoft.PowerShell.Utility\Measure-Object -Property capacity -Sum | ForEach-Object { $_.Sum }

    $MemoryInfo = @{
        TotalMemory = $TotalMemory;
        FreeMemory  = $OsWmiObject.FreePhysicalMemory * 1024;
    }

    $MemoryInfo
}

function Get-DataFromYamlFile {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $Path,
        [Parameter(Mandatory = $true)]
        [String]
        $SearchString
    )

    $yamlContent = Get-Content -Path $Path -Raw -ErrorAction Stop
    $startIndex = $yamlContent.IndexOf($SearchString) + $SearchString.Length
    $endIndex = $yamlContent.IndexOf("`n", $startIndex)
    return $yamlContent.Substring($startIndex, $endIndex - $startIndex)
}

function Get-AzCliInfo {
    $AzCliInfo = @{
        AzCliInstalled                      = $null;
        AzCliVersion                        = $null;
        AzCliLoggedIn                       = $null;
        AzCliSubscription                   = $null;
        ArcApplianceResourceGroup           = $null;
        ArcApplianceName                    = $null;
        ArcApplianceVersion                 = $null;
        ArcApplianceProvisioningState       = $null;
        ArcApplianceStatus                  = $null;
        k8sExtensionVersion                 = $null;
        k8sExtensionProvisioningState       = $null;
        CustomLocationName                  = $null;
        CustomLocationVersion               = $null;
        CustomLocationProvisioningState     = $null;
        CustomLocationExtendedLocationUri   = $null;
        CustomLocationAzureLocation         = $null;
    }

    $CliInstalled = Get-Command "az" -ErrorAction SilentlyContinue
    $AzCliInfo.AzCliInstalled = $null -ne $CliInstalled

    # We cannot fetch more info if the CLI is not installed
    if ($AzCliInfo.AzCliInstalled -eq $true) {
        $AzCliInfo.AzCliVersion = (az version | ConvertFrom-Json).'azure-cli'
        $AzCliInfo.AzCliLoggedIn = $false

        # Check if az cli is logged in
        try {
            # This command throws an error if account not logged in
            $account = az account show --only-show-errors | ConvertFrom-Json

            if ($null -ne $account) {
                $AzCliInfo.AzCliLoggedIn = $true
                $AzCliInfo.AzCliSubscription = $account.id
            }
        } catch {}

        # If the CLI is logged in, we can find info about the required extensions
        if ($AzCliInfo.AzCliLoggedIn -eq $true) {

            # Check which extensions are installed, all three are required
            $extensions = az extension list --only-show-errors | ConvertFrom-Json
            $extensionsToFind = @{
                "arcappliance"= "arcappliance";
                "k8s-extension"= "k8s-extension";
                "customlocation"= "customlocation"
            }

            # Go through each extension and check if has the list of all extensions we want
            foreach ($extension in $extensions) {
                if ($extensionsToFind.contains($extension.name)) {
                    $extensionsToFind.remove($extension.name)

                    if ($extension.name -eq "arcappliance") {
                        $AzCliInfo.ArcApplianceVersion = $extension.version
                    }

                    if ($extension.name -eq "k8s-extension") {
                        $AzCliInfo.k8sExtensionVersion = $extension.version
                    }

                    if ($extension.name -eq "customlocation") {
                        $AzCliInfo.CustomLocationVersion = $extension.version
                    }
                }

                # Break the loop if we've found all the extensions we are looking for
                if ($extensionsToFind.Count -eq 0) {
                    $foundExtensions = $true
                    break
                }
            }

            # Get the name, resource group and subcription from the hci-resource.yaml file, since there can be multiple
            # Resource Bridges per subscription and even multiple Resource Bridges per resource group
            $ArcConfig = Get-ArcHciConfig
            $Path = $($ArcConfig.workingDir + "\hci-resource.yaml")
            $resourceGroup = Get-DataFromYamlFile -Path $Path -SearchString "resource_group: "
            $resourceBridgeName = Get-DataFromYamlFile -Path $Path -SearchString "name: "

            $ErrorActionPreference = 'Stop'

            # We only look at the state if the extension exists
            if ($null -ne $AzCliInfo.ArcApplianceVersion) {
                # Find the arcappliance / resource bridge status
                $arcappliances = az arcappliance list --resource-group $resourceGroup --subscription $account.id --only-show-errors | ConvertFrom-Json
                foreach ($arcappliance in $arcappliances) {
                    if ($arcappliance.name -Match $resourceBridgeName) {
                        $AzCliInfo.ArcApplianceResourceGroup = $arcappliance.resourceGroup
                        $AzCliInfo.ArcApplianceName = $arcappliance.name

                        $AzCliInfo.ArcApplianceProvisioningState = $arcappliance.provisioningState
                        $AzCliInfo.ArcApplianceStatus = $arcappliance.status

                        # We break as there can only be one resource bridge created
                        break
                    }
                }
            }

            if (($null -ne $AzCliInfo.k8sExtensionVersion) -and
                ($null -ne $AzCliInfo.ArcApplianceResourceGroup) -and
                ($null -ne $AzCliInfo.ArcApplianceName)) {
                $k8sExtensions = az k8s-extension list --resource-group $AzCliInfo.ArcApplianceResourceGroup --cluster-name $AzCliInfo.ArcApplianceName --cluster-type appliances --only-show-errors | ConvertFrom-Json
                foreach ($k8sExtension in $k8sExtensions) {
                    if ($k8sExtension.id -Match $account.id) {
                        $AzCliInfo.k8sExtensionProvisioningState = $k8sExtension.provisioningState;
                        break
                    }
                }
            }

            if ($null -ne $AzCliInfo.CustomLocationVersion) {
                $customLocations = az customlocation list --resource-group $AzCliInfo.ArcApplianceResourceGroup --subscription $account.id --only-show-errors | ConvertFrom-Json
                foreach ($customLocation in $customLocations) {
                    if ($customLocation.hostResourceId -Match $AzCliInfo.ArcApplianceName) {
                        $AzCliInfo.CustomLocationName = $customLocation.name
                        $AzCliInfo.CustomLocationProvisioningState = $customLocation.provisioningState
                        $AzCliInfo.CustomLocationAzureLocation = $customLocation.location
                        $AzCliInfo.CustomLocationExtendedLocationUri = $customLocation.id
                        break
                    }
                }
            }
        }
    }

    $AzCliInfo
}

function main {
    Import-Module $MocModuleName -ErrorAction SilentlyContinue -ErrorVariable ImportErr
    if ($null -ne $ImportErr -and
        $ImportErr.Count -gt 0 -and
        $null -ne $ImportErr[0]) {
        if ($ImportErr.FullyQualifiedErrorId -ne "Modules_ModuleNotFound,Microsoft.PowerShell.Commands.ImportModuleCommand") {
            throw $ImportErr[0]
        }
    }

    $Inventory = @{
        InstallationState                   = 'NotInstalled';
        StoragePath                         = $null;
        TotalDiskSpace                      = $null;
        FreeDiskSpace                       = $null;
        TotalMemory                         = $null;
        FreeMemory                          = $null;
        MocVersion                          = $null;
        Networking                          = $null;
        OperatingSystem                     = $null;
        ModuleVersion                       = $null;
        AzCliInstalled                      = $null;
        AzCliVersion                        = $null;
        AzCliLoggedIn                       = $null;
        AzCliSubscription                   = $null;
        ArcApplianceResourceGroup           = $null;
        ArcApplianceName                    = $null;
        ArcApplianceVersion                 = $null;
        ArcApplianceProvisioningState       = $null;
        ArcApplianceStatus                  = $null;
        k8sExtensionVersion                 = $null;
        k8sExtensionProvisioningState       = $null;
        CustomLocationName                  = $null;
        CustomLocationVersion               = $null;
        CustomLocationProvisioningState     = $null;
        CustomLocationExtendedLocationUri   = $null;
        CustomLocationAzureLocation         = $null;
    }

    $MocConfig = Get-MocSetupConfig

    if ($null -ne $MocConfig -and $null -ne $MocConfig.InstallState) {
        $ValidInstallStates = @(
            [InstallState]::Updating,
            [InstallState]::UpdateFailed,
            [InstallState]::Installed
        )

        $CurrentInstallState = [InstallState]$MocConfig.installState

        if ($CurrentInstallState -in $ValidInstallStates) {
            $Inventory.InstallationState = [Enum]::ToObject([InstallState], $MocConfig.installState).toString()
            $Inventory.StoragePath = $MocConfig.imageDir

            $InstalledVersions = Get-Module -Name $MocModuleName | Sort-Object -Property Version -Descending
            $Inventory.ModuleVersion = $InstalledVersions[0].version.ToString()

            $Inventory.MocVersion = $MocConfig.version

            $DiskSpaceInfo = Get-DiskInfo -IsFailoverCluster $IsFailoverCluster
            $Inventory.TotalDiskSpace = $DiskSpaceInfo.TotalDiskSpace
            $Inventory.FreeDiskSpace = $DiskSpaceInfo.FreeDiskSPace

            $MemoryInfo = Get-MemoryInfo -IsFailoverCluster $IsFailoverCluster
            $Inventory.TotalMemory = $MemoryInfo.TotalMemory
            $Inventory.FreeMemory = $MemoryInfo.FreeMemory

            try {
                $ErrorActionPreference = 'Continue'
                $AzCliInfo = Get-AzCliInfo
                if (($AzCliInfo.AzCliInstalled -eq $false) -or ($AzCliInfo.AzCliLoggedIn -eq $false)) {
                    $Inventory.InstallationState = 'NotInstalled';
                }
            }
            catch [System.Management.Automation.ActionPreferenceStopException] {
                # Reaching this point in the script means that AKS is installed but Resource Bridge is not installed
                if ($_.Exception.Message -like '*Cannot find path *\hci-resource.yaml'' because it does not exist.*') {
                    $AzCliInfo = @{
                        AzCliInstalled                      = $null;
                        AzCliVersion                        = $null;
                        AzCliLoggedIn                       = $null;
                        AzCliSubscription                   = $null;
                        ArcApplianceResourceGroup           = $null;
                        ArcApplianceName                    = $null;
                        ArcApplianceVersion                 = $null;
                        ArcApplianceProvisioningState       = $null;
                        ArcApplianceStatus                  = $null;
                        k8sExtensionVersion                 = $null;
                        k8sExtensionProvisioningState       = $null;
                        CustomLocationName                  = $null;
                        CustomLocationVersion               = $null;
                        CustomLocationProvisioningState     = $null;
                        CustomLocationExtendedLocationUri   = $null;
                        CustomLocationAzureLocation         = $null;
                    }
                    $Inventory.InstallationState = 'NotInstalled';
                } else {
                    throw
                }
            }
            # Save each variable to returned data object
            $AzCliInfo.Keys | ForEach-Object { $Inventory[$_] = $AzCliInfo[$_] }
        }
    }

    $Inventory.OperatingSystem = (Get-CimInstance -Class Win32_OperatingSystem -ErrorAction SilentlyContinue).caption

    $Inventory
}

main

}
## [END] Get-WACAKResourceBridgeInventory ##
function Get-WACAKServerInventoryAks {
<#

.SYNOPSIS
Gets information about a server for host configuration settings.

.DESCRIPTION
Retrieves inventory based on the requirements of the server.

.EXAMPLE
./Get-ServerInventoryAks.ps1
Return a server inventory object with data related to its configuration.

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Readers

#>

Param(
  [Parameter(Mandatory = $true)]
  [object]
  $hostRequirements
)

function Get-ResultAsArray {
  Param(
    [AllowNull()]
    [Parameter(Mandatory = $true)]
    [object]$result
  )

  if ($null -eq $result) {
    return @{ Value = @() }
  }
  elseif (-Not($result -is [array])) {
    return @{ Value = @($result) }
  }

  return @{ Value = $result }
}

function Get-RolesAndFeatures {
  Param(
    [Parameter(Mandatory = $true)]
    [object[]]$requirements
  )
  $FeaturesToRetrieve = $requirements | ForEach-Object { $_.Name }
  $result = Get-WindowsFeature -Name $FeaturesToRetrieve

  $rolesAndFeatures = (Get-ResultAsArray -result $result).Value
  $outputFeatures = @()

  foreach ($requirement in $requirements) {
    $roleFeature = $rolesAndFeatures | Where-Object { $requirement.Name -eq $_.Name }
    if ($null -eq $roleFeature) {
      $outputFeatures += , @{ Name = $requirement.Name; InstallState = 'NotPresent'; }
    }
    else {
      $outputFeatures += , @{ Name = $requirement.Name; InstallState = $roleFeature.InstallState; }
    }
  }
  return Get-ResultAsArray -result $outputFeatures
}

function Get-FirewallRules {
  Param(
    [Parameter(Mandatory = $true)]
    [object[]]$requirements
  )

  Import-Module netsecurity

  $stores = @('PersistentStore', 'RSOP');
  $allRules = @()
  foreach ($store in $stores) {
    $rules = (Get-NetFirewallRule -PolicyStore $store)

    $rulesHash = @{}
    $rules | ForEach-Object {
      $newRule = ($_ | Microsoft.PowerShell.Utility\Select-Object `
          instanceId, `
          name, `
          displayName, `
          description, `
        @{Name = "enabled"; Expression = { $_.Enabled -eq [Microsoft.PowerShell.Cmdletization.GeneratedTypes.NetSecurity.Enabled]::True } }, `
          direction, `
          status, `
        @{Name = "portFilter"; Expression = { $null } })

      $rulesHash[$_.CreationClassName] = $newRule
      $allRules += $newRule
    }

    $portFilters = (Get-NetFirewallPortFilter  -PolicyStore $store)

    $portFilters | ForEach-Object {
      $newPortFilter = $_ | Microsoft.PowerShell.Utility\Select-Object dynamicTransport, icmpType, localPort, remotePort, protocol;
      $newPortFilter.localPort = @($newPortFilter.localPort);
      $newPortFilter.remotePort = @($newPortFilter.remotePort);
      $newPortFilter.icmpType = @($newPortFilter.icmpType);
      $rule = $rulesHash[$_.CreationClassName];
      if ($rule -and $null -ne $newPortFilter) {
        $rule.portFilter = $newPortFilter
      }
    }
  }

  $outputRules = @()
  foreach ($requirement in $requirements) {
    $rule = $allRules | Where-Object { $_.Direction -eq $requirement.Direction -and $null -ne $_.portFilter -and $_.portFilter.localPort -eq $requirement.LocalPort -and $_.DisplayName -eq $requirement.Name }
    if ($null -eq $rule) {
      $outputRules += , @{ Exists = $false; Direction = $requirement.Direction; LocalPort = $requirement.LocalPort; Name = $requirement.Name }
    }
    else {
      $outputRules += , @{
        Exists    = $true;
        Direction = $requirement.Direction;
        LocalPort = $requirement.LocalPort;
        Name      = $requirement.Name;
        Data      = @{
          InstanceId  = $rule.InstanceId;
          DisplayName = $rule.DisplayName;
          Enabled     = $rule.Enabled;
          Status      = $rule.Status;
        }
      }
    }
  }

  return  Get-ResultAsArray -result $outputRules
}

function Get-VirtualSwitches {
  Param(
    [Parameter(Mandatory = $true)]
    [object]$requirement
  )

  $VirtualSwitches = @()

  Import-Module Hyper-V -ErrorAction SilentlyContinue
  $hyperVModule = Get-Module Hyper-V -ErrorAction SilentlyContinue -ErrorVariable +err
  if ($hyperVModule) {
    try {
      $result = Get-VMSwitch -SwitchType $requirement.types -ErrorAction Stop
    }
    catch {
      $result = @()
    }

    $VirtualSwitches = (Get-ResultAsArray -result $result).Value
  }

  return  Get-ResultAsArray -result $VirtualSwitches
}

function Get-PowerShellModule {
  Param(
    [Parameter(Mandatory = $true)]
    [object]$requirement
  )
  Import-Module -Name $requirement.Name -ErrorAction SilentlyContinue
  $InstalledVersions = Get-InstalledModule -Name $requirement.Name -AllVersions -ErrorAction SilentlyContinue `
  | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty Version `
  | ForEach-Object { [System.Version]$_ } | Sort-Object -Descending

  try {
    $AllAvailableVersions = Find-Module -Repository $requirement.Repository -Name $requirement.Name -AllVersions `
    | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty Version `
    | ForEach-Object { [System.Version]$_ } | Sort-Object -Descending

    $HighestVersion = $AllAvailableVersions[0]
    $AvailableSubVersions = $AllAvailableVersions.Where{ $_.Major -eq $requirement.majorVersion }
    $HighestSubVersion = $AvailableSubVersions[0]

    if ($null -eq $HighestSubVersion) {
      throw
    }
  }
  catch {
    # Unlikely to happen but could happen in cases where there's no internet or major version is not available(maybe during testing)
    throw "Couldn't find $($requirement.Name)(version $($requirement.majorVersion).*.*) in $($requirement.Repository)"
  }

  $NewerVersion = $null
  if ($HighestVersion -gt $HighestSubVersion) {
    [string]$NewerVersion = $HighestVersion
  }

  if ($null -eq $InstalledVersions) {
    return  @{
      installed    = $False
      version      = [string]$HighestSubVersion
      newerVersion = $NewerVersion
    }
  }

  $CurrentVersion = $InstalledVersions[0]
  $CorrectSubVersion = $null

  # Correct the version if we have a higher subversion or a different major version
  if (($HighestSubVersion -gt $CurrentVersion) -or ($HighestSubVersion.Major -ne $CurrentVersion.Major)) {
    [string]$CorrectSubVersion =$HighestSubVersion
  }
  return  @{
    installed         = $True
    version           = [string]$CurrentVersion
    newerVersion      = $NewerVersion
    correctSubVersion = $CorrectSubVersion;
  }
}

function Get-ServerMemory {
  # return the available memory in GB

  $availableMemory = Get-Counter -Counter "\Hyper-V Dynamic Memory Balancer(*)\Available Memory" -SampleInterval 1 -MaxSamples 1 -ErrorAction SilentlyContinue

  if ($null -eq $availableMemory) {
    return $availableMemory
  }

  $freeMemory = $availableMemory.CounterSamples.CookedValue / 1KB

  return $freeMemory
}

function Get-StorageVolumes {
  # return the available storage volumes with their free space in GB
  $FreeSpace = @{n = 'FreeSpace'; e = { [math]::Round($_.Free / 1GB) } };
  $isSystemDrive = @{n = 'IsSystemDrive'; e = { $env:SystemDrive -eq "$($_.Name):" } };
  $path = @{n = 'Path'; e = { $_.Root.TrimEnd('\') } }
  $adminShare = @{
    n = 'AdminShare';
    e = {
      @{
        Name    = "$($_.Name)$";
        UncPath = "\\$(Get-FQDN)\$($_.Name)$"
      }
    }
  }
  return Get-PSDrive -PSProvider 'FileSystem' | Microsoft.PowerShell.Utility\Select-Object Name, $path, $FreeSpace, $isSystemDrive, $adminShare | Where-Object { $_.FreeSpace -GT 0 };
}

function Get-FQDN {
  if ($env:userdnsdomain) {
    return ("$env:computername.$env:userdnsdomain").ToLower();
  }
  return $env:computername;
}

function Get-NumberOfLogicalProcessors() {
  return (
    Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object NumberOfLogicalProcessors
  ).NumberOfLogicalProcessors
}

function Get-OperatingSystem {
  return (Get-CimInstance -Class Win32_OperatingSystem -ErrorAction SilentlyContinue).caption
}

#### Main ####

$rolesAndFeatures = @()
$firewallRules = @()
$virtualSwitches = @()
$powershellModule = $null
$storageVolumes = @()
if ($null -ne $hostRequirements.Server) {
  if ($null -ne $hostRequirements.Server.RolesAndFeatures -and $hostRequirements.Server.RolesAndFeatures.length -ne 0) {
    $rolesAndFeatures = (Get-RolesAndFeatures -requirements $hostRequirements.Server.RolesAndFeatures).Value
  }

  if ($null -ne $hostRequirements.Server.FirewallRules -and $hostRequirements.Server.FirewallRules.length -ne 0) {
    $firewallRules = (Get-FirewallRules -requirements $hostRequirements.Server.FirewallRules).Value
  }

  if ($null -ne $hostRequirements.Server.VirtualSwitch -and $hostRequirements.Server.VirtualSwitch.Exists) {
    $virtualSwitches = (Get-VirtualSwitches -requirement $hostRequirements.Server.VirtualSwitch).Value
  }

  if ($null -ne $hostRequirements.Server.PowerShellModule) {
    $powershellModule = Get-PowerShellModule -requirement $hostRequirements.Server.PowerShellModule
  }
}

if ($null -ne $hostRequirements.StorageSpace -and $hostRequirements.StorageSpace.count -gt 0) {
  $storageVolumes = @(Get-StorageVolumes)
}

$freeMemory = Get-ServerMemory
$numberOfLogicalProcessors = Get-NumberOfLogicalProcessors
$operatingSystem = Get-OperatingSystem

return  @{
  RolesAndFeatures  = $rolesAndFeatures
  FirewallRules     = $firewallRules
  VirtualSwitches   = $virtualSwitches
  AvailableMemory   = $freeMemory
  StorageVolumes    = $storageVolumes
  LogicalProcessors = $numberOfLogicalProcessors
  PowerShellModule  = $powershellModule
  OperatingSystem   = $operatingSystem
}

}
## [END] Get-WACAKServerInventoryAks ##
function Get-WACAKSystemDrivePath {
<#
.SYNOPSIS
Gets the system drive path.

.DESCRIPTION
Gets the system drive path.

.EXAMPLE
./Get-SystemDrivePath.ps1
Returns C or D etc...

.ROLE
Readers

#>

$env:SystemDrive
}
## [END] Get-WACAKSystemDrivePath ##
function Get-WACAKUserProfilePath {
<#
.SYNOPSIS
Gets the user profile path.

.DESCRIPTION
Gets the user profile path.

.EXAMPLE
./Get-UserProfilePath.ps1
Returns C:\Users\testUser etc

.ROLE
Readers

#>

$env:UserProfile
}
## [END] Get-WACAKUserProfilePath ##
function Get-WACAKValidationTestList {
<#

.SYNOPSIS
Gets the list of validation tests run with Set-AksHciConfig

.DESCRIPTION
Gets the list of MOC, KVA, and DownloadSDK validation tests run with Set-AksHciConfig

.EXAMPLE
./Get-ValidationTestList.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Readers

#>
Import-Module AksHci -ErrorAction Stop

$testList = @()
# $downloadSDKTests = Test-DownloadSDKConfiguration -list
$mocTests = Test-MocConfiguration -list
$kvaTests = Test-KvaConfiguration -list
$testList += "Validate DownloadSDK Host Firewall URL Requirements"
# foreach ($test in $downloadSDKTests) {
#     $testList += $test.TestName
# }
foreach ($test in $mocTests) {
    $testList += $test.TestName
}
foreach ($test in $kvaTests) {
    $testList += $test.TestName
}
return $testList

}
## [END] Get-WACAKValidationTestList ##
function Install-WACAKAdAuth {
<#

.SYNOPSIS
Installs AD Authentication on the target cluster.

.DESCRIPTION
Installs AD Authentication on the target cluster.

.EXAMPLE
./Install-AdAuth.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $ClusterName,
    [Parameter(Mandatory = $true)]
    [string]
    $Spn,
    [Parameter(Mandatory = $true)]
    [string]
    $KeytabFileContents,
    [Parameter(Mandatory = $true)]
    [string]
    $Username,
    [Parameter(Mandatory = $true)]
    [boolean]
    $IsAdminGroup
)

Import-Module AksHci

$KeytabFilePath = $(New-Item -Path "$env:TEMP\current.keytab" -ItemType File -Force).FullName
[System.Convert]::FromBase64String(($KeytabFileContents)) | Set-Content -Path $KeytabFilePath -Encoding Byte
$SidRegex = '^([sS]-1-\d(-\d+)+)$'
$IsSidFormat = $Username -match $SidRegex

if ($IsSidFormat -and $IsAdminGroup) {
    Install-AksHciAdAuth -Name $ClusterName -keytab $KeytabFilePath -SPN $Spn -adminGroupSID $Username
} elseif ($IsSidFormat) {
    Install-AksHciAdAuth -Name $ClusterName -keytab $KeytabFilePath -SPN $Spn -adminUserSID $Username
} elseif ($IsAdminGroup) {
    Install-AksHciAdAuth -Name $ClusterName -keytab $KeytabFilePath -SPN $Spn -adminGroup $Username
} else {
    Install-AksHciAdAuth -Name $ClusterName -keytab $KeytabFilePath -SPN $Spn -adminUser $Username
}
Remove-Item -Path $KeytabFilePath -Recurse -Force

}
## [END] Install-WACAKAdAuth ##
function Install-WACAKAksHci {
<#

.SYNOPSIS
Setup management cluster on host.

.DESCRIPTION
Setup management cluster on host.

.EXAMPLE
./Install-AksHci.ps1.

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>
Param(
    [Parameter(Mandatory = $true)]
    [string]
    $LogSource
)

Import-Module AksHci

$VerbosePreference = "continue";
$LogName = "Microsoft-ServerManagementExperience";
New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
function writeInfoLog($logMessage) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
        -Message $logMessage  -ErrorAction SilentlyContinue
}

function writeSuccessLog($logMessage) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType SuccessAudit `
        -Message $logMessage  -ErrorAction SilentlyContinue
}

function writeErrorLog($errorMessage) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
        -Message $errorMessage -ErrorAction SilentlyContinue
}

$infoLog = "[Install-AksHci]: Installing"
writeInfoLog $infoLog

try {
    AksHci\Install-AksHci
}
catch {
    $err = $_.Exception.Message
    $errorLog = "[Install-AksHci]:" + $err
    writeErrorLog $errorLog
    throw $err
}

$successLog = "[Install-AksHci]: Done"
writeSuccessLog $successLog
}
## [END] Install-WACAKAksHci ##
function Install-WACAKAksHciModulePreview {
<#

.SYNOPSIS
Installs the preview version of AksHci powershell module.

.DESCRIPTION
Installs the preview version of AksHci powershell module.

.EXAMPLE
./Install-AksHciModulePreview.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $Pat
)

Unregister-PSRepository -Name "AksHciPSGalleryPreview" -ErrorAction:SilentlyContinue

$patToken = $Pat | ConvertTo-SecureString -AsPlainText -Force
$credsAzureDevopsServices = New-Object System.Management.Automation.PSCredential("test@foo.com", $patToken)
Register-PSRepository -Name "AksHciPSGalleryPreview" -SourceLocation "https://pkgs.dev.azure.com/msazure/msk8s/_packaging/PSGalleryPreview/nuget/v2" -PublishLocation "https://pkgs.dev.azure.com/msazure/msk8s/_packaging/PSGalleryPreview/nuget/v2" -InstallationPolicy Trusted -Credential $credsAzureDevopsServices

install-module -Name Az.Resources -RequiredVersion 3.2.0 -Repository PSGallery -Force
install-module -Name Az.Accounts -RequiredVersion 2.2.4 -Repository PSGallery -Force
install-module -Name AzureAD -RequiredVersion 2.0.2.128 -Repository PSGallery -Force

Install-Module -Repository AksHciPSGalleryPreview -Credential $credsAzureDevopsServices -Name DownloadSDK -AllowPrerelease
Install-Module -Repository AksHciPSGalleryPreview -Credential $credsAzureDevopsServices -Name Moc -AllowPrerelease
Install-Module -Repository AksHciPSGalleryPreview -Credential $credsAzureDevopsServices -Name Kva -AllowPrerelease
Install-Module -Repository AksHciPSGalleryPreview -Credential $credsAzureDevopsServices -Name AksHci -AllowPrerelease
}
## [END] Install-WACAKAksHciModulePreview ##
function Install-WACAKAksHciPrerequisites {
<#

.SYNOPSIS
Install the prerequisites needed before installing the AksHci powershell module.

.DESCRIPTION
Install the prerequisites needed before installing the AksHci powershell module.

.EXAMPLE
./Install-AksHciPrerequisites.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

$ResetExecutionPolicy = $False
$PreviousExecutionPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($PreviousExecutionPolicy -eq 'Undefined' -or $PreviousExecutionPolicy -eq 'Restricted') {
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    $ResetExecutionPolicy = $True
}

Install-PackageProvider -Name NuGet -Scope CurrentUser -Force
PowerShellGet\Install-Module -Name PowershellGet -Repository PSGallery -Force -Confirm:$false -SkipPublisherCheck

if ($ResetExecutionPolicy) {
    Set-ExecutionPolicy -ExecutionPolicy $PreviousExecutionPolicy -Scope CurrentUser -Force
}

}
## [END] Install-WACAKAksHciPrerequisites ##
function Install-WACAKArcHciModule {
<#

.SYNOPSIS
Installs the powershell module specified in $moduleName from $repositoryName at version $moduleVersion

.DESCRIPTION
Installs the powershell module specified in $moduleName from $repositoryName at version $moduleVersion

.EXAMPLE
./Install-ArcHciModule.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $repositoryName,
    [Parameter(Mandatory = $true)]
    [string]
    $moduleName
)

PowerShellGet\Install-Module -Name $moduleName -Force -Confirm:$false -SkipPublisherCheck -AcceptLicense


}
## [END] Install-WACAKArcHciModule ##
function Install-WACAKAzCliAndPowerShellGet {
<#

.SYNOPSIS
Install az cli if not already installed and install the latest PowerShellGet version

.DESCRIPTION
Install az cli if not already installed and install the latest PowerShellGet version

.ROLE
Administrators

#>
$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine')


$installedAzCli = $false
# If the az cli is not installed, the following commands will silently install it.
# If the az cli is already installed, the same command is used to silently upgrade it to the latest version.
# There is currently no way to use the az upgrade command without prompting the user to click through the interactive dialog.

New-Item -Path $env:Temp -Name 'install-azcli' -ItemType 'Directory' -Force
$scriptDir = $env:Temp + "\install-azcli"

$azCliOutfile = $($env:Temp) + "\install-azcli\AzureCLI.msi"
$logPath = $env:Temp + "\install-azcli\out.txt"
Invoke-WebRequest -Uri https://aka.ms/installazurecliwindows -OutFile $azCliOutfile
Start-Process -Wait -NoNewWindow -FilePath  "msiexec.exe" -ArgumentList "/i ""$($azCliOutfile)"" /qn /L*v ""$($logPath)"""

$originalPath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
$pathAddition = $originalPath + ";C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin"
$newpath = [Environment]::SetEnvironmentVariable("Path", $pathAddition, 'Machine')

# Sets the path for the current session, does not require reboot to show up
$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine')
Remove-Item $azCliOutfile

$installedAzCli = $true

$powershellGet = Get-Module PowershellGet -errorAction SilentlyContinue
if ($null -eq $powershellGet -or ($powershellGet -and $powershellGet.Version -and $powershellGet.Version.Major -lt 2)) {
    Install-PackageProvider -Name NuGet -Force
    PowerShellGet\Install-Module -Name PowerShellGet -Force -Confirm:$false -SkipPublisherCheck
}

# We return whether we installed az cli for the reboot
return $installedAzCli
}
## [END] Install-WACAKAzCliAndPowerShellGet ##
function Install-WACAKModule {
<#

.SYNOPSIS
Installs the powershell module specified in $moduleName from $repositoryName at version $moduleVersion

.DESCRIPTION
Installs the powershell module specified in $moduleName from $repositoryName at version $moduleVersion

.EXAMPLE
./Install-Module.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $repositoryName,
    [Parameter(Mandatory = $true)]
    [string]
    $moduleName,
    [Parameter(Mandatory = $true)]
    [string]
    $moduleVersion
)

PowerShellGet\Install-Module -Name $moduleName -RequiredVersion $moduleVersion -Repository $repositoryName -AcceptLicense -Force

}
## [END] Install-WACAKModule ##
function Install-WACAKResourceBridge {
<#

.SYNOPSIS
Setup Resource Bridge.

.DESCRIPTION
Setup Resource Bridge.

.EXAMPLE
./Install-ResourceBridge.ps1.

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>
Param(
    [parameter(Mandatory = $true)] [string]$LogSource,

    [parameter(Mandatory = $true)] [string] $resourceGroup,
    [parameter(Mandatory = $true)] [string] $subscription,

    # Name of the resource bridge
    [parameter(Mandatory = $true)] [string] $resourceBridgeName,

    [parameter(Mandatory = $false)] [string] $customLocationName = ($resourceBridgeName + "-cl"),
    [parameter(Mandatory = $true)] [string] $controlPlaneIP = "192.168.0.200",
    [AllowEmptyString()][parameter(Mandatory = $true)] [string] $cloudServiceIP,
    [parameter(Mandatory = $true)] [string] $vnetName = "ArcHciVnet",
    [parameter(Mandatory = $true)] [string] $vswitchName = "ComputeSwitch",
    [parameter(Mandatory = $true)] [string] $vnetType = "External",

    # Static IP parameters
    [Parameter(Mandatory = $true)][bool] $isDHCPSetup,
    [AllowNull()][AllowEmptyString()][Parameter(Mandatory = $true)][string] $IpAddressPrefix,
    [AllowNull()][AllowEmptyString()][Parameter(Mandatory = $true)][string] $Gateway,
    [AllowNull()][AllowEmptyString()][AllowEmptyCollection()][Parameter(Mandatory = $true)][string[]] $DnsServers,
    [AllowNull()][AllowEmptyString()][Parameter(Mandatory = $true)][string] $ArcApplianceVmIPStart,
    [AllowNull()][AllowEmptyString()][Parameter(Mandatory = $true)][string] $ArcApplianceVmIPEnd,

    [parameter(Mandatory = $true)] [string] $location = "eastus",
    [parameter(Mandatory = $true)] [string] $imageDir,
    [parameter(Mandatory = $true)] [string] $cloudConfigLocation,
    [parameter(Mandatory = $true)] [string] $workingDir,
    [parameter(Mandatory = $true)] [string] $vlanID,
    [parameter(Mandatory = $false)] [string] $operatorName = "vmss-hci",
    [parameter(Mandatory = $false)] [string] $operatorNameAKShybrid = "aks-hybrid-vmss-hci",
    [parameter(Mandatory = $true)] [string] $catalogName,
    [parameter(Mandatory = $true)] [string] $ringName,
    [AllowEmptyString()] [parameter(Mandatory = $true)] [string] $mocConfigVersion
)

Set-StrictMode -Version 5.0

$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine')

# Logging functions
$LogName = "Microsoft-ServerManagementExperience";
New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
function writeInfoLog($logMessage) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
        -Message $logMessage  -ErrorAction SilentlyContinue
}

function writeSuccessLog($logMessage) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType SuccessAudit `
        -Message $logMessage  -ErrorAction SilentlyContinue
}

function writeErrorLog($errorMessage) {
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Error `
        -Message $errorMessage -ErrorAction SilentlyContinue
}

# Create our cli path
$global:mocMetadataDirectory = $($env:USERPROFILE + "\.wssd")
$global:accessFileLocation = [io.Path]::Combine($global:mocMetadataDirectory, "cloudconfig")
$workDirectory = $workingDir

if (-not(Test-Path $workDirectory)) {
    New-Item -ItemType Directory -Path $workDirectory | Out-Null
  }

Import-Module -Name ArcHci
Import-Module -Name Moc

function Invoke-AzCommandLine
{
    <#
    .DESCRIPTION
        Executes a command and optionally ignores errors.

    .PARAMETER command
        Comamnd to execute.

    .PARAMETER arguments
        Arguments to pass to the command.

    .PARAMETER ignoreError
        Optionally, ignore errors from the command (don't throw).

    .PARAMETER showOutput
        Optionally, show live output from the executing command.

    .PARAMETER showOutputAsProgress
        Optionally, show output from the executing command as progress bar updates.

    .PARAMETER progressActivity
        The activity name to display when showOutputAsProgress was requested.
    #>

    param (
        [String]$command,
        [String]$arguments,
        [Switch]$ignoreError,
        [Switch]$showOutput,
        [Switch]$showOutputAsProgress,
        [String]$progressActivity
    )

    try {
		$previousErrorAction = $errorActionPreference
		$errorActionPreference = "Continue"

        if ($showOutputAsProgress.IsPresent)
        {
            $errorResult = $($result = (& $command $arguments.Split(" ")  | ForEach-Object { $status = $_ -replace "`t"," - " })) 2>&1
        }
        elseif ($showOutput.IsPresent)
        {
            $errorResult = $($result = (& $command $arguments.Split(" ") | ForEach-Object { Write-Information "$_" -InformationAction Continue; return $_ })) 2>&1
        }
        else
        {
            $errorResult = $($result = (& $command $arguments.Split(" "))) 2>&1
        }
        $previousExitCode = $LASTEXITCODE
    }
    catch {
        if ($ignoreError.IsPresent)
        {
            return
        }
        throw
    }
    finally {
		$errorActionPreference = $previousErrorAction
    }

    if ($null -ne $errorResult -and -not ($errorResult.Exception.Message -match "Please let us know how we are doing") `
           -and -not ($errorResult.Exception.Message -match "The installed extension '.+' is experimental") `
           -and -not ($errorResult.Exception.Message -match "The installed extension '.+' is in preview.") `
           -and -not ($errorResult.Exception.Message -match "Setting GA feature gate arcmonitoring=true") `
           -and -not ($errorResult.Exception.Message -match "Command group '.+' is in preview and under development."))
    {
        $stack = Get-PSCallStack
        #An error message was returned, just throw that message
        $errMessage = "$command $arguments returned a non empty error stream [$($errorResult.ToString())] at [$($stack)]"
        throw $errMessage
    }

    $out = $result | Where-Object {$_.gettype().Name -ne "ErrorRecord"}  # On a non-zero exit code, this may contain the error

    if ($previousExitCode)
    {

        $errMessage = "$command $arguments returned a non zero exit code $previousExitCode [$errorResult]"

        if ($ignoreError.IsPresent)
        {
            $ignoreMessage = "[IGNORED ERROR] $errMessage"
            return
        }
        throw $errMessage
    }
    return $out
}

function Invoke-AzCommand {
    <#
    .DESCRIPTION
        Executes an az cli command.
    .PARAMETER arguments
        Arguments to pass to az cli.
    .PARAMETER ignoreError
        Optionally, ignore errors from the command (don't throw).
    .OUTPUTS
        N/A
    .EXAMPLE
        Invoke-AzCommand -arguments "provider show --namespace 'Microsoft.Resources'"
    #>
    param (
        [Parameter(Mandatory=$true)]
        [String]$arguments,

        [Parameter(Mandatory=$false)]
        [Switch]$ignoreError
    )

    $azCliFullPath = (Get-Command "az").Source
    if (-not $azCliFullPath) {
        throw $("Unable to find the `"az`" command ")
    }
    $response = Invoke-AzCommandLine -Command $azCliFullPath -Arguments $arguments -ignoreError:$ignoreError -showOutput

    return $response
}

function Get-AksHciSetupConfig {
    try {
        return Get-AksHciConfig;
    }
    catch {
        return $null
    }
}

function Get-MocSetupConfig {
    try {
        return Get-MocConfig;
    }
    catch {
        return $null
    }
}

function Install-MocPrerequisites {
    <#
    .DESCRIPTION
        Install MOC pre-requisites.
    .PARAMETER cloudServiceIP
        IP address of the cloud agent (optional parameter)
    .PARAMETER workingDir
        Path to the working directory (optional parameter)
    .PARAMETER imageDir
        Path to the folder the VM files should be created in (optional parameter)
    .PARAMETER cloudConfigLocation
        Path to the folder the cloud agent config files should be created in (optional parameter)
    .OUTPUTS
        N/A
    .EXAMPLE
        Install-MocPrerequisites -workingDir "C:\workingDir" -imageDir "C:\imageDir" -cloudConfigLocation "C:\cloudConfig" -catalog "catalog" -ring "ring"
    #>
    Param(
        [AllowEmptyString()]
        [Parameter(Mandatory=$true)]
        [string] $cloudServiceIP,

        [Parameter(Mandatory=$true)]
        [String] $workingDir,

        [Parameter(Mandatory=$true)]
        [String] $imageDir,

        [Parameter(Mandatory=$true)]
        [String] $cloudConfigLocation,

        [Parameter(Mandatory=$true)]
        [String] $catalogName,

        [Parameter(Mandatory=$true)]
        [String] $ringName
    )

    $infoLog = "Install-ResourceBridge: Initializing Moc node"
    writeInfoLog $infoLog

    Initialize-MocNode

    $infoLog = "Install-ResourceBridge: Setting Moc config"
    writeInfoLog $infoLog

    if ([string]::IsNullOrEmpty($cloudServiceIP)) {
        Set-MocConfig -workingDir $workingDir -imageDir $imageDir -cloudConfigLocation $cloudConfigLocation -catalog $catalogName -ring $ringName -Version $mocConfigVersion -createAutoConfigContainers $false -skipHostLimitChecks
    } else {
        Set-MocConfig -workingDir $workingDir -imageDir $imageDir -cloudConfigLocation $cloudConfigLocation -catalog $catalogName -ring $ringName -CloudServiceIp $cloudServiceIP -Version $mocConfigVersion -createAutoConfigContainers $false -skipHostLimitChecks
    }

    $infoLog = "Install-ResourceBridge: Installing Moc"
    writeInfoLog $infoLog

    # Install-Moc can sometimes give errors that don't break the installation, will keep it as info log until this issue is fixed
    Install-Moc
}

function Deploy-ArcAppliance {
    Param (
        [Parameter(Mandatory=$true)]
        [String] $workDirectory,

        [Parameter(Mandatory=$true)]
        [String] $resourceGroup,

        [Parameter(Mandatory=$true)]
        [String] $resourceBridgeName,

        [Parameter(Mandatory=$false)]
        [String] $sleepDuration = 60
    )

    Invoke-AzCommand "arcappliance validate hci --config-file ""$($workDirectory)\hci-appliance.yaml"""
    Invoke-AzCommand "arcappliance prepare hci --config-file ""$($workDirectory)\hci-appliance.yaml"""
    Invoke-AzCommand "arcappliance deploy hci --config-file ""$($workDirectory)\hci-appliance.yaml"" --outfile ""$workDirectory\kubeconfig"""
    Invoke-AzCommand "arcappliance create hci --config-file ""$($workDirectory)\hci-appliance.yaml"" --kubeconfig ""$workDirectory\kubeconfig"""

    $infoLog = "Install-ResourceBridge: Finished running az arcappliance validate + prepare + deploy + create, waiting to be in Succeeded state..."
    writeInfoLog $infoLog

    $timer = [system.diagnostics.stopwatch]::StartNew()
    while ($true)
    {
        $res = Invoke-AzCommand "arcappliance show --resource-group ""$($resourceGroup)"" --name ""$($resourceBridgeName)""" -ignoreError
        $res = $res | ConvertFrom-Json
        try {
            if ($res.status -eq "Running") {
                # Arc appliance has been seen to transition from a Running state back to a Configuring state in some cases
                # Adding a 2 min sleep here to make sure that the state becomes stable before moving to the next step
                Start-Sleep $sleepDuration * 2
                break
            }

            if ($res.status -eq "Failed") {
                throw "arcappliance deployment Failed."
            }
        } catch {
            # Do nothing, since we may have to catch errors where $res.status is null before the arcappliance is
            # successfully created. If the arcappliance never gets created, then the timer will time out.
        }

        if ($timer.Elapsed.TotalSeconds -gt 30 * 60) {
            # 30 minutes have passed
            throw "arcappliance deployment did not complete in time."
        }

        Start-Sleep $sleepDuration
    }
}

function Create-K8sExtensionForVMs {
    Param (
        [Parameter(Mandatory=$true)]
        [String] $workingDir,

        [Parameter(Mandatory=$true)]
        [String] $operatorName,

        [Parameter(Mandatory=$true)]
        [String] $resourceGroup,

        [Parameter(Mandatory=$true)]
        [String] $resourceBridgeName,

        [Parameter(Mandatory=$false)]
        [String] $sleepDuration = 60
    )

    $hciClusterId = (Get-AzureStackHci).AzureResourceUri

    Invoke-AzCommand "k8s-extension create --cluster-type appliances --cluster-name ""$($resourceBridgeName)"" --resource-group ""$($resourceGroup)"" --name ""$($operatorName)"" --extension-type Microsoft.AZStackHCI.Operator --scope cluster --release-namespace helm-operator2 --configuration-settings Microsoft.CustomLocation.ServiceAccount=hci-operators --config-protected-file ""$($workingDir)\hci-config.json"" --configuration-settings HCIClusterID=$hciClusterId --auto-upgrade true"
    $infoLog = "Install-ResourceBridge: Finished running az k8s-extension create for VMs, waiting to be in Installed state..."
    writeInfoLog $infoLog

    $timer = [system.diagnostics.stopwatch]::StartNew()
    while ($true)
    {
        $res = Invoke-AzCommand "k8s-extension show --cluster-type appliances --cluster-name ""$($resourceBridgeName)"" --resource-group ""$($resourceGroup)"" --name ""$($operatorName)""" -ignoreError
        $res = $res | ConvertFrom-Json
        try {
            if ($res.provisioningState -eq "Succeeded") {
                break
            }

            if ($res.provisioningState -eq "Failed") {
                throw "k8s-extension creation for VMs Failed."
            }
        } catch {
            # Do nothing, since we may have to catch errors where $res.provisioningState is null before the k8s-extension
            # is successfully created. If the k8s-extension never gets created, then the timer will time out.
        }

        if ($timer.Elapsed.TotalSeconds -gt 5 * 60) {
            # 5 minutes have passed
            throw "k8s-extension creation for VMs did not complete in time."
        }

        Start-Sleep $sleepDuration
    }
}

function Create-K8sExtensionForAKS {
    Param (
        [Parameter(Mandatory=$true)]
        [String] $workingDir,

        [Parameter(Mandatory=$true)]
        [String] $operatorNameAKShybrid,

        [Parameter(Mandatory=$true)]
        [String] $resourceGroup,

        [Parameter(Mandatory=$true)]
        [String] $resourceBridgeName,

        [Parameter(Mandatory=$false)]
        [String] $sleepDuration = 60
    )

    Invoke-AzCommand "k8s-extension create --cluster-type appliances --cluster-name ""$($resourceBridgeName)"" --resource-group ""$($resourceGroup)"" --name ""$($operatorNameAKShybrid)"" --extension-type Microsoft.HybridAKSOperator --config Microsoft.CustomLocation.ServiceAccount=""default"""
    $infoLog = "Install-ResourceBridge: Finished running az k8s-extension create for AKS hybrid clusters, waiting to be in Installed state..."
    writeInfoLog $infoLog

    $timer = [system.diagnostics.stopwatch]::StartNew()
    while ($true)
    {
        $res = Invoke-AzCommand "k8s-extension show --cluster-type appliances --cluster-name ""$($resourceBridgeName)"" --resource-group ""$($resourceGroup)"" --name ""$($operatorNameAKShybrid)""" -ignoreError
        $res = $res | ConvertFrom-Json
        try {
            if ($res.provisioningState -eq "Succeeded") {
                break
            }

            if ($res.provisioningState -eq "Failed") {
                throw "k8s-extension creation for AKS hybrid clusters Failed."
            }
        } catch {
            # Do nothing, since we may have to catch errors where $res.provisioningState is null before the k8s-extension
            # is successfully created. If the k8s-extension never gets created, then the timer will time out.
        }

        if ($timer.Elapsed.TotalSeconds -gt 5 * 60) {
            # 5 minutes have passed
            throw "k8s-extension creation for AKS hybrid clusters did not complete in time."
        }

        Start-Sleep $sleepDuration
    }
}

function Create-CustomLocation {
    Param (
        [Parameter(Mandatory=$true)]
        [String] $resourceGroup,

        [Parameter(Mandatory=$true)]
        [String] $resourceBridgeName,

        [Parameter(Mandatory=$true)]
        [String] $customLocationName,

        [Parameter(Mandatory=$true)]
        [String] $operatorName,

        [Parameter(Mandatory=$true)]
        [String] $location
    )
    $applianceID = Invoke-AzCommand "arcappliance show -g ""$($resourceGroup)"" --name ""$($resourceBridgeName)"" --query id -o tsv" -ignoreError
    $extensionID = Invoke-AzCommand "k8s-extension show --cluster-type appliances -c ""$($resourceBridgeName)"" -g ""$($resourceGroup)"" --name ""$($operatorName)"" --query id -o tsv" -ignoreError
    $customLocNamespace = "hci-operators"
    Invoke-AzCommand "customlocation create --resource-group ""$($resourceGroup)"" --name ""$($customLocationName)"" --location ""$($location)"" --cluster-extension-ids ""$($extensionID)"" --namespace ""$($customLocNamespace)"" --host-resource-id ""$($applianceID)"""
    $infoLog = "Install-ResourceBridge: Finished running az customlocation create, waiting to be in Succeeded state..."
    writeInfoLog $infoLog
}

try {
    $AksHciConfig = Get-AksHciSetupConfig
    $MocConfig = Get-MocSetupConfig
    if (($null -eq $AksHciConfig) -and ($null -eq $MocConfig)) {
        $infoLog = "Install-ResourceBridge: Setting up and installing Moc"
        writeInfoLog $infoLog

        Install-MocPrerequisites -cloudServiceIP $cloudServiceIP -workingDir $workingDir -imageDir $imageDir `
        -cloudConfigLocation $cloudConfigLocation -catalogName $catalogName -ringName $ringName

        $infoLog = "Install-ResourceBridge: Finished setting up Moc"
        writeInfoLog $infoLog
    }

    $infoLog = "Install-ResourceBridge: Creating config files"
    writeInfoLog $infoLog

    # vnetName can be any string that fits the regex [a-zA-Z-]+
    $vnetName = "ArcHciVnet"
    if ($isDHCPSetup) {
        New-ArcHciConfigFiles -subscriptionID $subscription -location $location -resourceGroup $resourceGroup `
        -resourceName $resourceBridgeName -workDirectory $workDirectory -controlPlaneIP $controlPlaneIP `
        -vipPoolStart $controlPlaneIP -vipPoolEnd $controlPlaneIP -vswitchName $vswitchName -vlanID $vlanID -vnetName $vnetName
    } else {
        New-ArcHciConfigFiles -subscriptionID $subscription -location $location -resourceGroup $resourceGroup `
        -resourceName $resourceBridgeName -workDirectory $workDirectory -controlPlaneIP $controlPlaneIP `
        -vipPoolStart $controlPlaneIP -vipPoolEnd $controlPlaneIP -k8snodeippoolstart $ArcApplianceVmIPStart `
        -k8snodeippoolend $ArcApplianceVmIPEnd -gateway $Gateway -dnsservers $DnsServers `
        -ipaddressprefix $IpAddressPrefix -vswitchName $vswitchName -vlanID $vlanID -vnetName $vnetName
    }

    $infoLog = "Install-ResourceBridge: Finished creating config files"
    writeInfoLog $infoLog

    $infoLog = "Install-ResourceBridge: Deploying arc appliance"
    writeInfoLog $infoLog

    Invoke-AzCommand "account set -s ""$subscription"""

    Deploy-ArcAppliance $workDirectory $resourceGroup $resourceBridgeName

    $infoLog = "Install-ResourceBridge: Finished deploying arc appliance"
    writeInfoLog $infoLog

    $infoLog = "Install-ResourceBridge: Creating k8s-extension for VMs"
    writeInfoLog $infoLog

    Create-K8sExtensionForVMs $workDirectory $operatorName $resourceGroup $resourceBridgeName

    $infoLog = "Install-ResourceBridge: Finished creating k8s-extension for VMs"
    writeInfoLog $infoLog

    $infoLog = "Install-ResourceBridge: Creating k8s-extension for AKS hybrid clusters"
    writeInfoLog $infoLog

    Create-K8sExtensionForAKS $workDirectory $operatorNameAKShybrid $resourceGroup $resourceBridgeName

    $infoLog = "Install-ResourceBridge: Finished creating k8s-extension for AKS hybrid clusters"
    writeInfoLog $infoLog

    $infoLog = "Install-ResourceBridge: Creating custom location"
    writeInfoLog $infoLog

    Create-CustomLocation $resourceGroup $resourceBridgeName $customLocationName $operatorName $location

    $infoLog = "Install-ResourceBridge: Finished creating custom location"
    writeInfoLog $infoLog
} catch {
    $err = $_.Exception.Message
    $errorLog = "Install-ResourceBridge: " + $err + " Please check your system and retry deployment."
    writeErrorLog $errorLog
    throw $err
}

$successLog = "Install-ResourceBridge: Done"
writeSuccessLog $successLog

}
## [END] Install-WACAKResourceBridge ##
function Invoke-WACAKAzCliLogin {
<#

.SYNOPSIS
Creates new volume using SDDC Management methods

.DESCRIPTION
Creates new volume using SDDC Management methods

.ROLE
Administrators

#>

Param(
    [Parameter(Mandatory = $true)] [string] $appId,
    [Parameter(Mandatory = $true)] [string] $password,
    [Parameter(Mandatory = $true)] [string] $tenant
)

$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine')

az login --service-principal -u $appId -p $password --tenant $tenant
}
## [END] Invoke-WACAKAzCliLogin ##
function New-WACAKAksHciCluster {
<#

.SYNOPSIS
Creates an Aks HCI cluster through the AksHci module.

.DESCRIPTION
Creates an Aks HCI cluster through the AksHci module.

.EXAMPLE
./New-AksHciCluster.ps1

.NOTES
The supported operating systems are Azure Stack HCI, Windows Server 2019, and Windows Server 2016.

.ROLE
Administrators

#>
Param(
    [Parameter(Mandatory=$true)]
    [string]
    $ClusterName,
    [Parameter(Mandatory=$true)]
    [string]
    $K8sVersion,
    [Parameter(Mandatory=$true)]
    [int]
    $ControlPlaneReplicaCount,
    [Parameter(Mandatory=$true)]
    [string]
    $ControlPlaneVmSize,
    [Parameter(Mandatory=$true)]
    [string]
    $LoadBalancerVmSize,
    [Parameter(Mandatory=$true)]
    [string]
    $NodePoolName,
    [Parameter(Mandatory=$true)]
    [int]
    $NodeCount,
    [Parameter(Mandatory=$true)]
    [string]
    $NodeVmSize,
    [Parameter(Mandatory=$true)]
    [string]
    $OSType,
    [Parameter(Mandatory=$true)]
    [int]
    $NodeMaxPodCount,
    [AllowNull()]
    [AllowEmptyCollection()]
    [Parameter(Mandatory=$true)]
    [string[]]
    $Taints,
    [Parameter(Mandatory=$true)]
    [bool]
    $EnableADAuth,
    [Parameter(Mandatory=$true)]
    [Object]
    $Vnet,
    [Parameter(Mandatory=$true)]
    [string]
    $NetworkConfigurationType
)

Import-Module AksHci

# We need to convert the vnet to a valid network object
$VnetSettingParams = @{}
$Vnet.psobject.properties | ForEach-Object {
    $VnetSettingParams[$_.Name] = $_.Value
}
$VnetObject = AksHci\New-AksHciNetworkSetting @VnetSettingParams

AksHci\New-AksHciCluster -Name $ClusterName `
-kubernetesVersion $K8sVersion `
-controlPlaneNodeCount $ControlPlaneReplicaCount `
-controlPlaneVmSize $ControlPlaneVmSize `
-loadBalancerVmSize $LoadBalancerVmSize `
-nodePoolName $NodePoolName `
-nodeCount $NodeCount `
-nodeVmSize $NodeVmSize `
-osType $OSType `
-nodeMaxPodCount $NodeMaxPodCount `
-taints $Taints `
-enableADAuth:$EnableADAuth `
-vnet $VnetObject `
-primaryNetworkPlugin $NetworkConfigurationType
}
## [END] New-WACAKAksHciCluster ##
function New-WACAKAksHciClusterNetwork {
<#
.SYNOPSIS
Creates a new virtual network on the AKS-HCI host.

.DESCRIPTION
Creates a new virtual network on the AKS-HCI host.

.EXAMPLE
./New-AksHciClusterNetwork.ps1
Returns the virtual network that was just created.

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators
#>

Param(
    [Parameter(Mandatory = $true)]
    [bool]
    $isDHCPSetup,
    [Parameter(Mandatory = $true)]
    [string]
    $VnetName,
    [Parameter(Mandatory = $true)]
    [string]
    $VswitchName,
    [AllowNull()]
    [AllowEmptyString()]
    [Parameter(Mandatory = $true)]
    [string]
    $IpAddressPrefix,
    [AllowNull()]
    [AllowEmptyString()]
    [Parameter(Mandatory = $true)]
    [string]
    $Gateway,
    [AllowNull()]
    [AllowEmptyString()]
    [AllowEmptyCollection()]
    [Parameter(Mandatory = $true)]
    [string[]]
    $DnsServers,
    [AllowNull()]
    [AllowEmptyString()]
    [Parameter(Mandatory = $true)]
    [string]
    $K8snodeIPPoolStart,
    [AllowNull()]
    [AllowEmptyString()]
    [Parameter(Mandatory = $true)]
    [string]
    $K8snodeIPPoolEnd,
    [Parameter(Mandatory = $true)]
    [string]
    $VipPoolStart,
    [Parameter(Mandatory = $true)]
    [string]
    $VipPoolEnd,
    [Parameter(Mandatory = $true)]
    [int]
    $Vlanid
)

Import-Module AksHci

if ($isDHCPSetup) {
    New-AksHciClusterNetwork -name $VnetName -vswitchName $VswitchName `
    -vlanID $Vlanid -vippoolstart $VipPoolStart -vippoolend $VipPoolEnd
} else {
    New-AksHciClusterNetwork -name $VnetName -vswitchName $VswitchName `
    -vlanID $Vlanid -ipaddressprefix $IpAddressPrefix -gateway $Gateway `
    -dnsservers $DnsServers -vippoolstart $VipPoolStart -vippoolend $VipPoolEnd `
    -k8snodeippoolstart $K8snodeIPPoolStart -k8snodeippoolend $K8snodeIPPoolEnd
}
}
## [END] New-WACAKAksHciClusterNetwork ##
function New-WACAKAksHciNodePool {
<#

.SYNOPSIS
Creates a node pool for AKS on Azure Stack HCI.

.DESCRIPTION
Creates a node pool for AKS on Azure Stack HCI.

.EXAMPLE
./New-AksHciNodePool.ps1

.NOTES
The supported operating systems are Azure Stack HCI, Windows Server 2019, and Windows Server 2016.

.ROLE
Administrators

#>
Param(
    [Parameter(Mandatory=$true)]
    [string]
    $ClusterName,
    [Parameter(Mandatory=$true)]
    [string]
    $NodePoolName,
    [Parameter(Mandatory=$true)]
    [int]
    $NodeCount,
    [Parameter(Mandatory=$true)]
    [string]
    $OSType,
    [Parameter(Mandatory=$true)]
    [string]
    $VmSize,
    [Parameter(Mandatory=$true)]
    [int]
    $NodeMaxPodCount,
    [AllowNull()]
    [AllowEmptyCollection()]
    [Parameter(Mandatory=$true)]
    [string[]]
    $Taints
)

Import-Module AksHci

AksHci\New-AksHciNodePool -clusterName $ClusterName -name $NodePoolName -count $NodeCount -osType $OSType -vmSize $VmSize -maxPodCount $NodeMaxPodCount -taints $Taints
}
## [END] New-WACAKAksHciNodePool ##
function Remove-WACAKAksHciCluster {
<#
.SYNOPSIS
Removes/Deletes given AKS HCI clusters
.DESCRIPTION
Removes/Deletes given AKS HCI clusters
.EXAMPLE
./Remove-AksHciCluster.ps1
.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.
.ROLE
Administrators
#>
Param(
    [Parameter(Mandatory = $true)]
    [string[]]
    $ClusterNames
)

Import-Module AksHci

foreach ($ClusterName in $ClusterNames) {
    Remove-AksHciCluster -name $ClusterName -Confirm:$false
}

}
## [END] Remove-WACAKAksHciCluster ##
function Remove-WACAKAksHciClusterNetwork {
<#
.SYNOPSIS
Removes the virtual network from the AKS-HCI host.

.DESCRIPTION
Removes the virtual network from the AKS-HCI host.

.EXAMPLE
./Remove-AksHciClusterNetwork.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators
#>

Param(
    [Parameter(Mandatory = $true)]
    [string]
    $VnetName
)

Import-Module AksHci

Remove-AksHciClusterNetwork -name $VnetName

}
## [END] Remove-WACAKAksHciClusterNetwork ##
function Remove-WACAKAksHciModule {
<#

.SYNOPSIS
Removes the AksHci powershell module.

.DESCRIPTION
Removes the AksHci powershell module and all of the dependent modules.

.EXAMPLE
./Remove-AksHciModule.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Uninstall-Module -Name AksHci -AllVersions -Force -ErrorAction:SilentlyContinue
Uninstall-Module -Name Kva -AllVersions -Force -ErrorAction:SilentlyContinue
Uninstall-Module -Name Moc -AllVersions -Force -ErrorAction:SilentlyContinue
Uninstall-Module -Name MSK8SDownloadAgent -AllVersions -Force -ErrorAction:SilentlyContinue
Uninstall-Module -Name DownloadSdk -AllVersions -Force -ErrorAction:SilentlyContinue
Unregister-PSRepository -Name WSSDRepo -ErrorAction:SilentlyContinue
Unregister-PSRepository -Name AksHciPSGallery -ErrorAction:SilentlyContinue
Unregister-PSRepository -Name AksHciPSGalleryPreview -ErrorAction:SilentlyContinue
}
## [END] Remove-WACAKAksHciModule ##
function Remove-WACAKAzCliExtensions {
<#

.SYNOPSIS
Removes a list of Az Cli extensions

.DESCRIPTION
Removes a list of Az Cli extensions

.ROLE
Administrators

#>
Param(
    [parameter(Mandatory = $true)] [string[]] $extensions
)

$env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine')

# If no account is available will send back null -> will then parse to UI that needs az login
$value = az account show

# Az account is logged in and valid
if ($null -ne $value) {
    foreach ($extension in $extensions) {
        az extension remove --name $extension 2>$null
    }
}
}
## [END] Remove-WACAKAzCliExtensions ##
function Remove-WACAKOldResourceBridgeFolders {
<#

.SYNOPSIS
Removes stale folders related to Resource Bridge deployment from the machine

.DESCRIPTION
Removes the .wssd\python and kva\.ssh folders from the machine to prevent corrupted Resource Bridge deployments

.EXAMPLE
./Remove-OldResourceBridgeFolders.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>

Remove-Item -Path "$env:USERPROFILE\.wssd\python" -Recurse -Force 2>$null
Remove-Item -Path "C:\ProgramData\kva\.ssh" -Recurse -Force 2>$null
}
## [END] Remove-WACAKOldResourceBridgeFolders ##
function Resolve-WACAKFilePath {
<#

.SYNOPSIS
Resolves and expands a file path.

.DESCRIPTION
Resolves and expands a file path

.EXAMPLE
./Resolve-FilePath.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Readers

#>

Param(
    [Parameter(Mandatory=$true)]
    [string]
    $FilePath
)

$ExpandedPath = $ExecutionContext.InvokeCommand.ExpandString($FilePath)
return $ExpandedPath
}
## [END] Resolve-WACAKFilePath ##
function Set-WACAKAksHciConfiguration {
<#

.SYNOPSIS
Sets up configuration for AKS HCI platform

.DESCRIPTION
Sets up configuration for AKS HCI platform

.EXAMPLE
./Set-AksHciConfiguration.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>
Param(
    [Parameter(Mandatory = $true)]
    [Object]
    $NetworkSettings,
    [Parameter(Mandatory = $true)]
    [Object]
    $AksHciConfigurations,
    [Parameter(Mandatory = $false)]
    [Object]
    $ProxySettings
)

Import-Module AksHci -ErrorAction Stop

$LogName = "Microsoft-ServerManagementExperience";
$LogSource = "Azure Kubernetes Service hybrid deployment Set-AksHciConfig";

if ([System.Diagnostics.EventLog]::SourceExists($LogSource) -eq $False) {
    New-EventLog -LogName $LogName -Source $LogSource
}

function HandleTestOutput() {
    Param(
        [AllowEmptyString()]
        [Parameter(Mandatory = $true)]
        [String]
        $Line
    )

    if ($Line -match 'Test \([0-9]+ of [0-9]+\): "([^"]+)".*') {
        $CurrentTestResults.Name = $matches[1]
        $CurrentTestResults.Status = 'In Progress'
        $CurrentTestResults.Details = ''
        $CurrentTestResults.Recommendation = ''
    } elseif ($Line -match 'Test succeeded') {
        $CurrentTestResults.Status = 'Succeeded'
    } elseif ($Line -match 'Test failed') {
        $CurrentTestResults.Status = 'Failed'
    } elseif ($Line -match 'Details: (.+)') {
        $CurrentTestResults.Details = $matches[1]
    } elseif ($Line -match 'Recommendation: (.+)') {
        $CurrentTestResults.Recommendation = $matches[1]
    } else {
        return
    }

    $ValueToLog = "Set-AksHciConfig: " + ($CurrentTestResults | ConvertTo-Json)
    Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Message $ValueToLog
}

# Always clean up before setting configuration
Uninstall-AksHci -Confirm:$false -WarningAction SilentlyContinue

$NetworkParams = @{}
$NetworkSettings.psobject.properties | ForEach-Object { $NetworkParams[$_.Name] = $_.Value }

$Vnet = AksHci\New-AksHciNetworkSetting @NetworkParams
Add-Member -InputObject $AksHciConfigurations -NotePropertyName 'Vnet' -NotePropertyValue $Vnet

if ($null -ne $ProxySettings) {
    $ProxyParams = @{}
    if ($ProxySettings.username -and $ProxySettings.password) {
        $SecureStringPassword = ConvertTo-SecureString -String $ProxySettings.password -AsPlainText -Force
        $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ProxySettings.username, $SecureStringPassword
        $ProxyParams['credential'] = $Credential
    }

    $ProxySettings.psobject.properties | ForEach-Object {
        if (-Not @('username', 'password').Contains($_.Name)) {
            $ProxyParams[$_.Name] = $_.Value
        }
    }

    $Proxy = New-AksHciProxySetting @ProxyParams
    Add-Member -InputObject $AksHciConfigurations -NotePropertyName 'proxySettings' -NotePropertyValue $Proxy
}

$ConfigParams = @{}
$AksHciConfigurations.psobject.properties | ForEach-Object { $ConfigParams[$_.Name] = $_.Value }

$CurrentTestResults = @{}
Set-AksHciConfig @ConfigParams -ErrorAction Stop 6>&1 | ForEach-Object { HandleTestOutput($_) }

}
## [END] Set-WACAKAksHciConfiguration ##
function Set-WACAKAksHciRegistration {
<#

.SYNOPSIS
Sets up registration for AKS HCI platform

.DESCRIPTION
Sets up registration for AKS HCI platform

.EXAMPLE
./Set-AksHciRegistration.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Administrators

#>
Param(
    [Parameter(Mandatory = $true)]
    [Object]
    $AksHciRegistrationSettings,
    [Parameter(Mandatory = $true)]
    [String]
    $servicePrincipalUsername,
    [Parameter(Mandatory = $true)]
    [String]
    $servicePrincipalPass
)

Import-Module AksHci -ErrorAction Stop

$RegistrationParams = @{}
$AksHciRegistrationSettings.psobject.properties | ForEach-Object { $RegistrationParams[$_.Name] = $_.Value }

$SecurePass = ConvertTo-SecureString -String $servicePrincipalPass -AsPlainText -Force
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $servicePrincipalUsername, $SecurePass

Disconnect-AzAccount
AksHci\Set-AksHciRegistration @RegistrationParams -Credential $Credential -WarningAction SilentlyContinue

}
## [END] Set-WACAKAksHciRegistration ##
function Set-WACAKFileContent {
<#

.SYNOPSIS
Sets the content of a file

.DESCRIPTION
Sets the content of a file

.EXAMPLE
./Set-FileContent.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Readers

#>

Param(
    [Parameter(Mandatory=$true)]
    [string]
    $Path,
    [Parameter(Mandatory=$true)]
    $Value,
    [Parameter(Mandatory=$true)]
    [bool]
    $Override,
    [Parameter(Mandatory=$false)]
    [bool]
    $Base64Decode = $false
)

$Path = $ExecutionContext.InvokeCommand.ExpandString($Path);

if($Base64Decode) {
    $Value = [Text.Encoding]::Utf8.GetString([Convert]::FromBase64String($Value))
}
if ($Override) {
    Set-Content -Path $Path -Value $Value -Force
} else {
    Set-Content -Path $Path -Value $Value
}

}
## [END] Set-WACAKFileContent ##
function Sync-WACAKAksHciBilling {
<#

.SYNOPSIS
Trigger Aks Hci billing sync

.DESCRIPTION
Trigger Aks Hci billing sync

.EXAMPLE
./Sync-AksHciBilling.ps1

.NOTES
Returns a list of available updates for the current version

.ROLE
Readers

#>

AksHci\Sync-AksHciBilling
}
## [END] Sync-WACAKAksHciBilling ##
function Test-WACAKPaths {
<#

.SYNOPSIS
Determines if a list of paths exist on a server.

.DESCRIPTION
Determines if a list of paths exist on a server.

.EXAMPLE
./Test-Paths.ps1

.NOTES
The supported Operating Systems are Windows Server 2016, Windows Server 2019.

.ROLE
Readers

#>

Param(
    [Parameter(Mandatory=$true)]
    [string[]]
    $Paths,
    [Parameter()]
    [bool]
    $ShouldBeLeaf = $False
)

$Results = @()
$Type = if($ShouldBeLeaf){'Leaf'} else {'Any'}
foreach($Path in $Paths) {
    $ExpandedPath =  $ExecutionContext.InvokeCommand.ExpandString($Path)
    $results += @{
        Path = $Path;
        Exists = Test-Path -Path $ExpandedPath -PathType $Type
    }
}

return $Results

}
## [END] Test-WACAKPaths ##
function Update-WACAKAksHci {
<#

.SYNOPSIS
Update to the latest version

.DESCRIPTION
Update AKSHCI setup to the latest version

.EXAMPLE
./Update-AksHci.ps1

.ROLE
Administrators

#>

[CmdletBinding()]
param ()

Import-Module AksHci

AksHci\Update-AksHci

}
## [END] Update-WACAKAksHci ##
function Update-WACAKAksHciCluster {
<#

.SYNOPSIS
Update cluster to the given version

.DESCRIPTION
Update AksHci cluster to the given version

.EXAMPLE
./Update-AksHciCluster.ps1

.ROLE
Administrators

#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]
    [string]
    $ClusterName,
    [Parameter(Mandatory = $true)]
    [string]
    $KubernetesVersion
)

Import-Module AksHci

AksHci\Update-AksHciCluster -Name $ClusterName -kubernetesVersion $KubernetesVersion -Confirm:$false -operatingSystem

}
## [END] Update-WACAKAksHciCluster ##
function Write-WACAKUserSettings {
<#
.SYNOPSIS
Dumps provided text in a json file on the machine

.DESCRIPTION
Writes user settings into the file $env:USERPROFILE/Windows Admin Center/aks-hci-settings.json

.ROLE
Readers
#>

Param(
    [Parameter(Mandatory=$true)]
    [string]
    $UserSettings
)

New-Item -ItemType Directory -Force -Path "$env:USERPROFILE/Windows Admin Center" | Out-Null;
Write-Output $UserSettings > "$env:USERPROFILE/Windows Admin Center/aks-hci-settings.json"

}
## [END] Write-WACAKUserSettings ##
function Get-WACAKCimWin32LogicalDisk {
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
## [END] Get-WACAKCimWin32LogicalDisk ##
function Get-WACAKCimWin32NetworkAdapter {
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
## [END] Get-WACAKCimWin32NetworkAdapter ##
function Get-WACAKCimWin32PhysicalMemory {
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
## [END] Get-WACAKCimWin32PhysicalMemory ##
function Get-WACAKCimWin32Processor {
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
## [END] Get-WACAKCimWin32Processor ##
function Get-WACAKClusterInventory {
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
## [END] Get-WACAKClusterInventory ##
function Get-WACAKClusterNodes {
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
## [END] Get-WACAKClusterNodes ##
function Get-WACAKDecryptedDataFromNode {
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
## [END] Get-WACAKDecryptedDataFromNode ##
function Get-WACAKEncryptionJWKOnNode {
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
## [END] Get-WACAKEncryptionJWKOnNode ##
function Get-WACAKServerInventory {
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
## [END] Get-WACAKServerInventory ##

# SIG # Begin signature block
# MIIoOQYJKoZIhvcNAQcCoIIoKjCCKCYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAGpkMpC4VVQqRV
# 2sGYTxuvW4H/fBTqBawWjfMW7gjcpaCCDYUwggYDMIID66ADAgECAhMzAAADTU6R
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
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIGWV
# G1VlYxifpJZv7TpGxDgT0ak1gXXe2Dh86tI83y9eMEIGCisGAQQBgjcCAQwxNDAy
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20wDQYJKoZIhvcNAQEBBQAEggEAZeVsq4QD+j6cWa5n/BHVVeJ9ptNwKZHwtQGS
# 2I7B6JbvbJzX6E1fk6sQiwywd3Ibl1u97P8BJ+HnTqE2fTnXouroeV11C6+NUTFw
# JZx1YTSsjceamGFl6rao2DXPhSCFoLqoU7PBcKLc3rnCuUPbBSAUmKzU407vFVDH
# VtGfddtsz0q1x3z5n5QdqWpVmk0fdiAkBPDA2Z40tNpU/E60851qrZmMM6RmNqUg
# rOf41fSc52fMHJsZNOiu14BRocx3tH1wzuYcQJNESMSofGV8BYpzjcWNsdbTxGfE
# jtbgUSjX+E6M+fCK3v9D8HzwoF9giFHlHBJ6vYmHL+Ecnnu9XKGCF5QwgheQBgor
# BgEEAYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZI
# AWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGE
# WQoDATAxMA0GCWCGSAFlAwQCAQUABCCjalCjtZvuVqlfnjfOLicZDX4nacB3lAVQ
# lZpgo5gFGgIGZVbCx+afGBMyMDIzMTIwNjIzNTE1Ny4xMTNaMASAAgH0oIHRpIHO
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
# c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6Rr0jTAi
# GA8yMDIzMTIwNjEzMjQyOVoYDzIwMjMxMjA3MTMyNDI5WjB0MDoGCisGAQQBhFkK
# BAExLDAqMAoCBQDpGvSNAgEAMAcCAQACAiW7MAcCAQACAhPwMAoCBQDpHEYNAgEA
# MDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAI
# AgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBAGmcd8NPOQupwLSgzOFTxGWCQElp
# CUMjxBxxGIdhuYhdVDA9zaOLTMo0FGyyRUOyPOvIR7gWAjXUxTYb1mCuFjHLl6D+
# 3EEotxGpCJ6Q6dtNtA/CM2t4RfxbUt3gLBlsFqXcNNw6Xsv6R/XeMqzp1c5nK02b
# XKNebNWyjyy86Ts8QtsIBklFGRmhs5rj+jQixrdVj/wGnBfj0P/d3kCnqpdBvR/R
# adGUjao74+CZ3yTV4Put3OKyKbIN156Mbyd+HbXkdIrljyXdoNeafEtaCHnYU0uv
# P0jvsE+uIGTtiMnRbj8R3OGbreh/VNQ/aGOxtn1f5fyKfkeY9UjUx+8n94MxggQN
# MIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdYn
# af9yLVbIrgABAAAB1jANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
# CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCDWw91h89bxQaDovdZLRl8SDdzZ
# 1m7lEjp33y0Dfd2vSTCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EINbLTQ1X
# eNM+EBinOEJMjZd0jMNDur+AK+O8P12j5ST8MIGYMIGApH4wfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAHWJ2n/ci1WyK4AAQAAAdYwIgQgQLOvfvDO
# qqPid9EWY9iO5XSoScUy2WrTytwAI+yxXt4wDQYJKoZIhvcNAQELBQAEggIAFRbh
# F4qiINYnWULF9LalOTRwpEY0Ak7y47m+a3iwVjIQ9pDXla+svnBVhAFWFOewHfhv
# 6xIM1fmkx12f5W4VyAAtC3f6smPl2Azo1gBQScgGhxTf8ygruXk0MsIRSRVTDZka
# DW4D0vqjLjqvASyZWcuGzAsZKAqLIuTAIqJkBY99XpuePrbkln0WGwuerB4NY+te
# dLLlhh7Xp3A+1yARwmRrth2AEUp/b9ZiITEl1+MTYcL6KD//o3YvTFBJdoINoxX5
# bqQVzmZKGFPrc8+eb7Itji9RKPD/0chV8HIIeZCjHmAHCVv4vop/yBhG1WbdsPm8
# m/tb91irwPBFkZS06AslEBnYjlDBYSU8lhjHjQlh3mwlB8kCEnny/jyOQ3PVCEV6
# Mc5kLjT8V9E6E/u/muEYV32UrHDsgEcLB3fYw/0REitBFOq1ov4AM8X46oPyx7eE
# bPXLXp7vwUyAPbLtjgRgDXObkaWDwi/6LnsoKsqZrjEwNfNa65CJjrOJVsHVlDIL
# 88q3NvTrW+tJGBpzN6lmkQa9Zls0HhCkeQI7/jIdobWk+DsF0sLw7X5NTqF6qyrO
# 1dHgOsZvuORc+vbwaGAi/pRXUKbARaZohYyGYKGKSSG4lsw60wFgG1NCcYg+D78N
# 5/82T6QCNr9yEq9ZyfYXImHL16UcR4eo4+2/5Ho=
# SIG # End signature block
