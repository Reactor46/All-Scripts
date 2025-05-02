[CmdletBinding()]
param (
    [Parameter(ParameterSetName='StartFarm')]
    [switch]$Start,
    [Parameter(ParameterSetName='StopFarm')]
    [switch]$Stop
)

begin {
    <#
    $resourcesAvailable = $true
    if (Test-Path (Join-Path $PSScriptRoot "Stop-DistributedCacheCluster.ps1") -ne $true)
    {
        $resourcesAvailable = $false
    } 
    if (Test-Path (Join-Path $PSScriptRoot "Change-SQLAlias.ps1") -ne $true)
    {
        $resourcesAvailable = $false
    } 
    if (Test-Path (Join-Path $PSScriptRoot "Start-DistributedCacheCluster.ps1") -ne $true)
    {
        $resourcesAvailable = $false
    }
    if (Test-Path (Join-Path $PSScriptRoot "RestartTimerServicesInFarm.ps1") -ne $true)
    {
        $resourcesAvailable = $false
    }
    if (!$resourcesAvailable)
    {
        Throw " - required resources are not available!"
    }
    #>

    #Loading external functions
    Import-Module .\ManageSPServices.psm1
    Import-Module ..\Helpers\ManageSQLAlias.psm1
}

process {
    if ($Stop)
    {
        Stop-SPTimerService
        ..\DistributedCacheMgmt\Stop-DistributedCacheCluster.ps1
        Stop-IISService

        Get-SPLogLevel | Set-SPLogLevel -TraceSeverity None -EventSeverity Critical
        $servers = Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}

        foreach ($server in $servers)
        {
            Set-SQLAlias -AliasName "SP13TESTSQL" -DBServer "DBSQL_SP2013_TEST-DISABLED" -PortNumber 45001 -RemoteComputer $server.Name
        }
    }

    if ($Start)
    {
        $servers = Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}
        
        
        foreach ($server in $servers)
        {
            Set-SQLAlias -AliasName "SP13TESTSQL" -DBServer "DBSQL_SP2013_TEST" -PortNumber 45001 -RemoteComputer $server.Name
        }
        Get-SPLogLevel | Clear-SPLogLevel
        Start-IISService
        ..\DistributedCacheMgmt\Start-DistributedCacheCluster.ps1
        Start-SPTimerServices
    }
}