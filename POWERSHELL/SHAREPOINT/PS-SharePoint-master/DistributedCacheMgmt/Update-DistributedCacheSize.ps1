<#
.SYNOPSIS
    Updates the Distributed Cache Service Cache Szie


.DESCRIPTION

    Stops all Cache Hosts, then updates the size, starts all cache hosts


    spence@harbar.net
    25/06/2015
    
.NOTES
    File Name  : Update-DistributedCacheSize.ps1
    Author     : Spencer Harbar (spence@harbar.net)
    Requires   : PowerShell Version 3.0 
					Stop-DistributedCacheCluster.ps1
					Start-DistributedCacheCluster.ps1
.LINK
.PARAMETER File  
    The configuration file

#>
[CmdletBinding()]
#region PARAMS
param (
    [Parameter(Mandatory=$true,
               ValueFromPipelineByPropertyName=$true,
               Position=0)]
    [ValidateNotNullorEmpty()]
    [Int]
    $CacheSizeInMB,

    [Parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true,
               Position=1)]
    [ValidateNotNullorEmpty()]
    [PSCredential]
    $Credential
)
#endregion PARAMS

begin {
    Write-Output "$(Get-Date -Format T) : Initiated Distributed Cache Service Cache Size change."
    Add-PSSnapin -Name "Microsoft.SharePoint.PowerShell"
    $StopDistributedCacheScript = "$PSScriptRoot\Stop-DistributedCacheCluster.ps1"
    $StartDistributedCacheScript = "$PSScriptRoot\Start-DistributedCacheCluster.ps1"
}

process {
    try {
        # Get the servers running DC
        $DistributedCacheServers = Get-SPServer | `
                                   Where-Object {($_.ServiceInstances | ForEach TypeName) -eq "Distributed Cache"} | `
                                   ForEach Address

        # Stop service instances
        Invoke-Expression $StopDistributedCacheScript

        # Update the Cache Size
        Write-Output "$(Get-Date -Format T) : Changing Distributed Cache Service Cache Size..."
        Invoke-Command -ComputerName $DistributedCacheServers[0] -Credential $Credential -Authentication Credssp `
                       -ArgumentList $CacheSizeInMB `
                       -ScriptBlock {
                            Add-PSSnapin -Name "Microsoft.SharePoint.PowerShell"
                            Update-SPDistributedCacheSize -CacheSizeInMB ($args[0])
                        }
        # Start service instances
        Invoke-Expression $StartDistributedCacheScript
    }
    catch {
        Write-Output "ERROR: We failed during changing the Distributed Cache Size."
        $_.Exception.Message
        Exit
    }
}

end {
    Write-Output "$(Get-Date -Format T) : Completed Distributed Cache Service Cache Size Change!"
}