<#
.SYNOPSIS
    Adds a host to the cache cluster
#>
[CmdletBinding()]
#region PARAMS
param () 
#endregion PARAMS

begin {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell
    $server = $env:ComputerName
}

process {
    try {
        
        Write-Verbose "$(Get-Date -Format T) : Adding server $server as Distributed Cache host..."
        Add-SPDistributedCacheServiceInstance -ErrorAction Stop
        # If DC memory allocation is not default, you have to change the allocation for entire cluster after adding a new host
        Write-Verbose "$(Get-Date -Format T) : Added server $server as Distributed Cache host!"
    }
    catch {
        Write-Error "We failed during adding cache host cluster on $server with the following error: `n $($_.Exception.Message)"
        Exit
    }
}

end {}

#EOF