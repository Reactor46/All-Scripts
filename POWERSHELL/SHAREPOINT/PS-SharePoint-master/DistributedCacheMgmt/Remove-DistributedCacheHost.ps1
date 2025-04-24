<#
.SYNOPSIS
    Removes host from cache cluster
    

.DESCRIPTION

    spence@harbar.net
    25/06/2015
    
.NOTES
    File Name  : Remove-DistributedCache.ps1
    Author     : Spencer Harbar (spence@harbar.net)
    Requires   : PowerShell Version 3.0  
.LINK

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
        
        Write-Output "$(Get-Date -Format T) : Removing server $server as Distributed Cache host..."
        Remove-SPDistributedCacheServiceInstance -ErrorAction Stop
   
        $Count = 0
        $MaxCount = 5
        While ( ($Count -lt $MaxCount) -and (Get-NetTCPConnection -LocalPort 22234 -ErrorAction SilentlyContinue).Count -gt 0) {
            Write-Output "$(Get-Date -Format T) : Waiting on port to free up..."
            Start-Sleep 30
            $Count++
        }
        Write-Output "$(Get-Date -Format T) : Removed server $server as Distributed Cache host!"
    }
    catch {
        Write-Output "ERROR: We failed during removing cache host cluster on $server."
        $_.Exception.Message
        Exit
    }
}

end {}

#EOF