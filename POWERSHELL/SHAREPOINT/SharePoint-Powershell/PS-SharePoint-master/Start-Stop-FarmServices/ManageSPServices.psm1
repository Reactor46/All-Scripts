<#
.SYNOPSIS
  Functions to manage SharePoint-related Windows services across farm servers.

.INPUTS
  None

.OUTPUTS
  None

.NOTES
  Version:        2.0
  Author:         Egor Yudkin
  Creation Date:  03/05/2018
  Purpose/Change: Added functions for Search-related services
                  Cleaned up the code
                  Improved compliance with code standards

  Version:        1.0
  Author:         Egor Yudkin
  Creation Date:  12/27/2017
  Purpose/Change: Initial development

#>

#Requires -Version 3
#Requires -PSSnapin Microsoft.SharePoint.PowerShell

Function StartStop-WindowsService
{
    [CmdletBinding()]
    param 
    (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ServiceName,

        [ValidateSet("Start", "Stop", "Restart")]
        [string]
        $Action,

        [parameter()]
        [string]
        $ServiceTitle,

        [parameter()]
        [string[]]
        $ServerNames
    )

    if (!$ServerNames) 
    {
        $ServerNames = @($env:COMPUTERNAME)
    }
    if ($null -eq $ServiceTitle -or $ServiceTitle -eq [string]::Empty)
    {
        $ServiceTitle = $ServiceName
    }

	foreach ($server in $ServerNames) 
    {
        try
        {
            $service = Get-Service -ComputerName $server -Name $ServiceName -ErrorAction Stop
        }
        catch
        {
            Write-Warning "Service '$ServiceName' was not found on $server"
            continue
        }

        try 
        {
            switch ($Action)
            {
                "Start"
                {
                    Write-Verbose "Starting '$ServiceTitle' service on $server"
                    Start-Service -InputObject $service -ErrorAction Stop
                }
                "Stop"
                {
                    Write-Verbose "Stopping '$ServiceTitle' service on $server"
                    Stop-Service -InputObject $service -ErrorAction Stop
                }
                "Restart"
                {
                    Write-Verbose "Restarting '$ServiceTitle' service on $server"
                    Restart-Service -InputObject $service -ErrorAction Stop
                }
            }
	        
        }
        catch
        {
            Write-Warning "Service '$ServiceName' failed to $Action"
        }
	}
}

Function Set-WindowsServiceState
{
}

Function Set-WindowsServiceStartupType
{
}

Function Restart-SPTimerService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
	if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "sptimerv4" -ServiceTitle "SharePoint Timer" -Action Restart -ServerNames $ServerNames
}

Function Stop-SPTimerService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
	if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
	StartStop-WindowsService -ServiceName "sptimerv4" -ServiceTitle "SharePoint Timer" -Action Stop -ServerNames $ServerNames
}

Function Start-SPTimerService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
	StartStop-WindowsService -ServiceName "sptimerv4" -ServiceTitle "SharePoint Timer" -Action Start -ServerNames $ServerNames
}

Function Enable-SPTimerService
{
    [CmdletBinding()]
    param (
        [string[]]$ServerNames
    ) 
    process {
	    if (!$ServerNames) {$ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name}
	    foreach ($server in $ServerNames) {
	     Write-Verbose "Enabling SharePoint Timer service on $server"
	     Set-Service -InputObject (Get-Service -ComputerName $server -Name "sptimerv4") -StartupType Automatic
	    }
    }
}

Function Disable-SPTimerService
{
    [CmdletBinding()]
    param (
        [string[]]$ServerNames
    ) 
    process {
	    if (!$ServerNames) {$ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name}
	    foreach ($server in $ServerNames) {
	     Write-Verbose "Disabling SharePoint Timer service on $server"
	     Set-Service -InputObject (Get-Service -ComputerName $server -Name "sptimerv4") -StartupType Disabled
	    }
    }
}

Function Restart-IISService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "w3svc" -ServiceTitle "IIS" -Action Restart -ServerNames $ServerNames
}

Function Stop-IISService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "w3svc" -ServiceTitle "IIS" -Action Stop -ServerNames $ServerNames
}

Function Start-IISService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "w3svc" -ServiceTitle "IIS" -Action Start -ServerNames $ServerNames
}

Function Restart-SPAdminService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "spadminv4" -ServiceTitle "SharePoint Administration" -Action Restart -ServerNames $ServerNames
}

Function Stop-SPAdminService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "spadminv4" -ServiceTitle "SharePoint Administration" -Action Stop -ServerNames $ServerNames
}

Function Start-SPAdminService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "spadminv4" -ServiceTitle "SharePoint Administration" -Action Start -ServerNames $ServerNames
}

Function Enable-SPAdminService
{
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    )
	if (!$ServerNames) {$ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name}
	foreach ($server in $ServerNames) {
	    Write-Verbose "Enabling SharePoint Administration service on $server"
	    Set-Service -InputObject (Get-Service -ComputerName $server -Name "spadminv4") -StartupType Automatic
	}
}

Function Disable-SPAdminService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    )
	if (!$ServerNames) {$ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name}
	foreach ($server in $ServerNames) {
	    Write-Verbose "Disabling SharePoint Administration service on $server"
	    Set-Service -InputObject (Get-Service -ComputerName $server -Name "spadminv4") -StartupType Disabled
	}
}

Function Stop-SPSearchService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "OSearch15" -ServiceTitle "SharePoint Server Search" -Action Stop -ServerNames $ServerNames
}

Function Start-SPSearchService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "OSearch15" -ServiceTitle "SharePoint Server Search" -Action Start -ServerNames $ServerNames
}

Function Enable-SPSearchService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    )
	if (!$ServerNames) {$ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name}
	foreach ($server in $ServerNames) {
	    Write-Verbose "Enabling SharePoint Search service on $server"
	    Set-Service -InputObject (Get-Service -ComputerName $server -Name "OSearch15") -StartupType Automatic
	}
}

Function Disable-SPSearchService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    )
	if (!$ServerNames) {$ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name}
	foreach ($server in $ServerNames) {
	    Write-Verbose "Disabling SharePoint Search service on $server"
	    Set-Service -InputObject (Get-Service -ComputerName $server -Name "OSearch15") -StartupType Disabled
	}
}

Function Stop-SPSearchHostService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "SPSearchHostController" -ServiceTitle "SharePoint Search Host Controller" -Action Stop -ServerNames $ServerNames
}

Function Start-SPSearchHostService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    ) 
    if (!$ServerNames) 
    {
        $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
    }
    StartStop-WindowsService -ServiceName "SPSearchHostController" -ServiceTitle "SharePoint Search Host Controller" -Action Start -ServerNames $ServerNames
}

Function Enable-SPSearchHostService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    )
    process 
    {
	    if (!$ServerNames) 
        {
            $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
        }
	    foreach ($server in $ServerNames) 
        {
	        Write-Verbose "Enabling SharePoint Search Host Controller service on $server"
	        Set-Service -InputObject (Get-Service -ComputerName $server -Name "SPSearchHostController") -StartupType Automatic
	    }
    }
}

Function Disable-SPSearchHostService
{
    [CmdletBinding()]
    param 
    (
        [parameter()]
        [string[]]
        $ServerNames
    )
    process 
    {
	    if (!$ServerNames) 
        {
            $ServerNames = (Get-SPServer | ?{$_.Role -ne [Microsoft.SharePoint.Administration.SPServerRole]::Invalid}).Name
        }
	    foreach ($server in $ServerNames) 
        {
	        Write-Verbose "Disabling SharePoint Search Host Controller service on $server"
	        Set-Service -InputObject (Get-Service -ComputerName $server -Name "SPSearchHostController") -StartupType Disabled
	    }
    }
}

Export-ModuleMember Restart-SPTimerService,
                    Stop-SPTimerService, 
                    Start-SPTimerService,
                    Enable-SPTimerService,
                    Disable-SPTimerService, 
                    Restart-IISService, 
                    Stop-IISService, 
                    Start-IISService, 
                    Restart-SPAdminService, 
                    Stop-SPAdminService, 
                    Start-SPAdminService,
                    Enable-SPAdminService,
                    Disable-SPAdminService,
                    Stop-SPSearchHostService, 
                    Start-SPSearchHostService,
                    Enable-SPSearchHostService,
                    Disable-SPSearchHostService,
                    Stop-SPSearchService, 
                    Start-SPSearchService,
                    Enable-SPSearchService,
                    Disable-SPSearchService
