Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# ====================================================================================
# Func: Get-AdministratorsGroup
# Desc: Returns the actual (localized) name of the built-in Administrators group
# From: Taken from the AutoSPInstaller project: http://autospinstaller.codeplex.com
# ====================================================================================

Function Get-AdministratorsGroup
{
    If(!$builtinAdminGroup)
    {
        $builtinAdminGroup = (Get-WmiObject -Class Win32_Group -computername $env:COMPUTERNAME -Filter "SID='S-1-5-32-544' AND LocalAccount='True'" -errorAction "Stop").Name
    }
    Return $builtinAdminGroup
}

# ====================================================================================
# ScriptBlock: AlterDB
# Description: PowerShell script block passed to the SQL server
# ====================================================================================

[ScriptBlock] $global:AlterDB = {

    param ([string] $DBName,[string] $LoginName)

    # Create our SQL query to fix db permissions
    [string] $dbCommand = "ALTER AUTHORIZATION ON DATABASE::[$DBname] TO [$LoginName];"

    # Attempt to get SQL cmdlets
    $SQLSnapins = Get-PSSnapin -registered | Where-Object {$_.Name -like "SqlServer*"}
    $SQLModules = Get-Module | Where-Object {$_.Name -like "SQL*"}
    $SQLcmdlets = $false

    # Check to see if a snapin for the SQL cmdlets exists (versions prior to SQL 2012)
    if($SQLSnapins -ne $null)
    {
        foreach($Snapin in $SQLSnapins)
        {
            Add-PSSnapin $Snapin | Out-Null
        }
        $SQLcmdlets = $true
    }

    # If no snapin, check for SQL module (SQL 2012+)
    elseif($SQLModules -ne $null)
    {
        foreach($Module in $SQLModules)
        {
            Import-Module SQLPS | Out-Null
        } 
        $SQLcmdlets = $true            
    }

    # Display warning if no SQL cmdlets available
    else
    {
        Write-Host "No SQL cmdlets found!" -ForegroundColor Yellow
    }

    # If SQL cmdlets available, execute our query
    if($SQLcmdlets -eq $true)
    {
        Invoke-SqlCmd -Query $dbCommand
    }
}

<#
.SYNOPSIS
Reprovisions the User Profile Synchronization Service	

.DESCRIPTION
This function fixes the SharePoint user profile synchronization service not starting by reprovisioning the User Profile Synchronization Service. DBcreator and SecurityAdmin access to the SharePoint SQL server are required, as well as local administrator on the SharePoint application server. You'll also need to be able to invoke a remote PowerShell session on the SQL server. In addition, you will have to recreate your synchronization connections afterward (I hope to automate this in a future release of this script).

.EXAMPLE
Fix-SPUserProfileSync -FarmAdmin "DOMAIN\SP_farm" -FarmPassword "FARM_password" -DBServer "DBServerName"

This command takes three required parameters and resets the synchronisation database. "DBServerName" can be an alias or the actual SQL server name.

.NOTES
NAME: Fix-SPUserProfileSync
AUTHOR: Wes Kroesbergen
CAVEATS: DBcreator and SecurityAdmin access to the SharePoint SQL server are required, as well as local administrator on the SharePoint application server. You'll also need to be able to invoke a remote PowerShell session on the SQL server. In addition, you will have to recreate your synchronization connections afterward (I hope to automate this in a future release of this script).

.LINK
http://www.kroesbergens.com
#>

Function Fix-SPUserProfileSync
{
    Param([Parameter(Position=0,Mandatory=$true)][string] $FarmAdmin,
        [Parameter(Position=1,Mandatory=$true)][string] $FarmPassword,
        [Parameter(Position=2,Mandatory=$true)][string] $DBServer)

    $FarmPassword = ConvertTo-SecureString $FarmPassword -AsPlainText -Force    
    $farmAcctDomain,$farmAcctUser = $farmAdmin -Split "\\"
    $builtinAdminGroup = Get-AdministratorsGroup

    # Elevate privileges if necessary
    Try
    {        
        ([ADSI]"WinNT://$env:COMPUTERNAME/$builtinAdminGroup,group").Add("WinNT://$farmAcctDomain/$farmAcctUser")
        If (-not $?) {Throw}
        # Restart the SPTimerV4 service if it's running, so it will pick up the new credential
        If ((Get-Service -Name SPTimerV4).Status -eq "Running")
        {
            Write-Host -ForegroundColor White " - Restarting SharePoint Timer Service..."
            Restart-Service SPTimerV4
        }
    }
    Catch {Write-Host -ForegroundColor White " - $FarmAdmin is already a member of `"$builtinAdminGroup`"."}

    # Stop the Timer Service
    Stop-Service sptimerv4

    # Get the Sync DB and clear the data
    $SyncDB=Get-SPDatabase | Where-Object {$_.Type -eq "Microsoft.Office.Server.Administration.SynchronizationDatabase"}
    $SyncDB.Unprovision()
    $SyncDB.Status='Offline'

    # Reset the Sync info in the User Profile Application
    $UPA=Get-SPServiceApplication | Where-Object {$_.TypeName -match "User Profile Service Application"}
    $UPA.ResetSynchronizationMachine()
    $UPA.ResetSynchronizationDatabase()

    # Reprovision the Sync DB
    $SyncDB.Provision()

    $UPSinstance = Get-SPServiceInstance | Where-Object {$_.TypeName -match "User Profile Synchronization Service" }
    $UPA.SetSynchronizationMachine($env:COMPUTERNAME, $UPSinstance.ID, $FarmAdmin, $FarmPassword)
    Start-SPServiceInstance -Identity $UPSinstance.ID 

    # Before we try to connect remotely, let's confirm DBServer is not an alias
    $AliasPath = "HKLM:\Software\Microsoft\MSSQLServer\Client\ConnectTo\"
    $Item = Get-ItemProperty -Path $AliasPath -Name $DBServer
    if ((test-path -path $AliasPath) -eq $True)
    {
        write-host "Alias exists!"

        # If object exists, we parse the definitions and retrieve the one with a TCP alias value
        $AliasProperties = ($Item | Get-Member | Where-Object {$_.Definition -like "*DBMSSOCN*"}).Definition.Split(",")
        $DBServer = $AliasProperties[1]
        Write-Host "Actual server name is $DBServer!" -ForegroundColor Yellow
    }

    # Make sure the farm account has DBO rights to Sync DB
    Invoke-Command -ComputerName $DBServer -ScriptBlock $AlterDB -ArgumentList $SyncDB,$FarmAdmin

    # Start the Timer Service
    Start-Service sptimerv4

    # Remove farm account from local admin
    Write-Host -ForegroundColor White " - Removing $FarmAdmin from local group `"$builtinAdminGroup`"..."
        
    try
    {
        ([ADSI]"WinNT://$env:COMPUTERNAME/$builtinAdminGroup,group").Remove("WinNT://$farmAcctDomain/$farmAcctUser")
        If (-not $?) {throw}
    }
    catch {Write-Host -ForegroundColor White " - $FarmAdmin already removed from `"$builtinAdminGroup.`""}

    # Restart SPTimerV4 so it can now run under non-local Admin privileges and avoid Health Analyzer warning
    Write-Host -ForegroundColor White " - Starting SharePoint Timer Service..."

    Restart-Service SPTimerV4
}