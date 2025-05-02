# Copyright (c) 2005 Microsoft Corporation.  All rights reserved.
#
# ServiceControl.ps1
#
# This script performs 3 functions:
# 1.  Saves or removes the state of currently running functions
# 2.  Stops services that use exchange files so they can be updated
# 3.  Starts services based on the state taken in step 1
# See the Usage function for details

param(
    [string]$Operation,
    [string[]]$Roles,
    [string]$Version = $RoleTargetVersion,
    [string]$SetupScriptsDirectory = $null
)

# disable ExecutionPolicy in this PowerShell context
($context = $ExecutionContext.GetType().getfield("_context","nonpublic,instance").getvalue($ExecutionContext)).GetType().getfield("_authorizationManager","nonpublic,instance").setvalue($context, (New-Object System.Management.Automation.AuthorizationManager "Microsoft.Powershell"))

# Disable the confirmation for destructive tasks
$confirmPreference = "None"

#
# "Member variables" for the service control functions
#
$script:logDir = "$env:SYSTEMDRIVE\ExchangeSetupLogs"

# E14: 254664: Flag to determine to log to ExchangeSetup.log or a separate file.
$script:writeToExchangeSetupLog = $true

$script:serviceStateFile = "$script:logDir\ServiceState$Version.xml"
$script:serviceStartupModeFile = "$script:logDir\ServiceStartupMode$Version.xml"

$script:services = @( )
$script:servicesRegistryData = @( )
$script:previouslyStartedServices = @{ }

$script:knownRoles = @( 'AdminTools', 'Bridgehead', 'ClientAccess', 'Gateway', 'Mailbox', 'UnifiedMessaging', 'FrontendTransport', 'Cafe', 'Monitoring', 'CentralAdmin', 'Critical', "LanguagePacks", 'OSP' )
$script:serviceControlError = $null

# DEVS:  Put services for your role here
#
# NOTES:
#        MSExchangeHM service should be last in the list if it is applicable to the Exchange role
#        This is so that Active Monitoring responder won't interfere with your service before your service is fully initialized
#
#        O15:3254791: MSExchangeMonitoring service was removed by sdv 1099096, but we needed it back due to B2B upgrade from E15 RTM to later version.
#
$script:servicesToControl = @{}
$script:servicesToControl['Common']             = @( 'WinMgmt', 'RemoteRegistry', 'HealthService', 'OnePoint', 'MOM', 'OMCFG', 'pla' )
$script:servicesToControl['ClientAccess']       = @( 'MSExchangeMonitoring', 'MSExchangeIMAP4', 'MSExchangePOP3' , 'MSExchangeADTopology' ,'MSExchangeTopologyService', 'MSExchangeFDS', 'IISAdmin', 'MSExchangeServiceHost', 'W3Svc', 'MSExchangeRPC', 'MSExchangeIMAP4BE', 'MSExchangePOP3BE', 'MSExchangeMailboxReplication', 'MSExchangeFBA', 'MSExchangeProtectedServiceHost', 'MSExchangeDiagnostics', 'MSExchangeHM', 'MSExchangeHMRecovery')
$script:servicesToControl['Gateway']            = @( 'MSExchangeMonitoring', 'WorkerService', 'MSExchangeTransport', 'MSExchangeTransportLogSearch', 'MSExchangeEdgeSync', 'MSExchangeAntispamUpdate', 'MSExchangeEdgeCredential', 'MSExchangeServiceHost', 'MSExchangeHM', 'MSExchangeHMRecovery', 'MSExchangeDiagnostics')
$script:servicesToControl['Mailbox']            = @( 'MSExchangeMonitoring', 'IISAdmin', 'MSExchangeIS', 'MSExchangeMailboxAssistants', 'MSFTESQL-Exchange', 'MSExchangeThrottling', 'MSExchangeADTopology' ,'MSExchangeTopologyService', 'MSExchangeRepl', 'MSExchangeDagMgmt', 'MSExchangeWatchDog', 'MSExchangeTransportLogSearch', 'MSExchangeRPC', 'MSExchangeServiceHost', 'W3Svc', 'HTTPFilter', 'wsbexchange', 'MSExchangeTransportSyncManagerSvc', 'MSExchangeFastSearch', 'hostcontrollerservice', 'SearchExchangeTracing', 'MSExchangeSubmission', 'MSExchangeDelivery', 'MSExchangeMigrationWorkflow', 'MSExchangeDiagnostics', 'MSExchangeProcessUtilizationManager', 'MSExchangeHM', 'MSExchangeHMRecovery')
$script:servicesToControl['Bridgehead']         = @( 'MSExchangeMonitoring', 'AdminService', 'FMS', 'MSExchangeAntimalwareSvc', 'MSExchangeAntimalwareUpdateSvc', 'MSExchangeTransport' , 'MSExchangeADTopology' ,'MSExchangeTopologyService',  'MSExchangeEdgeSync', 'MSExchangeProtectedServiceHost', 'MSExchangeTransportLogSearch', 'MSExchangeTransportStreamingOptics', 'MSExchangeAntispamUpdate', 'MSExchangeServiceHost', 'hostcontrollerservice', 'SearchExchangeTracing', 'W3Svc', 'shm', 'MSMessageTracingClient', 'MSExchangeFileUpload', 'MSExchangeDiagnostics', 'MSExchangeHM', 'MSExchangeHMRecovery', 'MSExchangeStreamingOptics')
$script:servicesToControl['UnifiedMessaging']   = @( 'MSExchangeMonitoring', 'Exchange UM Service' , 'MSExchangeADTopology' ,'MSExchangeTopologyService',  'MSExchangeFDS', 'MSExchangeUM', 'MSExchangeServiceHost', 'W3Svc', 'MSExchangeDiagnostics', 'MSExchangeHM', 'MSExchangeHMRecovery')
$script:servicesToControl['FrontendTransport']  = @( 'MSExchangeMonitoring', 'AdminService', 'MSExchangeTransport' , 'MSExchangeADTopology' ,'MSExchangeTopologyService',  'MSExchangeEdgeSync', 'MSExchangeProtectedServiceHost', 'MSExchangeTransportLogSearch', 'MSExchangeAntispamUpdate', 'MSExchangeServiceHost', 'W3Svc', 'MSExchangeFrontendTransport', 'shm', 'MSMessageTracingClient', 'MSExchangeFileUpload', 'MSExchangeDiagnostics', 'MSExchangeHM', 'MSExchangeHMRecovery')
$script:servicesToControl['Cafe']               = @( 'MSExchangeMonitoring', 'MSExchangeDiagnostics', 'MSExchangeProcessUtilizationManager', 'MSExchangeHM', 'MSExchangeHMRecovery')
$script:servicesToControl['Monitoring']         = @( 'MSExchangeCAMOMConnector', 'MSExchangeMonitoringCorrelation' )
$script:servicesToControl['CentralAdmin']       = @( 'MSExchangeCentralAdmin', 'MSExchangeMonitoringCorrelation', 'WDSServer', 'MSDTC', 'MSExchangeDiagnostics', 'MSExchangeHM', 'MSExchangeHMRecovery')
$script:servicesToControl['OSP']                = @( 'IISAdmin', 'W3Svc','MSExchangeADTopology' ,'MSExchangeTopologyService', 'MSExchangeMonitoring', 'MSExchangeHM', 'MSExchangeHMRecovery')

$script:servicesToControl['LanguagePacks']      = $script:servicesToControl['AdminTools'] +
                                                  $script:servicesToControl['ClientAccess'] +
                                                  $script:servicesToControl['Gateway'] +
                                                  $script:servicesToControl['Mailbox'] +
                                                  $script:servicesToControl['Bridgehead'] +
                                                  $script:servicesToControl['UnifiedMessaging'] +
                                                  $script:servicesToControl['Cafe'] +
                                                  $script:servicesToControl['FrontendTransport'] +
                                                  $script:servicesToControl['OSP']


# List of critical services required for prereqs.
$script:servicesToControl['Critical']           = @( 'WinMgmt', 'RemoteRegistry', 'W3Svc', 'IISAdmin' )

# list of installed services, this should be calculated based on the registry values
$script:installedRoles = @()

# snapin key
$script:snapinKey = 'HKLM:\SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapIns\Microsoft.Exchange.Management.PowerShell.Setup'

# setup key
$script:setupKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'

# If log file folder doesn't exist, create it
if (!(Test-Path $logDir)){
	New-Item $logDir -type directory
}

# Get-ServiceToControl
#	Returns list of service(s) to control.
# Arguments:
#   $Roles - list of Exchange roles.
#   $Active - indicates that only non-stopped service should be returned.
# Returns:
#	Service(s) to control.
function Get-ServiceToControl ([string[]]$Roles, [switch]$Active)
{
    # 1. Populate full list of services for all roles.
    & {
        # 1.a. Get common ones.
        if (($Roles -notcontains 'Critical') -and ($script:servicesToControl['Common']))
        {
            $script:servicesToControl['Common']
        }
        # 1.b. Get services for each role.
        $Roles |
        foreach {
            if ($script:servicesToControl[$_])
            {
                $script:servicesToControl[$_]
            }
        }
    } |
    # 2. Eliminate duplicates.
    sort | unique |
    # 3. Filter only those which are installed
    # and (optionally) running.
    where {
        $serviceName = $_
        # 3.a. Check if installed.
        # Note the trick of requesting by pattern prevents Get-Service
        # from failing in case service is not installed.
        Get-Service "$serviceName*" |
        ?{$_.Name -eq $serviceName} |
        # 3.b. If $Active is specified, check that service is not stopped.
        ?{!$Active -or $_.Status -ne 'Stopped'}
    }
}

###############################################################################
#
# Service state management functions
#	These functions are used to determine if a service was running or not before setup started
#

# Export-ServiceState
#	Saves the state of all services to a well known location
# Returns:
#	void
function Export-ServiceState([switch]$Overwrite)
{
    if (Test-Path $script:serviceStateFile)
    {
        Log "State file $script:serviceStateFile already exists."
        if ($Overwrite)
        {
            Log "Overwrite is specified. File $script:serviceStateFile is going to be overwritten with a new state."
        }
    }

    if (!(Test-Path $script:serviceStateFile) -or $Overwrite)
    {
	    Log "Saving service state to '$script:serviceStateFile'..."
	    get-service | export-clixml $script:serviceStateFile
	}
}

# Export-ServiceStartupMode
#	Saves the registry data from the Services path
# Returns:
#	void
function Export-ServiceStartupMode([switch]$Overwrite)
{
    if (Test-Path $script:serviceStartupModeFile)
    {
        Log "State file $script:serviceStartupModeFile already exists."
        if ($Overwrite)
        {
            Log "Overwrite is specified. File $script:serviceStartupModeFile is going to be overwritten with a new state."
        }
    }

    if (!(Test-Path $script:serviceStartupModeFile) -or $Overwrite)
    {
	    Log "Saving services startup mode."
	    Get-WmiObject Win32_Service -ea SilentlyContinue | %{@{Name=$_.Name; StartMode=$_.StartMode}} | Export-Clixml $script:serviceStartupModeFile
	}
}

# Import-ServiceState
#	Retrieves the previously saved service state
# Returns:
#	void
function Import-ServiceState
{
	if( test-path $script:serviceStateFile)
	{
		Log "Reading service state from '$script:serviceStateFile'..."
		$script:services = import-clixml $script:serviceStateFile
		Log ("Loaded state for {0} services" -F $script:services.Length)
	}
	else
	{
		Log "Service state file not present:  $script:serviceStateFile"
	}
}

# Import-ServiceStartMode
#	Retrieves the previously saved service registry data
# Returns:
#	void
function Import-ServiceStartMode
{
	if( test-path $script:serviceStartupModeFile)
	{
		Log "Reading service registry data from '$script:serviceStartupModeFile'"
		$script:servicesRegistryData = import-clixml $script:serviceStartupModeFile
		Log ("Loaded state for {0} services registry" -F $script:servicesRegistryData.Length)
	}
	else
	{
		Log "Service registry data file not present: $script:serviceStartupModeFile"
	}
}

# Remove-ServiceState
#	Removes previously saved service state
# Returns:
#	void
function Remove-ServiceState
{
	Log "Removing service state from '$script:serviceStateFile'..."
	remove-item $script:serviceStateFile -ea SilentlyContinue
}


# Remove-ServiceStartMode
#	Removes previously saved service registry data file
# Returns:
#	void
function Remove-ServiceStartMode
{
	Log "Removing service registry data from '$script:serviceStartupModeFile'..."
	remove-item $script:serviceStartupModeFile -ea SilentlyContinue
}

# Test-ServiceWasRunning( serviceName )
#	Determines if the service was running when the state was recorded
# Params:
#	serviceName	- Name of the service to check
# Returns:
#	True if the service did not exist or service existed and was in the running state.
#	False on any other condition including the service not existing
#	or starting, but not running
function Test-ServiceWasRunning([String] $serviceName)
{
	$ret = $false

	$s = $script:services | where { $_.ServiceName -ieq $serviceName }
	if( $s.ServiceName -eq $null )
	{
		Log "'$serviceName' did not exist."
		$ret = $true
	}
	elseif( $s.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running )
	{
		Log "'$serviceName' was running."
		$ret = $true
	}
	else
	{
		Log "'$serviceName' was not running."
	}

	return $ret
}

###############################################################################
#
# Main service control functions
#	These functions start and stop services for a given server role
#

# StopServices( roleName )
#	Stops services for the specified role
# Params:
#	Roles - name of the role to stop services for
# Returns:
#	True if was able to stop all services for the specified role
#	False on any error
# Notes:
#	It is not an error to stop services that do not exist because
#	they may be created in the build that is about to install
function StopServices([string[]]$Roles)
{
	Status "Stopping services for '$Roles'..."
	$services = Get-ServiceToControl $Roles -Active
	[array]::Reverse($services)
	foreach( $serviceName in $services )
	{
		Log "Stopping service '$serviceName'."
		Stop-SetupService -ServiceName $serviceName -ev script:serviceControlError
		if( $script:serviceControlError[0] -ne $null )
		{
			Log ($script:serviceControlError[0].ToString())
			return $false
		}
	}
	return $true
}

# DisableServices
#	Disables services for a particular role
# Params:
#	Roles - name of the role to disable services for
# Returns:
#	True if was able to disable all services for the specified role
#	False on any error
function DisableServices([string[]]$Roles)
{
	Status "Disabling services for '$Roles'..."
	foreach( $serviceName in (Get-ServiceToControl $Roles) )
	{
		Log "Disabling service '$serviceName'."
		$serviceHandle = get-service $serviceName -ea SilentlyContinue

		if($serviceHandle -eq $null)
		{
			Log "Ignoring service '$serviceName'"
		}
		else
		{
			set-service $serviceName -startuptype disabled -ev script:serviceControlError
			if( $script:serviceControlError[0] -ne $null )
			{
				Log ($script:serviceControlError[0].ToString())
				return $false
			}
		}
	}
	return $true
}


# EnableServices( roleName )
#	Enables services for the specified role. We don't need any additional logic to take care
#   of dependencies between services as all the services are enabled for a role.
# Params:
#	Roles - name of the role to enable services for
# Returns:
#	True, if for all existing services starup type can be restored. False otherwise.
function EnableServices([string[]]$Roles)
{
	Status "Enabling services for '$Roles'..."
	Import-ServiceStartMode

	foreach( $serviceName in (Get-ServiceToControl $Roles) )
	{
		Log "Enabling service '$serviceName'."
		$service = get-service $serviceName -ea SilentlyContinue
		if($service -ne $null)
		{
			$serviceRegistryData = $script:servicesRegistryData | where {$_['Name'] -ieq $serviceName}

			if ($serviceRegistryData)
			{
				switch ($serviceRegistryData['StartMode'])
				{
					'Auto' { set-service $serviceName -startuptype automatic -ErrorAction SilentlyContinue -ev setServiceError }
					'Manual' { set-service $serviceName -startuptype manual -ErrorAction SilentlyContinue -ev setServiceError }
					'Disabled' { set-service $serviceName -startuptype disabled -ErrorAction SilentlyContinue -ev setServiceError }
					default { Log "'$serviceName' has unrecognized '$_'" }
				}
			}
			else
			{
				# this is not an error as the sevice might have just been installed
				LogWarning "'$serviceName' did not exist, this is not an error as the sevice might have just been installed"
			}

			if(($setServiceError -ne $null) -and ($setServiceError[0] -ne $null))
			{
				Log ($setServiceError[0].ToString())
				if($servicesNotEnabled -ne $null)
				{
					$servicesNotEnabled+=', ' + $serviceName
				}
				else
				{
					$servicesNotEnabled=$serviceName
				}
			}
		}
	}

	if($servicesNotEnabled -ne $null)
	{
		Status "Unable to enable service(s): '$servicesNotEnabled'."
		return $false
	}

	return $true
}


# StartServices( roleName )
#	Restores Exchange services and dependents for the specified role
#	to the state recorded during Export-ServiceState
# Params:
#	Roles - name of the role to restore services for
# Returns:
#	True if was able to restore all services for the specified role
#	False on any error
# Notes:
#	If the saved service state is not available, this function
#	assumes the services were stared
function StartServices([string[]]$Roles, [bool]$IgnoreTimeout = $false)
{
	$ret = $true;
	Status "Starting services for $Roles..."
	Import-ServiceState

	[bool]$skipDependents = ($Roles.Count -eq 1) -and ($Roles[0] -eq 'Critical')
	if ($skipDependents)
	{
		Log "Dependent services will be skipped."
	}
	else
	{
		Log "Starting services and their dependents."
	}

	foreach( $serviceName in (Get-ServiceToControl $Roles) )
	{
		$depth = [int]0
		if( -not( RestoreServiceAndDependents $serviceName $depth -IgnoreTimeout:$IgnoreTimeout -SkipDependents:$skipDependents) )
		{
			Status "Unable to restore service '$serviceName'."
			$ret = $false;
		}
	}

	return $ret
}


# RestoreServiceAndDependents(serviceName, depth)
#	Recursively start a service and its dependents if they
#	were started when the snapshot was made
# Params:
#	Args[0] - Name of the service to start
#	Args[1] - The recursion depth of this call
# Returns:
#	True if was able to restore all services and dependents
#	False on any error
# Notes:
#	If the saved service state is not available, this function
#	assumes the services were stared
function RestoreServiceAndDependents($ServiceName, $Depth, [bool]$IgnoreTimeout = $false, [bool]$SkipDependents = $false)
{
	Log "($depth) Enter: RestoreServiceAndDependents '$serviceName'."
	$service = get-service $serviceName -ea SilentlyContinue

	# Filter out cases where we don't need to start the service
	if( -Not(Test-ServiceWasRunning $serviceName) )
	{
		Log "($depth) $serviceName was not previously running."
	}
	elseif( $script:previouslyStartedServices[$serviceName] )
	{
		Log "($depth) Already started $serviceName..."
	}
	elseif( $service -eq $null )
	{
		Log "($depth) Ignoring non-existent service '$serviceName'."
	}
	else
	{
		# We actually need to start this service.  Use two steps:
		# 1.  Start the service
		# 2.  Recursively do the same thing for each dependent service

		# Start the service
		Log "($depth) Starting $serviceName..."
		start-setupservice -serviceName $serviceName -ev script:serviceControlError -IgnoreTimeout:$IgnoreTimeout
		if( $script:serviceControlError[0] -ne $null )
		{
			Log ("{0}: {1}" -f $depth, $script:serviceControlError[0].ToString())
			return $false
		}
		$script:previouslyStartedServices[$serviceName] = $true

        if (!$SkipDependents)
        {
		    # Start its dependents
		    Log ("({2}) {0} has {1} dependents." -F $service.Name, $service.DependentServices.Length, $depth)
		    foreach( $d in $service.DependentServices )
		    {
			    $newdepth = (1+$depth)
			    if( -not( RestoreServiceAndDependents $d.Name $newdepth $IgnoreTimeout ) )
			    {
				    Log "($depth) Unable to start dependent of '$serviceName'"
				    return $false
			    }
		    }
		}
	}
	Log "($depth) Exit: RestoreServiceAndDependents '$serviceName'"
	return $true
}

# DisableOneCopyAlertScheduledTask
#	Disables the one copy alert windows scheduled task (without deleting it).
#	Note that if the task is not installed, this is simply a noop and no error
#	is returned.
# Returns:
#	True if was able to disable the scheduled task (or task didn't exist)
#	False on any error
function DisableOneCopyAlertScheduledTask
{
	Log "Disabling the one copy alert windows scheduled task..."

	$ScriptsDirPath = GetExchangeScriptsPath
	if(($ScriptsDirPath -eq $null) -or !(test-path $ScriptsDirPath) )
	{
		Status ("Exchange server is not installed. Could not find the Scripts directory.")
		return $false
	}

	# $RoleBinPath will be something like:
	#    \\exrel\release\qfe\14.01.0202\14.01.0202.000\exchange14\all\retail\amd64\Setup\ServerRoles\Common
	$InstallerScriptFileName = 'ManageScheduledTask.ps1'
	$PsInstallerScriptPath = Join-Path $ScriptsDirPath $InstallerScriptFileName

	& $PsInstallerScriptPath -RunFromSetup -Disable -ServerName $env:COMPUTERNAME `
							 -TaskName "Database One Copy Alert"

	return $true
}

###############################################################################
#
# Utility Functions
#	These functions factor out common logic
#

# IsKnownRole( role )
#	Determine if the specified name is a valid role name
# Params:
#	Args[0] - Name of the role in question
# Returns:
#	True if the name is a valid role
#	False otherwise
function IsKnownRole
{
	$role = $Args[0]

	if( $script:knownRoles -ieq $role )
	{
		Log "$role is a known role"
		return $true
	}
	else
	{
		Log "$role is an unknown role"
		return $false
	}
}

# Log( $entry )
#	Append a string to ExchangeSetupLog, or
#	Append a string to a well known text file with a time stamp
#	as type of info
# Params:
#	Args[0] - Entry to write to log
# Returns:
#	void
function Log
{
	[string]$entry = $Args[0]
	if($entry)
	{
		if ($script:writeToExchangeSetupLog -eq $true)
		{
			$line = "[{0}] {1}" -F "ServiceControl.ps1", $entry
			Write-ExchangeSetupLog -Info $line
		}
		else
		{
			$line = "[{0}] {1}" -F $(get-date).ToString("HH:mm:ss"), $entry
			add-content -Path "$logDir\ServiceControl.log" -Value $line
		}
	}
}

# LogWarning( $entry )
#	Append a string to ExchangeSetupLog, or
#	Append a string to a well known text file with a time stamp
#	as type of warning
# Params:
#	Args[0] - Entry to write to log
# Returns:
#	void
function LogWarning
{
	[string]$entry = $Args[0]
	if($entry)
	{
		if ($script:writeToExchangeSetupLog -eq $true)
		{
			Write-ExchangeSetupLog -Warning $entry
		}
		else
		{
			$line = "[{0}] [Warning] {1}" -F $(get-date).ToString("HH:mm:ss"), $entry
			add-content -Path "$logDir\ServiceControl.log" -Value $line
		}
	}
}

# LogException( $entry )
#	Append an exception to ExchangeSetupLog, or
#	Append an exception to a well known text file with a time stamp
#	as type of error
# Params:
#	Args[0] - Entry to write to log
# Returns:
#	void
function LogException
{
	$entry = $Args[0]
	if($entry -and ($entry -is [System.Exception]))
	{
		if ($script:writeToExchangeSetupLog -eq $true)
		{
			Write-ExchangeSetupLog -Error $entry
		}
		else
		{
			$line = "[{0}] [Error] {1}" -F $(get-date).ToString("HH:mm:ss"), $entry.ToString()
			add-content -Path "$logDir\ServiceControl.log" -Value $line
		}
	}
}

# Status( $entry )
#	Log and write entry to console
# Params:
#	Args[0] - Entry to write to log
# Returns:
#	void
function Status
{
	$entry = $Args[0]

	Log $entry
	#write-host $entry
}

# Usage
#	Display the usage of this script
# Returns:
#	Void
function Usage
{
	write-output 'Usage:  ServiceControl.ps1 <op> <role>[,<role>...]'
	write-output "<op> `tOne of {Start Stop Save Remove}"
	write-output "`t- Start`tStarts the services for the specified role(s)"
	write-output "`t- Stops`tStarts the services for the specified role(s)"
	write-output "`t- Save`tSaves a snapshot of the state of all services"
	write-output "`t- Remove`tRemoves the previously taken snapshot"
	write-output "<role> `tOne of {$knownRoles}"
	write-output "`r`n`tA role only needs to be specified for the start and stop operations."
}

# GetExchangeInstallPath
#	Gets the install path for Exchange Server. If Exchange is not installed then
#	null is returned.
# Returns
#	string
function GetExchangeInstallPath
{
	$path = $null

	if(test-path $script:setupKey)
	{
		$msiInstallPath = (get-itemproperty -Path $script:setupKey).MsiInstallPath
		if($msiInstallPath -ne $null)
		{
			$path = [System.IO.Path]::Combine( `
			$msiInstallPath, `
			'bin')
		}
	}

	return $path
}

# GetExchangeScriptsPath
#	Gets the path to the Exchange Server's Scripts directory. If Exchange is not
# 	installed, null is returned.
# Returns
#	string
function GetExchangeScriptsPath
{
	$path = $null

	# If a directory has been passed in, use that
	if ($SetupScriptsDirectory)
	{
		return $SetupScriptsDirectory
	}

	# Otherwise, use the locally installed "Exchange Server\v15\Scripts" directory
	if(test-path $script:setupKey)
	{
		$msiInstallPath = (get-itemproperty -Path $script:setupKey).MsiInstallPath
		if($msiInstallPath -ne $null)
		{
			$path = Join-Path $msiInstallPath "Scripts"
		}
	}

	return $path
}

# GetInstalledBuildVersion
#	Gets the major build version for Exchange Server. If Exchange is not installed
#	then 0 is returned.
# Returns:
#	int
function GetInstalledBuildVersion
{
	$build = 0

	if(test-path $script:setupKey)
	{
		$build = $(get-itemproperty -Path $script:setupKey).MsiBuildMajor
	}

	return $build
}

# SetSnapinRegistryValues
#	Creates the snapin key if not present and sets the values.
# Params:
#	Args[0] - Exchange server install path
#	Args[1] - Exchange server build version
# Returns:
#	bool (true if snapin key was created, false otherwise)
function SetSnapinRegistryValues
{
	$snapinKeyCreated = $false

	$path = $Args[0]
	$build = $Args[1]

	$snapinValues =
	@{
		 "ApplicationBase"="$path"
		 "AssemblyName"="Microsoft.Exchange.PowerShell.Configuration, Version=15.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
		 "Description"="Setup Tasks for the Exchange Server (v$build)"
		 "CustomPSSnapInType"="Microsoft.Exchange.Management.PowerShell.SetupPSSnapIn"
		 "ModuleName"="$path\Microsoft.Exchange.PowerShell.configuration.dll"
		 "PowerShellVersion"="1.0"
		 "Vendor"="Microsoft"
		 "Version"="15.0.0.0"
	}

	if(!(test-path $script:snapinKey))
	{
		new-item -path $script:snapinKey -Force
		$snapinKeyCreated = $true

		foreach($setting in $snapinValues.Keys)
		{
			set-itemproperty -Path $script:snapinKey -Name $setting -Value $snapinValues[$setting] -Force
		}
	}

	return $snapinKeyCreated
}

# Test-WriteExchangeSetupLog
#	Tests if Write-ExchangeSetupLog command is available.
# Returns:
#	bool (true if Write-ExchangeSetupLog command is available, false otherwise)
function Test-WriteExchangeSetupLog
{
	$checkCommand = Get-Command Write-ExchangeSetupLog -ErrorAction SilentlyContinue
	return ($checkCommand -ne $null)
}

# AddSnapins
#	Adds the required snapins to access start-setupservice and stop-setupservice
# Returns:
#	void
function AddSnapins
{
    add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Setup -ea SilentlyContinue
}

# RemoveSnapins
#	Removes snapins loaded by AddSnapins
# Returns:
#	void
function RemoveSnapins
{
	remove-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Setup
}

# GetInstalledRoles
# 	Gets the installed roles from the registry and stores the information in script:installedRoles
# Returns
# 	Void
function GetInstalledRoles
{
	# Exchange might not be installed on the machine, so there won't be an existing
	# registry. We silentlycontinue on errors.
	$roleRegistryInfoSet = get-itemproperty HKLM:SOFTWARE\Microsoft\ExchangeServer\v15\* -ea SilentlyContinue

	foreach( $roleRegistryInfo in $roleRegistryInfoSet )
	{
		switch( $roleRegistryInfo.PSChildName)
		{
			{$_ -ieq "AdminTools"}
			{
				$script:installedRoles += "AdminTools"
				Log "Adding to installed roles list: AdminTools"
				break
			}

			{$_ -ieq "EdgeTransportRole"}
			{
				$script:installedRoles += "Gateway"
				Log "Adding to installed roles list: Gateway"
				break
			}

			{$_ -ieq "HubTransportRole"}
			{
				$script:installedRoles += "Bridgehead"
				Log "Adding to installed roles list: Bridgehead"
				break
			}

			{$_ -ieq "ClientAccessRole"}
			{
				$script:installedRoles += "ClientAccess"
				Log "Adding to installed roles list: ClientAccessMailboxRole"
				break
			}

			{$_ -ieq "UnifiedMessagingRole"}
			{
				$script:installedRoles += "UnifiedMessaging"
				Log "Adding to installed roles list: UnifiedMessaging"
				break
			}

			{$_ -ieq "MailboxRole"}
			{
				$script:installedRoles += "Mailbox"
				Log "Adding to installed roles list: Mailbox"
				break
			}

			{$_ -ieq "MonitoringRole"}
			{
				$script:installedRoles += "Monitoring"
				Log "Adding to installed roles list: Monitoring"
				break
			}

                        {$_ -ieq "FrontendTransportRole"}
			{
				$script:installedRoles += "FrontendTransport"
				Log "Adding to installed roles list: Mailbox"
				break
			}

			{$_ -ieq "Cafe"}
			{
				$script:installedRoles += "Cafe"
				Log "Adding to installed roles list: Cafe"
				break
			}

			default
			{
				# There are other things like 'Hygiene' in the registry as well, so we ignore this
				# information
				break
			}
		}
	}
}

#
# Main
#

# Command Write-ExchangeSetupLog is not available at this point during patching.
# So we could log stuff into a separate file.
# Must set this flag before everything else
$script:writeToExchangeSetupLog = Test-WriteExchangeSetupLog

Log "-----------------------------------------------"
Log ("* ServiceControl.ps1: {0}" -F $(get-date) )
Log "Performing service control with options: $Args"

$ret = $false

if (!$Operation)
{
	Status "Must specify an operation"
	Usage
	return
}

if ($Roles)
{
	$Roles | foreach {
		if (-not (IsKnownRole $_))
		{
			Status ("Unknown role '{0}'." -F $_)
			Usage
			return
		}
	}
}
else
{
    if ($Operation -match "Stop|DisableServices|Start|EnableServices")
    {
		Status  "Must specify one or more roles."
		Usage
		return
    }
}

$activity = $null

write-Progress -Activity 'Service Control' -Id 0 -Status 'In Progress' -PercentComplete 0
switch ($Operation)
{
	{$_ -ieq "Save"}
	{
		$activity = 'Archiving Service State'
		write-Progress -Activity $activity -Id 0 -Status 'In Progress' -PercentComplete 10
		Export-ServiceState
		Export-ServiceStartupMode
		$ret = $true
		break
	}

	{$_ -ieq "Remove"}
	{
		$activity = 'Removing State Archive'
		write-Progress -Activity $activity -Id 0 -Status 'In Progress' -PercentComplete 10
		Remove-ServiceState
		Remove-ServiceStartMode
		$ret = $true
		break
	}

	{$_ -ieq "Stop" }
	{
		$activity = 'Stopping Services'
		write-Progress -Activity $activity -Id 0 -Status 'In Progress' -PercentComplete 10

		# Loop through each role and terminate on error
		$ret = $true
		Log "Will stop services for the roles: $roles."
		if( -not( StopServices $roles ) )
		{
			$ret = $false
			Status "Unable to stop all services for roles: $roles."
		}

		break
	}

	{$_ -ieq "DisableServices" }
	{
		$activity = 'Disabling Services'
		write-Progress -Activity $activity -Id 0 -Status 'In Progress' -PercentComplete 10

		# Loop through each role
		$ret = $true
		Log "Will disable services for the roles: $roles."

		if( -not( DisableOneCopyAlertScheduledTask ) )
		{
			$ret = $false
			Status "Unable to disable one copy windows scheduled task for roles: $roles."
		}

		if( -not( DisableServices $roles ) )
		{
			$ret = $false
			Status "Unable to disable all services for roles: $roles."
		}

		break
	}

	{$_ -ieq "Start" }
	{
		$activity = 'Stating Services'
		write-Progress -Activity $activity -Id 0 -Status 'In Progress' -PercentComplete 10

		# Loop through each role and terminate on error
		$ret = $true
		if( -not( StartServices $roles -IgnoreTimeout $true ) )
		{
			$ret = $false
			Status "Unable to restore all services for roles: $roles."
		}

		break
	}

	{$_ -ieq "EnableServices" }
	{
		$activity = 'Enabling Services'
		write-Progress -Activity $activity -Id 0 -Status 'In Progress' -PercentComplete 10

		# Loop through each role and terminate on error
		$ret = $true
		Log "Will enable services for the roles: $roles."
		if( -not( EnableServices $roles ) )
		{
			$ret = $false
			Status "Unable to enable all services for roles: $roles."
		}

		break
	}

	{$_ -ieq "BeforePatch"}
	{
		# This is the set of stuff we need to do before applying a patch
		# We stop all the services and disable them. Before doing this we need
		# to store the service state and registry information
		$ret = $true

		$path = GetExchangeInstallPath
		$build = GetInstalledBuildVersion

		if(($path -eq $null) -or ($build -eq 0))
		{
			$ret = $false
			Status ("Exchange server is not installed")
		}
		else
		{
			# Set the snapin registry values so that the script can access start-setupservice
			# and stop-startupservice
			$keyCreated = SetSnapinRegistryValues $path $build
			AddSnapins

			trap
			{
				LogException $_.Exception
				break
			}

			# Save the service information
			Log "Saving service and registry data"
			Export-ServiceState -Overwrite
			Export-ServiceStartupMode -Overwrite

			GetInstalledRoles

			# Stop and disable the services

			Log "Stopping services for the following roles: $script:installedRoles"
			if( -not( StopServices $script:installedRoles ) )
			{
				$ret = $false
				Status "Unable to stop all services for $script:installedRoles."
			}
			elseif( -not( DisableServices $script:installedRoles ) )
			{
				$ret = $false
				Status "Unable to disable all services for $script:installedRoles."
			}
			RemoveSnapins
		}

		break
	}

	{$_ -ieq "AfterPatch"}
	{
		# Here we do stuff that is required after a patch is applied
		# We enable services for the installed roles, start the services based on
		# the data saved and then remove the files that store service related data
		$ret = $true

		$path = GetExchangeInstallPath
		$build = GetInstalledBuildVersion

		if(($path -eq $null) -or ($build -eq 0))
		{
			$ret = $false
			Status ("Exchange server is not installed")
		}
		else
		{
			# Set the snapin registry values so that the script can access start-setupservice
			# and stop-startupservice
			$keyCreated = SetSnapinRegistryValues $path $build
			AddSnapins

			trap
			{
				LogException $_.Exception
				break
			}

			GetInstalledRoles

			# Enable and start services
			Log "Enabling services for the following roles: $script:installedRoles"

			if( -not( EnableServices $script:installedRoles ) )
			{
				$ret = $false
				Status "Unable to enable all services for $script:installedRoles."
			}
			elseif( -not( StartServices $script:installedRoles ) )
			{
				$ret = $false
				Status "Unable to restore all services for $script:installedRoles."
			}

			# Remove saved service data
			Remove-ServiceState
			Remove-ServiceStartMode

			RemoveSnapins
		}
		break
	}

	default
	{
		Status "Unknown operation: '$_'"
		Usage
		break
	}

}

if ($activity -ne $null)
{
	write-Progress -Activity $activity -Id 0 -Status 'Completed' -Completed
}

if( $ret -eq $true )
{
	Log "Script completed succesfully."
}
else
{
	Log "Script completed with one or more errors."
}

# SIG # Begin signature block
# MIIdpAYJKoZIhvcNAQcCoIIdlTCCHZECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUyni0r9wguojVEN+Z0f0P0noo
# yYKgghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI4
# WhcNMTcwNjMwMTkyMTI4WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# Ojk4RkQtQzYxRS1FNjQxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAipCth86FRu1y
# rtsPu2NLSV7nv6A/oVAcvGrG7VQwRn+wGlXrBf4nyiybHGn9MxuB9u4EMvH8s75d
# kt73WT7lGIT1yCIh9VC9ds1iWfmxHZtYutUOM92+a22ukQW00T8U2yowZ6Gav4Q7
# +9M1UrPniZXDwM3Wqm0wkklmwfgEEm+yyCbMkNRFSCG9PIzZqm6CuBvdji9nMvfu
# TlqxaWbaFgVRaglhz+/eLJT1e45AsGni9XkjKL6VJrabxRAYzEMw4qSWshoHsEh2
# PD1iuKjLvYspWv4EBCQPPIOpGYOxpMWRq0t/gqC+oJnXgHw6D5fZ2Ccqmu4/u3cN
# /aAt+9uw4wIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFHbWEvi6BVbwsceywvljICto
# twQRMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBABbNYMMt3JjfMAntjQhHrOz4aUk970f/hJw1jLfYspFpq+Gk
# W3jMkUu3Gev/PjRlr/rDseFIMXEq2tEf/yp72el6cglFB1/cbfDcdimLQD6WPZQy
# AfrpEccCLaouf7mz9DGQ0b9C+ha93XZonTwPqWmp5dc+YiTpeAKc1vao0+ru/fuZ
# ROex8Zd99r6eoZx0tUIxaA5sTWMW6Y+05vZN3Ok8/+hwqMlwgNR/NnVAOg2isk9w
# ox9S1oyY9aRza1jI46fbmC88z944ECfLr9gja3UKRMkB3P246ltsiH1fz0kFAq/l
# 2eurmfoEnhg8n3OHY5a/Zzo0+W9s1ylfUecoZ4UwggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhAwggP4
# oAMCAQICEzMAAABkR4SUhttBGTgAAAAAAGQwDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNTEwMjgyMDMxNDZaFw0xNzAx
# MjgyMDMxNDZaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCTLtrY5j6Y2RsPZF9NqFhN
# FDv3eoT8PBExOu+JwkotQaVIXd0Snu+rZig01X0qVXtMTYrywPGy01IVi7azCLiL
# UAvdf/tqCaDcZwTE8d+8dRggQL54LJlW3e71Lt0+QvlaHzCuARSKsIK1UaDibWX+
# 9xgKjTBtTTqnxfM2Le5fLKCSALEcTOLL9/8kJX/Xj8Ddl27Oshe2xxxEpyTKfoHm
# 5jG5FtldPtFo7r7NSNCGLK7cDiHBwIrD7huTWRP2xjuAchiIU/urvzA+oHe9Uoi/
# etjosJOtoRuM1H6mEFAQvuHIHGT6hy77xEdmFsCEezavX7qFRGwCDy3gsA4boj4l
# AgMBAAGjggF/MIIBezAfBgNVHSUEGDAWBggrBgEFBQcDAwYKKwYBBAGCN0wIATAd
# BgNVHQ4EFgQUWFZxBPC9uzP1g2jM54BG91ev0iIwUQYDVR0RBEowSKRGMEQxDTAL
# BgNVBAsTBE1PUFIxMzAxBgNVBAUTKjMxNjQyKzQ5ZThjM2YzLTIzNTktNDdmNi1h
# M2JlLTZjOGM0NzUxYzRiNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzcitW2oynUC
# lTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEGCCsGAQUF
# BwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0MAwGA1Ud
# EwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjiDGRDHd1crow7hSS1nUDWvWas
# W1c12fToOsBFmRBN27SQ5Mt2UYEJ8LOTTfT1EuS9SCcUqm8t12uD1ManefzTJRtG
# ynYCiDKuUFT6A/mCAcWLs2MYSmPlsf4UOwzD0/KAuDwl6WCy8FW53DVKBS3rbmdj
# vDW+vCT5wN3nxO8DIlAUBbXMn7TJKAH2W7a/CDQ0p607Ivt3F7cqhEtrO1Rypehh
# bkKQj4y/ebwc56qWHJ8VNjE8HlhfJAk8pAliHzML1v3QlctPutozuZD3jKAO4WaV
# qJn5BJRHddW6l0SeCuZmBQHmNfXcz4+XZW/s88VTfGWjdSGPXC26k0LzV6mjEaEn
# S1G4t0RqMP90JnTEieJ6xFcIpILgcIvcEydLBVe0iiP9AXKYVjAPn6wBm69FKCQr
# IPWsMDsw9wQjaL8GHk4wCj0CmnixHQanTj2hKRc2G9GL9q7tAbo0kFNIFs0EYkbx
# Cn7lBOEqhBSTyaPS6CvjJZGwD0lNuapXDu72y4Hk4pgExQ3iEv/Ij5oVWwT8okie
# +fFLNcnVgeRrjkANgwoAyX58t0iqbefHqsg3RGSgMBu9MABcZ6FQKwih3Tj0DVPc
# gnJQle3c6xN3dZpuEgFcgJh/EyDXSdppZzJR4+Bbf5XA/Rcsq7g7X7xl4bJoNKLf
# cafOabJhpxfcFOowMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkqhkiG9w0B
# AQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAG
# A1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTEw
# HhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
# aWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# q/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2Avw
# OMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akrrnoJr9eW
# WcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1
# eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy4BI6t0le
# 2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9sbKvkjh+
# 0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVfGCi2
# zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8kA/DRelsv
# 1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLn
# JN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmnEyimp31n
# gOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+Hgg
# WCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0wggHpMBAG
# CSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZ
# BgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/
# BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8E
# UzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9k
# dWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsGAQUFBwEB
# BFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNVHSAEgZcw
# gZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsGAQUFBwIC
# MDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBu
# AHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKbC5YR4WOS
# mUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQ
# VdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6I/MTfaaQ
# dION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0wI/zRive
# /DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560STkKxgrC
# xq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQamASooPoI/
# E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HNpZfQ
# 7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ahXJbYANah
# Rr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3
# S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33VtY5E90Z1W
# Tk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1t
# bWrJUnMTDXpQzTGCBKowggSmAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBvjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUdYrV+WlviVyQvg6O9oelDSYfMdgwXgYKKwYB
# BAGCNwIBDDFQME6gJoAkAFMAZQByAHYAaQBjAGUAQwBvAG4AdAByAG8AbAAuAHAA
# cwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZI
# hvcNAQEBBQAEggEAcmV+6x/93WKiLRlwQ+wxTSbwUF5j9IzvZhJ7L2FjFg2NNZ/O
# VQwfUEbUPzDsDX90u0onos6yaKdej4tph44IJxiqUSdEBINecIIzu+a7LJCYg8Fr
# Xq7JUh63RUW3SBryJIXlu/0l1Kvh58GeQMyrL40BMk0EpFacXWBKFGsr4a3ANm0g
# lyn1FtkOcRDRyspWQ0bLCll3irAU+7BYttWWj+oT90BizXoT2tOjzEY84dfKNTcR
# 4NtSQclW+iBAwV/v9937LkXlcuY0RZ3FA/Faj+p4FKhUG6cHJn2l/GZemb5s+qMl
# /HZS/gTSr15J7mEhtH7duDu/0JSNNFS5tKnw/KGCAigwggIkBgkqhkiG9w0BCQYx
# ggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAACZqsWB
# n4yifYoAAAAAAJkwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0B
# BwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDQ1OVowIwYJKoZIhvcNAQkEMRYE
# FKHBDFKgHNRmLy9Ro5IZ8l84QJ7pMA0GCSqGSIb3DQEBBQUABIIBABdbck2sS0+G
# oO6eECROv3wI/mUHv/kNNJTr213CA1UMqT0vbsnNOKADDq8lIjLbkoUfc9Wm7WVr
# ICEqyYGcXpgFFBQrLaTg3EzsHhEX61bTzcbpR6V45dIZlITSYxBUvL+0Tp04RzTI
# 09YNyRk/Hl+hVuwyDZgZ6YdXtCGUpyV+gOCyCLmo1jic6aPvYBx/i9wfDFIjqKkV
# IzoxNBetTjQiWTB4eJl4/y4r7whu10MpTKNhkLUJl2PKSy0Cghtmnb+yB9iriGXS
# 8vIY/AFAFO1YnVj4kDdJdJSsSZ2cr3BcYYfccPQICe2SiKhfeAi7GlpmLlLPGgl1
# 0XsFXYdHIe8=
# SIG # End signature block
