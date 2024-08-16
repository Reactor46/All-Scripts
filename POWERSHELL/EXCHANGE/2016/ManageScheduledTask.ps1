# Copyright (c) Microsoft Corporation. All rights reserved.

param(
	[Parameter(ParameterSetName="Install",Mandatory=$true)]
	[switch] $Install,
	
	[Parameter(ParameterSetName="Install")]
	[switch] $DeleteExisting = $false,
	
	[Parameter(Mandatory=$true)]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $ServerName,
	
	[Parameter(Mandatory=$true)]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $TaskName,
	
	[Parameter(ParameterSetName="Install",Mandatory=$true)]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $PsScriptPath,
	
	[Parameter(ParameterSetName="Install")]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $PsScriptArgs,
	
	[Parameter(ParameterSetName="Uninstall",Mandatory=$true)]
	[switch] $Uninstall,
	
	# Disable a scheduled task without unregistering it
	[Parameter(ParameterSetName="Disable",Mandatory=$true)]
	[switch] $Disable,
	
	# Re-enable a pre-configured scheduled task and run it
	[Parameter(ParameterSetName="Enable",Mandatory=$true)]
	[switch] $Enable,
	
	# Returns $true or $false based on whether the specified task exists on the given server
	[Parameter(ParameterSetName="TestExistence",Mandatory=$true)]
	[switch] $TestExistence,
	
	# True if run from Exchange Setup (will cause the script to log to 
	# the Exchange setup logs)
	[Parameter(Mandatory=$false)]
	[switch] $RunFromSetup = $false
)

Set-StrictMode -Version 2.0
Import-LocalizedData -BindingVariable ManageScheduledTask_LocalizedStrings -FileName ManageScheduledTask.Strings.psd1

#-------------------
# Script variables #
#-------------------

$script:exeOutput = $null
$script:taskQueried = $null
$script:writeToExchangeSetupLog = $false
$script:exLogDir = "$env:SYSTEMDRIVE\ExchangeSetupLogs"


# Returns the name of the user this script is running as: eg: SDIMEB-DOM\Administrator
function Get-CurrentUserName
{
	[System.Security.Principal.WindowsIdentity]$id = [System.Security.Principal.WindowsIdentity]::GetCurrent()

	# E14:398785: to support a case in which an ampersand is appeared in the NetBIOS domain name.
	[System.String]$fixedName = $id.Name.Replace('&','&amp;')
	return $fixedName
}

# 
function Get-TaskXml(
	[Parameter(Mandatory=$true)] [string] $powershellCommand
)
{
	# XML describing the scheduled task to create
	#
	# Parameters:
	# {0} : current date computed using: (get-date).ToString("yyyy-MM-ddTHH:mm:ss")
	# {1} : user name
	# {2} : start time of the task
	# {3} : windows directory path: eg: $env:windir = D:\Windows
	# {4} : powershell command to execute: eg: D:\test.ps1 -MonitoringContext -ErrorAction:Stop
	[string]$xmlContent = 
'<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <!-- <Date>2010-01-15T03:00:00</Date> -->
	<Date>{0}</Date>
    <!-- <Author>SDIMEB-DOM\administrator</Author> -->
	<Author>{1}</Author>
    <Description>Database redundancy monitoring task.</Description>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <Repetition>
        <Interval>PT1H</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <!-- <StartBoundary>2010-01-15T03:00:00</StartBoundary> -->
	  <StartBoundary>{2}</StartBoundary>
      <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>NT AUTHORITY\SYSTEM</UserId>
<!--      <LogonType>Password</LogonType>    -->
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>Queue</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>false</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>true</WakeToRun>
    <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
    <Priority>7</Priority>
	<!-- <RestartOnFailure>
      <Interval>PT5M</Interval>
      <Count>5</Count>
    </RestartOnFailure> -->
  </Settings>
  <Actions Context="Author">
    <Exec>
      <!-- <Command>D:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe</Command> -->
	  <Command>{3}\System32\WindowsPowerShell\v1.0\powershell.exe</Command>
      <!-- <Arguments>-NonInteractive -WindowStyle Hidden -command "D:\test.ps1 -MonitoringContext -ErrorAction:Continue"</Arguments> -->
	  <Arguments>-NonInteractive -WindowStyle Hidden -command "{4}"</Arguments>
      <!-- <WorkingDirectory>D:\</WorkingDirectory> -->
    </Exec>
  </Actions>
</Task>'

	[DateTime]$currentDate = Get-Date
	$currentDateString = $currentDate.ToString("yyyy-MM-ddTHH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
	$userName = Get-CurrentUserName
	[DateTime]$newDate = $currentDate.AddMinutes(2)
	$newDateString = $newDate.ToString("yyyy-MM-ddTHH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
	
	Log-Verbose "Get-TaskXml: `$currentDateString=$currentDateString, `$userName=$userName, `$newDateString=$newDateString, WindowsDir=$($env:windir), `$powershellCommand=$powershellCommand"
	[string]$returnXmlString = [string]::Format($xmlContent, $currentDateString, $userName, $newDateString, $env:windir, $powershellCommand)
	return $returnXmlString
}

# Returns $true or $false based on whether the specified task exists on the given server
function Is-TaskExisting(
	[Parameter(Mandatory=$true)] [string] $_serverName,
	[Parameter(Mandatory=$true)] [string] $_taskName
)
{
	Log-Verbose "Is-TaskExisting: Entering: `$_serverName=$_serverName, `$_taskName=$_taskName"
	
	# Create the task scheduler COM object
	$st = New-Object -ComObject "Schedule.Service"
	# connect to remote server
	$st.Connect($_serverName, $null, $null, $null)
	$rootFolder = $st.GetFolder("\")
	$script:taskQueried = $null
	
	$checkCommand = 
	{
		$script:taskQueried = $rootFolder.GetTask($_taskName)
	}
	[bool]$taskFound = TryExecute-ScriptBlock -runCommand $checkCommand -silentOnErrors $true
	
	Log-Verbose "Is-TaskExisting: Returning '$taskFound'."
	Write-Output $taskFound
}


# Enable a pre-configured scheduled task and start running it (uses the COM interface).
# Returns $true if the task already exists, or if it was successfully created; $false otherwise.
function Enable-TaskUsingCOM (
	[Parameter(Mandatory=$true)] [string] $_serverName,
	[Parameter(Mandatory=$true)] [string] $_taskName
)
{
	Log-Verbose "Enable-TaskUsingCOM: Entering: `$_serverName=$_serverName, `$_taskName=$_taskName"
	
	# Create the task scheduler COM object
	$st = New-Object -ComObject "Schedule.Service"
	# connect to remote server
	$st.Connect($_serverName, $null, $null, $null)
	$rootFolder = $st.GetFolder("\")
	$script:taskQueried = $null
	
	$checkCommand = 
	{
		$script:taskQueried = $rootFolder.GetTask($_taskName)
	}
	[bool]$taskFound = TryExecute-ScriptBlock -runCommand $checkCommand -silentOnErrors $true
	
	if ($taskFound)
	{
		# mark the task as enabled
		$script:taskQueried.Enabled = $true
	
		# now, run the task
		$script:taskQueried.Run($null) | out-null
		
		Log-Verbose "Enable-TaskUsingCOM: Task successfully enabled and started."
	}
	else
	{
		Log-Error ($ManageScheduledTask_LocalizedStrings.res_0000 -f $_taskName,$_serverName,"Enable-TaskUsingCOM")
	}
}


# Create a scheduled task using the COM scripting interface
# Returns $true if the task already exists, or if it was successfully created; $false otherwise.
function Create-TaskUsingCOM (
	[Parameter(Mandatory=$true)] [string] $_serverName,
	[Parameter(Mandatory=$true)] [string] $_taskName,
	[Parameter(Mandatory=$true)] [string] $powershellCommand
)
{
	Log-Verbose "Create-TaskUsingCOM: Entering: `$_serverName=$_serverName, `$_taskName=$_taskName, `$powershellCommand=$powershellCommand"
	
	[string]$xmlString = Get-TaskXml $powershellCommand
	
	Log-Verbose "Create-TaskUsingCOM : `$xmlString: `n$xmlString"
	
	# check if the task has already been registered
	[bool]$alreadyRegistered = Is-TaskRegistered $_serverName $_taskName $xmlString
	if ($alreadyRegistered)
	{
		Log-Verbose "Create-TaskUsingCOM: Task '$_taskName' has already been registered on server '$_serverName'."
		return $true
	}
	
	# Create the task scheduler COM object
	$st = New-Object -ComObject "Schedule.Service"
	# connect to remote server
	$st.Connect($_serverName, $null, $null, $null)
	# Get the root folder to create the task in
	$rootFolder = $st.GetFolder("\")
	# Define the new task - MSDN says 0 must be passed in.
	$taskDef = $st.NewTask(0)
	$taskDef.XmlText = $xmlString

	# Finally, register the task:
	# 	flags 		= TASK_CREATE_OR_UPDATE (6)
	# 	logonType 	= TASK_LOGON_SERVICE_ACCOUNT (5)
	$task = $rootFolder.RegisterTaskDefinition($_taskName, $taskDef, 6, $null, $null, 5)	
	
	# Check that the task was successfully registered (wait upto 10 seconds)
	Log-Verbose "Create-TaskUsingCOM: Waiting upto 10 seconds for task to be registered..."
	
	$condition = { return (Is-TaskRegistered $_serverName $_taskName $xmlString) }
	[bool]$isRegistered = WaitForCondition $condition 10
	if ($isRegistered)
	{
		Log-Verbose "Create-TaskUsingCOM: Task '$_taskName' registered successfully on server '$_serverName'."
		return $true
	}
	else
	{
		Log-Error ($ManageScheduledTask_LocalizedStrings.res_0001 -f $_taskName,$_serverName,"Create-TaskUsingCOM")
		return $false
	}
}


# Delete a scheduled task that has already been registered.
function Delete-RegisteredTask(
	[Parameter(Mandatory=$true)] [string] $_serverName,
	[Parameter(Mandatory=$true)] [string] $_taskName)
{
	StopAndDelete-RegisteredTaskInternal $_serverName $_taskName $true
}

# Stop and optionally delete a scheduled task that has already been registered.
function StopAndDelete-RegisteredTaskInternal(
	[Parameter(Mandatory=$true)]  [string] $_serverName,
	[Parameter(Mandatory=$true)]  [string] $_taskName,
	[Parameter(Mandatory=$false)] [bool]   $_delete=$false)
{
	Log-Verbose "StopAndDelete-RegisteredTaskInternal: Entering: `$_serverName=$_serverName, `$_taskName=$_taskName, `$_delete=$_delete"
	
	# Create the task scheduler COM object
	$st = New-Object -ComObject "Schedule.Service"
	# connect to remote server
	$st.Connect($_serverName, $null, $null, $null)
	$rootFolder = $st.GetFolder("\")
	$script:taskQueried = $null
	
	$checkCommand = 
	{
		$script:taskQueried = $rootFolder.GetTask($_taskName)
	}
	[bool]$taskFound = TryExecute-ScriptBlock -runCommand $checkCommand -silentOnErrors $true
	
	if ($taskFound)
	{
		Log-Verbose "StopAndDelete-RegisteredTaskInternal: Found the existing task."
			
		# stop the running instances
		$script:taskQueried.Stop(0)
		
		# mark the task as disabled
		$script:taskQueried.Enabled = $false
				
		if ($_delete)
		{
			# delete the task
			$rootFolder.DeleteTask($_taskName, 0)
			Log-Verbose "StopAndDelete-RegisteredTaskInternal: Task '$_taskName' successfully deleted."
		}
	}
	else
	{
		Log-Verbose "StopAndDelete-RegisteredTaskInternal: The task does not exist. Exiting."
	}
}


# Returns $true if a specified task is already registered and enabled on the given server.
function Is-TaskRegistered( 
	[Parameter(Mandatory=$true)] [string] $_serverName,
	[Parameter(Mandatory=$true)] [string] $_taskName,
	[Parameter(Mandatory=$true)] [string] $_taskXmlString
)
{
	[xml]$taskXml = [xml]$_taskXmlString
	
	# Create the task scheduler COM object
	$st = New-Object -ComObject "Schedule.Service"
	# connect to remote server
	$st.Connect($_serverName, $null, $null, $null)
	
	$checkCommand = 
	{
		$rootFolder = $st.GetFolder("\")
		$searchTask = $rootFolder.GetTask($_taskName)
	
		# Check that the task was successfully registered
		if ($searchTask)
		{
			Log-Verbose "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has state: $($searchTask.State), Enabled=$($searchTask.Enabled)"
			
			# First, check if some critical arguments are the same
			[xml]$oldXml = [xml]$searchTask.Xml
			if ($oldXml.Task.Actions.Exec.Command -ine $taskXml.Task.Actions.Exec.Command)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different Execution Command.`nOld: $($oldXml.Task.Actions.Exec.Command)`nNew: $($taskXml.Task.Actions.Exec.Command)"
			}
			if ($oldXml.Task.Actions.Exec.Arguments -ine $taskXml.Task.Actions.Exec.Arguments)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different Execution Arguments.`nOld: $($oldXml.Task.Actions.Exec.Arguments)`nNew: $($taskXml.Task.Actions.Exec.Arguments)"
			}
			if ($oldXml.Task.Triggers.TimeTrigger.Repetition.Interval -ine $taskXml.Task.Triggers.TimeTrigger.Repetition.Interval)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different Repetion Interval Arguments.`nOld: $($oldXml.Task.Triggers.TimeTrigger.Repetition.Interval)`nNew: $($taskXml.Task.Triggers.TimeTrigger.Repetition.Interval)"
			}
			if ($oldXml.Task.Triggers.TimeTrigger.ExecutionTimeLimit -ine $taskXml.Task.Triggers.TimeTrigger.ExecutionTimeLimit)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different ExecutionTimeLimit Arguments.`nOld: $($oldXml.Task.Triggers.TimeTrigger.ExecutionTimeLimit)`nNew: $($taskXml.Task.TimeTrigger.Triggers.ExecutionTimeLimit)"
			}
			if ($oldXml.Task.Settings.ExecutionTimeLimit -ine $taskXml.Task.Settings.ExecutionTimeLimit)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different ExecutionTimeLimit Settings Arguments.`nOld: $($oldXml.Task.Settings.ExecutionTimeLimit)`nNew: $($taskXml.Task.Settings.ExecutionTimeLimit)"
			}		
			if ($oldXml.Task.Settings.MultipleInstancesPolicy -ine $taskXml.Task.Settings.MultipleInstancesPolicy)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different MultipleInstancesPolicy Arguments.`nOld: $($oldXml.Task.Settings.MultipleInstancesPolicy)`nNew: $($taskXml.Task.Settings.MultipleInstancesPolicy)"
			}
			if ($oldXml.Task.Settings.AllowHardTerminate -ine $taskXml.Task.Settings.AllowHardTerminate)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different AllowHardTerminate Arguments.`nOld: $($oldXml.Task.Settings.AllowHardTerminate)`nNew: $($taskXml.Task.Settings.AllowHardTerminate)"
			}
			# The RestartOnFailure settings are being removed
			if ( ( $oldXml.GetElementsByTagName("RestartOnFailure") | Measure-Object ).Count -gt 0 )
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has RestartOnFailure Arguments."
			}
			
			# Now, check that it is enabled and running
			# State can be Unknown=0,Disabled=1,Queued=2,Ready=3,Running=4
			if ( !$searchTask.Enabled -or ($searchTask.State -lt 2) )
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' is not registered properly."
			}
		}
		else
		{
			Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' was not found."
		}
	}
	
	[bool]$success = TryExecute-ScriptBlock -runCommand $checkCommand -silentOnErrors $true
	return $success
}

# Common function to run a scriptblock, log any error that occurred, and return 
# a boolean to indicate whether it was successful or not.
# NOTE: ErrorActionPreference of "Stop" is used to catch all errors.
#
# Optional parameters:
#
# 	cleanupCommand 
#		This scriptblock will be executed with ErrorActionPreference of "Continue", 
#		if an error occurred while running $runCommand.
#
#	throwOnError
#		If true, the error from $runCommand will be rethrown. Otherwise 'false' is returned on error.
#
#	silentOnErrors
#		If true, the error from $runCommand will not be logged via Log-ErrorRecord (i.e. Write-Error)
function TryExecute-ScriptBlock ([ScriptBlock]$runCommand, [ScriptBlock]$cleanupCommand={}, [bool]$throwOnError=$false, [bool]$silentOnErrors=$false)
{
	# Run the following in a separate script block so that we can change
    # ErrorActionPerefence without affecting the rest of the script.
	&{
		$ErrorActionPreference = "Stop"
		[bool]$success = $false;
		
		try
		{
			$ignoredObjects = @(&$runCommand)
			$success = $true;
		}
		catch
		{
			# Any error will end up in this catch block
			# For some reason, PS does not write out any errors unless I use this
			# scriptblock with "Continue" ErrorActionPreference.
			&{
				$ErrorActionPreference = "Continue"
				
				if (!$silentOnErrors)
				{
					Log-ErrorRecord $_
				}
				
				# Run the cleanup scriptblock
				$ignoredObjects = @(&$cleanupCommand)
			}
			
			if ($throwOnError)
			{
				throw
			}
		}
		if (!$throwOnError -or $success)
		{
			return $success
		}
	}
}

# Common function to wait for condition
function WaitForCondition {
    param ([ScriptBlock] $condition, [int] $seconds)
     
    $endTime = [DateTime]::Now.addseconds($seconds)
    while ([DateTime]::Now -lt $endTime)
    {
	   if ( &$condition )
	   {
	       return $true
	   }
	   Sleep-ForSeconds 2
    }
     
    # Check one last time
    if ( &$condition )
    {
        return $true
    }
    return $false
}

# Is-WriteExchangeSetupLogPresent
#	Tests if Write-ExchangeSetupLog command is available.
# Returns:
#	bool (true if Write-ExchangeSetupLog command is available, false otherwise)
function Is-WriteExchangeSetupLogPresent
{
	$checkCommand = Get-Command Write-ExchangeSetupLog -ErrorAction SilentlyContinue
	return ($checkCommand -ne $null)
}


# Sleep for the specified duration (in seconds)
function Sleep-ForSeconds ( [int]$sleepSecs )
{
	Log-Verbose "Sleeping for $sleepSecs seconds..."
	Start-Sleep $sleepSecs
}

# Common function to retrieve the current UTC time string
function Get-CurrentTimeString
{
	return [DateTime]::UtcNow.ToString("[HH:mm:ss.fff UTC]")
}

# Common function for verbose logging
function Log-Verbose ( [string]$msg )
{	
	if (!$RunFromSetup)
	{
		$timeStamp = Get-CurrentTimeString
		Write-Verbose "$timeStamp $msg"
	}
	elseif ($script:writeToExchangeSetupLog)
	{
		$line = "[{0}] {1}" -F "ManageScheduledTask.ps1", $msg
		Write-ExchangeSetupLog -Info $line
	}
	else
	{
		$line = "[{0}] [{1}] {2}" -F $(get-date).ToString("HH:mm:ss"), "ManageScheduledTask.ps1", $msg
		add-content -Path "$($script:exLogDir)\ServiceControl.log" -Value $line
	}
}

# Common function for warning logging
function Log-Warning ( [string]$msg )
{
	if (!$RunFromSetup)
	{
		$timeStamp = Get-CurrentTimeString
		Write-Warning "$timeStamp $msg"
	}
	elseif ($script:writeToExchangeSetupLog)
	{
		$line = "[{0}] {1}" -F "ManageScheduledTask.ps1", $msg
		Write-ExchangeSetupLog -Warning $line
	}
	else
	{
		$line = "[{0}] [{1}] [Warning] {2}" -F $(get-date).ToString("HH:mm:ss"), "ManageScheduledTask.ps1", $msg
		add-content -Path "$($script:exLogDir)\ServiceControl.log" -Value $line
	}
}

# Common function for error logging
function Log-Error ( [string]$msg, [switch]$Stop)
{
	if (!$RunFromSetup)
	{
		$timeStamp = Get-CurrentTimeString

		if (!$Stop)
		{
			Write-Error "$timeStamp $msg"
		}
		else
		{
			Write-Error "$timeStamp $msg" -ErrorAction:Stop
		}
	}
	elseif ($script:writeToExchangeSetupLog)
	{
		$line = "[{0}] {1}" -F "ManageScheduledTask.ps1", $msg
		$exception = New-Object -TypeName System.Exception -ArgumentList $line
		Write-ExchangeSetupLog -Error $exception
	}
	else
	{
		$line = "[{0}] [{1}] [Error] {2}" -F $(get-date).ToString("HH:mm:ss"), "ManageScheduledTask.ps1", $msg
		add-content -Path "$($script:exLogDir)\ServiceControl.log" -Value $line
	}
}

# Common function for logging an error, given an ErrorRecord
function Log-ErrorRecord( [System.Management.Automation.ErrorRecord] $errRecord, [switch]$Stop )
{
	# Trim the message so it will not display the "ErrorActionPreference is set to Stop" message
	#
	$failedMessage = $errRecord.ToString()
	if ($failedMessage.IndexOf("ErrorActionPreference") -ne -1)
    {
    	$failedMessage = $failedMessage.Substring($failedMessage.IndexOf("set to Stop: ") + 13)
	}
	$failedMessage = $failedMessage -replace "`r"
	$failedMessage = $failedMessage -replace "`n"
	$failedCommand = $errRecord.InvocationInfo.MyCommand
	Log-Error ($ManageScheduledTask_LocalizedStrings.res_0002 -f $failedCommand,$failedMessage) -Stop:$Stop
}



###################################################################
###  Entry point for the script itself
###################################################################


# Command Write-ExchangeSetupLog is not available unless when run during setup.
# So we could log stuff into a separate file.
# Must set this flag before everything else
$script:writeToExchangeSetupLog = Is-WriteExchangeSetupLogPresent

if ($RunFromSetup)
{
	# If log file folder doesn't exist, create it
	if (!(Test-Path $script:exLogDir))
	{
		New-Item $script:exLogDir -type directory	
	}
}


$Command = $MyInvocation.MyCommand
Log-Verbose "Starting: $($Command.Path)"
# The command below is useful to see what parameters are defined in this script cmdlet.
# $Command | fl Path, CommandType, Parameters, ParameterSets


if ($PSCmdlet.ParameterSetName  -eq "Install" )
{
	# Get the UNC script path on the remote server
	$remotePath = "\\" + $ServerName + "\" + $PsScriptPath.Replace(":", "$")
	if (Test-Path $remotePath)
	{
		Log-Verbose "Found script: '$remotePath'"
	}
	else
	{
		Log-Error ($ManageScheduledTask_LocalizedStrings.res_0003 -f $remotePath)
		return
	}
	
	if ($DeleteExisting)
	{
		# First delete the existing task
		Delete-RegisteredTask $ServerName $TaskName
	}
		
	# Using '&amp;' instead of '&' because the $psCommand will be passed into an XML block
	$psCommand = "&amp; '" + $PsScriptPath + "' " + $PsScriptArgs
	[bool]$success = Create-TaskUsingCOM $ServerName $TaskName $psCommand
	if ($success)
	{
		Log-Verbose "Task created successfully."
	}
	else
	{
		Log-Error $ManageScheduledTask_LocalizedStrings.res_0004
	}
}
elseif ($PSCmdlet.ParameterSetName  -eq "Uninstall" )
{
	# Delete the existing task 
	Delete-RegisteredTask $ServerName $TaskName
}
elseif ($PSCmdlet.ParameterSetName  -eq "Disable" )
{
	# Disable the existing task but don't delete it
	StopAndDelete-RegisteredTaskInternal $ServerName $TaskName $false
}
elseif ($PSCmdlet.ParameterSetName  -eq "Enable" )
{
	# Enable the existing task
	Enable-TaskUsingCOM $ServerName $TaskName
}
elseif ($PSCmdlet.ParameterSetName  -eq "TestExistence" )
{
	# returns $true or $false
	Is-TaskExisting $ServerName $TaskName
}

# SIG # Begin signature block
# MIIdrgYJKoZIhvcNAQcCoIIdnzCCHZsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU6XPKcJnoaZjOGWTLcQM7YCDP
# K52gghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBLQwggSwAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCByDAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUipR5l0rjusT7zBdbWO3ytoKRjGUwaAYKKwYB
# BAGCNwIBDDFaMFigMIAuAE0AYQBuAGEAZwBlAFMAYwBoAGUAZAB1AGwAZQBkAFQA
# YQBzAGsALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFu
# Z2UgMA0GCSqGSIb3DQEBAQUABIIBAF5jS9A3q5jl8YIxw1QwzlCS3UjnikxcQ1xC
# HlpbxgUKzgEVdRirzNu9CkiFdXVtyJoPUgnSdzj9Dwc0SzRovpYRzwyEkX7F8AQb
# /OP3vYQBYKsWo6o7kFBwDTwqU6bJyuZjyrsr/ybnl3qyCkeNA07FI4eNoLWZSR8F
# QqLpf8sTA0XDM7cNOM3c2wmg0NSbZPfrln29V/NyeTThsFylHTJc6q4martUBX/8
# PoH6Gq0xiBYru7hCUwR10ORSGA3Rnk5epc7Qe08Wqe0fO6cwByLPILnZxkxsOMyZ
# S4pcgQtPcSsr3sIiq04iqt6HbaMPFa6KSieAC/Gq4++ZU9VsHUehggIoMIICJAYJ
# KoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
# AhMzAAAAmarFgZ+Mon2KAAAAAACZMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMx
# CwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQ1MDRaMCMGCSqG
# SIb3DQEJBDEWBBTsTZGpQSLEqd9JOzE+Z/mQCS9GBzANBgkqhkiG9w0BAQUFAASC
# AQBzRTkahhTIdxL5NOoujDlzxDHqGxQx3mYmUwcM6igot0RmOXco2GpyNGTt237U
# /EsQItjfkSQtDaas0MCKurIEoTiIJR+weR1SJMiI3JCmyfTajpe2P14+HV4qF4M6
# nbRGpuB2AnH+kIW2GKoNBlZLLpbAU3R1rN/iMYjTx/vpt22Le3OgRiGPUiYJhHAZ
# 6HUkGtFbS5RmT/stFe3s9235tdxsymaeMFITpoLVKJx3CyqzxcJKucN7EbjcqdHQ
# JoV5SUm7mauj+ToYW0sFMVXZDmkClwTdlRUEEr1BYNO4MzzGDJGIlQJqg2NiqBy/
# QZuHcOsLG0cFLFqOhtChhv3h
# SIG # End signature block
