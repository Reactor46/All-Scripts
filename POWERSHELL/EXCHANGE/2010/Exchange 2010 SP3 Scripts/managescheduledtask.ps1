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
      <!-- <Arguments>-version 2.0 -NonInteractive -WindowStyle Hidden -command "D:\test.ps1 -MonitoringContext -ErrorAction:Continue"</Arguments> -->
	  <Arguments>-version 2.0 -NonInteractive -WindowStyle Hidden -command "{4}"</Arguments>
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
	
	Log-Verbose "Create-TaskUsingCOM: `$xmlString: `n$xmlString"
	
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
		finally
		{
			# Curious PS behavior: It appears that 'return' trumps 'throw', so don't return...
			if (!$throwOnError -or $success)
			{
				return $success
			}
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
# MIIadAYJKoZIhvcNAQcCoIIaZTCCGmECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU4NfHe2n2tOvN1k29+NPxUq2f
# XAWgghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
# 9w0BAQUFADB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMw
# IQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQTAeFw0xMjA5MDQyMTQy
# MDlaFw0xMzAzMDQyMTQyMDlaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC6pElsEPsi
# nGWiFpg7y2Fi+nQprY0GGdJxWBmKXlcNaWJuNqBO/SJ54B3HGmGO+vyjESUWyMBY
# LDGKiK4yHojbfz50V/eFpDZTykHvabhpnm1W627ksiZNc9FkcbQf1mGEiAAh72hY
# g1tJj7Tf0zXWy9kwn1P8emuahCu3IWd01PZ4tmGHmJR8Ks9n6Rm+2bpj7TxOPn0C
# 6/N/r88Pt4F+9Pvo95FIu489jMgHkxzzvXXk/GMgKZ8580FUOB5UZEC0hKo3rvMA
# jOIN+qGyDyK1p6mu1he5MPACIyAQ+mtZD+Ctn55ggZMDTA2bYhmzu5a8kVqmeIZ2
# m2zNTOwStThHAgMBAAGjggENMIIBCTATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQU3lHcG/IeSgU/EhzBvMOzZSyRBZgwHwYDVR0jBBgwFoAUyxHoytK0FlgB
# yTcuMxYWuUyaCh8wVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljQ29kU2lnUENBXzA4LTMxLTIwMTAu
# Y3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNDb2RTaWdQQ0FfMDgtMzEtMjAxMC5jcnQw
# DQYJKoZIhvcNAQEFBQADggEBACqk9+7AwyZ6g2IaeJxbxf3sFcSneBPRF1MoCwwA
# Qj84D4ncZBmENX9Iuc/reomhzU+p4LvtRxD+F9qHiRDRTBWg8BH/2pbPZM+B/TOn
# w3iT5HzVbYdx1hxh4sxOZLdzP/l7JzT2Uj9HQ8AOgXBTwZYBoku7vyoDd3tu+9BG
# ihcoMaUF4xaKuPFKaRVdM/nff5Q8R0UdrsqLx/eIHur+kQyfTwcJ7SaSbrOUGQH4
# X4HnrtqJj39aXoRftb58RuVHr/5YK5F/h9xGH1GVzMNiobXHX+vJaVxxkamNViAs
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggTDMIIDq6ADAgEC
# AhMzAAAAKzkySMGyyUjzAAAAAAArMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQTAeFw0xMjA5MDQyMTEyMzRaFw0xMzEyMDQyMTEyMzRaMIGz
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRN
# T1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QzBGNC0zMDg2LURFRjgxJTAj
# BgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCmtjAOA2WuUFqGa4WfSKEeycDuXkkHheBwlny+
# uV9iXwYm04s5uxgipS6SrdhLiDoar5uDrsheOYzCMnsWeO03ODrxYvtoggJo7Ou7
# QIqx/qEsNmJgcDlgYg77xhg4b7CS1kANgKYNeIs2a4aKJhcY/7DrTbq7KRPmXEiO
# cEY2Jv40Nas04ffa2FzqmX0xt00fV+t81pUNZgweDjIXPizVgKHO6/eYkQLcwV/9
# OID4OX9dZMo3XDtRW12FX84eHPs0vl/lKFVwVJy47HwAVUZbKJgoVkzh8boJGZaB
# SCowyPczIGznacOz1MNOzzAeN9SYUtSpI0WyrlxBSU+0YmiTAgMBAAGjggEJMIIB
# BTAdBgNVHQ4EFgQUpRgzUz+VYKFDFu+Oxq/SK7qeWNAwHwYDVR0jBBgwFoAUIzT4
# 2VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1w
# UENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# dDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAfsywe+Uv
# vudWtc9z26pS0RY5xrTN+tf+HmW150jzm0aIBWZqJoZe/odY3MZjjjiA9AhGfCtz
# sQ6/QarLx6qUpDfwZDnhxdX5zgfOq+Ql8Gmu1Ebi/mYyPNeXxTIh+u4aJaBeDEIs
# ETM6goP97R2zvs6RpJElcbmrcrCer+TPAGKJcKm4SlCM7i8iZKWo5k1rlSwceeyn
# ozHakGCQpG7+kwINPywkDcZqJoFRg0oQu3VjRKppCMYD6+LPC+1WOuzvcqcKDPQA
# 0yK4ryJys+fEnAsooIDK4+HXOWYw50YXGOf6gvpZC3q8qA3+HP8Di2OyTRICI08t
# s4WEO+KhR+jPFTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEF
# BQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jv
# c29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5MB4XDTEwMDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENv
# ZGUgU2lnbmluZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCy
# cllcGTBkvx2aYCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPV
# cgDbNVcKicquIEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlc
# RdyvrT3gKGiXGqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZ
# C/6SdCnidi9U3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgG
# hVxOVoIoKgUyt0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdc
# pReejcsRj1Y8wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBTLEejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYB
# BAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy8
# 2C0wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBW
# J5flJRP8KuEKU5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNy
# b3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3Js
# MFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcN
# AQEFBQADggIBAFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGj
# I8x8UJiAIV2sPS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbN
# LeNK0rxw56gNogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y
# 4k74jKHK6BOlkU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnp
# o1hW3ZsCRUQvX/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6
# H0q70eFW6NB4lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20O
# E049fClInHLR82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8
# Z4L5UrKNMxZlHg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9G
# uwdgR2VgQE6wQuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXri
# lUEnacOTj5XJjdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEv
# mtzjcT3XAH5iR9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCC
# A++gAwIBAgIKYRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPy
# LGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRN
# aWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1
# MzA5WhcNMjEwNDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZ
# USNQrc7dGE4kD+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cO
# BJjwicwfyzMkh53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn
# 1yjcRlOwhtDlKEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3
# U21StEWQn0gASkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG
# 7bfeI0a7xC1Un68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMB
# AAGjggGrMIIBpzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A
# +3b7syuwwzWzDzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1Ud
# IwSBkDCBjYAUDqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1p
# Y3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0
# BxMuZTBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUH
# AQEESDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# L2NlcnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcD
# CDANBgkqhkiG9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwT
# q86+e4+4LtQSooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+
# jwoFyI1I4vBTFd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwY
# Tp2OawpylbihOZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPf
# wgphjvDXuBfrTot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5
# ZlizLS/n+YWGzFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8cs
# u89Ds+X57H2146SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUw
# ZuhCEl4ayJ4iIdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHu
# diG/m4LBJ1S2sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9
# La9Zj7jkIeW1sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g
# 74TKIdbrHk/Jmu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUW
# a8kTo/0xggSvMIIEqwIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIHIMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBRMW/6DkLVA/v0ALLFv+rSrS37gIjBoBgorBgEEAYI3AgEMMVow
# WKAwgC4ATQBhAG4AYQBnAGUAUwBjAGgAZQBkAHUAbABlAGQAVABhAHMAawAuAHAA
# cwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZI
# hvcNAQEBBQAEggEAKfSodGicsxaTfVxVxi2uvL4KPfwbRTPL9RfiiQ+eLjJcxyTO
# MkBjx5IgqwZKYF5F2qTxBeO4q1z5iFVOZX0mi4wi4P6f7JWtSoq+okkBPabhvjkN
# hcXz6PjHVdW8YG4r+uPBONNbGaZmU3TI68xSIi8maXS+amkfeg1I9EGlhRC9uXwC
# pTVcSnu1CmizE0d5Kye1ibDyfNA3DYUDX8bf1IBL1lhJxx1k1AXQFP9QSct3sysV
# 1AfwdIlL8dh5aXRH3w26V0/qRzZmaPpm2N3+jhXsC5N07vBLzP/k05ICgcq6NFWl
# D8CbmOHb8CXQ9PyZgQm8v7jtO3xelo4eGkQVi6GCAigwggIkBgkqhkiG9w0BCQYx
# ggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAAArOTJI
# wbLJSPMAAAAAACswCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0B
# BwEwHAYJKoZIhvcNAQkFMQ8XDTEzMDEwODA4NDY1NlowIwYJKoZIhvcNAQkEMRYE
# FLHzqbI4PJ4lOj5jIGihFtrPKbNKMA0GCSqGSIb3DQEBBQUABIIBAJDPnBG1pMUy
# gdD1wszSwhZ4TyotCR9OAn+eSUt+0OQ2LQtwC9uyhvBoZSqYGf05e8xLa5xSoyBS
# w1GfBgO9chvGIqxsWGmXcjcm+rIYu08trMVsc6pzf1vRD84B7jl7m1NKxG/Us6KO
# qyi9BM9fy9+2qJXmUJBbFu3wdJPfR4w/PC2uvLB3d8BF9aBWv5LBIO2icE7sU4TB
# hp43CbxV9qRaIniVqy+lCzM6CKkil+MGwY5XDqzfeiCJdHPhmfwmnCwK8NxWuX3e
# mxb4gK94wQdIRvTCaFCueJcg+8IodmNf0rN+vRq9wWr3pjmB41DyWwIqFL5MTQNj
# CC5+5Z3rKRY=
# SIG # End signature block
