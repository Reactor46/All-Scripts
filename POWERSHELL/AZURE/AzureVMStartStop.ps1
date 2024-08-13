<#
.SYNOPSIS
	Starts or stops virtual machines based on the information found in csv configuration files.
	
.DESCRIPTION
	Starts or stops virtual machines based on the information found in csv configuration files.  The configuration files are located in the script directory and include 
	1) Subscriptions.csv to put user friendly names to the subscription id and management certificate thumbprint pair.  
	2) Jobs.csv to define the subscriptions that contain virtual machines to be started and stopped at various points in time.
	3) Exemptions.csv to list cloud services that are exempt from the start and stop script.

.NOTES
	Author: Chris Clayton
	Date: 2013/08/30
	Revision: 1.1

.EXAMPLE
	./AzureVMStartStop.ps1
#>


Import-Module "C:\Program Files (x86)\Microsoft SDKs\Windows Azure\PowerShell\Azure\Azure.psd1"

<#
================================================================================================================================================================
														Type Creation
================================================================================================================================================================
#>

<#
.SYNOPSIS
	Creates a .NET enumeration and registers it with PowerShell.
	
.DESCRIPTION
	Creates a .NET enumeration and registers it with PowerShell.  This will fail silently if the enumeration already exists.  This should be called before 
	any of these script methods are leveraged.

.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	Initialize-StartStopTypes
#>
function Initialize-StartStopTypes
{	
	[string]$azureRoleSizes = 'public enum AzureRoleSize { ExtraSmall = 0, Small = 1, Medium = 2, Large = 3, ExtraLarge = 4, A6 = 5, A7 = 6 }'

	# This was added to silently continue.  In the event that the type has already been created closing and reopening powershell is the only way to release it
	try { Add-Type -TypeDefinition $azureRoleSizes -ErrorAction SilentlyContinue } catch { }
}

<#
================================================================================================================================================================
														Configure the subscriptions to monitor
================================================================================================================================================================
#>

<#
.SYNOPSIS
	Deletes the Windows Azure subscriptions file used by PowerShell in the context of this script. 
	
.DESCRIPTION
	Deletes the Windows Azure subscriptions file used by PowerShell in the context of this script. It is recommended that this is called at the beginning and 
	end of the script.  This file is to be encrypted based on the Windows Azure PowerShell documentation, but based on actual use it is plain text XML.

.PARAMETER subscriptionDataFile
	The name of the data file that stores the subscription information.

.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	Clear-SubscriptionData 'c:\subscriptiondata.xml'
#>
function Clear-SubscriptionData
{	
	param
	(
		[string]$subscriptionDataFile
	)

	if(Test-Path -PathType Leaf -Path $subscriptionDataFile)
	{
		Remove-Item $subscriptionDataFile
	}
}

<#
.SYNOPSIS
	Adds a new subscription to the Windows Azure subscription file that is used in this scripts context.
	
.DESCRIPTION
	Adds a new subscription to the Windows Azure subscription file that is used in this scripts context.  Calling this function
	adds entries that can later be used with a call to Select-AzureSubscription.  Specify the name as the SubscriptionName 
	parameter and the file using the SubscriptionDataFile.	

.PARAMETER subscriptionName
	The name of the subscription entry that will be used later to reference the subscription id and X509 certificate pair.
	
.PARAMETER subscriptionId
	The subscription id as seen in the Windows Azure Developer portal.

.PARAMETER thumbprint
	The thumbprint of the X509 certificate to be used for Windows Azure management functions.  This must already be uploaded to the
	portal.
	
.PARAMETER storageAccount
	If provided this is the name of the storage account within the specified subscription that is used as the default.  Some
	PowerShell cmdlets require this to be set (ex. creating virtual machines).  If none is provided no current storage account
	setting will be assigned for this subscription.
	
.PARAMETER isDefault
	If $true this subscription is the default to be used when no other Windows Azure subscription has been defined.  If this 
	parameter is not specified it will default to not being the default.
	
.PARAMETER subscriptionDataFile
	The name of the data file that stores the subscription information.
	
.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	New-SubscriptionByCertificate 'MySubscription' 'XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX' 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' 'MyStorageAccount' $false 'c:\subscriptiondata.xml'

.LINK
	http://manage.windowsazure.com
#>
function New-SubscriptionByCertificate
{
	param
	(	
		[string]$subscriptionName,		
		[string]$subscriptionId,
		[string]$thumbprint,
		[string]$storageAccount = $null,
		[bool]$isDefault = $false,
		[string]$subscriptionDataFile
	)
		
	$managementCertificate = Get-Item "cert:\CurrentUser\MY\$($thumbprint)"	
	
	if($managementCertificate -eq $null)
	{
		$managementCertificate = Get-Item "cert:\LocalMachine\MY\$($thumbprint)"	
	}
	
	if(($storageAccount -ne $null) -and (![string]::IsNullOrWhiteSpace($storageAccount)) -and ($managementCertificate -ne $null))
	{
		Set-AzureSubscription -Certificate $managementCertificate -SubscriptionId $subscriptionId -CurrentStorageAccount $storageAccount -SubscriptionName $subscriptionName -SubscriptionDataFile $subscriptionDataFile
	}
	else
	{
		Set-AzureSubscription -Certificate $managementCertificate -SubscriptionId $subscriptionId -SubscriptionName $subscriptionName -SubscriptionDataFile $subscriptionDataFile
	}
}

<#
================================================================================================================================================================
														Job Control Files
================================================================================================================================================================
#>

<#
.SYNOPSIS
	Retrieves a list of subscriptions being monitored.
	
.DESCRIPTION
	Retrieves a list of subscriptions being monitored.  This information is taken from the csv file.

.PARAMETER subscriptionsFile
	A CSV file containing the subscriptions being monitored.  Each line of the file represents a seperate subscription.

.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	Get-MonitoredSubscriptions 'c:\subscriptions.csv'
#>
function Get-MonitoredSubscriptions
{
	param
	(
		[string]$subscriptionsFile
	)
	
	[System.Array]$subscriptionList = $null
	
	if(Test-Path -Path $subscriptionsFile -PathType Leaf)
	{
		$subscriptionList = Import-Csv $subscriptionsFile
	}
	
	return $subscriptionList
}

<#
.SYNOPSIS
	Retrieves a list of monitoring jobs.
	
.DESCRIPTION
	Retrieves a list of monitoring jobs.  This information is taken from the csv file.

.PARAMETER subscriptionsFile
	A CSV file containing the monitoring jobs.  Each line of the file represents a seperate job.

.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	Get-MonitoringJobFile 'c:\jobs.csv'
#>
function Get-MonitoringJobFile
{
	param
	(
		[string]$jobFile
	)
	
	[System.Array]$monitoringJobList = $null
	
	if(Test-Path -Path $jobFile -PathType Leaf)
	{
		$monitoringJobList = Import-Csv $jobFile
	}
	
	return $monitoringJobList
} 

<#
.SYNOPSIS
	Retrieves a list of cloud services that are exempt from monitoring.
	
.DESCRIPTION
	Retrieves a list of cloud services that are exempt from monitoring.  This information is taken from the csv file.

.PARAMETER subscriptionsFile
	A CSV file containing the exempt services.  Each line of the file represents a seperate cloud service.

.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	Get-MonitoringJobFile 'c:\exemptions.csv'
#>
function Get-ExemptionFile
{
	param
	(
		[string]$exemptionFile
	)
	
	[System.Array]$exemptionList = $null
	
	if(Test-Path -Path $exemptionFile -PathType Leaf)
	{
		$exemptionList = Import-Csv $exemptionFile
	}
	
	return $exemptionList
} 

<#
================================================================================================================================================================
														Virtual Machine List Checks
================================================================================================================================================================
#>

<#
.SYNOPSIS
	Retrieves a virtual machines within a cloud service that should be shutdown and deallocated.
	
.DESCRIPTION
	Retrieves a virtual machines within a cloud service that should be shutdown and deallocated.  If leave one is set a single virtual machine
	with the smallest instance size will not be deallocated to reduce cost and maintain the IP address of the cloud service.

.PARAMETER serviceName
	The name of the Cloud Service containing the virtual machines being queried.

.PARAMETER leaveOne
	If $true a single virtual machine will be selected to leave running.  This will stop the cloud service IP address from being lost.  The
	virtual machine selected will be the smallest size.

.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	Get-ShutdownList 'cloudservice1' $true
#>
function Get-ShutdownList
{
	param
	(
		[string]$serviceName,
		[bool]$leaveOne = $true
	)
	$results = @()
	[string]$smallestRunningVM = $null
	[string]$lastSize = $null
		
	$virtualMachines = Get-AzureVM -ServiceName $service.ServiceName
	
	# If one is to be left find the smallest virtual machine and leave it to reduce cost.
	if($leaveOne)
	{
		foreach($virtualMachine in $virtualMachines)
		{
			if($virtualMachine.InstanceStatus -ne 'StoppedDeallocated')
			{
				if(($lastSize -eq $null -or [string]::IsNullOrWhiteSpace($lastSize)) -or (([AzureRoleSize]$lastSize) -gt ([AzureRoleSize]$virtualMachine.InstanceSize)))
				{
					$lastSize = $virtualMachine.InstanceSize
					$smallestRunningVM = $virtualMachine.Name
				}
			}
		}
	}
	
	# Iterate through the virtual machines and create a list of ones to be shutdown
	foreach($virtualMachine in $virtualMachines)
	{
		if($virtualMachine.InstanceStatus -ne 'StoppedDeallocated')
		{
			if((-not $leaveOne) -or ($smallestRunningVM -ne $virtualMachine.Name))
			{
				$results += $virtualMachine
			}
		}	
	}

	
	return $results	
}

<#
.SYNOPSIS
	Retrieves a virtual machines within a cloud service that should be started.
	
.DESCRIPTION
	Retrieves a virtual machines within a cloud service that should be started.  If a virtual machine is not deallocated it will not be
	started.  This allows for users stopping their VMs without unexpected starts.

.PARAMETER serviceName
	The name of the Cloud Service containing the virtual machines being queried.

.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	Get-StartupList 'cloudservice1' 
#>
function Get-StartupList
{
	param
	(
		$serviceName
	)
	$results = @()	
	$virtualMachines = Get-AzureVM -ServiceName $service.ServiceName
	
	foreach($virtualMachine in $virtualMachines)
	{
		if($virtualMachine.InstanceStatus -eq 'StoppedDeallocated')
		{
			$results += $virtualMachine
		}	
	}

	
	return $results	
}

<#
================================================================================================================================================================
														Process Monitoring Job
================================================================================================================================================================
#>


<#
.SYNOPSIS
	Executes a single start / stop job as defined in the monitoring job object.
	
.DESCRIPTION
	Executes a single start / stop job as defined in the monitoring job object.  This value must not be null.

.PARAMETER monitoringJob
	The job object to be executed.  If the job is outside of the defined hours it will shutdown, otherwise it will start.

.PARAMETER subscriptionDataFile
	The windows azure subscription data file used for calls to select service.

.PARAMETER exemptServices
	An array containing the list of exempt services.  This is typically derived from a call to Get-ExemptionFile.

.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	Process-MonitoringJob $job 'c:\subscriptiondata.xml' $exemptServices
#>
function Process-MonitoringJob
{
	param
	(
		$monitoringJob,
		[string]$subscriptionDataFile,
		[System.Array]$exemptServices = $null
	)
	[System.DateTime]$processingTime = [System.DateTime]::Now
	
	[System.TimeSpan]$hourOffset = [System.TimeSpan]::Parse($monitoringJob.StartTime)	
	[System.DateTime]$startTime = $processingTime.Date.Add($hourOffset)
	[System.DateTime]$endTime = $startTime.AddHours($monitoringJob.RunHours)
	
	Select-AzureSubscription -SubscriptionName $monitoringJob.SubscriptionName -SubscriptionDataFile $subscriptionDataFile	
	$services = Get-AzureService 
	
	foreach($service in $services)
	{
		# Check to see if the cloud service is on the exemption list
		if(-not (Test-ServiceExempt $monitoringJob.SubscriptionName $service.ServiceName $exemptServices))
		{
			if(($processingTime -lt $startTime) -or ($processingTime -ge $endTime))
			{	
				if($monitoringJob.PerformStop -eq 'TRUE')
				{
					[bool]$leaveOneInService = $monitoringJob.LeaveOne -eq 'True'		
					$machinesToShutdown = Get-ShutdownList $service.ServiceName $leaveOneInService
					
					foreach($machine in $machinesToShutdown)
					{
						Stop-AzureVM -Name $machine.Name -ServiceName $service.ServiceName -Force
					}
				}
			}
			else
			{
				if($monitoringJob.PerformStart -eq 'TRUE')
				{
					foreach($service in $services)
					{
						$machinesToStart = Get-StartupList $service.ServiceName 
						
						foreach($machine in $machinesToStart)
						{
							Start-AzureVM -Name $machine.Name -ServiceName $service.ServiceName
						}
					}
				}
			}
		}
	}	
}

<#
.SYNOPSIS
	Checks for the value in the exemption list that matches the subscription id and service name.
	
.DESCRIPTION
	Checks for the value in the exemption list that matches the subscription id and service name.  If the list is null or empty it will always return false.

.PARAMETER subscriptionName
	The name of the subscription that the cloud service is contained in.

.PARAMETER serviceName
	The cloud service name that is being tested for the existence of.

.PARAMETER exemptServices
	An array containing the list of exempt services.  This is typically derived from a call to Get-ExemptionFile.

.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	Test-ServiceExempt 'MySubscription' 'cloudservice1' $exemptionList
#>
function Test-ServiceExempt
{
	param
	(
		[string]$subscriptionName,
		[string]$serviceName,
		[System.Array]$exemptServices = $null
	)
	[bool]$isExempt = $false

	if(($exemptServices -ne $null) -and ($exemptServices.Count -gt 0))
	{
		foreach($exemption in $exemptServices)
		{
			if(($subscriptionName -eq $exemption.SubscriptionName) -and ($serviceName -eq $exemption.CloudServiceName))
			{
				$isExempt = $true
			}
		}
	}
	
	return $isExempt
}

<#
================================================================================================================================================================
														Core Startup / Shutdown Entry Point
================================================================================================================================================================
#>

<#
.SYNOPSIS
	This is the entry point to the main script logic.  
	
.DESCRIPTION
	This is the entry point to the main script logic.  Encapsulation of this into a single method allows for multiple jobs to run in the same scheduled task.

.PARAMETER azureSubscriptionFileName
	The XML file that contains the subscription file that is maintained by the Windows Azure PowerShell cmdlets.

.PARAMETER subscriptionFileName
	The fully qualified name of the file containing the subscription definitions.

.PARAMETER jobFileName
	The fully qualified name of the file containing the monitoring job definitions.

.PARAMETER exemptionFileName
	The fully qualified name of the file containing the exemption list.

.NOTES
	Author: Chris Clayton
	Date: 2013/07/11
	Revision: 1.0

.EXAMPLE
	Execute-JobMonitoring 'c:\StartStop\SubscriptionData.xml' 'c:\StartStop\JobSubscriptions.csv' 'c:\StartStop\jobs.csv' 'c:\StartStop\exemptions.csv'
#>
function Execute-JobMonitoring
{
	param
	(
		[string]$azureSubscriptionFileName,
		[string]$subscriptionFileName,
		[string]$jobFileName,
		[string]$exemptionFileName
	)
	
	# Prepare .NET types 
	Initialize-StartStopTypes

	# If previous runs did not finish clear them out
	Clear-SubscriptionData $azureSubscriptionFileName

	# Read in the jobs to be monitored and the subscription details
	[System.Array]$monitoredSubscriptions = Get-MonitoredSubscriptions $subscriptionFileName
	[System.Array]$monitoringJobs = Get-MonitoringJobFile $jobFileName
	[System.Array]$exemptServices = Get-ExemptionFile $exemptionFileName
			

	# Setup the subscriptions to be used
	foreach($subscription in $monitoredSubscriptions)
	{
		New-SubscriptionByCertificate $subscription.SubscriptionName $subscription.SubscriptionId $subscription.CertificateThumbprint $null $false	$azureSubscriptionFileName
	}

	# Execute each of the jobs
	foreach($monitoringJob in $monitoringJobs)
	{
		Process-MonitoringJob $monitoringJob $azureSubscriptionFileName $exemptServices
	}

	# Cleanup the subscription information
	Clear-SubscriptionData $azureSubscriptionFileName
}

<#
================================================================================================================================================================
														Script Body
================================================================================================================================================================
#>

# Must remain in the body of the script to execute correctly.  This determines the directory thescript was executed from
[string]$startStopScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# Setup file paths
[string]$startStopSubscriptionDataFile = [System.IO.Path]::Combine($startStopScriptDirectory, 'SubscriptionData.xml')
[string]$monitoredSubscriptionFile = [System.IO.Path]::Combine($startStopScriptDirectory, 'Subscriptions.csv')
[string]$monitoringJobsFile = [System.IO.Path]::Combine($startStopScriptDirectory, 'Jobs.csv')
[string]$exemptionsFile = [System.IO.Path]::Combine($startStopScriptDirectory, 'Exemptions.csv')


Execute-JobMonitoring $startStopSubscriptionDataFile $monitoredSubscriptionFile $monitoringJobsFile $exemptionsFile

