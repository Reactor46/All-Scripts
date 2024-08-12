<# 

.SYNOPSIS
	Purpose of this script to help Exchange administrators with the process of resetting Exchange 2013 and 2016 Health Mailboxes.

	For more details, please see this post:
		https://blogs.technet.microsoft.com/rmilne/2016/06/16/script-to-automate-exchange-health-manager-service-for-monitoring-mailbox-reset/
	
	There is logic to loop through multiple Exchange servers. A filter will be used to exclude non Exchange 2013 servers in the default configuration.  
    
    	Filters can be adjusted to suit the particular task that is required.  There are filtering multiple examples here:
   	http://blogs.technet.com/b/rmilne/archive/2014/03/24/powershell-filtering-examples.aspx

	This is the line that you would need to change:
		$ExchangeServers = Get-ExchangeServer $ServerString | Where-Object {$_.AdminDisplayVersion -match "^Version 15"}

	Version 2 added an additional qualifier so that if you had a consistent naming strategy, you could select servers from a particular site using a wildcard.  
	Format to use would be:
	   ExchYYZ*  

	If you want to select all servers use *   

	Note that only the act of stopping the Health Manager service was automated.  This was for a couple of reasons:
		This is not considered normal maintenance and thus did not want to automated end to end, else someone may be tempted to schedule this as a task every Sunday morning
		One script does not have the knowledge to deal with every environment.  Thus you will still need to do the AD work etc, and confirm the results in your environment.

	Refer to the blog post for all the details.
	
	Script to be executed on an Exchange 2013 or 2016 server


.DESCRIPTION
	Script will use WMI object to work with remote windows services.  Get-Service, Set-Service, Stop-Service and Start-Service can be used for this, though WMI was used here.
	As an example, the below returns a reference to the service on a machine called DC1 

	$Service = Get-WmiObject -ComputerName DC1 -Class Win32_Service -Filter "Name='MSExchangeHM'"

	Script requires a command line parameter.  You must specify the start/stop action to take, though in current version of the script it will default to check.  
	Accepted actions are:
		Start
		Stop
		Check

	Stop  -- will stop the Microsoft Exchange Health Manager service (MSExchangeHM) on the servers returned in the collection
	Start -- will stop the Microsoft Exchange Health Manager service (MSExchangeHM) on the servers returned in the collection
	Check -- will display the status of the Microsoft Exchange Health Manager service (MSExchangeHM) on the servers returned in the collection


	CmdletBinding was added to check for the command line paramater being added, and to check the values entered.
	It also allows you to tab complete them, as an added bonus.  Sweet! 


.ASSUMPTIONS
	Script is being executed with sufficient permissions to access the server(s) targeted.
	Script is executed in the Exchange Management Shell
	Scritp is executed on Exchange 2013 or 2016 server
	Script is for Exchange 2013 and Exchange 2016
	Thus you have the correct version of PowerShell and WMF installed
	You can live with the Write-Host cmdlets :) 
	You can add your error handling if you need it.  

	

.VERSION
  
	1.0  12-6-2016 -- Initial script
	1.1  16-6-2016 -- Initial script released to the scripting gallery 
	1.2  4-10-2016 -- Updated with Feedback from Feras 

    



This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, 
provided that You agree: 
(i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; 
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and 
(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within the Premier Customer Services Description.
This posting is provided "AS IS" with no warranties, and confers no rights. 

Use of included script samples are subject to the terms specified at http://www.microsoft.com/info/cpyright.htm.

#>



[CmdletBinding()]
Param(
	[ValidateSet('Start','Stop','Check')]
	[Parameter(Mandatory=$false,Position=1)]
	[string]$Action = 'Check'	

)

[console]::ForegroundColor = "Yellow"
Write-Host 
[string]$ServerString = Read-Host ' Type the Server Name or Servers Names (wildcard accepted.  Use * for all servers)'
[console]::ResetColor()

# State what we are going to do on the screen.  <AC/DC>For those who are about to rock, we salute you!! </AC/DC> 
Write-Host 
Write-Host -ForegroundColor DarkYellow "Specified action to take was: $Action" 
Write-Host

# Return only Exchange 2013 servers.  The version statement can be changed as needed.  Sort the list alphabetically as I have obsessive compulsive issues.....
# Select only the Name attribute, as the Exchange Sever Name is all we need for later.  
$ExchangeServers = Get-ExchangeServer $ServerString | Where-Object {$_.AdminDisplayVersion -match "^Version 15"}  | Sort-Object  Name | Select Name

#######################################################  Section to Stop Services ######################################################

ForEach ($Server In $ExchangeServers)
    {
	Write-Host -Foregroundcolor Magenta "Processing Server: " -NoNewLine; Write-Host -Foregroundcolor White $Server.Name
	$Service = Get-WmiObject -ComputerName $Server.Name -Class Win32_Service -Filter "Name='MSExchangeHM'"
	Write-Host -ForegroundColor Cyan "`tCurrent Service State: " $Service.State
	
	switch ($Action)    {
	
	"STOP" {
			
			IF  ($Service.State -eq "Stopped")
			{
				# Nothing for us to do here.  The service was already stopped
				Write-Host -ForegroundColor Yellow "Service Already Stopped. No action taken."
			}
			ElseIf ($Service.State -eq "Running" -or "Paused" -or "Stopping")
			{
				Write-Host -ForegroundColor Green "HM Service is Running.  Stopping Service..."
				# Piping the response to Out-Null to avoid filling screen with muck.  Edit if you want to reverse that....
				$Service.StopService() | out-null
	
				# Check to see that the service stopped.  Sleeps and then re-checks the state 
				While ($Service.State -ne "Stopped")
				{
					Write-Host -ForegroundColor Yellow "`tWaiting for HM Service to stop..."
					$Service = Get-WmiObject -ComputerName $Server.name -Class Win32_Service -Filter "Name='MSExchangeHM'"
					Start-Sleep -Seconds 5
					# Did not add a counter to move on if the service is not stopping.  The admin will see a tonne of yellow text and be alerted that way.
					# After they correct the issue, the script can be re-run.  
				}				
			}
			Else
			{
				Write-Host -ForegroundColor Red "Unexpected service state detected on stop operation.  Please remediate before proceeding"
			}
		
			Write-Host
	}
	
	"START" {
			
			IF  ($Service.State -eq "Running")
			{
				# The service was already running.  
				# Why was it already running?  If running this script, the script should be responsible for stopping and starting services.
				# You may want to flag that as an error - uncomment line below if desired 
				#Write-Host -Foregroundcolor Red "Service Already Running"
				Write-Host -ForegroundColor Green "Service Already Running. No action taken."
			}
			# Service should be stopped when processing this section 
			ElseIf ($Service.State -eq "Stopped")
			{
				Write-Host -ForegroundColor Green "HM Service is stopped.  Starting Service..."
				# Piping the response to Out-Null to avoid filling screen with muck.  Edit if you want to reverse that....
				$Service.StartService() | out-null
	
				# Check to see that the service starts.  Sleeps and then re-checks the state after a brief hiatus to prevent spamming the display as that would be rude
				While ($Service.State -ne "Running")
				{
					Write-Host -ForegroundColor Yellow "`tWaiting for HM Service to start..."
					$Service = Get-WmiObject -ComputerName $Server.name -Class Win32_Service -Filter "Name='MSExchangeHM'"
					Start-Sleep -Seconds 5	
				}			
			}
			Else
			{
				Write-Host -ForegroundColor Red "Unexpected service state detected on start operation.  Please remediate before proceeding"
			}
		
			Write-Host
	}

	Default { 
			Write-Host 
	}
	
	}
}

