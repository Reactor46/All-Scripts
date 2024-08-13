#requires -version 2
<#
.SYNOPSIS
Test-ExchangeServerHealth.ps1 - Exchange Server 2010 Health Check Script.

.DESCRIPTION 
Performs a series of health checks on the specified Exchange servers
and outputs the results to screen, and optionally to log file, HTML report,
and HTML email.

Use the ignorelist.txt file to specify any servers you want the script to
ignore (eg test/dev servers).

.OUTPUTS
Results are output to screen, as well as optional log file, HTML report, and HTML email

.PARAMETER Server
Perform a health check of a single server

.PARAMETER ReportMode
Set to $true to generate a HTML report. A default file name is used if none is specified.

.PARAMETER ReportFile
Allows you to specify a different HTML report file name than the default. Implies -reportmode:$true

.PARAMETER SendEmail
Sends the HTML report via email using the SMTP configuration within the script. Implies -reportmode:$true

.EXAMPLE
.\Test-ExchangeServerHealth.ps1
Checks all servers in the organization and outputs the results to the shell window.

.EXAMPLE
.\Test-ExchangeServerHealth.ps1 -Server HO-EX2010-MB1
Checks the server HO-EX2010-MB1 and outputs the results to the shell window.

.EXAMPLE
.\Test-ExchangeServerHealth.ps1 -ReportMode -SendEmail
Checks all servers in the organization, outputs the results to the shell window, a HTML report, and
emails the HTML report to the address configured in the script.

.LINK
http://exchangeserverpro.com/powershell-script-health-check-report-exchange-2010

.NOTES
Written By: Paul Cunningham
Website:	http://exchangeserverpro.com
Twitter:	http://twitter.com/exchservpro

Change Log
V1.0, 5/07/2012 - Initial version
V1.1, 5/08/2012 - Minor bug fixes and removed Edge Tranport checks
V1.2, 5/5/2013 - A lot of bug fixes, updated SMTP to use Send-MailMessage, added DAG health check. 
#>

[CmdletBinding()]
param (
	[Parameter( Mandatory=$false)]
	[string]$Server,

	[Parameter( Mandatory=$false)]
	[string]$ServerList,	
	
	[Parameter( Mandatory=$false)]
	[string]$ReportFile="exchangeserverhealth.html",

	[Parameter( Mandatory=$false)]
	[switch]$ReportMode,
	
	[Parameter( Mandatory=$false)]
	[switch]$SendEmail

	)


#...................................
# Variables
#...................................

$now = Get-Date											#Used for timestamps
$date = $now.ToShortDateString()						#Short date format for email message subject
[array]$exchangeservers = @()							#Array for the Exchange server or servers to check
[int]$transportqueuehigh = 100							#Change this to set transport queue high threshold
[int]$transportqueuewarn = 80							#Change this to set transport queue warning threshold
$mapitimeout = 10										#Timeout for each MAPI connectivity test, in seconds
$pass = "Green"
$warn = "Yellow"
$fail = "Red"
$ip = $null
[array]$summaryreport = @()
[array]$report = @()
[bool]$alerts = $false
[array]$dags = @()										#Array for DAG health check
[int]$replqueuewarning = 8								#Threshold to consider a replication queue unhealthy
$dagreportbody = $null


#...................................
# Modify these Variables
#...................................

#$ignorelistfile = "C:\Scripts\ExchangeServerHealth\ignorelist.txt"	#Path to the txt file containing list of server names to ignore


#...................................
# Modify these Email Settings
#...................................

$smtpsettings = @{
	To =  "administrator@exchangeserverpro.net"
	From = "exchangeserver@exchangeserverpro.net"
	Subject = "Exchange Server Health Report - $now"
	SmtpServer = "smtp.exchangeserverpro.net"
	}


#...................................
# Error/Warning Strings
#...................................

$string0 = "Server is not an Exchange server. "
$string1 = "Server is not reachable. "
$string3 = "------ Checking"
$string4 = "Could not test service health. "
$string5 = "required services not running. "
$string6 = "Could not check queue. "
$string7 = "Public Folder database not mounted. "
$string8 = "Skipping Edge Transport server. "
$string9 = "Mailbox databases not mounted. "
$string10 = "MAPI tests failed. "
$string11 = "Mail flow test failed. "
$string12 = "No Exchange Server 2003 checks performed. "
$string13 = "Server not found in DNS. "
$string14 = "Sending email. "
$string15 = "Done."
$string16 = "------ Finishing"
$string17 = "Unable to retrieve uptime. "
$string18 = "Ping failed. "
$string19 = "No alerts found, and AlertsOnly switch was used. No email sent. "


#...................................
# Initialize
#...................................

Write-Host "Initializing..."

#Add Exchange 2010 snapin if not already loaded
#if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
#{
#	Write-Verbose "Loading the Exchange 2010 snapin"
#	try
#	{
#		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
#	}
#	catch
#	{
#		#Snapin not loaded
#		Write-Warning $_.Exception.Message
#		EXIT
#	}
#	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
#	Connect-ExchangeServer -auto -AllowClobber
#}


#Set recipient scope
if (!(Get-ADServerSettings).ViewEntireForest)
{
	Set-ADServerSettings -ViewEntireForest $true -WarningAction SilentlyContinue
}


#...................................
# Script
#...................................

#Check if a single server was specified
if ($server)
{
	#Run for single specified server
	try
	{
		$exchangeservers = Get-ExchangeServer $server -ErrorAction STOP
	}
	catch
	{
		#Couldn't find Exchange server of that name
		Write-Warning $_.Exception.Message
		
		#Exit because single server was specified and couldn't be found in the organization
		EXIT
	}
}
elseif ($serverlist)
{
	#Run for a list of servers in a text file
	try
	{
        $tmpservers = @(Get-Content $serverlist -ErrorAction STOP)
		$exchangeservers = @($tmpservers | Get-ExchangeServer)
    }
    catch
	{
        #Write-Host -ForegroundColor $warn "The file $serverlist could not be found."
		#EXIT
    }
}
else
{
	#This is the list of server names to never alert for
	try
	{
    #    $ignorelist = @(Get-Content $ignorelistfile -ErrorAction STOP)
    }
    catch
	{
        Write-Host -ForegroundColor $warn "The file $ignorelistfile could not be found. No servers will be ignored."
    }
    
	$tmpservers = @(Get-ExchangeServer | sort site,name)
	
	#Remove the servers that are ignored from the list of servers to check
	foreach ($tmpserver in $tmpservers)
	{
		if (!($ignorelist -icontains $tmpserver.name))
		{
			$exchangeservers = $exchangeservers += $tmpserver.identity
		}
	}
}

#Begin the health checks
foreach ($server in $exchangeservers)
{

	Write-Host -ForegroundColor White "$string3 $server"

	#Find out some details about the server
	try
	{
		$serverinfo = Get-ExchangeServer $server -ErrorAction Stop
	}
	catch
	{
		Write-Warning $_.Exception.Message
		$serverinfo = $null
	}

	if ($serverinfo -eq $null )
	{
		#Server is not an Exchange server
		Write-Host -ForegroundColor $warn $string0
	}
	elseif ( $serverinfo.IsEdgeServer )
	{
		Write-Host -ForegroundColor White $string8
	}
	else
	{
		#Server is an Exchange server, continue the health check

		#Custom object properties
		$serverObj = New-Object PSObject
		$serverObj | Add-Member NoteProperty -Name "Server" -Value $server
		
		$site = ($serverinfo.site.ToString()).Split("/")
		$serverObj | Add-Member NoteProperty -Name "Site" -Value $site[-1]
		
		$serverObj | Add-Member NoteProperty -Name "DNS" -Value $null
		$serverObj | Add-Member NoteProperty -Name "Ping" -Value $null
		$serverObj | Add-Member NoteProperty -Name "Uptime (hrs)" -Value $null
		$serverObj | Add-Member NoteProperty -Name "Version" -Value $null
		$serverObj | Add-Member NoteProperty -Name "Roles" -Value $null
		$serverObj | Add-Member NoteProperty -Name "Client Access Server Role Services" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Hub Transport Server Role Services" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Mailbox Server Role Services" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Unified Messaging Server Role Services" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Queue Length" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "PF DBs Mounted" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "MB DBs Mounted" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "MAPI Test" -Value "n/a"
		$serverObj | Add-Member NoteProperty -Name "Error Details" -Value ""
		$serverObj | Add-Member NoteProperty -Name "Warning Details" -Value ""


		#Check server name resolves in DNS
		Write-Host "DNS Check: " -NoNewline;
		try 
		{
			$ip = @([System.Net.Dns]::GetHostByName($server).AddressList | Select-Object IPAddressToString -ExpandProperty IPAddressToString)
		}
		catch
		{
			Write-Host -ForegroundColor $warn $_.Exception.Message
			$ip = $null
		}
		finally	{}

		if ( $ip -ne $null )
		{

			Write-Host -ForegroundColor $pass "Pass"
			$serverObj | Add-Member NoteProperty -Name "DNS" -Value "Pass" -Force

			#Is server online
			Write-Host "Ping Check: " -NoNewline; 
			
			$ping = $null
			try
			{
				$ping = Test-Connection $server -Quiet -ErrorAction Stop
			}
			catch
			{
				Write-Host -ForegroundColor $warn $_.Exception.Message
			}

			switch ($ping)
			{
				$true {	Write-Host -ForegroundColor $pass "Pass"; $serverObj | Add-Member NoteProperty -Name "Ping" -Value "Pass" -Force }
				default { Write-Host -ForegroundColor $fail "Fail"; $serverObj | Add-Member NoteProperty -Name "Ping" -Value "Fail" -Force; $serverObj | Add-Member NoteProperty -Name "Error Details" -Value ($($serverObj."Error Details") + $string18) -Force }
			}
			
			#Uptime check, even if ping fails
            $uptime = $null
			$laststart = $null
			
			try 
			{
				$laststart = [System.Management.ManagementDateTimeconverter]::ToDateTime((Get-WmiObject -Class Win32_OperatingSystem -computername $server -ErrorAction Stop).LastBootUpTime)
			}
			catch
			{
				Write-Host -ForegroundColor $warn $_.Exception.Message
			}
			finally	{}
			
            Write-Host "Uptime (hrs): " -NoNewline

			if ($laststart -eq $null)
			{
				[string]$uptime = $string17
				switch ($ping)
				{
                	$true { $serverObj | Add-Member NoteProperty -Name "Warning Details" -Value ($($serverObj."Warning Details") + $string17) -Force }
					default { $serverObj | Add-Member NoteProperty -Name "Error Details" -Value ($($serverObj."Warning Details") + $string17) -Force }
				}
			}
			else
			{
				[int]$uptime = (New-TimeSpan $laststart $now).TotalHours
				[int]$uptime = "{0:N0}" -f $uptime
			    Switch ($uptime -gt 23) {
				    $true { Write-Host -ForegroundColor $pass $uptime }
				    $false { Write-Host -ForegroundColor $warn $uptime }
				    default { Write-Host -ForegroundColor $warn $uptime }
			    }
			}

			$serverObj | Add-Member NoteProperty -Name "Uptime (hrs)" -Value $uptime -Force	
			
			if ($ping -or ($uptime -ne $string17))
			{
				#Determine the friendly version number
				$ExVer = $serverinfo.AdminDisplayVersion
				Write-Host "Server version: " -NoNewline;
				
				if ($ExVer -like "Version 6.*")
				{
					$version = "Exchange 2003"
				}
				
				if ($ExVer -like "Version 8.*")
				{
					$version = "Exchange 2007"
				}
				
				if ($ExVer -like "Version 14.*")
				{
					$version = "Exchange 2010"
				}
				
				Write-Host $version				
				$serverObj | Add-Member NoteProperty -Name "Version" -Value $version -Force
			
				if ($version -eq "Exchange 2003")
				{
					Write-Host $string12
					$report = $report + $serverObj
				}

				#START - Exchange 2007/2010 Health Checks
				if ($version -eq "Exchange 2007" -or $version -eq "Exchange 2010")
				{
					Write-Host "Roles:" $serverinfo.ServerRole
					$serverObj | Add-Member NoteProperty -Name "Roles" -Value $serverinfo.ServerRole -Force
					
					$IsEdge = $serverinfo.IsEdgeServer		
					$IsHub = $serverinfo.IsHubTransportServer
					$IsCAS = $serverinfo.IsClientAccessServer
					$IsMB = $serverinfo.IsMailboxServer

					#START - General Server Health Check
					#Skipping Edge Transports for the general health check, as firewalls usually get
					#in the way. If you want to include then, remove this If.
					if ($IsEdge -ne $true)
					{
							#Service health is an array due to how multi-role servers return Test-ServiceHealth status
                            $servicehealth = @()
							try {
								$servicehealth = @(Test-ServiceHealth $server -ErrorAction Stop)
							}
							catch {
								$serverObj | Add-Member NoteProperty -Name "Warning Details" -Value ($($serverObj."Warning Details") + $string4) -Force
								Write-Host -ForegroundColor $warn $string4 ":" $_.Exception
                                $serverObj | Add-Member NoteProperty -Name "Client Access Server Role Services" -Value $string4 -Force
		                        $serverObj | Add-Member NoteProperty -Name "Hub Transport Server Role Services" -Value $string4 -Force
		                        $serverObj | Add-Member NoteProperty -Name "Mailbox Server Role Services" -Value $string4 -Force
		                        $serverObj | Add-Member NoteProperty -Name "Unified Messaging Server Role Services" -Value $string4 -Force
							}
							
							if ($servicehealth)
							{
								foreach($s in $servicehealth)
								{
									$roleName = $s.Role
									Write-Host $roleName "Services: " -NoNewline;
									
									
									switch ($s.RequiredServicesRunning)
									{
										$true { $svchealth = "Pass"; Write-Host -ForegroundColor $pass "Pass" }
										$false {$svchealth = "Fail"; Write-Host -ForegroundColor $fail "Fail"; $serverObj | Add-Member NoteProperty -Name "Error Details" -Value ($($serverObj."Error Details") + ("$roleName $string5")) -Force }
                                        default {$svchealth = "Warn"; Write-Host -ForegroundColor $warn "Warning"; $serverObj | Add-Member NoteProperty -Name "Error Details" -Value ($($serverObj."Error Details") + ("$roleName $string5")) -Force }
									}

									switch ($s.Role)
									{
										"Client Access Server Role" { $serverinfoservices = "Client Access Server Role Services" }
										"Hub Transport Server Role" { $serverinfoservices = "Hub Transport Server Role Services" }
										"Mailbox Server Role" { $serverinfoservices = "Mailbox Server Role Services" }
										"Unified Messaging Server Role" { $serverinfoservices = "Unified Messaging Server Role Services" }
									}
									
									$serverObj | Add-Member NoteProperty -Name $serverinfoservices -Value $svchealth -Force
								}
							}
					}
					#END - General Server Health Check

					#START - Hub Transport Server Check
					if ($IsHub)
					{
						$q = $null
						Write-Host "Total Queue: " -NoNewline; 
						try {
							$q = Get-Queue -server $server -ErrorAction Stop
						}
						catch {
							$serverObj | Add-Member NoteProperty -Name "Warning Details" -Value ($($serverObj."Warning Details") + $string6) -Force
							Write-Host -ForegroundColor $warn $string6
							Write-Warning $_.Exception.Message
						}
						
						if ($q)
						{
							$qcount = $q | Measure-Object MessageCount -Sum
							[int]$qlength = $qcount.sum
							$serverObj | Add-Member NoteProperty -Name "Queue Length" -Value $qlength -Force
							if ($qlength -le $transportqueuewarn)
							{
								Write-Host -ForegroundColor $pass $qlength
								$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Pass" -Force
							}
							elseif ($qlength -gt $transportqueuewarn -and $qlength -lt $transportqueuehigh)
							{
								Write-Host -ForegroundColor $warn $qlength
								$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Warn" -Force
							}
							else
							{
								Write-Host -ForegroundColor $fail $qlength
								$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Fail" -Force
							}
						}
						else
						{
							$serverObj | Add-Member NoteProperty -Name "Transport Queue" -Value "Unknown" -Force
						}
					}
					#END - Hub Transport Server Check

					#START - Mailbox Server Check
					if ($IsMB)
					{
						#Get the PF and MB databases
						[array]$pfdbs = @(Get-PublicFolderDatabase -server $server -status -WarningAction SilentlyContinue)
						[array]$mbdbs = @(Get-MailboxDatabase -server $server -status | Where {$_.Recovery -ne $true})
                        
                        if ($version -eq "Exchange 2010")
                        {
						    [array]$activedbs = @(Get-MailboxDatabase -server $server -status | Where {$_.MountedOnServer -eq ($serverinfo.fqdn)})
                        }
                        else
                        {
                            [array]$activedbs = $mbdbs
                        }
						
						#START - Database Mount Check
						
						#Check public folder databases
						if ($pfdbs.count -gt 0)
						{
							Write-Host "Public Folder databases mounted: " -NoNewline;
							[string]$pfdbstatus = "Pass"
							[array]$alertdbs = @()
							foreach ($db in $pfdbs)
							{
								if (($db.mounted) -ne $true)
								{
									$pfdbstatus = "Fail"
									$alertdbs += $db.name
								}
							}

							$serverObj | Add-Member NoteProperty -Name "PF DBs Mounted" -Value $pfdbstatus -Force
							
							if ($alertdbs.count -eq 0)
							{
								Write-Host -ForegroundColor $pass $pfdbstatus
							}
							else
							{
								Write-Host -ForegroundColor $fail $pfdbstatus
								$serverObj | Add-Member NoteProperty -Name "Error Details" -Value ($($serverObj."Error Details") + $string7) -Force
								Write-Host "Offline databases:"
								foreach ($al in $alertdbs)
								{
									Write-Host -ForegroundColor $fail `t$al
								}
							}
						}
						
						#Check mailbox databases
						if ($mbdbs.count -gt 0)
						{
							[string]$mbdbstatus = "Pass"
							[array]$alertdbs = @()

							Write-Host "Mailbox databases mounted: " -NoNewline;
							foreach ($db in $mbdbs)
							{
								if (($db.mounted) -ne $true)
								{
									$mbdbstatus = "Fail"
									$alertdbs += $db.name
								}
							}

							$serverObj | Add-Member NoteProperty -Name "MB DBs Mounted" -Value $mbdbstatus -Force							
							
							if ($alertdbs.count -eq 0)
							{
								Write-Host -ForegroundColor $pass $mbdbstatus
							}
							else
							{
								$serverObj | Add-Member NoteProperty -Name "Error Details" -Value ($($serverObj."Error Details") + $string9) -Force
								Write-Host -ForegroundColor $fail $mbdbstatus
								Write-Host "Offline databases: "
								foreach ($al in $alertdbs)
								{
									Write-Host -ForegroundColor $fail `t$al
								}
							}
						}
						
						#END - Database Mount Check
						
						#START - MAPI Connectivity Test
						if ($activedbs.count -gt 0 -or $pfdbs.count -gt 0 -or $version -eq "Exchange 2007")
						{
							[string]$mbdbstatus = "Pass"
							[array]$alertdbs = @()
							Write-Host "MAPI connectivity: " -NoNewline;
							foreach ($db in $mbdbs)
							{
								$mapistatus = Test-MapiConnectivity -Database $db.Identity -PerConnectionTimeout $mapitimeout
                                if ($mapistatus.Result.Value -eq $null)
                                {
                                    $mapiresult = $mapistatus.Result
                                }
                                else
                                {
                                    $mapiresult = $mapistatus.Result.Value
                                }
                                if (($mapiresult) -ne "Success")
								{
									$mbdbstatus = "Fail"
									$alertdbs += $db.name
								}
							}

							$serverObj | Add-Member NoteProperty -Name "MAPI Test" -Value $mbdbstatus -Force
							
							if ($alertdbs.count -eq 0)
							{
								Write-Host -ForegroundColor $pass $mbdbstatus
							}
							else
							{
								$serverObj | Add-Member NoteProperty -Name "Error Details" -Value ($($serverObj."Error Details") + $string10) -Force
								Write-Host -ForegroundColor $fail $mbdbstatus
								Write-Host "MAPI failed to: "
								foreach ($al in $alertdbs)
								{
									Write-Host -ForegroundColor $fail `t$al
								}
							}
						}
						#END - MAPI Connectivity Test
						
						#START - Mail Flow Test
						if ($activedbs.count -gt 0 -or ($version -eq "Exchange 2007" -and $mbdbs.count -gt 0))
						{
							$flow = $null
							$testmailflowresult = $null
							
							Write-Host "Mail flow test: " -NoNewline;
							try
							{
								$flow = Test-Mailflow $server -ErrorAction Stop
							}
							catch
							{
								$testmailflowresult = $_.Exception.Message
							}
							
							if ($flow)
							{
								$testmailflowresult = $flow.testmailflowresult
							}

							if ($testmailflowresult -eq "Success")
							{
								Write-Host -ForegroundColor $pass $testmailflowresult
								$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "Pass" -Force
							}
							else
							{
								$serverObj | Add-Member NoteProperty -Name "Error Details" -Value ($($serverObj."Error Details") + $string11) -Force
								Write-Host -ForegroundColor $fail $testmailflowresult
								$serverObj | Add-Member NoteProperty -Name "Mail Flow Test" -Value "Fail" -Force
							}
						}
						else
						{
							Write-Host "Mail flow test: No active mailbox databases"
						}
						#END - Mail Flow Test
					}
					#END - Mailbox Server Check

				}
				#END - Exchange 2007/2010 Health Checks
				$report = $report + $serverObj
			}
			else
			{
				#Server is not reachable and uptime could not be retrieved
				Write-Host -ForegroundColor $warn $string1
				$serverObj | Add-Member NoteProperty -Name "Error Details" -Value ($($serverObj."Error Details") + $string1) -Force
				$serverObj | Add-Member NoteProperty -Name "Ping" -Value "Fail" -Force
				$report = $report + $serverObj
			}
		}
		else
		{
			Write-Host -ForegroundColor $Fail "Fail"
			Write-Host -ForegroundColor $warn $string13
			$serverObj | Add-Member NoteProperty -Name "Error Details" -Value ($($serverObj."Error Details") + $string13) -Force
			$serverObj | Add-Member NoteProperty -Name "DNS" -Value "Fail" -Force
			$report = $report + $serverObj
		}
	}	
}

### Begin DAG Health Report
Write-Verbose "Retrieving Database Availability Groups"
$dags = @(Get-DatabaseAvailabilityGroup -Status)
Write-Verbose "$($dags.count) DAGs found"

if ($($dags.count) -gt 0)
{

	foreach ($dag in $dags)
	{

		#Strings for use in the HTML report/email
		$dagsummaryintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Health Summary:</p>"
		$dagdetailintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Health Details:</p>"
		$dagmemberintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Member Health:</p>"

		$dagdbcopyReport = @()		#Database copy health report
		$dagciReport = @()			#Content Index health report
		$dagmemberReport = @()		#DAG member server health report
		$dagdatabaseSummary = @()	#Database health summary report
		
		Write-Verbose "---- Processing DAG $($dag.Name)"
		
		$dagmembers = @($dag | Select-Object -ExpandProperty Servers | Sort-Object Name)
		Write-Verbose "$($dagmembers.count) DAG members found"
		
		$dagdatabases = @(Get-MailboxDatabase -Status | Where-Object {$_.MasterServerOrAvailabilityGroup -eq $dag.Name} | Sort-Object Name)
		Write-Verbose "$($dagdatabases.count) DAG databases found"
		
		foreach ($database in $dagdatabases)
		{
			Write-Verbose "---- Processing database $database"

			#Custom object for Database
			$objectHash = @{
				"Database" = $database.Identity
				"Mounted on" = "Unknown"
				"Preference" = $null
				"Total Copies" = $null
				"Healthy Copies" = $null
				"Unhealthy Copies" = $null
				"Healthy Queues" = $null
				"Unhealthy Queues" = $null
				"Lagged Queues" = $null
				"Healthy Indexes" = $null
				"Unhealthy Indexes" = $null
				}
			$databaseObj = New-Object PSObject -Property $objectHash

			$dbcopystatus = @($database | Get-MailboxDatabaseCopyStatus)
			Write-Verbose "$database has $($dbcopystatus.Count) copies"
			foreach ($dbcopy in $dbcopystatus)
			{
				#Custom object for DB copy
				$objectHash = @{
					"Database Copy" = $dbcopy.Identity
					"Database Name" = $dbcopy.DatabaseName
					"Mailbox Server" = $null
					"Activation Preference" = $null
					"Status" = $null
					"Copy Queue" = $null
					"Replay Queue" = $null
					"Replay Lagged" = $null
					"Truncation Lagged" = $null
					"Content Index" = $null
					}
				$dbcopyObj = New-Object PSObject -Property $objectHash
				
				Write-Verbose "Database Copy: $($dbcopy.Identity)"
				
				$mailboxserver = $dbcopy.MailboxServer
				Write-Verbose "Server: $mailboxserver"

				$pref = ($database | Select-Object -ExpandProperty ActivationPreference | Where-Object {$_.Key -eq $mailboxserver}).Value
				Write-Verbose "Activation Preference: $pref"

				$copystatus = $dbcopy.Status
				Write-Verbose "Status: $copystatus"
				
				[int]$copyqueuelength = $dbcopy.CopyQueueLength
				Write-Verbose "Copy Queue: $copyqueuelength"
				
				[int]$replayqueuelength = $dbcopy.ReplayQueueLength
				Write-Verbose "Replay Queue: $replayqueuelength"
				
				$contentindexstate = $dbcopy.ContentIndexState
				Write-Verbose "Content Index: $contentindexstate"

				#Checking whether this is a replay lagged copy
				$replaylagcopies = @($database | Select -ExpandProperty ReplayLagTimes | Where-Object {$_.Value -gt 0})
				if ($($replaylagcopies.count) -gt 0)
	            {
	                [bool]$replaylag = $false
	                foreach ($replaylagcopy in $replaylagcopies)
				    {
					    if ($replaylagcopy.Key -eq $mailboxserver)
					    {
						    Write-Verbose "$database is replay lagged on $mailboxserver"
						    [bool]$replaylag = $true
					    }
				    }
	            }
	            else
				{
				   [bool]$replaylag = $false
				}
	            Write-Verbose "Replay lag is $replaylag"
						
				#Checking for truncation lagged copies
				$truncationlagcopies = @($database | Select -ExpandProperty TruncationLagTimes | Where-Object {$_.Value -gt 0})
				if ($($truncationlagcopies.count) -gt 0)
	            {
	                [bool]$truncatelag = $false
	                foreach ($truncationlagcopy in $truncationlagcopies)
				    {
					    if ($truncationlagcopy.Key -eq $mailboxserver)
					    {
						    [bool]$truncatelag = $true
					    }
				    }
	            }
	            else
				{
				   [bool]$truncatelag = $false
				}
	            Write-Verbose "Truncation lag is $truncatelag"
				
				$dbcopyObj | Add-Member NoteProperty -Name "Mailbox Server" -Value $mailboxserver -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Activation Preference" -Value $pref -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Status" -Value $copystatus -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Copy Queue" -Value $copyqueuelength -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Replay Queue" -Value $replayqueuelength -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Replay Lagged" -Value $replaylag -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Truncation Lagged" -Value $truncatelag -Force
				$dbcopyObj | Add-Member NoteProperty -Name "Content Index" -Value $contentindexstate -Force
				
				$dagdbcopyReport += $dbcopyObj
			}
		
			$copies = @($dagdbcopyReport | Where-Object { ($_."Database Name" -eq $database) })
		
			$mountedOn = ($copies | Where-Object { ($_.Status -eq "Mounted") })."Mailbox Server"
			if ($mountedOn)
			{
				$databaseObj | Add-Member NoteProperty -Name "Mounted on" -Value $mountedOn -Force
			}
		
			$activationPref = ($copies | Where-Object { ($_.Status -eq "Mounted") })."Activation Preference"
			$databaseObj | Add-Member NoteProperty -Name "Preference" -Value $activationPref -Force

			$totalcopies = $copies.count
			$databaseObj | Add-Member NoteProperty -Name "Total Copies" -Value $totalcopies -Force
		
			$healthycopies = @($copies | Where-Object { (($_.Status -eq "Mounted") -or ($_.Status -eq "Healthy")) }).Count
			$databaseObj | Add-Member NoteProperty -Name "Healthy Copies" -Value $healthycopies -Force
			
			$unhealthycopies = @($copies | Where-Object { (($_.Status -ne "Mounted") -and ($_.Status -ne "Healthy")) }).Count
			$databaseObj | Add-Member NoteProperty -Name "Unhealthy Copies" -Value $unhealthycopies -Force

			$healthyqueues  = @($copies | Where-Object { (($_."Copy Queue" -lt $replqueuewarning) -and (($_."Replay Queue" -lt $replqueuewarning)) -and ($_."Replay Lagged" -eq $false)) }).Count
	        $databaseObj | Add-Member NoteProperty -Name "Healthy Queues" -Value $healthyqueues -Force

			$unhealthyqueues = @($copies | Where-Object { (($_."Copy Queue" -ge $replqueuewarning) -or (($_."Replay Queue" -ge $replqueuewarning) -and ($_."Replay Lagged" -eq $false))) }).Count
			$databaseObj | Add-Member NoteProperty -Name "Unhealthy Queues" -Value $unhealthyqueues -Force

			$laggedqueues = @($copies | Where-Object { ($_."Replay Lagged" -eq $true) -or ($_."Truncation Lagged" -eq $true) }).Count
			$databaseObj | Add-Member NoteProperty -Name "Lagged Queues" -Value $laggedqueues -Force

			$healthyindexes = @($copies | Where-Object { ($_."Content Index" -eq "Healthy") }).Count
			$databaseObj | Add-Member NoteProperty -Name "Healthy Indexes" -Value $healthyindexes -Force
			
			$unhealthyindexes = @($copies | Where-Object { ($_."Content Index" -ne "Healthy") }).Count
			$databaseObj | Add-Member NoteProperty -Name "Unhealthy Indexes" -Value $unhealthyindexes -Force
			
			$dagdatabaseSummary += $databaseObj
		
		}
		
		#Get Test-Replication Health results for each DAG member
		foreach ($dagmember in $dagmembers)
		{
			$memberObj = New-Object PSObject
			$memberObj | Add-Member NoteProperty -Name "Server" -Value $($dagmember.Name)
		
			Write-Verbose "---- Checking replication health for $($dagmember.Name)"
			$replicationhealth = $dagmember | Invoke-Command {Test-ReplicationHealth}
			foreach ($healthitem in $replicationhealth)
			{
				Write-Verbose "$($healthitem.Check) $($healthitem.Result)"
				$memberObj | Add-Member NoteProperty -Name $($healthitem.Check) -Value $($healthitem.Result)
			}
			$dagmemberReport += $memberObj
		}

		
		#Roll the HTML
		if ($SendEmail -or $ReportFile)
		{
		
			####Begin Summary Table HTML
			$dagdatabaseSummaryHtml = $null
			#Begin Summary table HTML header
			$htmltableheader = "<p>
							<table>
							<tr>
							<th>Database</th>
							<th>Mounted on</th>
							<th>Preference</th>
							<th>Total Copies</th>
							<th>Healthy Copies</th>
							<th>Unhealthy Copies</th>
							<th>Healthy Queues</th>
							<th>Unhealthy Queues</th>
							<th>Lagged Queues</th>
							<th>Healthy Indexes</th>
							<th>Unhealthy Indexes</th>
							</tr>"

			$dagdatabaseSummaryHtml += $htmltableheader
			#End Summary table HTML header
			
			#Begin Summary table HTML rows
			foreach ($line in $dagdatabaseSummary)
			{
				$htmltablerow = "<tr>"
				$htmltablerow += "<td><strong>$($line.Database)</strong></td>"
				
				#Warn if mounted server is still unknown
				switch ($($line."Mounted on"))
				{
					"Unknown" { $htmltablerow += "<td class=""warn"">$($line."Mounted on")</td>" }
					default { $htmltablerow += "<td>$($line."Mounted on")</td>" }
				}
				
				#Warn if DB is mounted on a server that is not Activation Preference 1
				if ($($line.Preference) -gt 1)
				{
					$htmltablerow += "<td class=""warn"">$($line.Preference)</td>"		
				}
				else
				{
					$htmltablerow += "<td class=""pass"">$($line.Preference)</td>"
				}
				
				$htmltablerow += "<td>$($line."Total Copies")</td>"
				
				#Show as info if health copies is 1 but total copies also 1,
	            #Warn if healthy copies is 1, Fail if 0
				switch ($($line."Healthy Copies"))
				{	
					0 {$htmltablerow += "<td class=""fail"">$($line."Healthy Copies")</td>"}
					1 {
						if ($($line."Total Copies") -eq $($line."Healthy Copies"))
						{
							$htmltablerow += "<td class=""info"">$($line."Healthy Copies")</td>"
						}
						else
						{
							$htmltablerow += "<td class=""warn"">$($line."Healthy Copies")</td>"
						}
					  }
					default {$htmltablerow += "<td class=""pass"">$($line."Healthy Copies")</td>"}
				}

				#Warn if unhealthy copies is 1, fail if more than 1
				switch ($($line."Unhealthy Copies"))
				{
					0 {	$htmltablerow += "<td class=""pass"">$($line."Unhealthy Copies")</td>" }
					1 {	$htmltablerow += "<td class=""warn"">$($line."Unhealthy Copies")</td>" }
					default { $htmltablerow += "<td class=""fail"">$($line."Unhealthy Copies")</td>" }
				}

				#Warn if healthy queues + lagged queues is less than total copies
				#Fail if no healthy queues
				if ($($line."Total Copies") -eq ($($line."Healthy Queues") + $($line."Lagged Queues")))
				{
					$htmltablerow += "<td class=""pass"">$($line."Healthy Queues")</td>"
				}
				else
				{
					switch ($($line."Healthy Queues"))
					{
						0 { $htmltablerow += "<td class=""fail"">$($line."Healthy Queues")</td>" }
						default { $htmltablerow += "<td class=""warn"">$($line."Healthy Queues")</td>" }
					}
				}
				
				#Fail if unhealthy queues = total queues
				#Warn if more than one unhealthy queue
				if ($($line."Total Queues") -eq $($line."Unhealthy Queues"))
				{
					$htmltablerow += "<td class=""fail"">$($line."Unhealthy Queues")</td>"
				}
				else
				{
					switch ($($line."Unhealthy Queues"))
					{
						0 { $htmltablerow += "<td class=""pass"">$($line."Unhealthy Queues")</td>" }
						default { $htmltablerow += "<td class=""warn"">$($line."Unhealthy Queues")</td>" }
					}
				}
				
				#Info for lagged queues
				switch ($($line."Lagged Queues"))
				{
					0 { $htmltablerow += "<td>$($line."Lagged Queues")</td>" }
					default { $htmltablerow += "<td class=""info"">$($line."Lagged Queues")</td>" }
				}
				
				#Pass if healthy indexes = total copies
				#Warn if healthy indexes less than total copies
				#Fail if healthy indexes = 0
				if ($($line."Total Copies") -eq $($line."Healthy Indexes"))
				{
					$htmltablerow += "<td class=""pass"">$($line."Healthy Indexes")</td>"
				}
				else
				{
					switch ($($line."Healthy Indexes"))
					{
						0 { $htmltablerow += "<td class=""fail"">$($line."Healthy Indexes")</td>" }
						default { $htmltablerow += "<td class=""warn"">$($line."Healthy Indexes")</td>" }
					}
				}
				
				#Fail if unhealthy indexes = total copies
				#Warn if unhealthy indexes 1 or more
				#Pass if unhealthy indexes = 0
				if ($($line."Total Copies") -eq $($line."Unhealthy Indexes"))
				{
					$htmltablerow += "<td class=""fail"">$($line."Unhealthy Indexes")</td>"
				}
				else
				{
					switch ($($line."Unhealthy Indexes"))
					{
						0 { $htmltablerow += "<td class=""pass"">$($line."Unhealthy Indexes")</td>" }
						default { $htmltablerow += "<td class=""warn"">$($line."Unhealthy Indexes")</td>" }
					}
				}
				
				$htmltablerow += "</tr>"
				$dagdatabaseSummaryHtml += $htmltablerow
			}
			$dagdatabaseSummaryHtml += "</table>
									</p>"
			#End Summary table HTML rows
			####End Summary Table HTML

			####Begin Detail Table HTML
			$databasedetailsHtml = $null
			#Begin Detail table HTML header
			$htmltableheader = "<p>
							<table>
							<tr>
							<th>Database Copy</th>
							<th>Database Name</th>
							<th>Mailbox Server</th>
							<th>Activation Preference</th>
							<th>Status</th>
							<th>Copy Queue</th>
							<th>Replay Queue</th>
							<th>Replay Lagged</th>
							<th>Truncation Lagged</th>
							<th>Content Index</th>
							</tr>"

			$databasedetailsHtml += $htmltableheader
			#End Detail table HTML header
			
			#Begin Detail table HTML rows
			foreach ($line in $dagdbcopyReport)
			{
				$htmltablerow = "<tr>"
				$htmltablerow += "<td><strong>$($line."Database Copy")</strong></td>"
				$htmltablerow += "<td>$($line."Database Name")</td>"
				$htmltablerow += "<td>$($line."Mailbox Server")</td>"
				$htmltablerow += "<td>$($line."Activation Preference")</td>"
				
				Switch ($($line."Status"))
				{
					"Healthy" { $htmltablerow += "<td class=""pass"">$($line."Status")</td>" }
					"Mounted" { $htmltablerow += "<td class=""pass"">$($line."Status")</td>" }
					"Failed" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
					"FailedAndSuspended" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
					"ServiceDown" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
					"Dismounted" { $htmltablerow += "<td class=""fail"">$($line."Status")</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line."Status")</td>" }
				}
				
				if ($($line."Copy Queue") -lt $replqueuewarning)
				{
					$htmltablerow += "<td class=""pass"">$($line."Copy Queue")</td>"
				}
				else
				{
					$htmltablerow += "<td class=""warn"">$($line."Copy Queue")</td>"
				}
				
				if (($($line."Replay Queue") -lt $replqueuewarning) -or ($($line."Replay Lagged") -eq $true))
				{
					$htmltablerow += "<td class=""pass"">$($line."Replay Queue")</td>"
				}
				else
				{
					$htmltablerow += "<td class=""warn"">$($line."Replay Queue")</td>"
				}
				

				Switch ($($line."Replay Lagged"))
				{
					$true { $htmltablerow += "<td class=""info"">$($line."Replay Lagged")</td>" }
					default { $htmltablerow += "<td>$($line."Replay Lagged")</td>" }
				}

				Switch ($($line."Truncation Lagged"))
				{
					$true { $htmltablerow += "<td class=""info"">$($line."Truncation Lagged")</td>" }
					default { $htmltablerow += "<td>$($line."Truncation Lagged")</td>" }
				}
				
				Switch ($($line."Content Index"))
				{
					"Healthy" { $htmltablerow += "<td class=""pass"">$($line."Content Index")</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line."Content Index")</td>" }
				}
				
				$htmltablerow += "</tr>"
				$databasedetailsHtml += $htmltablerow
			}
			$databasedetailsHtml += "</table>
									</p>"
			#End Detail table HTML rows
			####End Detail Table HTML
			
			
			####Begin Member Table HTML
			$dagmemberHtml = $null
			#Begin Member table HTML header
			$htmltableheader = "<p>
								<table>
								<tr>
								<th>Server</th>
								<th>Cluster Service</th>
								<th>Replay Service</th>
								<th>Active Manager</th>
								<th>Tasks RPC Listener</th>
								<th>TCP Listener</th>
								<th>DAG Members Up</th>
								<th>Cluster Network</th>
								<th>Quorum Group</th>
								<th>File Share Quorum</th>
								<th>DB Copy Suspended</th>
								<th>DB Initializing</th>
								<th>DB Disconnected</th>
								<th>DB Log Copy Keeping Up</th>
								<th>DB Log Replay Keeping Up</th>
								</tr>"
			
			$dagmemberHtml += $htmltableheader
			#End Member table HTML header
			
			#Begin Member table HTML rows
			foreach ($line in $dagmemberReport)
			{
				$htmltablerow = "<tr>"
				$htmltablerow += "<td><strong>$($line."Server")</strong></td>"

				Switch ($($line.ClusterService))
				{
					$null { $htmltablerow += "<td>$($line.ClusterService)</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.ClusterService)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.ClusterService)</td>" }
				}
				
				Switch ($($line.ReplayService))
				{
					$null { $htmltablerow += "<td>$($line.ReplayService)</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.ReplayService)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.ReplayService)</td>" }
				}

				Switch ($($line.ActiveManager))
				{
					$null { $htmltablerow += "<td>$($line.ActiveManager)</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.ActiveManager)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.ActiveManager)</td>" }
				}
				
				Switch ($($line.TasksRPCListener))
				{
					$null { $htmltablerow += "<td>$($line.TasksRPCListener)</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.TasksRPCListener)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.TasksRPCListener)</td>" }
				}			
				
				Switch ($($line.TCPListener))
				{
					$null { $htmltablerow += "<td>$($line.TCPListener)</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.TCPListener)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.TCPListener)</td>" }
				}
				
				Switch ($($line.DAGMembersUp))
				{
					$null { $htmltablerow += "<td>$($line.DAGMembersUp)</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.DAGMembersUp)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.DAGMembersUp)</td>" }
				}
				
				Switch ($($line.ClusterNetwork))
				{
					$null { $htmltablerow += "<td>$($line.ClusterNetwork)</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.ClusterNetwork)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.ClusterNetwork)</td>" }
				}
				
				Switch ($($line.QuorumGroup))
				{
					$null { $htmltablerow += "<td>$($line.QuorumGroup)</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.QuorumGroup)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.QuorumGroup)</td>" }
				}
				
				Switch ($($line.FileShareQuorum))
				{
					$null { $htmltablerow += "<td>n/a</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.FileShareQuorum)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.FileShareQuorum)</td>" }
				}
				
				Switch ($($line.DBCopySuspended))
				{
					$null { $htmltablerow += "<td>n/a</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.DBCopySuspended)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.DBCopySuspended)</td>" }
				}
				
				Switch ($($line.DBInitializing))
				{
					$null { $htmltablerow += "<td>n/a</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.DBInitializing)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.DBInitializing)</td>" }
				}
				
				Switch ($($line.DBDisconnected))
				{
					$null { $htmltablerow += "<td>n/a</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.DBDisconnected)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.DBDisconnected)</td>" }
				}
				
				Switch ($($line.DBLogCopyKeepingUp))
				{
					$null { $htmltablerow += "<td>n/a</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.DBLogCopyKeepingUp)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.DBLogCopyKeepingUp)</td>" }
				}
				Switch ($($line.DBLogReplayKeepingUp))
				{
					$null { $htmltablerow += "<td>n/a</td>" }
					"Passed" { $htmltablerow += "<td class=""pass"">$($line.DBLogReplayKeepingUp)</td>" }
					default { $htmltablerow += "<td class=""warn"">$($line.DBLogReplayKeepingUp)</td>" }
				}
				$htmltablerow += "</tr>"
				$dagmemberHtml += $htmltablerow
			}
			$dagmemberHtml += "</table>
			</p>"
		}
		
		#Output the report objects to console, and optionally to email and HTML file
		#Forcing table format for console output due to issue with multiple output
		#objects that have different layouts

		Write-Host "---- Database Copy Health Summary ----"
		$dagdatabaseSummary | ft
				
		Write-Host "---- Database Copy Health Details ----"
		$dagdbcopyReport | ft
		
		Write-Host "`r`n---- Server Test-Replication Report ----`r`n"
		$dagmemberReport | ft
		
		if ($SendEmail -or $ReportFile)
		{
			$dagreporthtml = $dagsummaryintro + $dagdatabaseSummaryHtml + $dagdetailintro + $databasedetailsHtml + $dagmemberintro + $dagmemberHtml
			$dagreportbody += $dagreporthtml
		}
		
	}
}
else
{
	$dagreporthtml = "<p>No database availability groups found.</p>"
}
###End DAG Health Report

Write-Host $string16
#Generate the report
if ($ReportMode -or $SendEmail)
{
	#Get report generation timestamp
	$reportime = Get-Date

	#Generate report summary
	$summaryreport = $report | select Server,"Error Details","Warning Details" | Where {$_."Error Details" -ne "" -or $_."Warning Details" -ne ""}

	#Create HTML Report
	#Common HTML head and styles
	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 16px;}
				H2{font-size: 14px;}
				H3{font-size: 12px;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid black; padding: 5px; }
				td.pass{background: #7FFF00;}
				td.warn{background: #FFE600;}
				td.fail{background: #FF0000; color: #ffffff;}
				</style>
				<body>
				<h1 align=""center"">Exchange Server Health Check Report</h1>
				<h3 align=""center"">Generated: $reportime</h3>"

	if ($summaryreport)
	{
		$summaryhtml = "<h3>Exchange Server Health Check Summary</h3>
						<p>The following errors and warnings were detected.</p>
						<p>
						<table>
						<tr>
						<th>Server</th>
						<th>Errors</th>
						<th>Warnings</th>
						</tr>"
		foreach ($reportline in $summaryreport)
		{
			$htmltablerow = "<tr>"
						$htmltablerow += "<td>$($reportline.server)</td>"
						switch ($($reportline."Error Details"))
						{
							"" { $htmltablerow += "<td>$($reportline."Error Details")</td>" }
							default { $htmltablerow += "<td class=""fail"">$($reportline."Error Details")</td>" }
						}
						switch ($($reportline."Warning Details"))
						{
							"" { $htmltablerow += "<td>$($reportline."Warning Details")</td>" }
							default { $htmltablerow += "<td class=""warn"">$($reportline."Warning Details")</td>" }
						}
						
						$summaryhtml = $summaryhtml + $htmltablerow
		}
		$summaryhtml = $summaryhtml + "</table></p>"
		$alerts = $true
	}
	else
	{
		$summaryhtml = "<h3>Exchange Server Health Check Summary</h3>
						<p>No Exchange server health alerts or warnings.</p>"
		$alerts = $false
	}


	#Exchange 2007/2010 Report Table Header
	$htmltableheader = "<h3>Exchange Server 2007/2010 Health</h3>
						<p>
						<table>
						<tr>
						<th>Server</th>
						<th>Site</th>
						<th>Roles</th>
						<th>Version</th>
						<th>DNS</th>
						<th>Ping</th>
						<th>Uptime (hrs)</th>
						<th>Client Access Server Role Services</th>
						<th>Hub Transport Server Role Services</th>
						<th>Mailbox Server Role Services</th>
						<th>Unified Messaging Server Role Services</th>
						<th>Transport Queue</th>
						<th>PF DBs Mounted</th>
						<th>MB DBs Mounted</th>
						<th>MAPI Test</th>
						<th>Mail Flow Test</th>
						</tr>"

	#Exchange 2007/2010 Report Table
	$htmltable = $htmltable + $htmltableheader					
						
	foreach ($reportline in $report)
	{
		$htmltablerow = "<tr>"
		$htmltablerow += "<td>$($reportline.server)</td>"
		$htmltablerow += "<td>$($reportline.site)</td>"
		$htmltablerow += "<td>$($reportline.roles)</td>"
		$htmltablerow += "<td>$($reportline.version)</td>"					
						
		switch ($($reportline.dns))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline.dns)</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline.dns)</td>"}
		}
						
		switch ($($reportline.ping))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline.ping)</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline.ping)</td>"}
		}
		
		if ($($reportline."uptime (hrs)") -eq "Access Denied")
		{
			$htmltablerow += "<td class=""warn"">Access Denied</td>"		
		}
        elseif ($($reportline."uptime (hrs)") -eq $string17)
        {
            $htmltablerow += "<td class=""warn"">$string17</td>"
        }
		else
		{
			$hours = [int]$($reportline."uptime (hrs)")
			if ($hours -le 24)
			{
				$htmltablerow += "<td class=""warn"">$hours</td>"
			}
			else
			{
				$htmltablerow += "<td class=""pass"">$hours</td>"
			}
		}

		switch ($($reportline."Client Access Server Role Services"))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline."Client Access Server Role Services")</td>"}
			"Warn" {$htmltablerow += "<td class=""warn"">$($reportline."Client Access Server Role Services")</td>"}
			"Access Denied" {$htmltablerow += "<td class=""warn"">$($reportline."Client Access Server Role Services")</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline."Client Access Server Role Services")</td>"}
            "Could not test service health. " {$htmltablerow += "<td class=""warn"">$($reportline."Unified Messaging Server Role Services")</td>"}
			default {$htmltablerow += "<td>$($reportline."Client Access Server Role Services")</td>"}
		}
		
		switch ($($reportline."Hub Transport Server Role Services"))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline."Hub Transport Server Role Services")</td>"}
			"Warn" {$htmltablerow += "<td class=""warn"">$($reportline."Hub Transport Server Role Services")</td>"}
			"Access Denied" {$htmltablerow += "<td class=""warn"">$($reportline."Hub Transport Server Role Services")</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline."Hub Transport Server Role Services")</td>"}
            "Could not test service health. " {$htmltablerow += "<td class=""warn"">$($reportline."Unified Messaging Server Role Services")</td>"}
			default {$htmltablerow += "<td>$($reportline."Hub Transport Server Role Services")</td>"}
		}
		
		switch ($($reportline."Mailbox Server Role Services"))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline."Mailbox Server Role Services")</td>"}
			"Warn" {$htmltablerow += "<td class=""warn"">$($reportline."Mailbox Server Role Services")</td>"}
			"Access Denied" {$htmltablerow += "<td class=""warn"">$($reportline."Mailbox Server Role Services")</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline."Mailbox Server Role Services")</td>"}
            "Could not test service health. " {$htmltablerow += "<td class=""warn"">$($reportline."Unified Messaging Server Role Services")</td>"}
			default {$htmltablerow += "<td>$($reportline."Mailbox Server Role Services")</td>"}
		}
		
		switch ($($reportline."Unified Messaging Server Role Services"))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline."Unified Messaging Server Role Services")</td>"}
			"Warn" {$htmltablerow += "<td class=""warn"">$($reportline."Unified Messaging Server Role Services")</td>"}
			"Access Denied" {$htmltablerow += "<td class=""warn"">$($reportline."Unified Messaging Server Role Services")</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline."Unified Messaging Server Role Services")</td>"}
            "Could not test service health. " {$htmltablerow += "<td class=""warn"">$($reportline."Unified Messaging Server Role Services")</td>"}
			default {$htmltablerow += "<td>$($reportline."Unified Messaging Server Role Services")</td>"}
		}
						
		switch ($($reportline."Transport Queue"))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline."Transport Queue") ($($reportline."Queue Length"))</td>"}
			"Warn" {$htmltablerow += "<td class=""warn"">$($reportline."Transport Queue") ($($reportline."Queue Length"))</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline."Transport Queue") ($($reportline."Queue Length"))</td>"}
			"Unknown" {$htmltablerow += "<td class=""warn"">$($reportline."Transport Queue")</td>"}
			default {$htmltablerow += "<td >$($reportline."Transport Queue")</td>"}
		}

		switch ($($reportline."PF DBs Mounted"))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline."PF DBs Mounted")</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline."PF DBs Mounted")</td>"}
			default {$htmltablerow += "<td>$($reportline."PF DBs Mounted")</td>"}
		}

		switch ($($reportline."MB DBs Mounted"))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline."MB DBs Mounted")</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline."MB DBs Mounted")</td>"}
			default {$htmltablerow += "<td>$($reportline."MB DBs Mounted")</td>"}
		}

		switch ($($reportline."MAPI Test"))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline."MAPI Test")</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline."MAPI Test")</td>"}
			default {$htmltablerow += "<td>$($reportline."MAPI Test")</td>"}
		}
		
		switch ($($reportline."Mail Flow Test"))
		{
			"Pass" {$htmltablerow += "<td class=""pass"">$($reportline."Mail Flow Test")</td>"}
			"Fail" {$htmltablerow += "<td class=""fail"">$($reportline."Mail Flow Test")</td>"}
			default {$htmltablerow += "<td>$($reportline."Mail Flow Test")</td>"}
		}

		$htmltablerow += "</tr>"
		
		$htmltable = $htmltable + $htmltablerow
	}
	$htmltable = $htmltable + "</table></p>"

	
	$htmltail = "</body>
				</html>"
				

	$htmlreport = $htmlhead + $summaryhtml + $htmltable + $dagreportbody + $htmltail
	
	if ($ReportMode -or $ReportFile)
	{
		$htmlreport | Out-File $ReportFile
	}

	if ($SendEmail)
	{
		if ($alerts -eq $false -and $AlertsOnly -eq $true)
		{
			#Do not send email message
			Write-Host $string19
		}
		else
		{
			#Send email message
			Write-Host $string14
			Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
		}
	}
}

Write-Host $string15

