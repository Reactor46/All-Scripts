<#
 .SYNOPSIS 
This script is meant to get the health of an Exchange environment.

 .DESCRIPTION
Checks a number of things to determine overall Exchange environment health.  Current items checked:
	* Databases Unmounted
	* Databases mounted on unknown
	* Database bad copies
	* Databases not mounted on preferred server
	* Replication Issues on DAG members
	* Inactive Components on all server roles
	* Unhealthy Health Sets on all server roles
	* Automatic Exchange-related services that are not started on all server roles

 PARAMETER EmailOnly
 When this parameter is used, no reports are displayed in the console
 
 PARAMTER ConsoleOnly
 When this paramater is used, the reports are only displayed in the console. No email is sent.
 
 PARAMETER File
 When this paramter is used, a report is written to the filename specified in HTML format.
 
 VARIABLE Server
 Use this variable to specify the server from which to relay email message
 
 VARIABLE Sender
 Use this variable to specify the sender of the email message
 
 VARIABLE Recipient
 Use this variable to specify the recipient of the email message
 
 VARIABLE Server
 Use this variable to specify the server from which to relay alert message
 
 VARIABLE ReplQueueThreshold
 Use this variable to specify the length replication queue required to trigger an alert

 .EXAMPLE
.\Get-EnvironmentHealth.ps1
Does a basic script run through. It will email the report based on the script variables and output data to the console.

.\Get-EnvironmentHealth.ps1 -ConsoleOnly
Does not email report out and only displays report content in console

.\Get-EnvironmentHealth.ps1 -File "Myreport.html"
Will perform default actions and save a copy of the report to the file specified. If no full path is used, the report will be saved to the directory from which the script is run.

 .NOTES
Written By: Paul DiMarino
Database information gathering and some html formating gathered from Paul Cunningham's Get-DagHealth script.

 .LINK
Test-ExchangeHealth - https://talesfromtheshellscript.wordpress.com/2016/12/06/get-a-holistic-view-of-your-entire-microsoft-exchange-organization-with-a-single-script/
Get-DagHealth - http://exchangeserverpro.com/get-daghealth-ps1-database-availability-group-health-check-script/

 .CHANGE LOG
v1.00	- 2016/07/06	- First version started
#>

[CmdletBinding()]
param(
	[Parameter( Mandatory=$false)]
	[switch]$EmailOnly,
	
	[Parameter( Mandatory=$false)]
	[switch]$ConsoleOnly,
	
	[Parameter( Mandatory=$false)]
	[string]$File 
	)

#............................................
#Loads Exchange Shell
#............................................

. $env:ExchangeInstallPath\Bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

#............................................
# Script Variables
#............................................

#$currentdate = Get-Date -Uformat "%Y - %B %d - %A"
$currentdate = Get-Date -Format F
$alertmessagebody = $null

#............................................
# Mail Server Variables
#............................................
$server = "{Your EMail Relay here}"
$sender = "Global Exchange Alerting <{Your sending address here}>"
$recipient = "<{your email address here}>" 

#......................................................
# Script Begins Processing Standalone CAS Role servers
#......................................................

Write-Host "==========	Retrieving Standalone CAS Role Servers for analysis	==========`n" -ForegroundColor Cyan
$casservers = @(Get-exchangeserver | where {$_.serverrole -eq "ClientAccess"})

Write-Host "	$($casservers.count) Standalone CAS Role Servers found in Exchange Environment...`n" -ForegroundColor Green

If (!$casservers)
	{
		$NoCas = "True"
	}
Else
	{
		# Sets Email and File Information
		$cascompintro = "<p><strong>Standalone Client Access Server</strong> Component Alert Summary:</p>"
		$casserviceintro = "<p><strong>Standalone Client Acccess Server</strong> Service Alert Summary:</p>"
		$cashealthintro = "<p><strong>Standalone Client Access Server</strong> Health Set Alert Summary:</p>"	

		$cascomp = @()			# Component Report
		$casservicereport = @()		# Service Report
		$cashealthsetreport= @()	# Health Set Report

		ForEach ($cas in $casservers)
			{
				###Begins the Component Report###
				Write-Host "`n`n==========	Processing Component States on $($cas.name)		==========`n" -ForegroundColor Cyan
		
				$casCompObj = New-Object PSObject
				$casCompObj | Add-Member NoteProperty -Name "Server" -Value $($cas.Name)
				$site = (get-exchangeserver $cas.name).site
				$site = $site -replace ".*/"
				$casCompObj | Add-Member NoteProperty -Name "Site" -Value $site
				$compstate = ($cas | Invoke-Command {get-servercomponentstate})
				$inactive = ($compstate | where {$compstate.state -like "InActive"}).component
				If ($inactive)
					{
						$InactiveArray = [system.string]::Join(" - ", $inactive)
						ForEach ($badcomp in $inactive) 
						{
							Write-Host "	The $($badcomp) component is in an INACTIVE state!!!" -ForegroundColor Red
							Write-Host "	Attempting to start the $($badcomp) component..." -ForegroundColor Yellow
							Set-ServerComponentState $cas.name -Component $badcomp -Requester HealthAPI -State Active
							If (!$cascomperror)
								{
								$cascomperror = "TRUE"
								If ($subject)
									{
									$Subject += ",CAS Components"
									}
								Else
									{
									$subject = "Global Exchange Active Alerts: CAS Components"
									}
								}
						}
					}	
				ElseIf (!$inactive) 
					{
						Write-Host "	All components on $($cas.name) are in an Active state..." -ForegroundColor Green
						$InactiveArray = "None" 
					}
				$casCompObj | Add-Member NoteProperty -Name "InActive Components" -Value $InactiveArray
				$cascomp += $casCompObj

				###Begins the Service Report###
				Write-Host "`n`n==========	Processing Service States on $($cas.name)			==========`n" -ForegroundColor Cyan

				$ServiceObj = New-Object PSObject
				$ServiceObj | Add-Member NoteProperty -Name "Server" -Value $($cas.Name)
				$ServiceObj | Add-Member NoteProperty -Name "Site" -Value $site
				$Stoppedservices = (get-Service -computername $cas.name | where {($_.name -like "MSExch*" -OR $_.name -like "IIS*" -OR $_.name -like "W3SVC") -AND $_.status -like "Stopped"} | Select-Object Displayname)
				If ($StoppedServices)
					{
						$StoppedArray = [system.string]::Join(" - ", $StoppedServices.displayname)
						ForEach ($svc in $stoppedservices)
							{
								Write-Host "	The $($svc.displayname) service is not running!!!" -ForegroundColor Red
								If (!$casserviceerror)
									{
									$casserviceerror = "TRUE"
									If ($subject)
										{
										$Subject += ",CAS Services"
										}
									Else
										{
										$subject = "Global Exchange Active Alerts: CAS Services"
										}
									}
							}
					}
				ElseIf (!$StoppedServices)
					{
						Write-Host "	All Exchange-specific services on $($cas.name) are running..." -ForegroundColor Green
						$StoppedArray = "None"
					}
				$ServiceObj | Add-Member NoteProperty -Name "Stopped Services" -Value $StoppedArray
				$casservicereport += $ServiceObj
			

				###Begins the HealtSet Report###
				Write-Host "`n`n==========	Processing Health Set Status on $($cas.name)		==========`n" -ForegroundColor Cyan

				$casHealthObj = New-Object PSObject
				$casHealthObj | Add-Member NoteProperty -Name "Server" -Value $($cas.Name)
				$casHealthObj | Add-Member NoteProperty -Name "Site" -Value $site
				$UnhealthySets = (get-Serverhealth $cas.name | where {$_.alertvalue -like "Unhealthy"} | Select-Object Name)
				If ($UnhealthySets)
					{	
						$SetArray = [system.string]::Join(" - ", $Unhealthysets.name)
						ForEach ($setname in $unhealthysets)
							{
								Write-Host "	The $($setname.name) Health Set in in an unhealthy state!!!" -ForegroundColor Red
								If (!$cashealthseterror)
									{
									$cashealthseterror = "TRUE"
									If ($subject)
										{
										$Subject += ",CAS HealthSet"
										}
									Else
										{
										$subject = "Global Exchange Active Alerts: CAS HealthSet"
										}
									}
							}
					}
				ElseIf (!$UnhealthySets)
					{
						Write-Host "	All Health Sets on $($cas.name) are healthy..." -ForegroundColor Green
						$SetArray = "None"
					}
				$casHealthObj | Add-Member NoteProperty -Name "UnHealthy Sets" -Value $SetArray
				$cashealthsetreport += $casHealthObj	
			
				#.................	
				# Create the HTML
				#.................

				#..................................
				# Begin Component Report Table HTML
				#..................................

				$cascompsummaryHtml = $null

				###Begin Report table HTML header###
				$htmltableheader = "<p>
							<table>
							<tr>
							<th>Server</th>
							<th>Site</th>
							<th>InActive Components</th>
							</tr>"

				$cascompsummaryHtml += $htmltableheader
				###End Summary table HTML header###

				#Begin Component Summary table HTML rows
				ForEach ($line in $cascomp)
					{	
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line."Server")</strong></td>"
						$htmltablerow = $htmltablerow + "<td>$($line."Site")</td>"

						#Fail if state is inactive
						Switch ($($line."InActive Components"))
						{
							"None" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."InActive Components")</td>" }
							default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."InActive Components")</td>" }
						}
						$htmltablerow = $htmltablerow + "</tr>"
						$casCompsummaryHtml += $htmltablerow		
					}
				$cascompsummaryHtml += "</table>
				</p>"

				#...............................
				#Begin Service Report Table HTML
				#...............................

				$casservicesummaryHtml = $null

				###Begin Report table HTML header###
				$htmltableheader = "<p>
							<table>
							<tr>
							<th>Server</th>
							<th>Site</th>
							<th>Stopped Services</th>
							</tr>"

				$casservicesummaryHtml += $htmltableheader
				###End Summary table HTML header###

				###Begin Service Summary table HTML rows###
				ForEach ($line in $casservicereport)
					{	
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line."Server")</strong></td>"
						$htmltablerow = $htmltablerow + "<td>$($line."Site")</td>"

						#Fail if state is inactive
						Switch ($($line."Stopped Services"))
						{
							"None" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."Stopped Services")</td>" }
							default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."Stopped Services")</td>" }
						}
						$htmltablerow = $htmltablerow + "</tr>"
						$casservicesummaryHtml += $htmltablerow		
					}
				$casservicesummaryHtml += "</table>
				</p>"

				#...............................
				#Begin Service Report Table HTML
				#...............................

				$cashealthsummaryHtml = $null

				###Begin Report table HTML header###
				$htmltableheader = "<p>
							<table>
							<tr>
							<th>Server</th>
							<th>Site</th>
							<th>UnHealthy Sets</th>
							</tr>"

				$cashealthsummaryHtml += $htmltableheader
				###End Summary table HTML header###

				###Begin Component Summary table HTML rows###
				ForEach ($line in $cashealthsetreport) 
					{	
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line."Server")</strong></td>"
						$htmltablerow = $htmltablerow + "<td>$($line."Site")</td>"

						#Fail if state is inactive
						Switch ($($line."UnHealthy Sets"))
						{
							"None" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."UnHealthy Sets")</td>" }
							default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."UnHealthy Sets")</td>" }
						}
						$htmltablerow = $htmltablerow + "</tr>"
						$cashealthsummaryHtml += $htmltablerow		
					}
				$cashealthsummaryHtml += "</table>
				</p>"

			}

		#........................
		# Formats CAS Report HTML
		#.........................

		$casreporthtml = $cascompintro + $cascompsummaryHtml + $casserviceintro + $casservicesummaryHtml + $cashealthintro + $cashealthsummaryHtml 
		$casalertbody += $casreporthtml
		$reportmessagebody += $casalertbody
	}

#............................................
# Script Begins Processing Mailbox Servers
#............................................

$mbxservers = @()

Write-Host "`n`n==========	Retrieving Standalone Mailbox Servers for Analysis	==========`n" -ForegroundColor Cyan
$mbxservers = (get-mailboxserver | where {$_.databaseavailabilitygroup -eq $null})
Write-Host "	$($mbxservers.count) standalone mailbox servers found..." -ForegroundColor Green

If ($mbxservers.count -eq "0")
		{
			$NoMBX = "True"
		}
Else
	{

		#Sets info for use in HTML email
		$mbxsummaryintro = "<p>Standalone Mailbox Server Database Alert Summary:</p>"
		$mbxcompintro = "<p>Standalone Mailbox Server Component Alert Summary:</p>"
		$mbxserviceintro = "<p>Standalone Mailbox Server Service Alert Summary:</p>"
		$mbxhealthintro = "<p>Standalone Mailbox Server Health Set Alert Summary:</p>"	
				
		ForEach ($mbx in $mbxservers)
			{		
				#Opens the report hashes
				$mbxcomp = @()				# Component Report
				$mbxservicereport = @()		# Service Report
				$mbxhealthsetreport= @()	# Health Set Report
				$mbxdatabaseSummary = @()		# Database health summary report
				
				#Clears variables from previous MBX Server runs
				Remove-variable mbxdbmounterror -ErrorAction SilentlyContinue
				Remove-variable mbxdbindexerror -ErrorAction SilentlyContinue
				Remove-variable mbxcomperror -ErrorAction SilentlyContinue
				Remove-variable mbxserviceerror -ErrorAction SilentlyContinue
				Remove-variable mbxhealtherror -ErrorAction SilentlyContinue
				
				#Sets site and iscas variables for rest of mailbox server report
				$site = (get-exchangeserver $($mbx.name)).site
				$site = $site -replace ".*/"
				$CasCheck = (get-exchangeserver $($mbx.name)).IsClientAccessServer
					If ($CasCheck -eq "True")
						{
							$IsCas = "Yes"
						}
					Else
						{
							$IsCas = "No"
						}
					
				### Begin Database Summary Report ###
				Write-Host "`n`n==========	Retrieving $($mbx.name) Database Information	==========`n" -ForegroundColor Cyan
				$databases = (get-mailboxdatabase -server $mbx.name -Status | get-mailboxdatabasecopystatus)
				ForEach ($db in $databases)
					{
						$mbxdbobj = New-Object PSObject
						$mbxdbobj | Add-Member NoteProperty -Name "Database" -Value $db.databasename
						$mbxdbobj | Add-Member NoteProperty -Name "Server" -Value $($mbx.name)
						$mbxdbobj | Add-Member NoteProperty -Name "Site" -Value $site
						$MountState = $db.status
						If ($Mountstate -eq "Mounted")
							{
								$mbxdbobj | Add-Member NoteProperty -Name "Status" -Value $mountstate
							}
						Else
							{
								Write-Host "	$($db.databasename) is not mounted!!!" -ForegroundColor Red
								$mbxdbobj | Add-Member NoteProperty -Name "Status" -Value $mountstate
								If (!$mbxdbmounterror)
									{
										$mbxdbmounterror = "True"
										If ($subject)
											{
											$Subject += ",MBX Server Database Mount Status"
											}
										Else
											{
											$subject = "Global Exchange Active Alerts: MBX Server Database Mount Status"
											}
									}
							}
						$IndexState = $db.contentindexstate 
						If ($IndexState -eq "Healthy")
							{
								$mbxdbobj | Add-Member NoteProperty -Name "Index State" -Value $IndexState
							}
						Else
							{
								Write-Host "	$($db.databasename) has an unhealthy index!!!" -ForegroundColor Red
								$mbxdbobj | Add-Member NoteProperty -Name "Index State" -Value $IndexState
								If (!$mbxindexerror)
									{
										$mbxindexerror = "True"
										If ($subject)
											{
												$Subject += ",MBX Server Database Index"
											}
										Else
											{
												$subject = "Global Exchange Active Alerts: MBX Server Database Index"
											}
									}
					
							}
						$Mbxdatabasesummary += $mbxdbobj
					}
					$GlobalMBXdbsummaryreport += $Mbxdatabasesummary
					If (!$mbxdbmounterror -AND !$mbxdbindexerror)
						{
							Write-Host "	All databases on $($mbx.name) are mounted with healthy content indexes..." -ForegroundColor Green
						}
				
				###Begin Component Health Tests###
				Write-Host "`n==========	Processing Component States on $($mbx.name) 		==========`n" -ForegroundColor Cyan

				$mbxcompobj = New-Object PSObject
				$mbxcompobj | Add-Member NoteProperty -Name "Server" -Value $($mbx.Name)
				$mbxcompobj | Add-Member NoteProperty -Name "Site" -Value $site
				$mbxcompobj | Add-Member NoteProperty -Name "IsCAS" -Value $IsCas
				$compstate = ($mbx.name | Invoke-Command {get-servercomponentstate})
				$inactive = ($compstate | where {$compstate.state -like "InActive"}).component
				If ($inactive)
					{
						$InactiveArray = [system.string]::Join(" - ", $inactive)
						ForEach ($badcomp in $inactive) 
							{
								Write-Host "	The $($badcomp) component is in an INACTIVE state!!!" -ForegroundColor Red
								Write-Host "	Attempting to start the $($badcomp) component..." -ForegroundColor Yellow
								Set-ServerComponentState $mbx.name -Component $badcomp -Requester HealthAPI -State Active
								If (!$mbxcomperror)
									{
										$dagcomperror = "TRUE"
										If ($subject)
											{
												$Subject += ",MBX Server Components"
											}
										Else
											{
												$subject = "Global Exchange Active Alerts: MBX Server Components"
											}
									}
							}
					}	
				ElseIf (!$inactive) 
					{
						Write-Host "	All components on $($mbx.name) are in an Active state..." -ForegroundColor Green
						$InactiveArray = "None" 
					}
				$mbxCompObj | Add-Member NoteProperty -Name "InActive Components" -Value $InactiveArray
				$mbxcomp += $mbxCompObj
				$GlobalmbxComp += $mbxcomp 	

				###Begin Service Health Tests###
				Write-Host "`n==========	Processing Service States on $($mbx.name)			==========`n" -ForegroundColor Cyan

				$mbxServiceObj = New-Object PSObject
				$mbxServiceObj | Add-Member NoteProperty -Name "Server" -Value $($mbx.Name)
				$mbxServiceObj | Add-Member NoteProperty -Name "Site" -Value $site
				$mbxServiceObj | Add-Member NoteProperty -Name "IsCAS" -Value $IsCas
				$Stoppedservices = (get-Service -computername $mbx.name | where {($_.name -like "MSExch*" -OR $_.name -like "IIS*" -OR $_.name -like "W3SVC") -AND $_.status -like "Stopped"} | Select-Object Displayname)
					If ($StoppedServices)
						{
							$StoppedArray = [system.string]::Join(" - ", $StoppedServices.displayname)
							ForEach ($svc in $stoppedservices)
								{
									Write-Host "	The $($svc.displayname) service is not running!!!" -ForegroundColor Red
									If (!$mbxserviceerror)
										{
										$mbxserviceerror = "TRUE"
										If ($subject)
											{
											$Subject += ",MBX Server Services"
											}
										Else
											{
											$subject = "Global Exchange Active Alerts: MBX Server Services"
											}
										}
								}
							}
						ElseIf (!$StoppedServices)
							{
								Write-Host "	All Exchange-specific services on $($mbx.name) are running..." -ForegroundColor Green
								$StoppedArray = "None"
							}
				$mbxServiceObj | Add-Member NoteProperty -Name "Stopped Services" -Value $StoppedArray
				$mbxservicereport += $mbxServiceObj
				$GlobalmbxServiceReport += $mbxservicereport
					
				###Begins the HealtSet Report##
				Write-Host "`n==========	Processing Health Set Status on $($mbx.name)		==========`n" -ForegroundColor Cyan

				$mbxHealthObj = New-Object PSObject
				$mbxHealthObj | Add-Member NoteProperty -Name "Server" -Value $($mbx.Name)
				$mbxHealthObj | Add-Member NoteProperty -Name "Site" -Value $site
				$mbxHealthObj | Add-Member NoteProperty -Name "IsCAS" -Value $IsCas
				$UnhealthySets = (get-Serverhealth $mbx.name | where {$_.alertvalue -like "Unhealthy"} | Select-Object Name)
				If ($UnhealthySets)
					{	
						$SetArray = [system.string]::Join(" - ", $Unhealthysets.name)
						ForEach ($setname in $unhealthysets)
							{
								Write-Host "	The $($setname.name) Health Set in in an unhealthy state!!!" -ForegroundColor Red
								If (!$mbxhealtherror)
									{
										$mbxhealtherror = "TRUE"
										If ($subject)
											{
												$Subject += ",MBX Server HealthSet"
											}
										Else
											{
												$subject = "Global Exchange Active Alerts: MBX Server HealthSet"
											}
									}
							}
					}
				ElseIf (!$UnhealthySets)
					{
						Write-Host "	All Health Sets on $($mbx.name) are healthy..." -ForegroundColor Green
						$SetArray = "None"
					}
				$mbxHealthObj | Add-Member NoteProperty -Name "UnHealthy Sets" -Value $SetArray
				$mbxhealthsetreport += $mbxHealthObj
				$GlobalmbxHealthsetReport += $mbxhealthsetreport
				
				#....................
				#Create the HTML
				#....................
			
				#........................................
				#Begin Database Health Summary Table HTML
				#........................................

				$mbxdatabasesummaryHtml = $null
				###Begin Summary table HTML header###
				$htmltableheader = "<p>
								<table>
								<tr>
								<th>Database</th>
								<th>Server</th>
								<th>Site</th>
								<th>Status</th>
								<th>Index State</th>
								</tr>"

				$mbxdatabasesummaryHtml += $htmltableheader
				###End Summary table HTML header###
				
				###Begin Summary table HTML rows###
				ForEach ($line in $mbxdatabaseSummary)
					{
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line.Database)</strong></td>"
						
						$htmltablerow = $htmltablerow + "<td>$($line.Server)</td>"
						$htmltablerow = $htmltablerow + "<td>$($line.Site)</td>"
						
						#Fail if mounted status is still unmounted
						Switch ($($line."Status"))
						{
							"UnMounted" { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."Status")</td>" }
							default {$htmltablerow = $htmltablerow + "<td class=""pass"">$($line."Status")</td>"}
						}
					
						#Fail if index is unhealthy
						Switch ($($line."Index State"))
							{
								"Healthy" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."Index State")</td>" }
								default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."Index State")</td>" }
							}
					
						$htmltablerow = $htmltablerow + "</tr>"
						$mbxdatabasesummaryHtml += $htmltablerow
					}
				
				$mbxdatabasesummaryHtml += "</table>
										</p>"
				###End Database Summary table HTML rows###

				
				#.................................
				#Begin Component Report Table HTML
				#.................................

				$mbxcompsummaryHtml = $null
				###Begin Report table HTML header###
				$htmltableheader = "<p>
							<table>
							<tr>
							<th>Server</th>
							<th>IsCAS</th>
							<th>Site</th>
							<th>InActive Components</th>
							</tr>"

				$mbxcompsummaryHtml += $htmltableheader
				###End Summary table HTML header###
				
				###Begin Component Summary table HTML rows###
				ForEach ($line in $mbxcomp)
					{	
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line."Server")</strong></td>"
						$htmltablerow = $htmltablerow + "<td>$($line."IsCas")</td>"
						$htmltablerow = $htmltablerow + "<td>$($line."Site")</td>"

						#Fail if state is inactive
						Switch ($($line."InActive Components"))
						{
							"None" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."InActive Components")</td>" }
							default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."InActive Components")</td>" }
						}
						$htmltablerow = $htmltablerow + "</tr>"
						$mbxCompsummaryHtml += $htmltablerow		
					}
				$mbxcompsummaryHtml += "</table>
				</p>"
				
				#..................................
				#Begin Service Report Table HTML
				#..................................			

				$mbxservicesummaryHtml = $null
				###Begin Report table HTML header###
				$htmltableheader = "<p>
							<table>
							<tr>
							<th>Server</th>
							<th>IsCAS</th>
							<th>Site</th>
							<th>Stopped Services</th>
							</tr>"

				$mbxservicesummaryHtml += $htmltableheader
				###End Summary table HTML header###

				###Begin Service Summary table HTML rows###
				ForEach ($line in $mbxservicereport)
					{	
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line."Server")</strong></td>"
						$htmltablerow = $htmltablerow + "<td>$($line."IsCAS")</td>"
						$htmltablerow = $htmltablerow + "<td>$($line."Site")</td>"

						#Fail if state is inactive
						Switch ($($line."Stopped Services"))
							{
								"None" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."Stopped Services")</td>" }
								default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."Stopped Services")</td>" }
								}
						$htmltablerow = $htmltablerow + "</tr>"
						$mbxservicesummaryHtml += $htmltablerow		
					}
				$mbxservicesummaryHtml += "</table>
				</p>"

				#.............................
				#Begin Health Set Table HTML
				#.............................

				$mbxhealthsummaryHtml = $null
				###Begin Report table HTML header###
				$htmltableheader = "<p>
							<table>
							<tr>
							<th>Server</th>
							<th>IsCAS</th>
							<th>Site</th>
							<th>UnHealthy Sets</th>
							</tr>"

				$mbxhealthsummaryHtml += $htmltableheader
				###End Summary table HTML header###

				###Begin Health Set table HTML rows###
				ForEach ($line in $mbxhealthsetreport) 
					{	
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line."Server")</strong></td>"
						$htmltablerow = $htmltablerow + "<td>$($line."IsCAS")</td>"
						$htmltablerow = $htmltablerow + "<td>$($line."Site")</td>"

						#Fail if state is inactive
						Switch ($($line."UnHealthy Sets"))
							{
								"None" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."UnHealthy Sets")</td>" }
								default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."UnHealthy Sets")</td>" }
							}
						$htmltablerow = $htmltablerow + "</tr>"
						$mbxhealthsummaryHtml += $htmltablerow		
					}
				$mbxhealthsummaryHtml += "</table>
				</p>"
				
				#....................................
				#End All Table HTML
				#....................................
				

				
				#Creates report formating for all reporting
				$exchangereporthtml = $mbxsummaryintro + $mbxdatabasesummaryHtml + $mbxcompintro + $mbxcompsummaryHtml + $mbxserviceintro + $mbxservicesummaryHtml + $mbxhealthintro + $mbxhealthsummaryHtml
				$mbxalertbody += $exchangereporthtml
				$reportmessagebody += $mbxalertbody
			}

	
	}
	
#............................................
# Script Begins Processing DAG Members
#............................................

$dags = @()
[int]$replqueuethreshold = 5

Write-Host "`n`n==========	Retrieving DAGs in Environment for Analysis			==========`n" -ForegroundColor Cyan
$dags = @(Get-DatabaseAvailabilityGroup -Status)
Write-Host "	$($dags.count) DAGs found..." -ForegroundColor Green

If ($dags.count -eq "0")
	{
		$NoDAGs = "True"
	}
Else
	{
		ForEach ($dag in $dags)
			{
				#Sets info for use in HTML email
				$summaryintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Database Alert Summary:</p>"
				$dagcompintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Component Alert Summary:</p>"
				$dagserviceintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Service Alert Summary:</p>"
				$daghealthintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Health Set Alert Summary:</p>"	
				$memberintro = "<p>Database Availability Group <strong>$($dag.Name)</strong> Member Replication Health Summary:</p>"
				
				# Opens the report hashes
				$dagcomp = @()			# Component Report
				$dagservicereport = @()		# Service Report
				$daghealthsetreport= @()	# Replication Health Set Report
				$dbcopyReport = @()		# Database copy health report
				$databaseSummary = @()		# Database health summary report
				$memberReport = @()		# DAG member replication health report
				
				#Clears variables from previous DAG runs
				Remove-variable databasemounterror -ErrorAction SilentlyContinue
				Remove-variable dbhealtherror -ErrorAction SilentlyContinue
				Remove-variable dbindexerror -ErrorAction SilentlyContinue
				Remove-variable dbqueueerror -ErrorAction SilentlyContinue
				
				#Begins collecting DAG information for reports
				Write-Host "`n`n==========	Retrieving $($dag.name) DAG Information				==========`n" -ForegroundColor Cyan
				$dagmembers = @($dag | Select-Object -ExpandProperty Servers | Sort-Object Name)
				Write-Host "	$($dagmembers.count) DAG members found..." -ForegroundColor Green

				Write-Host "`n`n==========	Retrieving $($dag.name) DAG Database Information			==========`n	" -ForegroundColor Cyan	
				$dagdatabases = @(Get-MailboxDatabase -Status | Where-Object {$_.MasterServerOrAvailabilityGroup -eq $dag.Name} | Sort-Object Name)
				Write-Host "	$($dagdatabases.count) databases found in DAG..." -ForegroundColor Green
				Write-Host "`n`n==========	Processing DAG Databases					==========`n" -ForegroundColor Cyan
				
				ForEach ($database in $dagdatabases)
					{					
						###Custom object for Database###
						$objectHash = @{
							"Database" = $database.Identity
							"Mounted on" = "UnKnown"
							"Site" = $null
							"Mount State" = "UnMounted"
							"Preference" = $null
							"Total Copies" = $null
							"Unhealthy Copies" = $null
							"Unhealthy Queues" = $null
							"Unhealthy Indexes" = $null
								}
						$databaseObj = New-Object PSObject -Property $objectHash
					
						$dbcopystatus = @($database | Get-MailboxDatabaseCopyStatus)
						#Write-Host "$database has $($dbcopystatus.Count) copies..." -ForegroundColor Green
						ForEach ($dbcopy in $dbcopystatus)
							{
								###Custom object for DB copy###
								$objectHash = @{
									"Database Copy" = $dbcopy.Identity
									"Database Name" = $dbcopy.DatabaseName
									"Mailbox Server" = $null
									"Activation Preference" = $null
									"Status" = $null
									"Copy Queue" = $null
									"Content Index" = $null
										}
								$dbcopyObj = New-Object PSObject -Property $objectHash
					
								$mailboxserver = $dbcopy.MailboxServer
								$pref = ($database | Select-Object -ExpandProperty ActivationPreference | Where-Object {$_.Key -eq $mailboxserver}).Value
								$copystatus = $dbcopy.Status
								[int]$copyqueuelength = $dbcopy.CopyQueueLength
								$contentindexstate = $dbcopy.ContentIndexState
					
								$dbcopyObj | Add-Member NoteProperty -Name "Mailbox Server" -Value $mailboxserver -Force
								$dbcopyObj | Add-Member NoteProperty -Name "Activation Preference" -Value $pref -Force
								$dbcopyObj | Add-Member NoteProperty -Name "Status" -Value $copystatus -Force
								$dbcopyObj | Add-Member NoteProperty -Name "Copy Queue" -Value $copyqueuelength -Force
								$dbcopyObj | Add-Member NoteProperty -Name "Replay Queue" -Value $replayqueuelength -Force
								$dbcopyObj | Add-Member NoteProperty -Name "Replay Lagged" -Value $replaylag -Force
								$dbcopyObj | Add-Member NoteProperty -Name "Truncation Lagged" -Value $truncatelag -Force
								$dbcopyObj | Add-Member NoteProperty -Name "Content Index" -Value $contentindexstate -Force
					
								$dbcopyReport += $dbcopyObj
							}
					
						$copies = @($dbcopyReport | Where-Object { ($_."Database Name" -eq $database) })
						$mountedOn = ($copies | Where-Object { ($_.Status -eq "Mounted") })."Mailbox Server"
						If ($mountedOn)
							{
								$databaseObj | Add-Member NoteProperty -Name "Mounted on" -Value $mountedOn -Force
								$databaseObj | Add-Member NoteProperty -Name "Mount State" -Value "Mounted" -Force
							}
						Else
							{
								Write-Host "	$database is not mounted!!!" -ForegroundColor Red
								If (!$databasemounterror)
									{
										$databasemounterror = "True"
										If ($subject)
											{
											$Subject += ",DAG Database Mount Status"
											}
										Else
											{
											$subject = "Global Exchange Active Alerts: DAG Database Mount Status"
											}
									}
							}
						$site = (get-exchangeserver $mountedOn).site
						$site = $site -replace ".*/"
						$databaseObj | Add-Member NoteProperty -Name "Site" -Value $site -Force
						$activationPref = ($copies | Where-Object { ($_.Status -eq "Mounted") })."Activation Preference"
						$databaseObj | Add-Member NoteProperty -Name "Preference" -Value $activationPref -Force
					
						$totalcopies = $copies.count
						$databaseObj | Add-Member NoteProperty -Name "Total Copies" -Value $totalcopies -Force
					
						$unhealthycopies = @($copies | Where-Object { (($_.Status -ne "Mounted") -and ($_.Status -ne "Healthy")) }).Count
						If ($unhealthycopies -ne "0")
							{
								Write-host "	$Database has $($unhealthycopies) UnHealthy copies!!!" -ForegroundColor Orange
								If (!$dbhealtherror)
									{
										$dbhealtherror = "True"
										If ($subject)
											{
											$Subject += ",DAG Database Copy Health"
											}
										Else
											{
											$subject = "Global Exchange Active Alerts: DAG Database Copy Health"
											}
									}
							}
						$databaseObj | Add-Member NoteProperty -Name "Unhealthy Copies" -Value $unhealthycopies -Force
					
						$unhealthyqueues = @($copies | Where-Object { ($_."Copy Queue" -ge $replqueuethreshold) }).Count
						If ($unhealthyqueues -ne "0")
							{
								Write-Host "	$database has UnHealthy queues!!!" -ForegroundColor Orange
								If (!$dbqueueerror)
									{
										$dbqueueerror = "True"
										If ($subject)
											{
											$Subject += ",DAG Database Queue Health"
											}
										Else
											{
											$subject = "Global Exchange Active Alerts: DAG Database Queue Health"
											}
									}
							}
						$databaseObj | Add-Member NoteProperty -Name "Unhealthy Queues" -Value $unhealthyqueues -Force
					
						$unhealthyindexes = @($copies | Where-Object { ($_."Content Index" -ne "Healthy") }).Count
						If ($unhealthyindexes -ne "0")
							{
								Write-Host "	$database has Unhealthy Indexes!!!" -ForegroundColor Orange
								If (!$dbindexerror)
									{
										$dbindexerror = "True"
										If ($subject)
											{
											$Subject += ",DAG Database Index Health"
											}
										Else
											{
											$subject = "Global Exchange Active Alerts: DAG Database Index Health"
											}
									}
							}
						$databaseObj | Add-Member NoteProperty -Name "Unhealthy Indexes" -Value $unhealthyindexes -Force
						$databaseSummary += $databaseObj
					}
					
				If (!$databasemounterror -AND !$dbhealtherror -AND !$dbindexerror -AND !$dbqueueerror)
							{
								Write-Host "	All $($dagdatabases.count) mailbox databases in the DAG have passed all health checks..." -ForegroundColor Green
							}
				$GlobalDatabase += $databasesummary
				
				ForEach ($dagmember in $dagmembers)
					{
						$site = (get-exchangeserver $($dagmember.name)).site
						$site = $site -replace ".*/"
						$CasCheck = (get-exchangeserver $($dagmember.name)).IsClientAccessServer
						If ($CasCheck -eq "True")
							{
							$IsCas = "Yes"
							}
						Else
							{
							$IsCas = "No"
							}
						
						###Begin Replication Health Tests###
						$memberObj = New-Object PSObject
						$memberObj | Add-Member NoteProperty -Name "Server" -Value $($dagmember.Name)
						$memberObj | Add-Member NoteProperty -Name "DAG" -Value $($dag.Name)
						$memberObj | Add-Member NoteProperty -Name "Site" -Value $site
				
						Write-Host "`n==========	Checking replication health for $($dagmember.Name)		==========`n" -ForegroundColor Cyan
						$replicationhealth = $dagmember | Invoke-Command {Test-ReplicationHealth}
						$badrepl = ($replicationhealth | where {$replicationhealth.result -notlike "Passed"}).check
						If ($badrepl)
							{
								$badreplarray = [system.string]::Join(" - ", $badrepl)
								ForEach ($baditem in $badrepl)
									{
										Write-Host "	The $($baditem) test did not pass the test!!!" -ForegroundColor Red
										If (!$dagreplerror)
											{
												$dagreplerror = "True"
											If ($subject)
												{
												$Subject += ",DAG Member Replication"
												}
											Else
												{
												$subject = "Global Exchange Active Alerts: DAG Member Replication"
												}
											}
									}
							}
						ElseIf (!$badrepl)
							{
								Write-Host "	All replication tests passed..." -ForegroundColor Green
								$badreplarray = "None"
							}
						$memberobj | Add-Member NoteProperty -Name "Failed Tests" -Value $badreplarray
						$memberReport += $memberObj
					
						###Begin Component Health Tests###
						Write-Host "`n==========	Processing Component States on $($dagmember.name) 		==========`n" -ForegroundColor Cyan

						$DAGCompObj = New-Object PSObject
						$DAGCompObj | Add-Member NoteProperty -Name "Server" -Value $($dagmember.Name)
						$DAGCompObj | Add-Member NoteProperty -Name "DAG" -Value $($dag.Name)
						$DAGCompObj | Add-Member NoteProperty -Name "Site" -Value $site
						$DAGCompObj | Add-Member NoteProperty -Name "IsCAS" -Value $IsCas
						$compstate = ($dagmember | Invoke-Command {get-servercomponentstate})
						$inactive = ($compstate | where {$compstate.state -like "InActive"}).component
						If ($inactive)
							{
								$InactiveArray = [system.string]::Join(" - ", $inactive)
								ForEach ($badcomp in $inactive) 
									{
										Write-Host "	The $($badcomp) component is in an INACTIVE state!!!" -ForegroundColor Red
										Write-Host "	Attempting to start the $($badcomp) component..." -ForegroundColor Yellow
										Set-ServerComponentState $dagmember.name -Component $badcomp -Requester HealthAPI -State Active
										If (!$dagcomperror)
											{
											$dagcomperror = "TRUE"
											If ($subject)
												{
												$Subject += ",DAG Member Components"
												}
											Else
												{
												$subject = "Global Exchange Active Alerts: DAG Member Components"
												}
											}
									}
							}	
						ElseIf (!$inactive) 
							{
								Write-Host "	All components on $($dagmember.name) are in an Active state..." -ForegroundColor Green
								$InactiveArray = "None" 
							}
						$DAGCompObj | Add-Member NoteProperty -Name "InActive Components" -Value $InactiveArray
						$DAGcomp += $DAGCompObj
					
						###Begin Service Health Tests###
						Write-Host "`n==========	Processing Service States on $($dagmember.name)			==========`n" -ForegroundColor Cyan

						$ServiceObj = New-Object PSObject
						$ServiceObj | Add-Member NoteProperty -Name "Server" -Value $($dagmember.Name)
						$ServiceObj | Add-Member NoteProperty -Name "DAG" -Value $($dag.Name)
						$ServiceObj | Add-Member NoteProperty -Name "Site" -Value $site
						$ServiceObj | Add-Member NoteProperty -Name "IsCAS" -Value $IsCas
						$Stoppedservices = (get-Service -computername $dagmember.name | where {($_.name -like "MSExch*" -OR $_.name -like "IIS*" -OR $_.name -like "W3SVC") -AND $_.status -like "Stopped"} | Select-Object Displayname)
						If ($StoppedServices)
							{
								$StoppedArray = [system.string]::Join(" - ", $StoppedServices.displayname)
								ForEach ($svc in $stoppedservices)
								{
									Write-Host "	The $($svc.displayname) service is not running!!!" -ForegroundColor Red
									If (!$dagserviceerror)
										{
										$dagserviceerror = "TRUE"
										If ($subject)
											{
											$Subject += ",DAG Member Services"
											}
										Else
											{
											$subject = "Global Exchange Active Alerts: DAG Member Services"
											}
										}
								}
							}
						ElseIf (!$StoppedServices)
							{
								Write-Host "	All Exchange-specific services on $($dagmember.name) are running..." -ForegroundColor Green
								$StoppedArray = "None"
							}
						$ServiceObj | Add-Member NoteProperty -Name "Stopped Services" -Value $StoppedArray
						$DAGservicereport += $ServiceObj
					
						###Begins the HealtSet Report##
						Write-Host "`n==========	Processing Health Set Status on $($dagmember.name)		==========`n" -ForegroundColor Cyan

						$DAGHealthObj = New-Object PSObject
						$DAGHealthObj | Add-Member NoteProperty -Name "Server" -Value $($dagmember.Name)
						$DAGHealthObj | Add-Member NoteProperty -Name "DAG" -Value $($dag.name)
						$DAGHealthObj | Add-Member NoteProperty -Name "Site" -Value $site
						$DAGHealthObj | Add-Member NoteProperty -Name "IsCAS" -Value $IsCas
						$UnhealthySets = (get-Serverhealth $dagmember.name | where {$_.alertvalue -like "Unhealthy"} | Select-Object Name)
						If ($UnhealthySets)
							{	
								$SetArray = [system.string]::Join(" - ", $Unhealthysets.name)
								ForEach ($setname in $unhealthysets)
									{
										Write-Host "	The $($setname.name) Health Set in in an unhealthy state!!!" -ForegroundColor Red
										If (!$daghealtherror)
										{
										$daghealtherror = "TRUE"
										If ($subject)
											{
											$Subject += ",DAG Member HealthSet"
											}
										Else
											{
											$subject = "Global Exchange Active Alerts: DAG Member HealthSet"
											}
										}
									}
							}
						ElseIf (!$UnhealthySets)
							{
								Write-Host "	All Health Sets on $($dagmember.name) are healthy..." -ForegroundColor Green
								$SetArray = "None"
							}
						$DAGHealthObj | Add-Member NoteProperty -Name "UnHealthy Sets" -Value $SetArray
						$DAGhealthsetreport += $DAGHealthObj
					}
				$GlobalRep += $memberreport
				$GlobalDAGComp += $Dagcomp
				$GlobalDAGServiceReport += $DAGservicereport
				$GlobalDAGHealthsetReport += $DAGhealthsetreport
				
					
				#....................
				#Create the HTML
				#....................
			
				#........................................
				#Begin Database Health Summary Table HTML
				#........................................

				$databasesummaryHtml = $null
				###Begin Summary table HTML header###
				$htmltableheader = "<p>
								<table>
								<tr>
								<th>Database</th>
								<th>Mounted on</th>
								<th>Site</th>
								<th>Mount State</th>
								<th>Preference</th>
								<th>Total Copies</th>
								<th>Unhealthy Copies</th>
								<th>Unhealthy Queues</th>
								<th>Unhealthy Indexes</th>
								</tr>"

				$databasesummaryHtml += $htmltableheader
				###End Summary table HTML header###
				
				###Begin Summary table HTML rows###
				ForEach ($line in $databaseSummary)
					{
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line.Database)</strong></td>"
					
						#Warn if mounted server is still unknown
						Switch ($($line."Mounted on"))
						{
							"Unknown" { $htmltablerow = $htmltablerow + "<td class=""warn"">$($line."Mounted on")</td>" }
							default { $htmltablerow = $htmltablerow + "<td>$($line."Mounted on")</td>" }
						}
						
						$htmltablerow = $htmltablerow + "<td>$($line.Site)</td>"
						
						#Fail if mounted status is still unmounted
						Switch ($($line."Mount State"))
						{
							"UnMounted" { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."Mount State")</td>" }
							default {$htmltablerow = $htmltablerow + "<td class=""pass"">$($line."Mount State")</td>"}
						}
					
						#Warn if DB is mounted on a server that is not Activation Preference 1
						If ($($line.Preference) -gt 1)
							{
							$htmltablerow = $htmltablerow + "<td class=""warn"">$($line.Preference)</td>"		
							}
						Else
							{
							$htmltablerow = $htmltablerow + "<td class=""pass"">$($line.Preference)</td>"
							}
					
						$htmltablerow = $htmltablerow + "<td>$($line."Total Copies")</td>"
					
						#Warn if unhealthy copies is 1, fail if more than 1
						Switch ($($line."Unhealthy Copies"))
							{
								0 {	$htmltablerow = $htmltablerow + "<td class=""pass"">$($line."Unhealthy Copies")</td>" }
								1 {	$htmltablerow = $htmltablerow + "<td class=""warn"">$($line."Unhealthy Copies")</td>" }
								default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."Unhealthy Copies")</td>" }
							}
					
						#Warn if unhealthy queues is 1, fail if more than 1 
						Switch ($($line."Unhealthy Queues"))
							{
								0 { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."Unhealthy Queues")</td>" }
								1 {	$htmltablerow = $htmltablerow + "<td class=""warn"">$($line."Unhealthy Queues")</td>" }
								default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."Unhealthy Queues")</td>" }
							}
					
						#Warn if unhealthy indexes is 1, fail if more than 1
						Switch ($($line."Unhealthy Indexes"))
							{
								0 { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."Unhealthy Indexes")</td>" }
								1 {	$htmltablerow = $htmltablerow + "<td class=""warn"">$($line."Unhealthy Indexes")</td>" }
								default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."Unhealthy Indexes")</td>" }
							}
					
						$htmltablerow = $htmltablerow + "</tr>"
						$databasesummaryHtml += $htmltablerow
					}
				
				$databasesummaryHtml += "</table>
										</p>"
				###End Database Summary table HTML rows###

				#....................................		
				#Begin Replication Health Table HTML
				#....................................

				$memberHtml = $null
				###Begin Member table HTML header##
				$htmltableheader = "<p>
									<table>
									<tr>
									<th>Server</th>
									<th>Site</th>
									<th>Failed Replication Tests</th>
									</tr>"
				
				$memberHtml += $htmltableheader
				###End Member table HTML header###
				
				####Begin Replication Health Table HTML rows###
				ForEach ($line in $memberReport)
					{
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line."Server")</strong></td>"
						$htmltablerow = $htmltablerow + "<td>$($line."Site")</td>"

						# Fail if tests don't pass
						Switch ($($line."Failed Tests"))
							{
								"None" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."Failed Tests")</td>" }
								default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."Failed Tests")</td>" }
							}
					

						$htmltablerow = $htmltablerow + "</tr>"
						$memberHtml += $htmltablerow
					}
				$memberHtml += "</table>
				</p>"
				
				#.................................
				#Begin Component Report Table HTML
				#.................................

				$DagcompsummaryHtml = $null
				###Begin Report table HTML header###
				$htmltableheader = "<p>
							<table>
							<tr>
							<th>Server</th>
							<th>IsCAS</th>
							<th>Site</th>
							<th>InActive Components</th>
							</tr>"

				$dagcompsummaryHtml += $htmltableheader
				###End Summary table HTML header###
				
				###Begin Component Summary table HTML rows###
				ForEach ($line in $dagcomp)
					{	
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line."Server")</strong></td>"
						$htmltablerow = $htmltablerow + "<td>$($line."IsCas")</td>"
						$htmltablerow = $htmltablerow + "<td>$($line."Site")</td>"

						#Fail if state is inactive
						Switch ($($line."InActive Components"))
						{
							"None" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."InActive Components")</td>" }
							default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."InActive Components")</td>" }
						}
						$htmltablerow = $htmltablerow + "</tr>"
						$dagCompsummaryHtml += $htmltablerow		
					}
				$dagcompsummaryHtml += "</table>
				</p>"
				
				#..................................
				#Begin Service Report Table HTML
				#..................................			

				$dagservicesummaryHtml = $null
				###Begin Report table HTML header###
				$htmltableheader = "<p>
							<table>
							<tr>
							<th>Server</th>
							<th>IsCAS</th>
							<th>Site</th>
							<th>Stopped Services</th>
							</tr>"

				$dagservicesummaryHtml += $htmltableheader
				###End Summary table HTML header###

				###Begin Service Summary table HTML rows###
				ForEach ($line in $dagservicereport)
					{	
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line."Server")</strong></td>"
						$htmltablerow = $htmltablerow + "<td>$($line."IsCAS")</td>"
						$htmltablerow = $htmltablerow + "<td>$($line."Site")</td>"

						#Fail if state is inactive
						Switch ($($line."Stopped Services"))
							{
								"None" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."Stopped Services")</td>" }
								default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."Stopped Services")</td>" }
								}
						$htmltablerow = $htmltablerow + "</tr>"
						$dagservicesummaryHtml += $htmltablerow		
					}
				$dagservicesummaryHtml += "</table>
				</p>"

				#.............................
				#Begin Health Set Table HTML
				#.............................

				$daghealthsummaryHtml = $null
				###Begin Report table HTML header###
				$htmltableheader = "<p>
							<table>
							<tr>
							<th>Server</th>
							<th>IsCAS</th>
							<th>Site</th>
							<th>UnHealthy Sets</th>
							</tr>"

				$daghealthsummaryHtml += $htmltableheader
				###End Summary table HTML header###

				###Begin Health Set table HTML rows###
				ForEach ($line in $daghealthsetreport) 
					{	
						$htmltablerow = "<tr>"
						$htmltablerow = $htmltablerow + "<td><strong>$($line."Server")</strong></td>"
						$htmltablerow = $htmltablerow + "<td>$($line."IsCAS")</td>"
						$htmltablerow = $htmltablerow + "<td>$($line."Site")</td>"

						#Fail if state is inactive
						Switch ($($line."UnHealthy Sets"))
							{
								"None" { $htmltablerow = $htmltablerow + "<td class=""pass"">$($line."UnHealthy Sets")</td>" }
								default { $htmltablerow = $htmltablerow + "<td class=""fail"">$($line."UnHealthy Sets")</td>" }
							}
						$htmltablerow = $htmltablerow + "</tr>"
						$daghealthsummaryHtml += $htmltablerow		
					}
				$daghealthsummaryHtml += "</table>
				</p>"
				
				#....................................
				#End All Table HTML
				#....................................
			
			#Creates report formating for all reporting
			$exchangereporthtml = $summaryintro + $databasesummaryHtml + $dagcompintro + $dagcompsummaryHtml + $dagserviceintro + $dagservicesummaryHtml + $daghealthintro + $daghealthsummaryHtml + $memberintro + $memberHtml
			$databasealertbody += $exchangereporthtml
			}
			
		#Sets the final DAG report for all DAGS
		$reportmessagebody += $databasealertbody
	}

If (!$EmailOnly)
	{
		#............................................
		# Output the reports to the console
		#............................................
		If (!$nocas)
			{
				Write-Host "`n ---- Global Standalone CAS Component Summary ---- " -ForegroundColor Yellow
				$cascomp				| ft

				Write-Host "`n ---- Global Standalone CAS Service Summary ---- " -ForegroundColor Yellow
				$casservicereport | ft

				Write-Host "`n ---- Global Standalone CAS Health Set Summary ---- " -ForegroundColor Yellow
				$cashealthsetreport | ft
			}
		If (!$NoMBX)
			{
				Write-Host "`n ---- Global MBX Server Database Summary---- " -ForegroundColor Yellow
				$globalmbxdbsummaryreport	| ft -Autosize
				
				Write-Host "`n ---- Global MBX Server Component Summary---- " -ForegroundColor Yellow
				$globalmbxcomp | ft
				
				Write-Host "`n ---- Gobal MBX Server Service Summary---- " -ForegroundColor Yellow
				$globalmbxservicereport | ft
				
				Write-Host "`n ---- Global MBX Server Health Summary---- " -ForegroundColor Yellow
				$globalmbxhealthsetreport | ft
			}
		If (!$NoDAGs)
			{
				Write-Host "`n ---- Global DAG Database Summary---- " -ForegroundColor Yellow
				$globaldatabase	| ft -Autosize
				
				Write-Host "`n ---- Global DAG Component Summary---- " -ForegroundColor Yellow
				$globaldagcomp | ft
				
				Write-Host "`n ---- Gobal DAG Service Summary---- " -ForegroundColor Yellow
				$globaldagservicereport | ft
				
				Write-Host "`n ---- Global DAG Health Summary---- " -ForegroundColor Yellow
				$globaldaghealthsetreport | ft
				
				Write-Host "`n ---- Global DAG Replication Health Summary---- " -ForegroundColor Yellow
				$globalrep | ft
			}
	}	

If (!$ConsoleOnly -OR $File)
	{
		#............................................
		# Format alert message the report to HTML 
		#............................................

		$htmlreporthead="<html>
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
			td.info{background: #85D4FF;}
			</style>
			<body>
			<h3 align=""center"">Exchange Environment Alert Report</h3>
			<p>Exchange Server Alerting Components as of $currentdate</p>"
		
		$htmlreporttail = "</body></html>"	

		$htmlreport = $htmlreporthead + $reportmessagebody + $htmlreporttail
		
		If (!$ConsoleOnly)
			{
				If (!$subject)
					{
					$Subject = "Global Exchange Report - No Alerts Active on "+$currentdate
					Send-MailMessage -To $recipient -Subject $subject -SmtpServer $server -From $sender -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
					}
				Else
					{
					$Subject += " on "+$currentdate
					Send-MailMessage -To $recipient -Subject $subject -SmtpServer $server -From $sender -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8)
					}
			}
		If ($File)
			{
				$htmlreport | out-file $file
			}
	}
