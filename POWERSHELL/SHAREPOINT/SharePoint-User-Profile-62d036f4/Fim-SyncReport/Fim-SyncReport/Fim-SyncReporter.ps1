<#
.SYNOPSIS
   Fim-SyncReporter v1 - The script can create report for your SharePoint 2013 User Profile Synchronization operations in FIM and e-mail it.
   Author: Ivan Yankulov, BulPros Consulting
   Downloaded From: http://spyankulov.blogspot.com/
.DESCRIPTION
	The script can create report for your SharePoint 2013 User Profile Synchronization operations in FIM and e-mail it. There should be defined how many hours the report should include. For example if you want a report for the last 4 hours you will supply value 4 for the parameter Hours. If you want a report for the last 3 days the value will be 72. 
	The script can be executed locally on the server where the FIM Synchronization service is running or you can run it from remote domain machine. But you should consider running it under identity with needed permissions. Supply of different credentials is not available at this moment. 
	You have options to use only ServerName and Hours parameter, in this case a only the CSV report will be created in the script execution directory or you can choose different location.
.PARAMETER ServerName
 The name of the server where SharePoint User Profile Synchronization service instance is running
.PARAMETER Hours
 The number of hours you want to include from the moment of launching of the script
.PARAMETER ReportLoction
 The localtion where the CSV report will be saved, the default is the folder where the script is running from
.PARAMETER MailAdress
 The e-Mail adresses of that should receive the mail report, multiple comma separated values can e provided
.PARAMETER FromMailAdress
 The eMail address which the report will be sended from, it is not mandatory to be real address with mailbox,...etc, Just make sure it is not blocked by your messaging infrastructure.
.PARAMETER SMTPServer
 SMTP service that can accept requests and send mails to your recipients.
.EXAMPLE
	
	.\Fim-SyncReport.ps1 -ServerName spfimsync.contoso.com -Hours 48
	
	This will create a report for the last two days and will store it in the script execution folder under the name SyncReport.csv
	
.EXAMPLE
	
	.\Fim-SyncReport.ps1 -ServerName spfimsync.contoso.net -Hours 12 -FromMailAdress spupsreporter@contoso.com -MailAdress aaronp@contoso.com,sp-suppor-team@contoso.com -SMTPServer mailer.contoso.com -ReportLoction \\fileserver\fimnightreport
	
	This will create the report with target FIM server  spfimsync.contoso.net for the last 12 h., will mail it to two e-mail adresses, will save the CSV report on custom location
.LINK
  http://spyankulov.blogspot.com
  http://spyankulov.blogspot.com/2014/06/user-profile-synchronization-reporting.html
#>
[CmdletBinding()]
param(

	[parameter(Mandatory=$true,HelpMessage="Enter the server name where FIM Syncronisation service is running",Position=0)]
	[string]$ServerName,

	[parameter(Mandatory=$true,HelpMessage="Enter the number of hours that you want to include in the report",Position=1)]
	[string]$Hours,

	[parameter(Mandatory=$false,Position=2)]
	[string]$ReportLoction = $PSScriptRoot,

	[parameter(Mandatory=$false)]
	[string[]]$MailAdress,

	[parameter(Mandatory=$false)]
	[string]$FromMailAdress,

	[parameter(Mandatory=$false)]
	[string]$SMTPServer
)
function Get-SPSyncMA
{
[CmdletBinding()]
Param(
 [parameter(Mandatory=$true)][string]$ComputerName
)
Process
{
	Get-WmiObject -ComputerName $ComputerName -Class MIIS_ManagementAgent -Namespace root/MicrosoftIdentityIntegrationServer | Where {$_.Name -like "MOSS*"}
}
}
function Get-FimMARuns
{
 [CmdletBinding()]
Param(
	[parameter(Mandatory=$true)]
	[string]$MaName,
	[parameter(Mandatory=$true)]
	[string]$Hours,
	[parameter(Mandatory=$true)]
	[string]$ComputerName
)
Process
{
	$timeSpan = New-TimeSpan -Hours $Hours
	$nowUTC = (Get-Date).ToUniversalTime()
	$timeToStart = $nowUTC.Add(-$timeSpan)
	$filter = ("MaName = '{0}'" -F $MaName)
	$allHistory = Get-WmiObject -ComputerName $ComputerName -Class MIIS_RunHistory -Namespace root/MicrosoftIdentityIntegrationServer -Filter $filter
	ForEach ($history in $allHistory)
	{
		#Converting the start of the sync operation in order to be easier for comparing with the report interval
		$startTimeinDateTime = $history.RunStartTime | Get-Date
		if ($startTimeinDateTime -gt $timeToStart)
		{
			Write-Output $history
		}
	}
}
}
function Send-MailReport
{
[CmdletBinding()]
Param(
[parameter(Mandatory=$true)]
[string[]]$Recipients,
[parameter(Mandatory=$true)]
[string]$From,
[parameter(Mandatory=$true)]
[string]$SMTP
)
Process
{
 [string]$MailBody = Get-Content -Path '.\mail\header.txt'
$MailBody += Get-Content '.\mail\middle.txt'
$MailBody += Get-Content -Path '.\mail\footer.txt'
Send-MailMessage -From $From -Subject "SharePoint User Profile Synchronization issue report" -BodyAsHtml $MailBody -To $Recipients -SmtpServer $SMTP
}
}
### End Function region ###

$faultyOperations =@()
$Report = "$ReportLoction\SyncReport.csv"
$syncAgents = Get-SPSyncMA -ComputerName $ServerName
ForEach ($syncAgent in $syncAgents)
{
	$faultyOperations += Get-FimMARuns -MaName $syncAgent.Name -ComputerName $ServerName -Hours $Hours | Where {$_.RunStatus -ne 'success'}
}
If($faultyOperations)
{
	Clear-Content -Path '.\mail\middle.txt'
	If(Test-Path -Path $Report)
	{
		Clear-Content -Path $Report
	}
	Add-Content -Path $Report -Value "Connection Name,Status,SyncProfile,Started,SyncErrors,DiscoveryErrors,RetryErrors"
	ForEach ($faultyOp in $faultyOperations)
	{
		[xml]$asXML = $faultyOp.RunDetails().ReturnValue
		$connName = $asXML.'run-history'.'run-details'.'ma-name'
		$profile = $asXML.'run-history'.'run-details'.'run-profile-name'
		$start = $asXML.'run-history'.'run-details'.'step-details'.'start-date'
		$status = $faultyOp.RunStatus
		$syncErrors = ($asXML.'run-history'.'run-details'.'step-details'.'synchronization-errors'.GetEnumerator() | Measure-Object).Count
		$disErrors = ($asXML.'run-history'.'run-details'.'step-details'.'ma-discovery-counters'.GetEnumerator() | Measure-Object).Count
		$retErrors = ($asXML.'run-history'.'run-details'.'step-details'.'mv-retry-errors'.GetEnumerator() | Measure-Object).Count

		Add-Content -Path $Report -Value "$connName,$status,$profile,$start,$syncErrors,$disErrors,$retErrors"
		Add-Content -Value "<tr><td>$connName</td><td>$profile</td><td>$status</td><td>$start</td><td>$syncErrors</td></tr>" -Path '.\mail\middle.txt' -Force
	}
	If(($MailAdress) -and ($SMTPServer) -and ($FromMailAdress))
	{
		Send-MailReport -Recipients $MailAdress -SMTP $SMTPServer -From $FromMailAdress
	}
}