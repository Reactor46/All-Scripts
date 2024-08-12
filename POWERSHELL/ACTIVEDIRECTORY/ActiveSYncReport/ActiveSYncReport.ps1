<#	
	.NOTES
	===========================================================================
	 Created on:   	8/8/2017 10:05 AM
	 Created by:   	Vikas SUkhija (http://SysCloudPro.com)
	 Organization: 	
	 Filename:     	ActiveSYncReport.ps1
	===========================================================================
	.DESCRIPTION
		This Script will fectch Active Sync Devices Report from Exchange Online
#>
###################Load Functions/Modules####################
function Write-Log
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[array]$Name,
		[Parameter(Mandatory = $true)]
		[string]$Ext,
		[Parameter(Mandatory = $true)]
		[string]$folder
	)
	
	$log = @()
	$date1 = get-date -format d
	$date1 = $date1.ToString().Replace("/", "-")
	$time = get-date -format t
	
	$time = $time.ToString().Replace(":", "-")
	$time = $time.ToString().Replace(" ", "")
	
	foreach ($n in $name)
	{
		
		$log += (Get-Location).Path + "\" + $folder + "\" + $n + "_" + $date1 + "_" + $time + "_.$Ext"
	}
	return $log
}
Function LaunchEOL
{
	
	$UserCredential = Get-Credential
	
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	
	Import-PSSession $Session -Prefix "EOL" -AllowClobber
	
}

Function RemoveEOL
{
	
	$Session = Get-PSSession | where { $_.ComputerName -like "outlook.office365.com" }
	Remove-PSSession $Session
	
}
##########################Variables/Log/report######################
$log = Write-Log -Name Transcript_Async -folder logs -Ext log
$report = Write-Log -Name Report_ActiveSYnc_Online -folder Report -Ext csv
$reportemail1 = "sukhija1@labtest.com"
$reportemail2 = "sukhija2@labtest.com"
$reportemail4 = "sukhija3@labtest.com"
$smtpserver = "smtp.labtest.com"
$from = "ActiveSyncOnlineReport@labtest.com"
$Collection = @()
#Start-Transcript -Path $log
#####Main script########
try
{
	LaunchEOL
}
catch
{
	$_.Exception
	Write-Host "Exception has occured" -ForegroundColor Yellow
	exit
}
Write-Host "Fetching all ACtive SYnc users" -ForegroundColor Magenta
$casm = Get-EOLCASMailbox -Filter { ActiveSyncEnabled -eq 'true' } -resultsize:unlimited
Write-Host "Fetched all ACtive SYnc users" -ForegroundColor Green
$casm | foreach-object{
	$asyncpolicy = $_.ActiveSyncMailboxPolicy
	$user = $_.name
	Write-Host "Processing user................. $user" -ForegroundColor Green
	$devices = Get-EOLMobileDeviceStatistics -Mailbox $_.Identity | select DeviceType, DeviceID, DeviceUserAgent, FirstSyncTime, LastSuccessSync,
																		   Identity, DeviceModel, DeviceFriendlyName, DeviceOS, DeviceAccessState, DeviceAccessStateReason
	$devices | foreach-object {
		
		$coll = "" | select Userid, DeviceType, DeviceID, DeviceUserAgent, FirstSyncTime, LastSuccessSync,
							Identity, DeviceModel, DeviceFriendlyName, DeviceOS, DevicePolicy, DeviceAccessState, DeviceAccessStateReason
		
		$coll.Userid = $user
		$coll.DeviceType = $_.DeviceType
		$coll.DeviceID = $_.DeviceID
		$coll.DeviceUserAgent = $_.DeviceUserAgent
		$coll.FirstSyncTime = $_.FirstSyncTime
		$coll.LastSuccessSync = $_.LastSuccessSync
		$coll.Identity = $_.Identity
		$coll.DeviceModel = $_.DeviceModel
		$coll.DeviceFriendlyName = $_.DeviceFriendlyName
		$coll.DeviceOS = $_.DeviceOS
		$coll.DevicePolicy = $asyncpolicy
		$coll.DeviceAccessState = $_.DeviceAccessState
		$coll.DeviceAccessStateReason = $_.DeviceAccessStateReason
		$Collection += $coll
	}
}
#export the collection to csv , change the path accordingly

$Collection | export-csv $report -notypeinformation
RemoveEOL
Timeout 10
Send-MailMessage -SmtpServer $smtpserver -From $from -To $reportemail1, $reportemail2, $reportemail4 -Subject "ActiveSync Online Report" -Attachments $report
#Stop-Transcript
#############################################################################


