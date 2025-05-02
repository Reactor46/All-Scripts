<#	
	.NOTES
	===========================================================================
	 Created on:   	3/29/2017 1:55 PM
	 Created by:   	sukhijv
	 Organization: 	
	 Filename:     	OOO.ps1
	===========================================================================
	.DESCRIPTION
		Setup ooo for Users from CSV File by reading old address & New address
#>
########################Functions ###################################
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
$log = Write-Log -Name "log1", "log2" -Ext log -folder Logs
Start-Transcript -Path $log[0]
##########################Import CSv & ADD OOO########################
$start = "1/3/2017"
$end = "1/3/2018"
$data = Import-Csv $args
foreach ($i in $data)
{
	$pemail = $i.previousEmail
	$pemail = $pemail.trim()
	$femail = $i.futureemail
	$femail = $femail.trim()
	$ooomessage = @"
Please contact me at my new email address, $femail. Also note my $pemail account is no longer monitored.
"@
	Write-Host $ooomessage -ForegroundColor Green
	if (get-mailbox $pemail -ea silentlycontinue)
	{
		Try
		{
			Set-MailboxAutoReplyConfiguration -identity $pemail -AutoReplyState Scheduled -StartTime $start -endtime $end -InternalMessage $ooomessage –ExternalMessage $ooomessage -ExternalAudience:All
			Write-Host "Processing...................... $pemail" -ForegroundColor Green
		}
		catch
		{
			Write-Host "exception occured Processing $pemail" -ForegroundColor Yellow
			$_.Exception.Message
		}
	}
	else
	{
		Write-Host "$pemail mailbox not found" -ForegroundColor RED
	}
}
Stop-Transcript
######################################################################
