<#	
	.NOTES
	===========================================================================
	 Created on:   	3/31/2017 9:04 AM
	 Created by:   	sukhijv
	 Organization: 	
	 Filename:     	MergerExportPst.ps1
	===========================================================================
	.DESCRIPTION
		This script will export the PST files for list of suers.
        Folder with name transfer would be extracted
#>
######################Functions########################
Function Write-Log
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
##########################Import CSv & START EXPORT################
$exportfolder = "transfer"
$location = "\\Server\e$\Merger\PST"
$data = Import-Csv $args
foreach ($i in $data)
{
	$pemail = $i.Previousemail
	$pemail = $pemail.trim()
	$networkid = $i.Networkid
	$networkid = $networkid.trim()
	$path = $location + "\" + $networkid + ".pst"
	try
	{
		Write-Host "Processing...........$pemail................$networkid" -ForegroundColor Green
		New-MailboxExportRequest -Mailbox $pemail -IncludeFolders $exportfolder/* -filepath $path -ExcludeDumpster	
	}
	catch
	{
		Write-Host "Exception has occured processing $pemail....$networkid" -ForegroundColor Yellow
		$_.Exception.Message
	}
}
Stop-Transcript
######################################################################