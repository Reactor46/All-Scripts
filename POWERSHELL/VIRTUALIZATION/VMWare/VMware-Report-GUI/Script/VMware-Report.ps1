<#
 NAME: VMware-Report.ps1

 AUTHOR: Zachary Loeber
 DATE:   03/29/2013
 EMAIL:  zloeber@gmail.com
 COMMENT: VMware reporting script
 VERSION HISTORY
 	0.0.1 - 02/12/2013
 		- First release
    0.0.2 - 03/29/2013
        - Removed clear text password from config file
        - Allowed for alternate credentials for wmi calls
        - Huge GUI updates and fixes.
        
 TO ADD
   - ???
#>
#region Parameters
[CmdletBinding()]
param
(
	[Parameter(Position=0,Mandatory=$false,ValueFromPipeline=$false,ValueFromPipelineByPropertyName=$true)]
	[String]$ConfigFile = ""
)
#Add-PSSnapin VMware.VimAutomation.Core

#endregion Parameters

#region Help
<#
.SYNOPSIS
   VMware-Report.ps1
.DESCRIPTION
Use this script to regularly monitor and/or report on your vmware environment.
This shoudl be used in conjunction with VMwareReport-GUI.ps1 to generate the
config.xml file utilized for all report settings. Both VMwareReport-GUI.ps1 
and this script require a common set of variables and functions found in Globals.ps1
as well.

.PARAMETER ConfigFile
	Full file and path to an xml configuration file. If you are using multiple config files you can use this option with your scheduled tasks.
.EXAMPLE
	.\Get-VMwareReport.ps1 -Configfile ".\Config2.xml"
	Run script with alternate configuration file.
	

.EXAMPLE
	.\Get-VMwareReport.ps1 -Configfile 
	
	Run script with default configuration file (.\config.xml)
.LINK
   http://the-little-things.net
#>
#endregion help

#region Configuration (local)
$ReportToBeGenerated = $false

#endregion Configuration (local)

#region Globals
$Snapins=@(’VMware.VimAutomation.Core’)
#endregion Globals

#region Module/Snapin/Dot Sourcing
$RequiredSnapinsLoaded=$True 
if ($Snapins.Count -ge 1) 
{
	Foreach ($Snapin in $Snapins)
	{
		Add-PSSnapin $Snapin –ErrorAction SilentlyContinue 
		if ((Get-PSSnapin $Snapin) –eq $NULL) 
		{
			$RequiredSnapinsLoaded=$false
		} 
 	}
}
if (!$RequiredSnapinsLoaded) 
{ 
	exit $LASTEXITCODE 
}

# Dot sourced files
. .\globals.ps1
#endregion Module/Snapin/Dot Sourcing

#region Functions (local)
Function Get-CustomHTML ($Header) {
	$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($Header)</title>
<META http-equiv=Content-Type content='text/html; charset=windows-1252'>

<meta name="save" content="history">

<style type="text/css">
DIV .expando {DISPLAY: block; FONT-WEIGHT: normal; FONT-SIZE: 10pt; RIGHT: 8px; COLOR: #ffffff; FONT-FAMILY: Tahoma; POSITION: absolute; TEXT-DECORATION: underline}
TABLE {TABLE-LAYOUT: fixed; FONT-SIZE: 100%; WIDTH: 100%}
*{margin:0}
.dspcont { BORDER-RIGHT: #bbbbbb 1px solid; BORDER-TOP: #bbbbbb 1px solid; PADDING-LEFT: 16px; FONT-SIZE: 8pt;MARGIN-BOTTOM: -1px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 0px; BORDER-LEFT: #bbbbbb 1px solid; WIDTH: 95%; COLOR: #000000; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #bbbbbb 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; BACKGROUND-COLOR: #f9f9f9}
.filler {BORDER-RIGHT: medium none; BORDER-TOP: medium none; DISPLAY: block; BACKGROUND: none transparent scroll repeat 0% 0%; MARGIN-BOTTOM: -1px; FONT: 100%/8px Tahoma; MARGIN-LEFT: 43px; BORDER-LEFT: medium none; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: medium none; POSITION: relative}
.save{behavior:url(#default#savehistory);}
.dspcont1{ display:none}
a.dsphead0 {BORDER-RIGHT: #bbbbbb 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #bbbbbb 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #bbbbbb 1px solid; CURSOR: hand; COLOR: #FFFFFF; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #bbbbbb 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #cc0000}
a.dsphead1 {BORDER-RIGHT: #bbbbbb 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #bbbbbb 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #bbbbbb 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #bbbbbb 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #7BA7C7}
a.dsphead2 {BORDER-RIGHT: #bbbbbb 1px solid; PADDING-RIGHT: 5em; BORDER-TOP: #bbbbbb 1px solid; DISPLAY: block; PADDING-LEFT: 5px; FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-BOTTOM: -1px; MARGIN-LEFT: 0px; BORDER-LEFT: #bbbbbb 1px solid; CURSOR: hand; COLOR: #ffffff; MARGIN-RIGHT: 0px; PADDING-TOP: 4px; BORDER-BOTTOM: #bbbbbb 1px solid; FONT-FAMILY: Tahoma; POSITION: relative; HEIGHT: 2.25em; WIDTH: 95%; BACKGROUND-COLOR: #A5A5A5}
a.dsphead1 span.dspchar{font-family:monospace;font-weight:normal;}
td {VERTICAL-ALIGN: TOP; FONT-FAMILY: Tahoma}
th {VERTICAL-ALIGN: TOP; COLOR: #cc0000; TEXT-ALIGN: left}
BODY {margin-left: 4pt} 
BODY {margin-right: 4pt} 
BODY {margin-top: 6pt} 
</style>
</head>
<body>
<b><font face="Arial" size="5">$($Header)</font></b><hr size="8" color="#cc0000">
<font face="Arial" size="1"><b>Generated on $($ENV:Computername)</b></font><br>
<font face="Arial" size="1">Report created on $(Get-Date)</font>
<div class="filler"></div>
<div class="filler"></div>
<div class="filler"></div>
<div class="save">
"@
	Return $Report
}

Function Get-CustomHeader0 ($Title) {
	$Report = @"
		<h1><a class="dsphead0">$($Title)</a></h1>
		<div class="filler"></div>
"@
	Return $Report
}

Function Get-CustomHeader ($Num, $Title) {
	$Report = @"
	<h2><a class="dsphead$($Num)">
	$($Title)</a></h2>
	<div class="dspcont">
"@
	Return $Report
}

Function Get-CustomHeaderClose {
	$Report = @"
		</DIV>
		<div class="filler"></div>
"@
Return $Report
}

Function Get-CustomHeader0Close {
 $Report = @"
</DIV>
"@
Return $Report
}

Function Get-CustomHTMLClose {

 $Report = @"
</div>

</body>
</html>
"@
Return $Report
}

Function Get-HTMLTable {
	param([array]$Content)
	$HTMLTable = $Content | ConvertTo-Html
	$HTMLTable = $HTMLTable -replace '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">', ""
	$HTMLTable = $HTMLTable -replace '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN"  "http://www.w3.org/TR/html4/strict.dtd">', ""
	$HTMLTable = $HTMLTable -replace '<html xmlns="http://www.w3.org/1999/xhtml">', ""
	$HTMLTable = $HTMLTable -replace '<html>', ""
	$HTMLTable = $HTMLTable -replace '<head>', ""
	$HTMLTable = $HTMLTable -replace '<title>HTML TABLE</title>', ""
	$HTMLTable = $HTMLTable -replace '</head><body>', ""
	$HTMLTable = $HTMLTable -replace '</body></html>', ""
	Return $HTMLTable
}

Function Get-HTMLDetail ($Heading, $Detail) {
	$Report = @"
<TABLE>
	<tr>
	<th width='25%'><b>$Heading</b></font></th>
	<td width='75%'>$($Detail)</td>
	</tr>
</TABLE>
"@
	Return $Report
}

function Get-VIServices {
	$WMIParameters = @{
			Class = 'win32_service'
			ComputerName = $varVIServer
		}
		   
		if ($WMICred -ne $Null) {
			$WmiParameters.Credential = $WMICred
		}  
	$Services = Get-WmiObject @WMIParameters -ErrorAction:SilentlyContinue | Where {$_.DisplayName -like "VMware*" }
	#$Services = get-wmiobject win32_service -Credential $Cred -ErrorAction:SilentlyContinue -ComputerName $varVIServer | Where {$_.DisplayName -like "VMware*" }
	$myCol = @()
	Foreach ($service in $Services){
		$MyDetails = "" | select-Object Name, State, StartMode, Health
		If ($service.StartMode -eq "Auto")
		{
			if ($service.State -eq "Stopped")
			{
				$MyDetails.Name = $service.Displayname
				$MyDetails.State = $service.State
				$MyDetails.StartMode = $service.StartMode
				$MyDetails.Health = "Unexpected State"
			}
		}
		If ($service.StartMode -eq "Auto")
		{
			if ($service.State -eq "Running")
			{
				$MyDetails.Name = $service.Displayname
				$MyDetails.State = $service.State
				$MyDetails.StartMode = $service.StartMode
				$MyDetails.Health = "OK"
			}
		}
		If ($service.StartMode -eq "Disabled")
		{
			If ($service.State -eq "Running")
			{
				$MyDetails.Name = $service.Displayname
				$MyDetails.State = $service.State
				$MyDetails.StartMode = $service.StartMode
				$MyDetails.Health = "Unexpected State"
			}
		}
		If ($service.StartMode -eq "Disabled")
		{
			if ($service.State -eq "Stopped")
			{
				$MyDetails.Name = $service.Displayname
				$MyDetails.State = $service.State
				$MyDetails.StartMode = $service.StartMode
				$MyDetails.Health = "OK"
			}
		}
		$myCol += $MyDetails
	}
	Write-Output $myCol
}

function Get-ThinDisk {
<#
.SYNOPSIS
  Returns all virtual disks that are Thin
.DESCRIPTION
  The function returns all the Thin disks it finds on the datastore(s)
	passed to the function
.NOTES
  Authors:	Luc Dekens
.PARAMETER Datastore
  On or more datastore objects returned by Get-Datastore
.PARAMETER Thin
  If this switch is $true, only Thin virtual disks are returned. If the
switch is set to $false, all non-Thin virtual disks are returned.
.EXAMPLE
  PS> Get-Datastore | Get-ThinDisk
.EXAMPLE
  PS> Get-ThinDisk -Datastore (Get-Datastore -Name "DS*")
#>

	param(
	[parameter(valuefrompipeline = $true, mandatory = $true,
	HelpMessage = "Enter a datastore")]
	[VMware.VimAutomation.ViCore.Impl.V1.DatastoreManagement.DatastoreImpl[]]$Datastore,
	[switch]$Thin = $true
	)

	begin{
		if((Get-PowerCLIVersion).Build -lt 264274){
			Write-Error "The script requires at least PowerCLI 4.1 !"
			exit
		}

		$searchspec = New-Object VMware.Vim.HostDatastoreBrowserSearchSpec
		$query = New-Object VMware.Vim.VmDiskFileQuery
		$query.Details = New-Object VMware.Vim.VmDiskFileQueryFlags
		$query.Details.capacityKb = $true
		$query.Details.controllerType = $true
		$query.Details.diskExtents = $true
		$query.Details.diskType = $true
		$query.Details.hardwareVersion = $true
		$query.Details.thin = $true
		$query.Filter = New-Object VMware.Vim.VmDiskFileQueryFilter
		$query.Filter.Thin = $Thin
		$searchspec.Query += $query
	}

	process{
		$Datastore | %{
			$dsBrowser = Get-View $_.Extensiondata.Browser
			$datastorepath = "[" + $_.Name + "]"

			$taskMoRef = $dsBrowser.SearchDatastoreSubFolders_Task($datastorePath, $searchSpec)

			$task = Get-View $taskMoRef
			while ("running","queued" -contains $task.Info.State){
				$task.UpdateViewData("Info.State")
			}
			$task.UpdateViewData("Info.Result")
			if($task.Info.Result){
				foreach ($folder in $task.Info.Result){
					if($folder.File){
						foreach($file in $folder.File){
							$record = "" | Select DSName,Path,VmdkName
							$record.DSName = $_.Name
							$record.Path = $folder.FolderPath
							$record.VmdkName = $file.Path
							$record
						}
					}
				}
			}
		}
	}
}

function Find-Username ($username) {
	if ($username -ne $null)
	{
		$root = [ADSI]""
		$filter = ("(&(objectCategory=user)(samAccountName=$Username))")
		$ds = new-object system.DirectoryServices.DirectorySearcher($root,$filter)
		$ds.PageSize = 1000
		$ds.FindOne()
	}
}
#endregion Functions (local)

#region Start Script
#region Connect to VI
#If ($VIConnection.IsConnected -ne $true){
#	# Fix for scheduled tasks not running.
#	$USER = $env:username
#	$APPPATH = "C:\Documents and Settings\" + $USER + "\Application Data"
#
#	#SET THE APPDATA ENVIRONMENT WHEN NEEDED
#	if ($env:appdata -eq $null -or $env:appdata -eq 0)
#	{
#		$env:appdata = $APPPATH
#	}
#	$VIConnection = Connect-VIServer $varVIServer -ErrorAction:SilentlyContinue
#}

Load-Config

# Connect to the server
if ($varUseCurrentUser)
{
	$VIConnection = Connect-VIServer $varVIServer -ErrorAction:SilentlyContinue
}
else
{
	#$VIConnection = Connect-VIServer $varVIServer -User $varVIUser -Password $varVIPassword -ErrorAction:SilentlyContinue
    $creds = Get-VICredentialStoreItem -file $CredFile 
    $VIConnection = Connect-VIServer -Server $creds.Host -User $creds.User -Password $creds.Password
	$WMICred = New-Object System.Management.Automation.PSCredential([string]$creds.User,([string]$creds.Password|ConvertTo-SecureString -asPlainText -Force))
}

#endregion Connect to VI
If ($VIConnection.IsConnected)
{
	#region Additional VMware Properties
	New-VIProperty -Name PercentFree -ObjectType Datastore -Value {"{0:N2}" -f ($args[0].FreeSpaceMB/$args[0].CapacityMB*100)} -Force
	#endregion Additional VMware Properties

	#region Report Scope
	# Scope = Whole Farm
	if ($varScopeWholeFarm)
	{
		$Datacenters = @(Get-Datacenter)
		$VMs = @(Get-VM)
		$Hosts = @(Get-VMHost)
		$Clusters = @(Get-Cluster)
		$Datastores = @(Get-Datastore)
		$FullVM = @(Get-View -ViewType VirtualMachine)
	}
	else
	{
		# Should at least have a datacenter defined
		$Datacenters = @(Get-Datacenter $varScopeDatacenter)
		
		if ($varScopeHost -eq '')
		{
			# Scope = Just Datacenter
			if ($varScopeCluster -eq '') # Scope = Datacenter Level
			{
				$Clusters = (Get-Datacenter $varScopeDatacenter | Get-Cluster)
			}
			# Scope = Datacenter + Cluster
			else 
			{
				$Clusters = @(Get-Cluster $varScopeCluster)
			}
			$Hosts = @()
			Foreach ($Cluster in $Clusters)
			{
				Foreach ($HostServer in (Get-Cluster $Cluster | Get-VMHost))
				{
					$Hosts += $HostServer
				}
			}
		}
		else # Scope = Datacenter + Cluster + Host
		{
			$Clusters = (Get-Datacenter | Get-Cluster $varScopeCluster)
			$Hosts = ($Clusters | Get-VMHost $varScopeHost)
		}
		$VMs = @()
		Foreach ($Cluster in $Clusters)
		{
			Foreach ($VM in (Get-Cluster $Cluster | Get-VM))
			{
				$VMs += $VM
			}
		}
		
		# Get all datastores related with hosts
		$Datastores = $Hosts | Get-Datastore | Select -Unique		
	}
#endregion Report Scope

	$MyReport = Get-CustomHTML "$VIServer Daily Report"
	$MyReport += Get-CustomHeader0 ($varVIServer)

	# ---- General Summary Info ----
	$MyReport += Get-CustomHeader "1" "General Details"
	$MyReport += Get-HTMLDetail "Number of Hosts:" (($Hosts).Count)
	$MyReport += Get-HTMLDetail "Number of VMs:" (($VMs).Count)
	$MyReport += Get-HTMLDetail "Number of Clusters:" (($Clusters).Count)
	$MyReport += Get-HTMLDetail "Number of Datastores:" (($Datastores).Count)
	$MyReport += Get-CustomHeaderClose
	
	#region Virtual Center
	# ---- VC Host Services
	if ($varReportVCServices -and $varScopeWholeFarm)
	{
		$ReportToBeGenerated = $true
		$MyReport += Get-CustomHeader "1" "$VIServer Service Details"
		$MyReport += Get-HTMLTable (Get-VIServices)
		$MyReport += Get-CustomHeaderClose
	}
	
	# ---- VC Errors ----	
	if ($varReportVCErrors -and $varScopeWholeFarm)
	{
		# ---- VC Errors
		$VIEvent = Get-VIEvent -maxsamples 10000 -Start (Get-Date).AddDays(-$varReportVCErrorsAge)
		$OutputErrors = $VIEvent  | `
			where {$_.Type -eq "Error"} | `
			Select createdTime, `
				   @{N="User";E={(Find-Username (($_.userName.split("\"))[1])).Properties.displayname}},`
				   fullFormattedMessage
		If (($OutputErrors | Measure-Object).count -gt 0) {
			$ReportToBeGenerated = $true
			$MyReport += Get-CustomHeader "1" "Error Events (Last $varReportVCErrorsAge Day(s))"
			$MyReport += Get-HTMLTable $OutputErrors
			$MyReport += Get-CustomHeaderClose 
		}
		
		# ---- VMs created or Cloned ----
		if ($varReportVCVMsCreated)
		{
			$OutputCreatedVMs = $VIEvent | where {$_.Gettype().Name -eq "VmCreatedEvent" -or $_.Gettype().Name -eq "VmBeingClonedEvent" -or $_.Gettype().Name -eq "VmBeingDeployedEvent"} | Select createdTime, @{N="User";E={(Find-Username (($_.userName.split("\"))[1])).Properties.displayname}}, fullFormattedMessage
			If (($OutputCreatedVMs | Measure-Object).count -gt 0) {
				$ReportToBeGenerated = $true
				$MyReport += Get-CustomHeader "1" "VMs Created or Cloned (Last $varReportVCErrorsAge Day(s))"
				$MyReport += Get-HTMLTable $OutputCreatedVMs
				$MyReport += Get-CustomHeaderClose
			}
		}
		
		# ---- VMs Removed ----
		if ($varReportVCVMsDeleted)
		{
			$OutputRemovedVMs = $VIEvent | where {$_.Gettype().Name -eq "VmRemovedEvent"}| Select createdTime, @{N="User";E={(Find-Username (($_.userName.split("\"))[1])).Properties.displayname}}, fullFormattedMessage
			If (($OutputRemovedVMs | Measure-Object).count -gt 0) {
				$ReportToBeGenerated = $true
				$MyReport += Get-CustomHeader "1" "VMs Removed (Last $varReportVCEventAge Day(s))"
				$MyReport += Get-HTMLTable $OutputRemovedVMs
				$MyReport += Get-CustomHeaderClose
			}
		}
	}
	#endregion Virtual Center
	
	# Windows Event Logs
	if ($varReportVCEvntlogs -and $varScopeWholeFarm)
	{	
		# Errors
		$ConvDate = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::Now.AddDays(-$varReportVCEventLogsAge))
		$WMIParameters = @{
			Query = ("Select * from Win32_NTLogEvent Where Type='Error' and TimeWritten >='" + $ConvDate + "'")
			ComputerName = $varVIServer
		}	   
		if ($WMICred -ne $Null) {
			$WmiParameters.Credential = $WMICred
		}  
		$ErrLogs = Get-WmiObject @WMIParameters | Select @{N="TimeGenerated";E={$_.ConvertToDateTime($_.TimeGenerated)}}, Message
		#$ErrLogs = Get-WmiObject @WMIParameters | Where {$_.Message -like "*VMware*"} | Select @{N="TimeGenerated";E={$_.ConvertToDateTime($_.TimeGenerated)}}, Message
		#$ErrLogs = Get-WmiObject -ErrorAction:SilentlyContinue -computer $varVIServer -query ("Select * from Win32_NTLogEvent Where Type='Error' and TimeWritten >='" + $ConvDate + "'") | Where {$_.Message -like "*VMware*"} | Select @{N="TimeGenerated";E={$_.ConvertToDateTime($_.TimeGenerated)}}, Message
		
		#Warnings
		$WMIParameters = @{
			Query = ("Select * from Win32_NTLogEvent Where Type='Warning' and TimeWritten >='" + $ConvDate + "'")
			ComputerName = $varVIServer
		}	   
		if ($WMICred -ne $Null) {
			$WmiParameters.Credential = $WMICred
		}  
		$WarnLogs = Get-WmiObject @WMIParameters | Select @{N="TimeGenerated";E={$_.ConvertToDateTime($_.TimeGenerated)}}, Message 
		#$WarnLogs = Get-WmiObject @WMIParameters | Where {$_.Message -like "*VMware*"} | Select @{N="TimeGenerated";E={$_.ConvertToDateTime($_.TimeGenerated)}}, Message 
		#$WarnLogs = Get-WmiObject -ErrorAction:SilentlyContinue -computer $varVIServer -query ("Select * from Win32_NTLogEvent Where Type='Warning' and TimeWritten >='" + $ConvDate + "'") | Where {$_.Message -like "*VMware*"} | Select @{N="TimeGenerated";E={$_.ConvertToDateTime($_.TimeGenerated)}}, Message 

		If (($ErrLogs | Measure-Object).count -gt 0) 
		{
			$ReportToBeGenerated = $true
			$MyReport += Get-CustomHeader "1" "$VIServer Windows Event Logs: Error"
			$MyReport += Get-HTMLTable ($ErrLogs)
			$MyReport += Get-CustomHeaderClose
		}
		# Warnings
		If (($WarnLogs | Measure-Object).count -gt 0) {
			$ReportToBeGenerated = $true
			$MyReport += Get-CustomHeader "1" "$VIServer Windows Event Logs: Warning"
			$MyReport += Get-HTMLTable ($WarnLogs)
			$MyReport += Get-CustomHeaderClose
		}
	}	

	#region VMware Hosts	
	# ---- Hosts Not responding ----
	if ($varReportHostNotResponding)
	{
		$RespondHosts = $Hosts | where {$_.State -match "Not"} | Select name,CustomFields
		If (($RespondHosts | Measure-Object).count -gt 0) {
			$ReportToBeGenerated = $true
			$MyReport += Get-CustomHeader "1" "Hosts not responding"
			$MyReport += Get-HTMLTable $RespondHosts
			$MyReport += Get-CustomHeaderClose
		}
	}
	
	# ---- Hosts in Maintenance Mode ----
	if ($varReportHostsInMaint)
	{

		$MaintHosts = $Hosts | where {$_.State -match "Maintenance"} | Select name,CustomFields
		If (($MaintHosts | Measure-Object).count -gt 0) {
			$ReportToBeGenerated = $true
			$MyReport += Get-CustomHeader "1" "Hosts in Maintenance Mode"
			$MyReport += Get-HTMLTable $MaintHosts
			$MyReport += Get-CustomHeaderClose
		}
	}

	# Datastore Report
	if ($varReportHostsDatastore)
	{
		if ($varReportSelective)
		{
			$OutputDatastores = @($Datastores | where {$_.PercentFree -le $varReportHostsDatastoreThreshold} | select-Object Name, Type, CapacityMB, FreeSpaceMB, PercentFree | Sort PercentFree)
		}
		else
		{
			$OutputDatastores = @($Datastores | select-Object Name, Type, CapacityMB, FreeSpaceMB, PercentFree | Sort PercentFree)
		}
		If (($OutputDatastores | Measure-Object).count -gt 0) 
		{
			$ReportToBeGenerated = $true
			if ($varReportSelective)
			{
				$MyReport += Get-CustomHeader "1" "Datastores (Less than $varReportDatastoreThreshold% Free)"				
			}
			else
			{
				$MyReport += Get-CustomHeader "1" "Datastores"
			}
			$MyReport += Get-HTMLTable $OutputDatastores
			$MyReport += Get-CustomHeaderClose
		}
	}
	#endregion VMware Hosts
	
	#region VMware Guests
	# ---- No VM Tools ----
	if ($varReportVMTools)
	{
		# ---- No VM Tools ----
		$NoTools = $FullVM | Where { $_.Runtime.PowerState -eq "poweredOn" } | Select Name, @{N="ToolsVersion"; E={$_.Config.tools.toolsVersion}} | Where { $_.ToolsVersion -eq 0} | Select Name
		If (($NoTools | Measure-Object).count -gt 0) {
			$ReportToBeGenerated = $true
			$MyReport += Get-CustomHeader "1" "No VMTools"
			$MyReport += Get-HTMLTable $NoTools
			$MyReport += Get-CustomHeaderClose
		}
	}
	
	# ---- CD-Roms Connected ----
	if ($varReportCDConnected)
	{
		# ---- CD-Roms Connected ----
		$CDConn = $VMs| Where { $_ | Get-CDDrive | Where { $_.ConnectionState.Connected -eq "true" } } | Select Name, Host
		If (($CDConn | Measure-Object).count -gt 0) {
			$ReportToBeGenerated = $true
			$MyReport += Get-CustomHeader "1" "VM: CD-ROM Connected - VMotion Violation"
			$MyReport += Get-HTMLTable $CDConn
			$MyReport += Get-CustomHeaderClose
		}
	}
	
	# ---- Floppys Connected ----
	if ($varReportFloppyConnected)
	{
		$Floppy = $VMs| Where { $_ | Get-FloppyDrive | Where { $_.ConnectionState.Connected -eq "true" } } | Select Name, Host
		If (($Floppy | Measure-Object).count -gt 0) {
			$ReportToBeGenerated = $true
			$MyReport += Get-CustomHeader "1" "VM:Floppy Drive Connected - VMotion Violation"
			$MyReport += Get-HTMLTable $Floppy
			$MyReport += Get-CustomHeaderClose
		}
	}

	# VM Snapshot Report
	if ($varReportVMSnapshots)
	{
		$SnapshotReport = [PSObject]@()
		$OldSnapshotCount = 0
		ForEach ($vm in $VMs)
		{
			# List snaphosts
			$snapshots = Get-Snapshot -VM (Get-VM -Name $vm.Name) -WarningAction SilentlyContinue

			if ($snapshots -ne $null)
			{
				ForEach ($snapshot in $snapshots)
				{
					$SnapshotCount = $SnapshotCount + 1
					if ($varReportSelective -and ($snapshot.created -lt (Get-Date).AddDays(-$varReportSnapshotsThreshold)))
					{
						$OldSnapshotCount = $OldSnapshotCount +1
						$SnapshotObj = New-Object PSObject
						$SnapshotObj | Add-Member NoteProperty -Name "Snapshot Name" -Value $snapshot.Name
						$SnapshotObj | Add-Member NoteProperty -Name "VM Name" -Value $vm.Name
						$SnapshotObj | Add-Member NoteProperty -Name "VM Powerstate" -Value $vm.PowerState 
						$SnapshotObj | Add-Member NoteProperty -Name "Created" -Value $snapshot.Created
						$SnapshotReport += $SnapshotObj
					}
					else
					{
						$SnapshotObj = New-Object PSObject
						$SnapshotObj | Add-Member NoteProperty -Name "Snapshot Name" -Value $snapshot.Name
						$SnapshotObj | Add-Member NoteProperty -Name "VM Name" -Value $vm.Name
						$SnapshotObj | Add-Member NoteProperty -Name "VM Powerstate" -Value $vm.PowerState 
						$SnapshotObj | Add-Member NoteProperty -Name "Created" -Value $snapshot.Created
						$SnapshotReport += $SnapshotObj
					}
				}
			}
		}
		if ($SnapshotReport -ne $null)
		{
			if ($OldSnapshotCount -ge 1)
			{
				$ReportToBeGenerated = $true
				$MyReport += Get-CustomHeader "1" "VM Snapshots (Over $varReportSnapshotsThreshold Days Old)"
				$MyReport += Get-HTMLTable $SnapshotReport
				$MyReport += Get-CustomHeaderClose

			}
			else
			{
				if (!$varReportSelective)
				{
					$ReportToBeGenerated = $true
					$MyReport += Get-CustomHeader "1" "VM Snapshots"
					$MyReport += Get-HTMLTable $SnapshotReport
					$MyReport += Get-CustomHeaderClose
				}
			}
		}
	}

	# VM Thin Provisioned Report
	if ($varReportVMThinProvisioned)
	{
		$ThinDisks = $Datastores | Get-Thindisk
		If (($ThinDisks | Measure-Object).count -gt 0) {
			$ReportToBeGenerated = $true
			$MyReport += Get-CustomHeader "1" "VM:Thin Providioned Disks"
			$MyReport += Get-HTMLTable $ThinDisks
			$MyReport += Get-CustomHeaderClose
		}
	}
	#endregion VMware Guests
	
	$MyReport += Get-CustomHeader0Close
	$MyReport += Get-CustomHTMLClose
	Disconnect-VIServer $varVIServer -Confirm:$false
}
else
{
	$varEmailSubject = "ERROR:" + $varEmailSubject
	$htmlBody = "<p>Script was unable to connect to the VI server, sorry :(</p>"
}
#endregion Start Script

#region Script Output
#  Final Report
#$ReportOutput = $htmlHead + $htmlBody + $htmlTail
if ($ReportToBeGenerated)
{
	if ($varEmailReport -or $varSaveReportsLocally)
	{
		# Send an email
		if ($varSendMail)
		{
			$HTMLmessage = $MyReport | Out-String
			$email= 
			@{
				From = $varEmailSender
				To = $varEmailRecipient
		#		CC = "EMAIL@EMAIL.COM"
				Subject = $varEmailSubject
				SMTPServer = $varSMTPServer
				Body = $HTMLmessage
				Encoding = ([System.Text.Encoding]::UTF8)
				BodyAsHTML = $true
			}
			Send-MailMessage @email
			Sleep -Milliseconds 200
		}
		
		# Save a report
		if ($varSaveReportsLocally)
		{
			$MyReport | Out-File "$varReportFolder\$varReportName"
		}
	}
	else
	{
	   Return $MyReport
	}
}
#endregion Script Output