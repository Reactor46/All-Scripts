<#
	.SYNOPSIS
	Configure Cisco UCS Service Profile for use with Citrix XenDesktop using Citrix XenServer, Microsoft Hyper-V, or VMware vSphere hypervisors
	
	.DESCRIPTION
	The purpose of the script is to quickly build the pools, policies, and templates necessary in a Citrix XenDesktop environment following
	the best practices of Cisco and Citrix.  It is built with the assumption that ESXi, XenServer, or Hyper-V are used as the hypervisor.
	It also assumes a relatively clean UCS environment, with very little configuration already completed (there are very little checks in the
	script for existing configurations).  The best practices were developed over a 3 day period with input from Cisco and Citrix TMEs, CSEs,
	and customer feedback.

	The script takes the name of the Excel Configuration file and, optional, a switch to output logs to the console, as input.
	Example:

		/UCS_Config_Excel.ps1 -ExcelFile XenDesktopBP.xlsx -ToConsole

	Inside the script (starting around line 350), there is a commented out section that could output all variables from the worksheet
	and derived variables (I added this for troubleshooting).  Starting around line 505, there is another commented out section that
	could be used to clean up a UCS Platform Emulator environment prior to creation of the all the pools, policies, and templates.

	The script creates the following using input from the Excel Configuration worksheet:
		Management IP Pool
		UUID Pool
		8 MAC Pools (1 each for the management network, virtual machine data network, motion network (i.e.., vMotion), and IP storage network; A and B fabrics)
		QoS System Class settings
		4 QoS Policies (1 each for the management network, virtual machine data network, motion network (i.e.., vMotion), and IP storage network)
		Network Control Policy
		4 VLANs (1 each for the management network, virtual machine data network, motion network (i.e.., vMotion), and IP storage network)
		8 vNIC Templates (1 each for the management network, virtual machine data network, motion network (i.e.., vMotion), and IP storage network; A and B fabrics)
		WWNN Pool
		4 WWPN Pools (for 4 vHBA's - only two are used by the template created in this script)
		2 VSANs (A and B fabrics)
		4 vHBA Templates (for 4 vHBA's - only two are used by the template created in this script)
		SAN Boot Policy
		Server Pool
		Server Pool Policy
		BIOS Policy
		Host Firmware Package Policy
		Management Firmware Package Policy
		Maintenance Policy
		Local Disk Configuration Policy
		Service Profile Template
	
	.PARAMETER ExcelFile
	The name of the Excel file. This file must be in the same directory as this script.
	
	.PARAMETER ToConsole
	A switch. If specified, the output will be printed on a console.
	
	.EXAMPLE
	./UCS_Config_Excel.ps1 -ExcelFile XenDesktopBP.xlsx -ToConsole
	This reads ucspe.xlsx file and outputs the status of configuration to the Powershell console.
	
	.EXAMPLE
	./UCS_Config_Excel.ps1 XenDesktopBP.xlsx
	This reads ucspe.xlsx file. An output file, UCSM_Configuration_Script_Log.txt, is created.
	
	.Cisco UCS PowerTool Version
	Developed for use with CiscoUcs-PowerTool-0.9.10.0
	
	.NOTES
	Name:		UCS_Config__CitrixXDBP.ps1
	Alias:		
	Author:		Chris Carter (much of the work completed by others)
	Created:	10/2/2012
	Update History:	10/2/2012	Chris Carter	First Release	
	                10/4/2012   Chris Carter    Updated for use with Cisco UCS-PowerTool-0.9.10.0
#>

param([parameter(mandatory=$true)][validateNotNullOrEmpty()]$excelFile, [switch]$toConsole)

function Out-FileOrConsole ($content)
{
	if ($toConsole)
	{ Write-Host "$(Get-Date) $content" -ForegroundColor "Yellow" -BackgroundColor "Black"}
	else
	{ Add-Content -Path $scriptLogFullPath -Value "$(Get-Date) $content" }
} ##### End of function Out-FileOrConsole

function Remove-File
{
	param($fileName)
	if (Test-Path($fileName)) { del $fileName }
} ##### End of function Remove-File

function Remove-ComObject
{
	[CmdletBinding()]
	param()
	
	END
	{
		Start-Sleep -Milliseconds 500
		[Management.Automation.ScopedItemOptions] $scopedOpt = 'ReadOnly, Constant'
		Get-Variable -Scope 1 | where { $_.Value.PsTypeNames -contains 'System.__ComObject' -and -not($scopedOpt -band $_.Options) } |
		Remove-Variable -Scope 1 -Verbose:([bool]$psBoundParameters['Verbose'].IsPresent)
		
		[gc]::Collect()
		
	} ##### end of END
} ##### end of function Remove-ComObject

##### set up script logging
$thisPath = Split-Path (Resolve-Path $MyInvocation.MyCommand.Path)
Set-Location $thisPath
$scriptLog = "UCSM_Configuration_Script_Log.txt"
$scriptLogFullPath = Join-Path $thisPath $scriptLog
if (!$toConsole)
{ Remove-File $scriptLogFullPath }
$start = Get-Date
Out-FileOrConsole "Starting script logging."

##### make sure the CiscoUcsPS module is loaded
if (!(Get-Module -Name CiscoUcsPs))
{
	Out-FileOrConsole "Import module CiscoUcsPs"
	try {Import-Module CiscoUcsPs }
	catch
	{
		Out-FileOrConsole "..Importing module CiscoUcsPs failed. Quit the script."
		exit(1)
	}
}

##### make sure the Excel file exists
if (!(Test-Path $excelFile))
{ Out-FileOrConsole "The Excel file, $excelFile, does not exist. Quit the script."; exit(2) }

Out-FileOrConsole "Read the excel file..."
$fullPathName = Join-Path $thisPath $excelFile
try { $excel = New-Object -ComObject Excel.Application}
catch {
	Out-FileOrConsole "..Failed to access to Excel application. Quit the script."
	exit(2)
}
$excel.Visible = $false
try { $wb = $excel.Workbooks.Open($fullPathName) }
catch {
	Out-FileOrConsole "..Failed to open the Excel file, $fullPathName. Quit the script."
	$excel.Quit()
	Remove-ComObject
	exit(3)
}

### go to XenDesktop sheet
$sheetName = "XenDesktop"
Out-FileOrConsole "Open worksheet $sheetName..."
try { $ws = $wb.Worksheets.Item($sheetName) }
catch {
	Out-FileOrConsole "..Cannot open worksheet $sheetName. Quit the script."
	$wb.Close()
	$excel.Quit()
	Remove-ComObject
	exit(4)
}
$ws.Activate()

Out-FileOrConsole "Read values from worksheet $sheetName..."
[string]$ucsIp = $ws.Cells.Item(2, 2).Value2; if (!$ucsIp) { "..UCSM IP is missing. Quit the script."; exit }; $ucsIp = $ucsIp.Trim()
[string]$mgmtIpFrom = $ws.Cells.Item(6, 1).Value2; if (!$mgmtIpFrom) { "..KVM Mgmt IP From is missing. Quit the script."; exit }; $mgmtIpFrom = $mgmtIpFrom.Trim()
[string]$mgmtIpTo = $ws.Cells.Item(6, 2).Value2; if (!$mgmtIpTo) { "..KVM Mgmt IP To is missing. Quit the script."; exit }; $mgmtIpTo = $mgmtIpTo.Trim()
[string]$mgmtIpSubmask = $ws.Cells.Item(6, 3).Value2; if (!$mgmtIpSubmask) { "..KVM Mgmt IP Subnet Mask is missing. Quit the script."; exit }; $mgmtIpSubmask = $mgmtIpSubmask.Trim()
[string]$mgmtIpDefgw = $ws.Cells.Item(6, 4).Value2; if (!$mgmtIpDefgw) { "..KVM Mgmt IP Default Gateway is missing. Quit the script."; exit }; $mgmtIpDefgw = $mgmtIpDefgw.Trim()
[string]$siteId = $ws.Cells.Item(10, 2).Value2; if (!$siteId) { "..Site ID is missing. Quit the script."; exit }; $siteId = $siteId.Trim()
[string]$siteDescr = $ws.Cells.Item(11, 2).Value2; if (!$siteDescr) { $siteDescr = "Site" } 
[string]$podId = $ws.Cells.Item(12, 2).Value2; if (!$podId) { "..POD ID is missing. Quit the script."; exit }; $podId = $podId.Trim()
[string]$podDescr = $ws.Cells.Item(13, 2).Value2; if (!$podDescr) { $podDescr = "env" }
[string]$serviceProfileTemplate_name = $ws.Cells.Item(16, 2).Value2; if (!$serviceProfileTemplate_name) { "..Service Profile Template Name is missing. Quit the script."; exit }; $serviceProfileTemplate_name = $serviceProfileTemplate_name.Trim()
[string]$serviceProfileTemplate_descr = $ws.Cells.Item(17, 2).Value2; if (!$serviceProfileTemplate_descr) { $serviceProfileTemplate_descr = "Service Profile Template" }; $serviceProfileTemplate_descr = $serviceProfileTemplate_descr.Trim()
[string]$vlan_mgmt_name = $ws.Cells.Item(21, 1).Value2; if (!$vlan_mgmt_name) { "..Management VLAN Name is missing. Quit the script."; exit }; $vlan_mgmt_name = $vlan_mgmt_name.Trim()
[int]$vlan_mgmt_id = $ws.Cells.Item(21, 2).Value2; if (!$vlan_mgmt_id) { "..Management VLAN ID is missing. Quit the script."; exit };
[string]$vlan_mgmt_descr = $ws.Cells.Item(21, 3).Value2; if (!$vlan_mgmt_descr) { $vlan_mgmt_descr = "VLAN used for Hypervisor Management Network" }; $vlan_mgmt_descr = $vlan_mgmt_descr.Trim()
[string]$vlan_vmdata_name = $ws.Cells.Item(22, 1).Value2; if (!$vlan_vmdata_name) { "..VM DATA VLAN Name is missing. Quit the script."; exit }; $vlan_vmdata_name = $vlan_vmdata_name.Trim()
[int]$vlan_vmdata_id = $ws.Cells.Item(22, 2).Value2; if (!$vlan_vmdata_id) { "..VM DATA VLAN ID is missing. Quit the script."; exit };
[string]$vlan_vmdata_descr = $ws.Cells.Item(22, 3).Value2; if (!$vlan_vmdata_descr) { $vlan_vmdata_descr = "VLAN used for Virtual Desktop Data Network" }; $vlan_vmdata_descr = $vlan_vmdata_descr.Trim()
[string]$vlan_vmotion_name = $ws.Cells.Item(23, 1).Value2; if (!$vlan_vmotion_name) { "..vMotion VLAN Name is missing. Quit the script."; exit }; $vlan_vmotion_name = $vlan_vmotion_name.Trim()
[int]$vlan_vmotion_id = $ws.Cells.Item(23, 2).Value2; if (!$vlan_vmotion_id) { "..vMotion VLAN ID is missing. Quit the script."; exit };
[string]$vlan_vmotion_descr = $ws.Cells.Item(23, 3).Value2; if (!$vlan_vmotion_descr) { $vlan_vmotion_descr = "VLAN used for VM Motion Network" }; $vlan_vmotion_descr = $vlan_vmotion_descr.Trim()
[string]$vlan_ipstorage_name = $ws.Cells.Item(24, 1).Value2; if (!$vlan_ipstorage_name) { "..VM Motion VLAN Name is missing. Quit the script."; exit }; $vlan_ipstorage_name = $vlan_ipstorage_name.Trim()
[int]$vlan_ipstorage_id = $ws.Cells.Item(24, 2).Value2; if (!$vlan_ipstorage_id) { "..VM Motion VLAN ID is missing. Quit the script."; exit };
[string]$vlan_ipstorage_descr = $ws.Cells.Item(24, 3).Value2; if (!$vlan_ipstorage_descr) { $vlan_ipstorage_descr = "VLAN used for VM Motion Network" }; $vlan_ipstorage_descr = $vlan_ipstorage_descr.Trim()
[string]$vSan_a_name = $ws.Cells.Item(28, 2).Value2; if (!$vSan_a_name) { $vSan_a_name = "VSAN-A" }; $vSan_a_name = $vSan_a_name.Trim()
[int]$vSan_a_id = $ws.Cells.Item(28, 3).Value2; if (!$vSan_a_id) { "..VSAN A ID is missing. Quit the script."; exit };
[int]$FCoE_Vlan_a = $ws.Cells.Item(28, 4).Value2; if (!$FCoE_Vlan_a) { "..FCoE VLAN A is missing. Quit the script."; exit };
[string]$vSan_b_name = $ws.Cells.Item(29, 2).Value2; if (!$vSan_b_name) { $vSan_b_name = "VSAN-B" }; $vSan_b_name = $vSan_b_name.Trim()
[int]$vSan_b_id = $ws.Cells.Item(29, 3).Value2; if (!$vSan_b_id) { "..VSAN B ID is missing. Quit the script."; exit };
[int]$FCoE_Vlan_b = $ws.Cells.Item(29, 4).Value2; if (!$FCoE_Vlan_b) { "..FCoE VLAN B is missing. Quit the script."; exit };
[string]$san_primary_target_primary = $ws.Cells.Item(34, 1).Value2; if (!$san_primary_target_primary) { "..SAN Primary Target Primary WWN is missing. Quit the script."; exit }; $san_primary_target_primary = $san_primary_target_primary.Trim()
[string]$san_primary_target_secondary = $ws.Cells.Item(34, 2).Value2; if (!$san_primary_target_secondary) { "..SAN Primary Target Secondary WWN is missing. Quit the script."; exit }; $san_primary_target_secondary = $san_primary_target_secondary.Trim()
[string]$san_secondary_target_primary = $ws.Cells.Item(34, 3).Value2; if (!$san_secondary_target_primary) { "..SAN Secondary Target Primary WWN is missing. Quit the script."; exit }; $san_secondary_target_primary = $san_secondary_target_primary.Trim()
[string]$san_secondary_target_secondary = $ws.Cells.Item(34, 4).Value2; if (!$san_secondary_target_secondary) { "..SAN Secondary Target Secondary WWN is missing. Quit the script."; exit }; $san_secondary_target_secondary = $san_secondary_target_secondary.Trim()

##### close Excel and cleanup
Out-FileOrConsole "Close Excel file..."
$wb.Close()
$excel.Quit()
Remove-ComObject
Start-Sleep -Seconds 2


##### derived variables
#### uuid pool
$uuidName = "CitrixXD_UUID_site_" + $siteId + "_pod_" + $podId
$uuidDescr = $siteDescr + " " + $podDescr
$uuidFrom = "00" + $siteId + $podId + "-000000000001"
$uuidTo = "00" + $siteId + $podId + "-0000000003E8" #### 1000 uuid's

##### mac pools
##### mac pool MAC-MGMT-A
$macPoolName_MGMT_A = "MAC-MGMT-A"
$macPoolDescr_MGMT_A = "MAC Pool used for Management Port Group - Fabric A" + "; " + $siteDescr + " " + $podDescr
$macPoolFrom_MGMT_A = "00:25:B5:" + $siteId + $podId + ":1A:00"
$macPoolTo_MGMT_A = "00:25:B5:" + $siteId + $podId + ":1A:FF" #### 256 mac addresses

##### mac pool MAC-VM_DATA-A
$macPoolName_VM_DATA_A = "MAC-VM_DATA-A"
$macPoolDescr_VM_DATA_A = "MAC Pool used for Guest Virtual Machine (DATA) Port Group - Fabric A" + "; " + $siteDescr + " " + $podDescr
$macPoolFrom_VM_DATA_A = "00:25:B5:" + $siteId + $podId + ":2A:00"
$macPoolTo_VM_DATA_A = "00:25:B5:" + $siteId + $podId + ":2A:FF" #### 256 mac addresses

##### mac pool MAC-VM_MOTION-A
$macPoolName_VM_MOTION_A = "MAC-VM_MOTION-A"
$macPoolDescr_VM_MOTION_A = "MAC Pool used for VM_MOTION Port Group - Fabric A" + "; " + $siteDescr + " " + $podDescr
$macPoolFrom_VM_MOTION_A = "00:25:B5:" + $siteId + $podId + ":3A:00"
$macPoolTo_VM_MOTION_A = "00:25:B5:" + $siteId + $podId + ":3A:FF" #### 256 mac addresses

##### mac pool MAC-IP_STORAGE-A
$macPoolName_IP_STORAGE_A = "MAC-IP_STORAGE-A"
$macPoolDescr_IP_STORAGE_A = "MAC Pool used for IP Storage Port Group - Fabric A" + "; " + $siteDescr + " " + $podDescr
$macPoolFrom_IP_STORAGE_A = "00:25:B5:" + $siteId + $podId + ":4A:00"
$macPoolTo_IP_STORAGE_A = "00:25:B5:" + $siteId + $podId + ":4A:FF" #### 256 mac addresses

##### mac pool MAC-MGMT-B
$macPoolName_MGMT_B = "MAC-MGMT-B"
$macPoolDescr_MGMT_B = "MAC Pool used for Management Port Group - Fabric B" + "; " + $siteDescr + " " + $podDescr
$macPoolFrom_MGMT_B = "00:25:B5:" + $siteId + $podId + ":1B:00"
$macPoolTo_MGMT_B = "00:25:B5:" + $siteId + $podId + ":1B:FF" #### 256 mac addresses

##### mac pool MAC-VM_DATA-B
$macPoolName_VM_DATA_B = "MAC-VM_DATA-B"
$macPoolDescr_VM_DATA_B = "MAC Pool used for Guest Virtual Machine (DATA) Port Group - Fabric B" + "; " + $siteDescr + " " + $podDescr
$macPoolFrom_VM_DATA_B = "00:25:B5:" + $siteId + $podId + ":2B:00"
$macPoolTo_VM_DATA_B = "00:25:B5:" + $siteId + $podId + ":2B:FF" #### 256 mac addresses

##### mac pool MAC-VM_MOTION-B
$macPoolName_VM_MOTION_B = "MAC-VM_MOTION-B"
$macPoolDescr_VM_MOTION_B = "MAC Pool used for VM_MOTION Port Group - Fabric B" + "; " + $siteDescr + " " + $podDescr
$macPoolFrom_VM_MOTION_B = "00:25:B5:" + $siteId + $podId + ":3B:00"
$macPoolTo_VM_MOTION_B = "00:25:B5:" + $siteId + $podId + ":3B:FF" #### 256 mac addresses

##### mac pool MAC-IP_STORAGE-B
$macPoolName_IP_STORAGE_B = "MAC-IP_STORAGE-B"
$macPoolDescr_IP_STORAGE_B = "MAC Pool used for IP Storage Port Group - Fabric B" + "; " + $siteDescr + " " + $podDescr
$macPoolFrom_IP_STORAGE_B = "00:25:B5:" + $siteId + $podId + ":4B:00"
$macPoolTo_IP_STORAGE_B = "00:25:B5:" + $siteId + $podId + ":4B:FF" #### 256 mac addresses

#### nic templates
#### Fabric A Templates
$vNicTemplate_a_mgmt_name = "MGMT-A"
$vNicTemplate_a_mgmt_descr = "Management vNIC - Fabric A"
$vNicTemplate_a_data_name = "VM_DATA-A"
$vNicTemplate_a_data_descr = "Client Virtual Machine Data vNIC - Fabric A"
$vNicTemplate_a_vmMotion_name = "VM_MOTION-A"
$vNicTemplate_a_vmMotion_descr = "VM_MOTION vNIC - Fabric A"
$vNicTemplate_a_ipStorage_name = "IP_STORAGE-A"
$vNicTemplate_a_ipStorage_descr = "IP Storage vNIC - Fabric A"

#### Fabric B Templates
$vNicTemplate_b_mgmt_name = "MGMT-B"
$vNicTemplate_b_mgmt_descr = "Management vNIC - Fabric B"
$vNicTemplate_b_data_name = "VM_DATA-B"
$vNicTemplate_b_data_descr = "Client Virtual Machine Data vNIC - Fabric B"
$vNicTemplate_b_vmMotion_name = "VM_MOTION-B"
$vNicTemplate_b_vmMotion_descr = "VM_MOTION vNIC - Fabric B"
$vNicTemplate_b_ipStorage_name = "IP_STORAGE-B"
$vNicTemplate_b_ipStorage_descr = "IP Storage vNIC - Fabric B"

#### wwnn pool
$wwnnPoolName = "CitrixXD_site_" + $siteId + "_pod_" + $podId
$wwnnPoolDescr = $siteDescr + " " + $podDescr
$wwnnPoolFrom = "20:00:00:25:B5:" + $siteId + $podId + ":00:00"
$wwnnPoolTo = "20:00:00:25:B5:" + $siteId + $podId + ":03:E7" #### 1000 wwnn addresses

#### wwpn pools
$wwpn_a1_name = "vHBA-A1_site_" + $siteId + "_pod_" + $podId
$wwpn_a1_descr = "WWPN Pool for Adapter 1 - Fabric A " + $siteDescr + " " + $podDescr
$wwpn_a1_from = "20:00:00:25:B5:" + $siteId + $podId + ":A1:00"
$wwpn_a1_to = "20:00:00:25:B5:" + $siteId + $podId + ":A1:FF" #### 256 wwpn addresses on fab_a
$wwpn_a2_name = "vHBA-A2_site_" + $siteId + "_pod_" + $podId
$wwpn_a2_descr = "WWPN Pool for Adapter 2 - Fabric A " + $siteDescr + " " + $podDescr
$wwpn_a2_from = "20:00:00:25:B5:" + $siteId + $podId + ":A2:00"
$wwpn_a2_to = "20:00:00:25:B5:" + $siteId + $podId + ":A2:FF" #### 256 wwpn addresses on fab_a
$wwpn_b1_name = "vHBA-B1_site_" + $siteId + "_pod_" + $podId
$wwpn_b1_descr = "WWPN Pool for Adapter 1 - Fabric B " + $siteDescr + " " + $podDescr
$wwpn_b1_from = "20:00:00:25:B5:" + $siteId + $podId + ":B1:00"
$wwpn_b1_to = "20:00:00:25:B5:" + $siteId + $podId + ":B1:FF" #### 256 wwpn addresses on fab_b
$wwpn_b2_name = "vHBA-B2_site_" + $siteId + "_pod_" + $podId
$wwpn_b2_descr = "WWPN Pool for Adapter 2 - Fabric A " + $siteDescr + " " + $podDescr
$wwpn_b2_from = "20:00:00:25:B5:" + $siteId + $podId + ":B2:00"
$wwpn_b2_to = "20:00:00:25:B5:" + $siteId + $podId + ":B2:FF" #### 256 wwpn addresses on fab_b

#### hba templates
$vHbaTemplate_a1_name = "vHBA-A1"
$vHbaTemplate_a1_descr = "vHBA-A1 - Fabric A " + $siteDescr + " " + $podDescr
$vHbaTemplate_a2_name = "vHBA-A2"
$vHbaTemplate_a2_descr = "vHBA-A2 - Fabric A " + $siteDescr + " " + $podDescr
$vHbaTemplate_b1_name = "vHBA-B1"
$vHbaTemplate_b1_descr = "vHBA-B1 - Fabric B " + $siteDescr + " " + $podDescr
$vHbaTemplate_b2_name = "vHBA-B2"
$vHbaTemplate_b2_descr = "vHBA-B2 - Fabric B " + $siteDescr + " " + $podDescr

##### polices
#### bios policy
$biosPolicy_name = "CitrixXD_Host"

#### network control policy
$networkControlPolicyName = "cdp-on_link-down"

#### local disk configuration
$localDiskPolicy_name = "XDNoLocal"
$localDiskPolicy_descr = "Citrix XenDesktop Local Disk Configuration Policy - No Local Disk"

#### server pool
$serverPoolName = "CitrixXD_Server_Pool"
$serverPoolPolicyName = "CitrixXD"

#### boot policy
$bootPolicy_name = "SAN_boot"
$bootPolicy_descr = $siteDescr + " " + $podDescr + " SAN Boot"
$san_Primary_hba = "fc0"
$san_secondary_hba = "fc1"

#### maintenance policy
$maintenancePolicy_name = "usr_ack"
$maintenancePolicy_descr = "User Acknowledge"

#### firmware policies
$hostFirmwarePackagePolicy_name = "CitrixXD"
$hostFirmwarePackagePolicy_descr = "Citrix XenDesktop Host Firmware Package"
$mgmtFirmwarePackagePolicy_name = "CitrixXD"
$mgmtFirmwarePackagePolicy_descr = "Citrix XenDesktop Management Firmware Package"

##### Service Profile templates
$serviceProfileTemplate_descr = $siteDescr + " " + $podDescr
$serviceProfileTemplate_vHba0_name = "fc0"
$serviceProfileTemplate_vHba1_name = "fc1"
$serviceProfileTemplate_vNic0_name = "MGMT-A"
$serviceProfileTemplate_vNic1_name = "MGMT-B"
$serviceProfileTemplate_vNic2_name = "VM_DATA-A"
$serviceProfileTemplate_vNic3_name = "VM_DATA-B"
$serviceProfileTemplate_vNic4_name = "IP_STORAGE-A"
$serviceProfileTemplate_vNic5_name = "IP_STORAGE-B"
$serviceProfileTemplate_vNic6_name = "VM_MOTION-A"
$serviceProfileTemplate_vNic7_name = "VM_MOTION-B"

<#
#### Worksheet Variables
Out-FileOrConsole $serviceProfileTemplate_name
Out-FileOrConsole $serviceProfileTemplate_descr
Out-FileOrConsole $ucsIp
Out-FileOrConsole $siteId
Out-FileOrConsole $siteDescr
Out-FileOrConsole $podId
Out-FileOrConsole $podDescr
Out-FileOrConsole $vlan_mgmt_name
Out-FileOrConsole $vlan_mgmt_id
Out-FileOrConsole $vlan_mgmt_descr
Out-FileOrConsole $vlan_vmdata_name
Out-FileOrConsole $vlan_vmdata_id
Out-FileOrConsole $vlan_vmdata_descr
Out-FileOrConsole $vlan_vmotion_name
Out-FileOrConsole $vlan_vmotion_id
Out-FileOrConsole $vlan_vmotion_descr
Out-FileOrConsole $vlan_ipstorage_name
Out-FileOrConsole $vlan_ipstorage_id
Out-FileOrConsole $vlan_ipstorage_descr
Out-FileOrConsole $vSan_a_name
Out-FileOrConsole $vSan_a_id
Out-FileOrConsole $FCoE_Vlan_a
Out-FileOrConsole $vSan_b_name
Out-FileOrConsole $vSan_b_id
Out-FileOrConsole $FCoE_Vlan_b
Out-FileOrConsole $san_primary_target_primary
Out-FileOrConsole $san_primary_target_secondary
Out-FileOrConsole $san_secondary_target_primary
Out-FileOrConsole $san_secondary_target_secondary
Out-FileOrConsole $mgmtIpFrom
Out-FileOrConsole $mgmtIpTo
Out-FileOrConsole $mgmtIpSubmask
Out-FileOrConsole $mgmtIpDefgw

#### Derived Variables
Out-FileOrConsole $uuidName
Out-FileOrConsole $uuidDescr
Out-FileOrConsole $uuidFrom
Out-FileOrConsole $uuidTo
Out-FileOrConsole $networkControlPolicyName
Out-FileOrConsole $localDiskPolicy_name
Out-FileOrConsole $localDiskPolicy_descr
Out-FileOrConsole $serverPoolName
Out-FileOrConsole $serverPoolPolicyName
Out-FileOrConsole $macPoolName_MGMT_A
Out-FileOrConsole $macPoolDescr_MGMT_A
Out-FileOrConsole $macPoolFrom_MGMT_A
Out-FileOrConsole $macPoolTo_MGMT_A
Out-FileOrConsole $macPoolName_VM_DATA_A
Out-FileOrConsole $macPoolDescr_VM_DATA_A
Out-FileOrConsole $macPoolFrom_VM_DATA_A
Out-FileOrConsole $macPoolTo_VM_DATA_A
Out-FileOrConsole $macPoolName_VM_MOTION_A
Out-FileOrConsole $macPoolDescr_VM_MOTION_A
Out-FileOrConsole $macPoolFrom_VM_MOTION_A
Out-FileOrConsole $macPoolTo_VM_MOTION_A
Out-FileOrConsole $macPoolName_IP_STORAGE_A
Out-FileOrConsole $macPoolDescr_IP_STORAGE_A
Out-FileOrConsole $macPoolFrom_IP_STORAGE_A
Out-FileOrConsole $macPoolTo_IP_STORAGE_A
Out-FileOrConsole $macPoolName_MGMT_B
Out-FileOrConsole $macPoolDescr_MGMT_B
Out-FileOrConsole $macPoolFrom_MGMT_B
Out-FileOrConsole $macPoolTo_MGMT_B
Out-FileOrConsole $macPoolName_VM_DATA_B
Out-FileOrConsole $macPoolDescr_VM_DATA_B
Out-FileOrConsole $macPoolFrom_VM_DATA_B
Out-FileOrConsole $macPoolTo_VM_DATA_B
Out-FileOrConsole $macPoolName_VM_MOTION_B
Out-FileOrConsole $macPoolDescr_VM_MOTION_B
Out-FileOrConsole $macPoolFrom_VM_MOTION_B
Out-FileOrConsole $macPoolTo_VM_MOTION_B
Out-FileOrConsole $macPoolName_IP_STORAGE_B
Out-FileOrConsole $macPoolDescr_IP_STORAGE_B
Out-FileOrConsole $macPoolFrom_IP_STORAGE_B
Out-FileOrConsole $macPoolTo_IP_STORAGE_B
Out-FileOrConsole $vNicTemplate_a_mgmt_name
Out-FileOrConsole $vNicTemplate_a_mgmt_descr
Out-FileOrConsole $vNicTemplate_a_data_name
Out-FileOrConsole $vNicTemplate_a_data_descr
Out-FileOrConsole $vNicTemplate_a_vmMotion_name
Out-FileOrConsole $vNicTemplate_a_vmMotion_descr
Out-FileOrConsole $vNicTemplate_a_ipStorage_name
Out-FileOrConsole $vNicTemplate_a_ipStorage_descr
Out-FileOrConsole $vNicTemplate_b_mgmt_name
Out-FileOrConsole $vNicTemplate_b_mgmt_descr
Out-FileOrConsole $vNicTemplate_b_data_name
Out-FileOrConsole $vNicTemplate_b_data_descr
Out-FileOrConsole $vNicTemplate_b_vmMotion_name
Out-FileOrConsole $vNicTemplate_b_vmMotion_descr
Out-FileOrConsole $vNicTemplate_b_ipStorage_name
Out-FileOrConsole $vNicTemplate_b_ipStorage_descr
Out-FileOrConsole $wwnnPoolName
Out-FileOrConsole $wwnnPoolDescr
Out-FileOrConsole $wwnnPoolFrom
Out-FileOrConsole $wwnnPoolTo
Out-FileOrConsole $wwpn_a1_name
Out-FileOrConsole $wwpn_a1_descr
Out-FileOrConsole $wwpn_a1_from
Out-FileOrConsole $wwpn_a1_to
Out-FileOrConsole $wwpn_a2_name
Out-FileOrConsole $wwpn_a2_descr
Out-FileOrConsole $wwpn_a2_from
Out-FileOrConsole $wwpn_a2_to
Out-FileOrConsole $wwpn_b1_name
Out-FileOrConsole $wwpn_b1_descr
Out-FileOrConsole $wwpn_b1_from
Out-FileOrConsole $wwpn_b1_to
Out-FileOrConsole $wwpn_b2_name
Out-FileOrConsole $wwpn_b2_descr
Out-FileOrConsole $wwpn_b2_from
Out-FileOrConsole $wwpn_b2_to
Out-FileOrConsole $vHbaTemplate_a1_name
Out-FileOrConsole $vHbaTemplate_a1_descr
Out-FileOrConsole $vHbaTemplate_a2_name
Out-FileOrConsole $vHbaTemplate_a2_descr
Out-FileOrConsole $vHbaTemplate_b1_name
Out-FileOrConsole $vHbaTemplate_b1_descr
Out-FileOrConsole $vHbaTemplate_b2_name
Out-FileOrConsole $vHbaTemplate_b2_descr
Out-FileOrConsole $biosPolicy_name
Out-FileOrConsole $bootPolicy_name
Out-FileOrConsole $bootPolicy_descr
Out-FileOrConsole $san_Primary_hba
Out-FileOrConsole $san_secondary_hba
Out-FileOrConsole $maintenancePolicy_name
Out-FileOrConsole $maintenancePolicy_descr
Out-FileOrConsole $hostFirmwarePackagePolicy_name
Out-FileOrConsole $hostFirmwarePackagePolicy_descr
Out-FileOrConsole $mgmtFirmwarePackagePolicy_name
Out-FileOrConsole $mgmtFirmwarePackagePolicy_descr
Out-FileOrConsole $serviceProfileTemplate_vHba0_name
Out-FileOrConsole $serviceProfileTemplate_vHba1_name
Out-FileOrConsole $serviceProfileTemplate_vNic0_name
Out-FileOrConsole $serviceProfileTemplate_vNic1_name
Out-FileOrConsole $serviceProfileTemplate_vNic2_name
Out-FileOrConsole $serviceProfileTemplate_vNic3_name
Out-FileOrConsole $serviceProfileTemplate_vNic4_name
Out-FileOrConsole $serviceProfileTemplate_vNic5_name
Out-FileOrConsole $serviceProfileTemplate_vNic6_name
Out-FileOrConsole $serviceProfileTemplate_vNic7_name
#>

##### connecting to ucsm
Out-FileOrConsole "Connect to ucsm..."
$ucs = Connect-Ucs $ucsIp

if (!$ucs)
{ Out-FileOrConsole "..Cannot connect to $ucsIp. Quit the script."; exit(3)}

<#
#####optional - remove default Pools, Policies, Profiles, and Templates
Out-FileOrConsole "Remove default Server, UUID, WWNN, WWPN and MAC Pools..."
Get-UcsServerPool -Ucs $ucs -Name default -LimitScope | Remove-UcsServerPool -Force | Out-Null
Get-UcsUuidSuffixPool -Ucs $ucs -Name default -LimitScope | Remove-UcsUuidSuffixPool -Force | Out-Null
Get-UcsWwnPool -Ucs $ucs -Name node-default -LimitScope | Remove-UcsWwnPool -Force | Out-Null
Get-UcsWwnPool -Ucs $ucs -Name default -LimitScope | Remove-UcsWwnPool -Force | Out-Null
Get-UcsMacPool -Ucs $ucs -Name default -LimitScope | Remove-UcsMacPool -Force | Out-Null
Get-UcsIqnPoolPool -Ucs $ucs -Name default -LimitScope | Remove-UcsIqnPoolPool -Force | Out-Null
#### following are only used with UCS Platform Emulator
Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "11" -LimitScope | Remove-UcsServiceProfile -Force | Out-Null
Get-UcsOrg -Level root | Get-UcsOrg -Name "Finance" -LimitScope | Remove-UcsOrg -Force | Out-Null
Get-UcsOrg -Level root | Get-UcsServerPool -Name "blade-pool-2" -LimitScope | Remove-UcsServerPool -Force | Out-Null
Get-UcsFiLanCloud -Id "A" | Get-UcsVlan -Name "default" | Remove-UcsVlan -Force | Out-Null
Get-UcsFiLanCloud -Id "A" | Get-UcsVlan -Name "finance" | Remove-UcsVlan -Force | Out-Null
Get-UcsFiLanCloud -Id "A" | Get-UcsVlan -Name "human-resource" | Remove-UcsVlan -Force | Out-Null
Get-UcsFiLanCloud -Id "B" | Get-UcsVlan -Name "finance" | Remove-UcsVlan -Force | Out-Null
Get-UcsFiLanCloud -Id "B" | Get-UcsVlan -Name "default" | Remove-UcsVlan -Force | Out-Null
Get-UcsFiLanCloud -Id "B" | Get-UcsVlan -Name "human-resource" | Remove-UcsVlan -Force | Out-Null
Get-UcsOrg -Level root | Get-UcsQosPolicy -Name "qos-1" -LimitScope | Remove-UcsQosPolicy -Force | Out-Null
Get-UcsOrg -Level root | Get-UcsMacPool -Name "mac-pool-1" -LimitScope | Remove-UcsMacPool -Force | Out-Null
#>

##### start configuration
Out-FileOrConsole "Create Management IP Pool..."
Get-UcsIpPool -Ucs $ucs -Name ext-mgmt -LimitScope | Add-UcsIpPoolBlock -From $mgmtIpFrom -To $mgmtIpTo -DefGw $mgmtIpDefgw | Out-Null

Out-FileOrConsole "Create UUID Pool..."
$uuid = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsUuidSuffixPool -Name $uuidName -Descr $uuidDescr
$uuid | Add-UcsUuidSuffixBlock -From $uuidFrom -To $uuidTo | Out-Null

Out-FileOrConsole "Create MAC Pool - MGMT_A"
$macPool = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsMacPool -Name $macPoolName_MGMT_A -Descr $macPoolDescr_MGMT_A
$macPool | Add-UcsMacMemberBlock -From $macPoolFrom_MGMT_A -To $macPoolTo_MGMT_A | Out-Null
Out-FileOrConsole "Create MAC Pool - MGMT_B"
$macPool = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsMacPool -Name $macPoolName_MGMT_B -Descr $macPoolDescr_MGMT_B
$macPool | Add-UcsMacMemberBlock -From $macPoolFrom_MGMT_B -To $macPoolTo_MGMT_B | Out-Null

Out-FileOrConsole "Create MAC Pool - VM_DATA_A"
$macPool = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsMacPool -Name $macPoolName_VM_DATA_A -Descr $macPoolDescr_VM_DATA_A
$macPool | Add-UcsMacMemberBlock -From $macPoolFrom_VM_DATA_A -To $macPoolTo_VM_DATA_A | Out-Null
Out-FileOrConsole "Create MAC Pool - VM_DATA_B"
$macPool = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsMacPool -Name $macPoolName_VM_DATA_B -Descr $macPoolDescr_VM_DATA_B
$macPool | Add-UcsMacMemberBlock -From $macPoolFrom_VM_DATA_B -To $macPoolTo_VM_DATA_B | Out-Null

Out-FileOrConsole "Create MAC Pool - VM_MOTION_A"
$macPool = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsMacPool -Name $macPoolName_VM_MOTION_A -Descr $macPoolDescr_VM_MOTION_A
$macPool | Add-UcsMacMemberBlock -From $macPoolFrom_VM_MOTION_A -To $macPoolTo_VM_MOTION_A | Out-Null
Out-FileOrConsole "Create MAC Pool - VM_MOTION_B"
$macPool = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsMacPool -Name $macPoolName_VM_MOTION_B -Descr $macPoolDescr_VM_MOTION_B
$macPool | Add-UcsMacMemberBlock -From $macPoolFrom_VM_MOTION_B -To $macPoolTo_VM_MOTION_B | Out-Null

Out-FileOrConsole "Create MAC Pool - IP_STORAGE_A"
$macPool = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsMacPool -Name $macPoolName_IP_STORAGE_A -Descr $macPoolDescr_IP_STORAGE_A
$macPool | Add-UcsMacMemberBlock -From $macPoolFrom_IP_STORAGE_A -To $macPoolTo_IP_STORAGE_A | Out-Null
Out-FileOrConsole "Create MAC Pool - IP_STORAGE_B"
$macPool = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsMacPool -Name $macPoolName_IP_STORAGE_B -Descr $macPoolDescr_IP_STORAGE_B
$macPool | Add-UcsMacMemberBlock -From $macPoolFrom_IP_STORAGE_B -To $macPoolTo_IP_STORAGE_B | Out-Null

Out-FileOrConsole "Set QoS System Class values"
$temp = Get-UcsQosclassDefinition | Set-UcsManagedObject -PropertyMap @{Descr=""; } -force
$QoS_besteffort = Get-UcsBestEffortQosClass | Set-UcsBestEffortQosClass -Mtu "normal" -MulticastOptimize "no" -Name "" -Weight "3" -force
$QoS_bronze = Get-UcsQosClass -Priority "bronze" | Set-UcsQosClass -AdminState "enabled" -Cos "1" -Drop "drop" -Mtu "normal" -MulticastOptimize "no" -Name "" -Weight "best-effort" -force
$QoS_gold = Get-UcsQosClass -Priority "gold" | Set-UcsQosClass -AdminState "enabled" -Cos "4" -Drop "drop" -Mtu "9216" -MulticastOptimize "no" -Name "" -Weight "3" -force
$QoS_silver = Get-UcsQosClass -Priority "silver" | Set-UcsQosClass -AdminState "enabled" -Cos "2" -Drop "drop" -Mtu "9216" -MulticastOptimize "no" -Name "" -Weight "best-effort" -force
$QoS_fc = Get-UcsFcQosClass -Priority "fc" | Set-UcsFcQosClass -Cos "3" -Name "" -Weight "2" -force

Out-FileOrConsole "Create QoS Policies..."
$QoS_IP_STORAGE = Get-UcsOrg -Level root  | Add-UcsQosPolicy -Descr "" -Name $vlan_ipstorage_name
$temp = $QoS_IP_STORAGE | Add-UcsVnicEgressPolicy -ModifyPresent -Burst 10240 -HostControl "none" -Name "" -Prio "gold" -Rate "line-rate"
$QoS_VM_MOTION = Get-UcsOrg -Level root  | Add-UcsQosPolicy -Descr "" -Name $vlan_vmotion_name
$temp = $QoS_VM_MOTION | Add-UcsVnicEgressPolicy -ModifyPresent -Burst 10240 -HostControl "none" -Name "" -Prio "silver" -Rate "line-rate"
$QoS_MGMT = Get-UcsOrg -Level root  | Add-UcsQosPolicy -Descr "" -Name $vlan_mgmt_name
$temp = $QoS_MGMT | Add-UcsVnicEgressPolicy -ModifyPresent -Burst 10240 -HostControl "none" -Name "" -Prio "bronze" -Rate "1000000"
$QoS_VM_DATA = Get-UcsOrg -Level root  | Add-UcsQosPolicy -Descr "" -Name $vlan_vmdata_name
$temp = $QoS_VM_DATA | Add-UcsVnicEgressPolicy -ModifyPresent -Burst 10240 -HostControl "none" -Name "" -Prio "best-effort" -Rate "line-rate"

Out-FileOrConsole "Create Network Control Policy..."
$networkControlPolicy_param = @{
	Name = $networkControlPolicyName
	Cdp = "enabled"
	MacRegisterMode = "only-native-vlan"
	UplinkFailAction = "link-down"
}
$networkControlPolicy = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsNetworkControlPolicy @networkControlPolicy_param
$networkControlPolicy | Add-UcsPortSecurityConfig -ModifyPresent -Forge allow | Out-Null

Out-FileOrConsole "Create VLANs..."
$vlan_mgmt = Get-UcsLanCloud -Ucs $ucs | Add-UcsVlan -DefaultNet no -Id $vlan_mgmt_id -Name $vlan_mgmt_name
$vlan_vmdata = Get-UcsLanCloud -Ucs $ucs | Add-UcsVlan -DefaultNet no -Id $vlan_vmdata_id -Name $vlan_vmdata_name
$vlan_vmotion = Get-UcsLanCloud -Ucs $ucs | Add-UcsVlan -DefaultNet no -Id $vlan_vmotion_id -Name $vlan_vmotion_name
$vlan_ipstorage = Get-UcsLanCloud -Ucs $ucs | Add-UcsVlan -DefaultNet no -Id $vlan_ipstorage_id -Name $vlan_ipstorage_name

Out-FileOrConsole "Create vNIC templates for management..."
$vNicTemplate_mgmt_a_params = @{
	Name = $vNicTemplate_a_mgmt_name
	Descr = $vNicTemplate_a_mgmt_descr
	IdentPoolName = $macPoolName_MGMT_A
	Target = "adaptor"
	NwCtrlPolicyName = $networkControlPolicyName
	SwitchId = "A"
    QosPolicyName = $vlan_mgmt_name
	TemplType = "updating-template"
}
$vNicTemplate_mgmt_a = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVnicTemplate @vNicTemplate_mgmt_a_params
$vNicTemplate_mgmt_a | Add-UcsVnicInterface -ModifyPresent -DefaultNet yes -Name $vlan_mgmt_name | Out-Null

$vNicTemplate_mgmt_b_params = @{
	Name = $vNicTemplate_b_mgmt_name
	Descr = $vNicTemplate_b_mgmt_descr
	IdentPoolName = $macPoolName_MGMT_B
	Target = "adaptor"
	NwCtrlPolicyName = $networkControlPolicyName
	SwitchId = "B"
    QoSPolicyName = $vlan_mgmt_name
	TemplType = "updating-template"
}
$vNicTemplate_mgmt_b = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVnicTemplate @vNicTemplate_mgmt_b_params
$vNicTemplate_mgmt_b | Add-UcsVnicInterface -ModifyPresent -DefaultNet yes -Name $vlan_mgmt_name | Out-Null

Out-FileOrConsole "Create vNIC templates for VM Data..."
$vNicTemplate_data_a_params = @{
	Name = $vNicTemplate_a_data_name
	Descr = $vNicTemplate_a_data_descr
	IdentPoolName = $macPoolName_VM_DATA_A
	Target = "adaptor"
	NwCtrlPolicyName = $networkControlPolicyName
	SwitchId = "A"
    QosPolicyName = $vlan_vmdata_name
	TemplType = "updating-template"
}
$vNicTemplate_data_a = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVnicTemplate @vNicTemplate_data_a_params
$vNicTemplate_data_a | Add-UcsVnicInterface -ModifyPresent -DefaultNet no -Name $vlan_vmdata_name | Out-Null

$vNicTemplate_data_b_params = @{
	Name = $vNicTemplate_b_data_name
	Descr = $vNicTemplate_b_data_descr
	IdentPoolName = $macPoolName_VM_DATA_B
	Target = "adaptor"
	NwCtrlPolicyName = $networkControlPolicyName
	SwitchId = "B"
    QosPolicyName = $vlan_vmdata_name
	TemplType = "updating-template"
}
$vNicTemplate_data_b = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVnicTemplate @vNicTemplate_data_b_params
$vNicTemplate_data_b | Add-UcsVnicInterface -ModifyPresent -DefaultNet no -Name $vlan_vmdata_name | Out-Null

Out-FileOrConsole "Create vNIC templates for VM_Motion..."
$vNicTemplate_vmotion_a_params = @{
	Name = $vNicTemplate_a_vmMotion_name
	Descr = $vNicTemplate_a_vmMotion_descr
	IdentPoolName = $macPoolName_VM_MOTION_A
	Target = "adaptor"
	NwCtrlPolicyName = $networkControlPolicyName
	SwitchId = "A"
    QosPolicyName = $vlan_vmotion_name
    Mtu = "9000"
	TemplType = "updating-template"
}
$vNicTemplate_vmotion_a = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVnicTemplate @vNicTemplate_vmotion_a_params
$vNicTemplate_vmotion_a | Add-UcsVnicInterface -ModifyPresent -DefaultNet no -Name $vlan_vmotion_name | Out-Null

$vNicTemplate_vmotion_b_params = @{
	Name = $vNicTemplate_b_vmMotion_name
	Descr = $vNicTemplate_b_vmMotion_descr
	IdentPoolName = $macPoolName_VM_MOTION_B
	Target = "adaptor"
	NwCtrlPolicyName = $networkControlPolicyName
	SwitchId = "B"
    QosPolicyName = $vlan_vmotion_name
    Mtu = "9000"
	TemplType = "updating-template"
}
$vNicTemplate_vmotion_b = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVnicTemplate @vNicTemplate_vmotion_b_params
$vNicTemplate_vmotion_b | Add-UcsVnicInterface -ModifyPresent -DefaultNet no -Name $vlan_vmotion_name | Out-Null

Out-FileOrConsole "Create vNIC templates for IP_STORAGE..."
$vNicTemplate_ipstorage_a_params = @{
	Name = $vNicTemplate_a_ipStorage_name
	Descr = $vNicTemplate_a_ipStorage_descr
	IdentPoolName = $macPoolName_IP_STORAGE_A
	Target = "adaptor"
	NwCtrlPolicyName = $networkControlPolicyName
	SwitchId = "A"
    QosPolicyName = $vlan_ipstorage_name
    Mtu = "9000"
	TemplType = "updating-template"
}
$vNicTemplate_ipstorage_a = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVnicTemplate @vNicTemplate_ipstorage_a_params
$vNicTemplate_ipstorage_a | Add-UcsVnicInterface -ModifyPresent -DefaultNet no -Name $vlan_ipstorage_name | Out-Null

$vNicTemplate_ipstorage_b_params = @{
	Name = $vNicTemplate_b_ipStorage_name
	Descr = $vNicTemplate_b_ipStorage_descr
	IdentPoolName = $macPoolName_IP_STORAGE_B
	Target = "adaptor"
	NwCtrlPolicyName = $networkControlPolicyName
	SwitchId = "B"
    QosPolicyName = $vlan_ipstorage_name
    Mtu = "9000"
	TemplType = "updating-template"
}
$vNicTemplate_ipstorage_b = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVnicTemplate @vNicTemplate_ipstorage_b_params
$vNicTemplate_ipstorage_b | Add-UcsVnicInterface -ModifyPresent -DefaultNet no -Name $vlan_ipstorage_name | Out-Null

Out-FileOrConsole "Create WWNN Pool..."
$wwnnPool = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsWwnPool -Name $wwnnPoolName -Descr $wwnnPoolDescr -Purpose node-wwn-assignment
$wwnnPool | Add-UcsWwnMemberBlock -From $wwnnPoolFrom -To $wwnnPoolTo | Out-Null

Out-FileOrConsole "Create WWPN Pools..."
$wwpnPoolA = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsWwnPool -Name $wwpn_a1_name -Descr $wwpn_a1_descr -Purpose port-wwn-assignment
$wwpnPoolA | Add-UcsWwnMemberBlock -From $wwpn_a1_from -To $wwpn_a1_to | Out-Null
$wwpnPoolA = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsWwnPool -Name $wwpn_a2_name -Descr $wwpn_a2_descr -Purpose port-wwn-assignment
$wwpnPoolA | Add-UcsWwnMemberBlock -From $wwpn_a2_from -To $wwpn_a2_to | Out-Null

$wwpnPoolB = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsWwnPool -Name $wwpn_b1_name -Descr $wwpn_b1_descr -Purpose port-wwn-assignment
$wwpnPoolB | Add-UcsWwnMemberBlock -From $wwpn_b1_from -To $wwpn_b1_to | Out-Null
$wwpnPoolB = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsWwnPool -Name $wwpn_b2_name -Descr $wwpn_b2_descr -Purpose port-wwn-assignment
$wwpnPoolB | Add-UcsWwnMemberBlock -From $wwpn_b2_from -To $wwpn_b2_to | Out-Null

Out-FileOrConsole "Set up VSAN A & B..."
Get-UcsFiSanCloud -Ucs $ucs -Id A | Add-UcsVsan -FcoeVlan $FCoE_Vlan_a -Id $vSan_a_id -Name $vSan_a_name | Out-Null
Get-UcsFiSanCloud -Ucs $ucs -Id B | Add-UcsVsan -FcoeVlan $FCoE_Vlan_b -Id $vSan_b_id -Name $vSan_b_name | Out-Null

Out-FileOrConsole "Create vHBA Templates..."
$vHbaTemplate_a1_params = @{
	Name = $vHbaTemplate_a1_name
	Descr = $vHbaTemplate_a1_descr
	IdentPoolName = $wwpn_a1_name
	SwitchId = "A"
	TemplType = "updating-template"
}
$vHbaTemplate_a1 = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVhbaTemplate @vHbaTemplate_a1_params
$vHbaTemplate_a1 | Add-UcsVhbaInterface -ModifyPresent -Name $vSan_a_name | Out-Null

$vHbaTemplate_a2_params = @{
	Name = $vHbaTemplate_a2_name
	Descr = $vHbaTemplate_a2_descr
	IdentPoolName = $wwpn_a2_name
	SwitchId = "A"
	TemplType = "updating-template"
}
$vHbaTemplate_a2 = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVhbaTemplate @vHbaTemplate_a2_params
$vHbaTemplate_a2 | Add-UcsVhbaInterface -ModifyPresent -Name $vSan_a_name | Out-Null

$vHbaTemplate_b1_params = @{
	Name = $vHbaTemplate_b1_name
	Descr = $vHbaTemplate_b1_descr
	IdentPoolName = $wwpn_b1_name
	SwitchId = "B"
	TemplType = "updating-template"
}
$vHbaTemplate_b1 = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVhbaTemplate @vHbaTemplate_b1_params
$vHbaTemplate_b1 | Add-UcsVhbaInterface -ModifyPresent -Name $vSan_b_name | Out-Null

$vHbaTemplate_b2_params = @{
	Name = $vHbaTemplate_b2_name
	Descr = $vHbaTemplate_b2_descr
	IdentPoolName = $wwpn_b2_name
	SwitchId = "B"
	TemplType = "updating-template"
}
$vHbaTemplate_b2 = Get-UcsOrg -Ucs $ucs -Level root -LimitScope | Add-UcsVhbaTemplate @vHbaTemplate_b2_params
$vHbaTemplate_b2 | Add-UcsVhbaInterface -ModifyPresent -Name $vSan_b_name | Out-Null

Out-FileOrConsole "Create SAN Boot policy..."
$bootPolicy = Get-UcsOrg -Level root -LimitScope | Add-UcsBootPolicy -Name $bootPolicy_name -Descr $bootPolicy_descr -RebootOnUpdate no
$bootPolicy_storage = $bootPolicy | Add-UcsLsbootStorage -Order 1
$bootPolicy_storage_primary = $bootPolicy_storage | Add-UcsLsbootSanImage -Type primary -VnicName $san_Primary_hba
$bootPolicy_storage_primary | Add-UcsLsbootSanImagePath -Type primary -Wwn $san_primary_target_primary | Out-Null
$bootPolicy_storage_primary | Add-UcsLsbootSanImagePath -Type secondary -Wwn $san_primary_target_secondary | Out-Null
$bootPolicy_storage_secondary = $bootPolicy_storage | Add-UcsLsbootSanImage -Type secondary -VnicName $san_secondary_hba
$bootPolicy_storage_secondary | Add-UcsLsbootSanImagePath -Type primary -Wwn $san_secondary_target_primary | Out-Null
$bootPolicy_storage_secondary | Add-UcsLsbootSanImagePath -Type secondary -Wwn $san_secondary_target_secondary | Out-Null

Out-FileOrConsole "Create Server Pool and Server Pool policy..."
$serverPool = Get-UcsOrg -Level root  | Add-UcsServerPool -Descr "Citrix XenDesktop Server Pool" -Name $serverPoolName
$serverPoolPolicy = Get-UcsOrg -Level root  | Add-UcsServerPoolPolicy -Descr "Citrix XenDesktop Server Pool Policy" -Name $serverPoolPolicyName -PoolDn "org-root/compute-pool-CitrixXD_Server_Pool" -Qualifier ""

Out-FileOrConsole "Create BIOS policy..."
$mo = Get-UcsOrg -Level root  | Add-UcsBiosPolicy -Descr "" -Name $biosPolicy_name -RebootOnUpdate "no"
$mo_1 = $mo | Set-UcsBiosVfCPUPerformance -VpCPUPerformance "enterprise" -force
$mo_2 = $mo | Set-UcsBiosVfCoreMultiProcessing -VpCoreMultiProcessing "all" -force
$mo_3 = $mo | Set-UcsBiosVfDirectCacheAccess -VpDirectCacheAccess "enabled" -force
$mo_4 = $mo | Set-UcsBiosEnhancedIntelSpeedStep -VpEnhancedIntelSpeedStepTech "disabled" -force
$mo_5 = $mo | Set-UcsBiosExecuteDisabledBit -VpExecuteDisableBit "enabled" -force
$mo_6 = $mo | Set-UcsBiosHyperThreading -VpIntelHyperThreadingTech "enabled" -force
$mo_7 = $mo | Set-UcsBiosTurboBoost -VpIntelTurboBoostTech "disabled" -force
$mo_8 = $mo | Set-UcsBiosIntelDirectedIO -VpIntelVTForDirectedIO "enabled" -force
$mo_9 = $mo | Set-UcsBiosVfIntelVirtualizationTechnology -VpIntelVirtualizationTechnology "enabled" -force
$mo_10 = $mo | Set-UcsBiosNUMA -VpNUMAOptimized "enabled" -force
$mo_11 = $mo | Set-UcsBiosVfProcessorCState -VpProcessorCState "disabled" -force
$mo_12 = $mo | Set-UcsBiosVfProcessorC1E -VpProcessorC1E "disabled" -force
$mo_13 = $mo | Set-UcsBiosVfProcessorC3Report -VpProcessorC3Report "disabled" -force
$mo_14 = $mo | Set-UcsBiosVfProcessorC6Report -VpProcessorC6Report "disabled" -force
$mo_15 = $mo | Set-UcsBiosVfProcessorC7Report -VpProcessorC7Report "disabled" -force
$mo_16 = $mo | Set-UcsBiosVfQuietBoot -VpQuietBoot "disabled" -force
$mo_17 = $mo | Set-UcsBiosVfSelectMemoryRASConfiguration -VpSelectMemoryRASConfiguration "maximum-performance" -force

Out-FileOrConsole "Create Firmware policies..."
Get-UcsOrg -Level root -LimitScope | Add-UcsFirmwareComputeHostPack -Name $hostFirmwarePackagePolicy_name -Descr $hostFirmwarePackagePolicy_descr | Out-Null
Get-UcsOrg -Level root -LimitScope | Add-UcsFirmwareComputeMgmtPack -Name $mgmtFirmwarePackagePolicy_name -Descr $mgmtFirmwarePackagePolicy_descr | Out-Null

Out-FileOrConsole "Create Maintenance policy..."
Get-UcsOrg -Level root -LimitScope | Add-UcsMaintenancePolicy -Name $maintenancePolicy_name -Descr $maintenancePolicy_descr -UptimeDisr user-ack | Out-Null

Out-FileOrConsole "Create Local Disk Configuration Policy"
$localDiskPolicy = Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr $localDiskPolicy_descr -Mode "no-local-storage" -Name $localDiskPolicy_name -ProtectConfig "yes"

Out-FileOrConsole "Create Service Profile Template..."
$serviceProfileTemplate_params = @{
	BiosProfileName = $biosPolicy_name
	BootPolicyName = $bootPolicy_name
	Name = $serviceProfileTemplate_name
	Descr = $serviceProfileTemplate_descr
	ExtIpState = "pooled"
	HostFwPolicyName = $hostFirmwarePackagePolicy_name
	IdentPoolName = $uuidName
	MaintPolicyName = $maintenancePolicy_name
	MgmtFwPolicyName = $mgmtFirmwarePackagePolicy_name
    LocalDiskPolicyName = $localDiskPolicy_name
	Type = "updating-template"
}
$serviceProfileTemplate = Get-UcsOrg -Level root -LimitScope | Add-UcsServiceProfile @serviceProfileTemplate_params

Out-FileOrConsole "..Add vNICs to Service Profile Template..."
$mo_1 = $serviceProfileTemplate | Add-UcsVnic -Addr "derived" -AdminVcon "any" -Name $serviceProfileTemplate_vNic0_name -NwTemplName $vNicTemplate_a_mgmt_name -Order "3"
$mo_2 = $serviceProfileTemplate | Add-UcsVnic -Addr "derived" -AdminVcon "any" -Name $serviceProfileTemplate_vNic1_name -NwTemplName $vNicTemplate_b_mgmt_name -Order "4"
$mo_3 = $serviceProfileTemplate | Add-UcsVnic -Addr "derived" -AdminVcon "any" -Name $serviceProfileTemplate_vNic2_name -NwTemplName $vNicTemplate_a_data_name -Order "5"
$mo_4 = $serviceProfileTemplate | Add-UcsVnic -Addr "derived" -AdminVcon "any" -Name $serviceProfileTemplate_vNic3_name -NwTemplName $vNicTemplate_b_data_name -Order "6"
$mo_5 = $serviceProfileTemplate | Add-UcsVnic -Addr "derived" -AdminVcon "any" -Name $serviceProfileTemplate_vNic4_name -NwTemplName $vNicTemplate_a_ipStorage_name -Order "7"
$mo_6 = $serviceProfileTemplate | Add-UcsVnic -Addr "derived" -AdminVcon "any" -Name $serviceProfileTemplate_vNic5_name -NwTemplName $vNicTemplate_b_ipStorage_name -Order "8"
$mo_7 = $serviceProfileTemplate | Add-UcsVnic -Addr "derived" -AdminVcon "any" -Name $serviceProfileTemplate_vNic6_name -NwTemplName $vNicTemplate_a_vmMotion_name -Order "9"
$mo_8 = $serviceProfileTemplate | Add-UcsVnic -Addr "derived" -AdminVcon "any" -Name $serviceProfileTemplate_vNic7_name -NwTemplName $vNicTemplate_b_vmMotion_name -Order "10"

Out-FileOrConsole "..Add vHBAs to Service Profile Template..."
$mo_9 = $serviceProfileTemplate | Add-UcsVhba -Addr "derived" -AdminVcon "any" -Name $serviceProfileTemplate_vHba0_name -NwTemplName $vHbaTemplate_a1_name -Order "1"
$mo_10 = $serviceProfileTemplate | Add-UcsVhba -Addr "derived" -AdminVcon "any" -Name $serviceProfileTemplate_vHba1_name -NwTemplName $vHbaTemplate_b1_name -Order "2"
$mo_11 = $serviceProfileTemplate | Add-UcsVnicFcNode -ModifyPresent -Addr "pool-derived" -IdentPoolName $wwnnPoolName | Out-Null

$mo_12 = $serviceProfileTemplate | Add-UcsServerPoolAssignment -ModifyPresent -Name $serverPoolName -RestrictMigration "no"
$mo_13 = $serviceProfileTemplate | Set-UcsServerPower -State "admin-down" -force

Out-FileOrConsole "Disconnecting from $($ucs.ucs)"
Disconnect-Ucs -Ucs $ucs

##### End of script logging
$end = Get-Date
$diff = New-TimeSpan $start $end
Out-FileOrConsole "It took $($diff.Hours) hour(s), $($diff.Minutes) minute(s) and $($diff.seconds) second(s) to run the script."
Out-FileOrConsole "Stopping script logging."
# SIG # Begin signature block
# MIIYnwYJKoZIhvcNAQcCoIIYkDCCGIwCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU935akjrXO07Jkp/4UOLoijJl
# 5figghSHMIIDnzCCAoegAwIBAgIQeaKlhfnRFUIT2bg+9raN7TANBgkqhkiG9w0B
# AQUFADBTMQswCQYDVQQGEwJVUzEXMBUGA1UEChMOVmVyaVNpZ24sIEluYy4xKzAp
# BgNVBAMTIlZlcmlTaWduIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EwHhcNMTIw
# NTAxMDAwMDAwWhcNMTIxMjMxMjM1OTU5WjBiMQswCQYDVQQGEwJVUzEdMBsGA1UE
# ChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xNDAyBgNVBAMTK1N5bWFudGVjIFRpbWUg
# U3RhbXBpbmcgU2VydmljZXMgU2lnbmVyIC0gRzMwgZ8wDQYJKoZIhvcNAQEBBQAD
# gY0AMIGJAoGBAKlZZnTaPYp9etj89YBEe/5HahRVTlBHC+zT7c72OPdPabmx8LZ4
# ggqMdhZn4gKttw2livYD/GbT/AgtzLVzWXuJ3DNuZlpeUje0YtGSWTUUi0WsWbJN
# JKKYlGhCcp86aOJri54iLfSYTprGr7PkoKs8KL8j4ddypPIQU2eud69RAgMBAAGj
# geMwgeAwDAYDVR0TAQH/BAIwADAzBgNVHR8ELDAqMCigJqAkhiJodHRwOi8vY3Js
# LnZlcmlzaWduLmNvbS90c3MtY2EuY3JsMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
# MDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AudmVyaXNp
# Z24uY29tMA4GA1UdDwEB/wQEAwIHgDAeBgNVHREEFzAVpBMwETEPMA0GA1UEAxMG
# VFNBMS0zMB0GA1UdDgQWBBS0t/GJSSZg52Xqc67c0zjNv1eSbzANBgkqhkiG9w0B
# AQUFAAOCAQEAHpiqJ7d4tQi1yXJtt9/ADpimNcSIydL2bfFLGvvV+S2ZAJ7R55uL
# 4T+9OYAMZs0HvFyYVKaUuhDRTour9W9lzGcJooB8UugOA9ZresYFGOzIrEJ8Byyn
# PQhm3ADt/ZQdc/JymJOxEdaP747qrPSWUQzQjd8xUk9er32nSnXmTs4rnykr589d
# nwN+bid7I61iKWavkugszr2cf9zNFzxDwgk/dUXHnuTXYH+XxuSqx2n1/M10rCyw
# SMFQTnBWHrU1046+se2svf4M7IV91buFZkQZXZ+T64K6Y57TfGH/yBvZI1h/MKNm
# oTkmXpLDPMs3Mvr1o43c1bCj6SU2VdeB+jCCA8QwggMtoAMCAQICEEe/GZXfjVJG
# Q/fbbUgNMaQwDQYJKoZIhvcNAQEFBQAwgYsxCzAJBgNVBAYTAlpBMRUwEwYDVQQI
# EwxXZXN0ZXJuIENhcGUxFDASBgNVBAcTC0R1cmJhbnZpbGxlMQ8wDQYDVQQKEwZU
# aGF3dGUxHTAbBgNVBAsTFFRoYXd0ZSBDZXJ0aWZpY2F0aW9uMR8wHQYDVQQDExZU
# aGF3dGUgVGltZXN0YW1waW5nIENBMB4XDTAzMTIwNDAwMDAwMFoXDTEzMTIwMzIz
# NTk1OVowUzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMSsw
# KQYDVQQDEyJWZXJpU2lnbiBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAqcqypMzNIK8KfYmsh3XwtE7x38EP
# v2dhvaNkHNq7+cozq4QwiVh+jNtr3TaeD7/R7Hjyd6Z+bzy/k68Numj0bJTKvVIt
# q0g99bbVXV8bAp/6L2sepPejmqYayALhf0xS4w5g7EAcfrkN3j/HtN+HvV96ajEu
# A5mBE6hHIM4xcw1XLc14NDOVEpkSud5oL6rm48KKjCrDiyGHZr2DWFdvdb88qiaH
# XcoQFTyfhOpUwQpuxP7FSt25BxGXInzbPifRHnjsnzHJ8eYiGdvEs0dDmhpfoB6Q
# 5F717nzxfatiAY/1TQve0CJWqJXNroh2ru66DfPkTdmg+2igrhQ7s4fBuwIDAQAB
# o4HbMIHYMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3Au
# dmVyaXNpZ24uY29tMBIGA1UdEwEB/wQIMAYBAf8CAQAwQQYDVR0fBDowODA2oDSg
# MoYwaHR0cDovL2NybC52ZXJpc2lnbi5jb20vVGhhd3RlVGltZXN0YW1waW5nQ0Eu
# Y3JsMBMGA1UdJQQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIBBjAkBgNVHREE
# HTAbpBkwFzEVMBMGA1UEAxMMVFNBMjA0OC0xLTUzMA0GCSqGSIb3DQEBBQUAA4GB
# AEpr+epYwkQcMYl5mSuWv4KsAdYcTM2wilhu3wgpo17IypMT5wRSDe9HJy8AOLDk
# yZNOmtQiYhX3PzchT3AxgPGLOIez6OiXAP7PVZZOJNKpJ056rrdhQfMqzufJ2V7d
# uyuFPrWdtdnhV/++tMV+9c8MnvCX/ivTO1IbGzgn9z9KMIIELzCCAxegAwIBAgIQ
# cUepDtq4dUao7/FgM/2etTANBgkqhkiG9w0BAQUFADBKMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMVGhhd3RlLCBJbmMuMSQwIgYDVQQDExtUaGF3dGUgQ29kZSBTaWdu
# aW5nIENBIC0gRzIwHhcNMTIwMjAxMDAwMDAwWhcNMTQwMzMwMjM1OTU5WjCBhDEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCkNhbGlmb3JuaWExETAPBgNVBAcTCFNhbiBK
# b3NlMRYwFAYDVQQKFA1DaXNjbyBTeXN0ZW1zMR0wGwYDVQQLFBRJTkZPUk1BVElP
# TiBTRUNVUklUWTEWMBQGA1UEAxQNQ2lzY28gU3lzdGVtczCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALv1yCSp03CrQZ7oflWAT3vDGTrfQcjTuByIET1S
# QmQwUzyh+i1VBkRO9BRw+k4hyfGxQGqa1aEMrFGvM0tLO/XHAAsQMzKnOjEHx6OR
# GcCypjfm1ou0UGsE0RX7WGNeW0b7bFNhdQ+GZ7+cv2J1xRyl5pvNcS5Ffi2YlLoI
# vlE3nWUKz3uW09s0Peebw+BjFd+6UcG02Nx88/XkVNBtH2qXjRmFIe9XfkvesGhV
# JCHO15SmKWt4ZJsEVxCgTgUYN/aGn6+SCwTP1tEXogzMLTPdwnf5lFbdNhzGRKau
# LL2rewzY1kQlwQ57E2Sj6ssSTHxt4uK4MJp8KL8qSfyvz/MCAwEAAaOB1TCB0jAM
# BgNVHRMBAf8EAjAAMDsGA1UdHwQ0MDIwMKAuoCyGKmh0dHA6Ly9jcy1nMi1jcmwu
# dGhhd3RlLmNvbS9UaGF3dGVDU0cyLmNybDAfBgNVHSUEGDAWBggrBgEFBQcDAwYK
# KwYBBAGCNwIBFjAdBgNVHQQEFjAUMA4wDAYKKwYBBAGCNwIBFgMCB4AwMgYIKwYB
# BQUHAQEEJjAkMCIGCCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBEG
# CWCGSAGG+EIBAQQEAwIEEDANBgkqhkiG9w0BAQUFAAOCAQEAMPGSFU2LPUHvJpbV
# TtL7BE0ZnvsX4rbqd6lCU5WboYesUBVB5pu9Yc/9ZMZisSeFqc9+81+RrnCu+c7C
# 3ikiUjvJkFXNWArTFEimKzVdj9GVg615RNOYlu1h0PUoCluhHZBeAqyLeKgLoRiQ
# 4enH05E1edL4bP6dnYlSqOZqP3/Hd4vWQl1Pcvc/4+SiwMwKsco5rXU8wEFcHeLx
# juljzGjOit0rtLuUhMwwNKb5yGpgPAJMtOgoxQ5pfL9IbZHhisaweZrPnBWJlvgS
# Jv5rZjr44NJ3rdsozCcdHx3SCXge1flOVlq3o2zEbdIq9PXqca8PleryF0xV856j
# 9BWOfjCCBEUwggOuoAMCAQICEDNlUAh5rXPiMLngHQ1/rJEwDQYJKoZIhvcNAQEF
# BQAwgc4xCzAJBgNVBAYTAlpBMRUwEwYDVQQIEwxXZXN0ZXJuIENhcGUxEjAQBgNV
# BAcTCUNhcGUgVG93bjEdMBsGA1UEChMUVGhhd3RlIENvbnN1bHRpbmcgY2MxKDAm
# BgNVBAsTH0NlcnRpZmljYXRpb24gU2VydmljZXMgRGl2aXNpb24xITAfBgNVBAMT
# GFRoYXd0ZSBQcmVtaXVtIFNlcnZlciBDQTEoMCYGCSqGSIb3DQEJARYZcHJlbWl1
# bS1zZXJ2ZXJAdGhhd3RlLmNvbTAeFw0wNjExMTcwMDAwMDBaFw0yMDEyMzAyMzU5
# NTlaMIGpMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMdGhhd3RlLCBJbmMuMSgwJgYD
# VQQLEx9DZXJ0aWZpY2F0aW9uIFNlcnZpY2VzIERpdmlzaW9uMTgwNgYDVQQLEy8o
# YykgMjAwNiB0aGF3dGUsIEluYy4gLSBGb3IgYXV0aG9yaXplZCB1c2Ugb25seTEf
# MB0GA1UEAxMWdGhhd3RlIFByaW1hcnkgUm9vdCBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAKyg8PuAWdScx6TPnaFZcwkQRQwNLG5o8WxbSGhJWTf8
# CzMZwnd/zBAtlTQc5utNCacc0rjJlzYCt4nUJF8GwMxElJSNAmJv61rdEY0omlyE
# kBB6Db10Zi9qOKDi1VRE6x0Hnwe6b+7p/U4LKfU+hKAB8Zyr+Bx+iaToodhxZQ2j
# UXvuvNIiYA25W53fuvxRWwuvmLLpLukE6GKH3ivI107BTGQe3c+HWLpKT8poBx0c
# nUrG1S+RzHxxchzFwGfrMv3JklyU2oXAm79TfSsJ9IydkR+XalLL3gk2pHfYe4dQ
# RNU+bilp+zlJJh4JpYB7QC3r6CeFyf5h/X7mfJcd1Z0CAwEAAaOBwjCBvzAPBgNV
# HRMBAf8EBTADAQH/MDsGA1UdIAQ0MDIwMAYEVR0gADAoMCYGCCsGAQUFBwIBFhpo
# dHRwczovL3d3dy50aGF3dGUuY29tL2NwczAOBgNVHQ8BAf8EBAMCAQYwHQYDVR0O
# BBYEFHtbRc+vzst6/TGSGmq280brV0hQMEAGA1UdHwQ5MDcwNaAzoDGGL2h0dHA6
# Ly9jcmwudGhhd3RlLmNvbS9UaGF3dGVQcmVtaXVtU2VydmVyQ0EuY3JsMA0GCSqG
# SIb3DQEBBQUAA4GBAISoTMk+Krya4syPC7Ild8RhiYljWtSjFUDU+14/tEPqYxcr
# a5l0ngmo3dRWFS56eTFfY5ZTGzTZFepPbXDKvvaCqe3ahXfMdhxqgQoh2EGZf14u
# gsHoqveTgQWqkrQft5rABxf1y8a0TA7XVtxxIHQ41nTG1o9rr4uNoGwpC2HgMIIE
# nDCCA4SgAwIBAgIQR5dNeHOlvKsNL7NwGS/OXjANBgkqhkiG9w0BAQUFADCBqTEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDHRoYXd0ZSwgSW5jLjEoMCYGA1UECxMfQ2Vy
# dGlmaWNhdGlvbiBTZXJ2aWNlcyBEaXZpc2lvbjE4MDYGA1UECxMvKGMpIDIwMDYg
# dGhhd3RlLCBJbmMuIC0gRm9yIGF1dGhvcml6ZWQgdXNlIG9ubHkxHzAdBgNVBAMT
# FnRoYXd0ZSBQcmltYXJ5IFJvb3QgQ0EwHhcNMTAwMjA4MDAwMDAwWhcNMjAwMjA3
# MjM1OTU5WjBKMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMVGhhd3RlLCBJbmMuMSQw
# IgYDVQQDExtUaGF3dGUgQ29kZSBTaWduaW5nIENBIC0gRzIwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQC3i891W58l2n45sJPbONOpI9CC+ukkflwLjoP4
# 5npZ5qPFmKeZ0kT/AKalOQSK2imI6tui8xyZFSbCsfT84QxHqQkRBgogkrnHoASM
# XJQZq1slLB1ifnANzmFs3SuCyc5dSF/3wr68QSMeTyld10+89MUq/GPmfCZOmad5
# QZ4QSnp5ycaG94aV0ibOPBgq1nzOr82tu/eCLHAmN0XlD0cixgEovS6DXGqkR8Hn
# 0NhrgUY/IRf1B8VDWqZnLLh7YBG1g+71dApycUQ9WP7oGqs4w1nbf244fXbHcmmY
# NpZX02Yc0lSRBC5UGbDcPbUiXobVKn4g313merFl/sUCTjEtAgMBAAGjggEcMIIB
# GDASBgNVHRMBAf8ECDAGAQH/AgEAMDQGA1UdHwQtMCswKaAnoCWGI2h0dHA6Ly9j
# cmwudGhhd3RlLmNvbS9UaGF3dGVQQ0EuY3JsMA4GA1UdDwEB/wQEAwIBBjAyBggr
# BgEFBQcBAQQmMCQwIgYIKwYBBQUHMAGGFmh0dHA6Ly9vY3NwLnRoYXd0ZS5jb20w
# HQYDVR0lBBYwFAYIKwYBBQUHAwIGCCsGAQUFBwMDMCkGA1UdEQQiMCCkHjAcMRow
# GAYDVQQDExFWZXJpU2lnbk1QS0ktMi0xMDAdBgNVHQ4EFgQU1A1lP3q9NMb+R+dM
# DcC98t4Vq3EwHwYDVR0jBBgwFoAUe1tFz6/Oy3r9MZIaarbzRutXSFAwDQYJKoZI
# hvcNAQEFBQADggEBAFb+U1zhx568p+1+U21qFEtRjEBegF+qpOgv7zjIBMnKPs/f
# OlhOsNS2Y8UpV/oCBZpFTWjbKhvUND2fAMNay5VJpW7hsMX8QU1BSm/Td8jXOI3k
# Gd4Y8x8VZYNtRQxT+QqaLqVdv28ygRiSGWpVAK1jHFIGflXZKWiuSnwYmnmIayMj
# 2Cc4KimHdsr7x7ZiIx/telZM3ZwyW/U9DEYYlTsqI2iDZEHZAG0PGSQVaHK9xXFn
# bqxM25DrUaUaYgfQvmoARzxyL+xPYT5zhc5aCre6wBwTdeMiOSjdbR0JRp1PuuhA
# gZHGpM6UchsBzypuFWeVia59t7fN+Qo9dbZrPCUxggOCMIIDfgIBATBeMEoxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxUaGF3dGUsIEluYy4xJDAiBgNVBAMTG1RoYXd0
# ZSBDb2RlIFNpZ25pbmcgQ0EgLSBHMgIQcUepDtq4dUao7/FgM/2etTAJBgUrDgMC
# GgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYK
# KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFjAjBgkqhkiG
# 9w0BCQQxFgQUqeWFuEOclD2znwBDVoB73ncX7FIwDQYJKoZIhvcNAQEBBQAEggEA
# uReGbEgoPrD3MbcVKLUuMbzt8xIfSvqmqzcI1o/wFl+xzhiJmExM3peTAmcU7i8H
# YZ1Wkjh3/GDSTHg6HzVNm5xE2Y8m3oVBwylU3s+JYdUIIl5KTwf1DFDy4XWvta+0
# dlHmty3a3HnWQIx4t7H21EXHfwZwoKMhc/gF4we502uJ4t+TqhEEHJWk9kmZ5+2T
# EmCM0h/r1ZmLT6j0EGVR+K9fVXxVboLtPAU17Ac3fy855tk3e6n49YCpRzv6aHl7
# v3F1BLoGR8BtUEioB/dV2jceEYVJCMIyKeqtnl3CxcOs41CmQMHEn2hfHrCNirIv
# wT0MMlGpl1dneLUgSDooHaGCAX8wggF7BgkqhkiG9w0BCQYxggFsMIIBaAIBATBn
# MFMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5WZXJpU2lnbiwgSW5jLjErMCkGA1UE
# AxMiVmVyaVNpZ24gVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBDQQIQeaKlhfnRFUIT
# 2bg+9raN7TAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAc
# BgkqhkiG9w0BCQUxDxcNMTIxMDA1MTI1ODI5WjAjBgkqhkiG9w0BCQQxFgQUdhu0
# OOO3vlftTXZihrH+08iqSc8wDQYJKoZIhvcNAQEBBQAEgYBHF8L5LeYC/0JoxZ0I
# nIGUKk3irMolcmMQS4VrB3NCSxhA+Gpcuu8OcNDUNwE4DouO0kAJAEEx6aE+EImz
# ytxS0JGU++tL9BpYjVwsHgPDOZ1R7OPhugIg7qbDW42c8OKH131upR/LBJUdhMz4
# X1+wO1i8D15N9VxPUn9BOpF8eQ==
# SIG # End signature block
