# ========================================================
#
# 	Script Information
#
#	Title:				Remote Computer Inventory No GUI
#	Author:				Assaf Miron
#	Originally created:	21/06/2008
#	Last Updated:		20/01/2009 
#	Original path:		ComputerInventory.PS1
#	Description:		Collects Remote Computer Data Using WMI and Registry Access	
#						Outputs all information to a Data Grid Form and to a CSV Log File						
#	
# ========================================================

param ($ComputerFile = $(Read-Host "Enter a path to a text file containing Computer Names"),[bool] $OutCSV = $False, [bool] $OutDB = $True)

#region Definitions
# Get Script Location 
$ScriptLocation = Split-Path -Parent $MyInvocation.MyCommand.Path


# Log File where the results are Saved
$LogFile = $ScriptLocation+"\Test-Monitor.csv"
# Error Log File where dead or no data computers are saved
$ErrorLogFile = $ScriptLocation+"\ErrorLog.txt"
# Check to see if the Log File Directory exists
If((Test-Path ($LogFile.Substring(0,$logFile.LastIndexof("\")))) -eq $False)
{ 
	# Create The Directory
	New-Item ($LogFile.Substring(0,$logFile.LastIndexof("\"))) -Type Directory
}

# Define Connection String
$ConnectionString = "packet size=4096;data source=SQLSRV;persist security info=True;initial catalog=Inventory"

#~~< Ping1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Ping1 = New-Object System.Net.NetworkInformation.Ping

#endregion

#region Functions
function Check-Empty( $Object )
{
#	Input		: An Object with values.
#	Output		: A Trimmed String of the Object or '-' if it's Null.
#	Description : Check the object if its null or not and return it's value.
	If($Object -eq $null -or $Object -eq "")
	{
		return "-"
	}
	else
	{
		return $Object.ToString().Trim()
	}
}

function StoredProc{
# Recieve Paramters
	param(
			$SP_ComputerName,
			$SP_DomainName="-",
			$SP_MainOU="-",
			$SP_SystemType="-",
			$SP_Manufacturer="-",
			$SP_Model="-",
			$SP_Chassis="-",
			$SP_SN="-",
			$SP_NumProcessors="0",
			$SP_CPUName="-",
			$SP_MemSlotCount=0,
			$SP_MemSlotUsed="-",
			$SP_TotalMemory="-",
			$SP_AvMB="-",
			$SP_UserName="-",
			$SP_OS="-",
			$SP_SP="-",
			$SP_OSCDKey="-",
			$SP_IPAddress="-",
			$SP_MACAddress="-",
			$SP_HotFixes="-",
			$SP_DiskDrives="-",
			$SP_NetDrives="-",
			$SP_DiskFreeSpace="-",
			$SP_DiskSize="-",
			$SP_RDPStatus="-",
			$SP_RAStatus="-",
			$SP_AUClient="-",
			$SP_AVServer="-",
			$SP_AVDefs="-",
			$SP_Printers="-",
			$SP_ComputerTotalHealth="-"
	)
# Stored Procedure doesnt recieve Null Values
# Check the Null Valued Paramters
# Check If one of the Parameters are Null, and return the Value Trimmed
	$SP_ComputerName = Check-Empty $SP_ComputerName
	$SP_DomainName = Check-Empty $SP_DomainName
	$SP_MainOU = Check-Empty $SP_MainOU
	$SP_SystemType = Check-Empty $SP_SystemType
	$SP_Manufacturer = Check-Empty $SP_Manufacturer 
	$SP_Model = Check-Empty $SP_Model 
	$SP_Chassis = Check-Empty $SP_Chassis 
	$SP_SN = Check-Empty $SP_SN
	if( $SP_NumProcessors -eq $Null) { $SP_NumProcessors = 0 }
	$SP_CPUName = Check-Empty $SP_CPUName
	if( $SP_MemSlotCount -eq $Null) { $SP_MemSlotCount = 0 }
	$SP_MemSlotUsed = Check-Empty $SP_MemSlotUsed
	$SP_TotalMemory = Check-Empty $SP_TotalMemory
	$SP_AvMB = Check-Empty $SP_AvMB
	$SP_UserName = Check-Empty $SP_UserName
	$SP_OS = Check-Empty $SP_OS 
	$SP_SP = Check-Empty $SP_SP 
	$SP_OSCDKey = Check-Empty $SP_OSCDKey 
	$SP_IPAddress = Check-Empty $SP_IPAddress 
	$SP_MACAddress = Check-Empty $SP_MACAddress
	$SP_HotFixes = Check-Empty $SP_HotFixes
	$SP_DiskDrives = Check-Empty $SP_DiskDrives
	$SP_NetDrives = Check-Empty $SP_NetDrives
	$SP_DiskFreeSpace = Check-Empty $SP_DiskFreeSpace
	$SP_DiskSize = Check-Empty $SP_DiskSize
	$SP_RDPStatus = Check-Empty $SP_RDPStatus
	$SP_RAStatus = Check-Empty $SP_RAStatus
	$SP_AUClient = Check-Empty $SP_AUClient
	$SP_AVServer = Check-Empty $SP_AVServer
	$SP_AVDefs = Check-Empty $SP_AVDefs
	$SP_Printers = Check-Empty $SP_Printers
	$SP_ComputerTotalHealth = Check-Empty $SP_ComputerTotalHealth

	$cmd = New-Object System.Data.SqlClient.SqlCommand("InsertComputerInfo" ,$conn)
	$cmd.CommandType = [System.data.CommandType]'StoredProcedure'
	
	$cmd.Parameters.Add("@ComputerName", $SP_ComputerName) | Out-Null
	$cmd.Parameters.Add("@DomainName", $SP_DomainName) | Out-Null
	$cmd.Parameters.Add("@MainOU", $SP_MainOU) | Out-Null
	$cmd.Parameters.Add("@SystemType", $SP_SystemType) | Out-Null
	$cmd.Parameters.Add("@Manufacturer", $SP_Manufacturer) | Out-Null
	$cmd.Parameters.Add("@Model", $SP_Model) | Out-Null
	$cmd.Parameters.Add("@Chassis", $SP_Chassis) | Out-Null
	$cmd.Parameters.Add("@SN", $SP_SN) | Out-Null
	$cmd.Parameters.Add("@NumProcessors", $SP_NumProcessors) | Out-Null
	$cmd.Parameters.Add("@CPUName", $SP_CPUName) | Out-Null
	$cmd.Parameters.Add("@MemorySlotCount", $SP_MemSlotCount) | Out-Null
	$cmd.Parameters.Add("@MemorySlotUsed", $SP_MemSlotUsed) | Out-Null
	$cmd.Parameters.Add("@TotalMemory", $SP_TotalMemory) | Out-Null
	$cmd.Parameters.Add("@AvailableMemory", $SP_AvMB) | Out-Null
	$cmd.Parameters.Add("@UserName", $SP_UserName) | Out-Null
	$cmd.Parameters.Add("@OS", $SP_OS) | Out-Null
	$cmd.Parameters.Add("@SP", $SP_SP) | Out-Null
	$cmd.Parameters.Add("@OSCDKey", $SP_OSCDKey) | Out-Null
	$cmd.Parameters.Add("@IPAddress", $SP_IPAddress) | Out-Null
	$cmd.Parameters.Add("@MACAddress", $SP_MACAddress) | Out-Null
	$cmd.Parameters.Add("@HotFixes", $SP_HotFixes) | Out-Null
	$cmd.Parameters.Add("@DiskDrives", $SP_DiskDrives) | Out-Null
	$cmd.Parameters.Add("@NetworkDrives", $SP_NetDrives) | Out-Null
	$cmd.Parameters.Add("@DiskSize", $SP_DiskSize) | Out-Null
	$cmd.Parameters.Add("@DiskFreeSpace", $SP_DiskFreeSpace) | Out-Null
	$cmd.Parameters.Add("@RDPStatus", $SP_RDPStatus) | Out-Null
	$cmd.Parameters.Add("@RAStatus", $SP_RAStatus) | Out-Null
	$cmd.Parameters.Add("@AUClient", $SP_AUClient) | Out-Null
	$cmd.Parameters.Add("@AVServer", $SP_AVServer) | Out-Null
	$cmd.Parameters.Add("@AVDefs", $SP_AVDefs) | Out-Null
	$cmd.Parameters.Add("@Printers", $SP_Printers) | Out-Null
	$cmd.Parameters.Add("@ComputerTotalHealth", $SP_ComputerTotalHealth) | Out-Null
	if($conn.State -eq "Open"){
		$cmd.ExecuteNonQuery() | Out-Null
	}
	else {
		$conn.Open()
		$cmd.ExecuteNonQuery() | Out-Null
	}
}
#endregion

# Open a new Connection and Create a new Command Type
# Connect To DB
$conn = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
$conn.Open()

# Create an Array of Computers Enterd in the Input File
$arrComputers = Get-Content -path $ComputerFile -encoding UTF8

# Create an Empty Array to Contain all The Data of all Scaned Computers
$AllComputerInfo = @()
# Init counter for Progress
$counter = (1 * $arrComputers.count) / 100
Write-Host $arrComputers.count	
# Scan all Computers in the Array $arrComputers
foreach ($strComputer in $arrComputers)
{ 
	Write-Progress	"Scanning Computers..." "% Complete:" -PercentComplete	$counter
	# Uses the Ping Command to check if the Computer is Alive
	$Alive=""
	$Alive = $Ping1.Send($strComputer).Status 
	if($Alive -eq "Success")
	{
		# Querying $strComputer For Opend Port (port 135)
		$cmdPortQry = "& '$ScriptLocation\portqry.exe' -n $strComputer -e 135 | find "+[char]34+"TCP port 135 (epmap service):"+[char]34
		$PortQuery = Invoke-Expression $cmdPortQry
		If($PortQuery.split(":")[1].Trim() -eq "LISTENING") # Check Ports in dest computer
		{
			Write-Host "Scanning $strComputer For Hardware Data" 
			
			$PSCommand = "$ScriptLocation\Collect-Data.ps1"
			$PSPath = "C:\WINDOWS\system32\windowspowershell\v1.0\powershell.exe -noprofile "
			$PSCommand = $pspath+[char]34+". '"+$PSCommand+"' $strComputer"+[char]34 
			$DataObject = New-Object psobject
			Invoke-Expression  $PSCommand
			$DataObject = Import-Clixml -Path C:\CompDet.xml
			# Check if the Computer had Errors - No Caption means no results - Enter to Black List
			if($DataObject.Caption -eq $null)
			{
				# Write the computer name in the Error Log File
				$strComputer | Out-File -Append $ErrorLogFile
			}
			# Change the Notify Icon to Show Exporting Text
			Write-Host "Exporting $strComputer Information"
			if($OutDB)
			{
				#region Exporting data - Stored Procedure
				# if Information is Valid - Send to Stored Procedure
				StoredProc -SP_ComputerName $DataObject.Caption -SP_DomainName $DataObject.Domain -SP_SystemType $DataObject.SystemType`
				-SP_Manufacturer $DataObject.Manufacturer -SP_Model $DataObject.Model -SP_Chassis $DataObject.'Chassis Type' -SP_Printers $DataObject.Printers`
				-SP_SN $DataObject.SerialNumber -SP_NumProcessors ([int]$DataObject.NumberOfProcessors) -SP_CPUName $DataObject.'CPU Names'`
				-SP_TotalMemory $DataObject.TotalPhysicalMemory -SP_AvMB $DataObject.AvailableMem -SP_UserName $DataObject.UserName -SP_MainOU $DataObject.MainOU -SP_MemSlotCount ([int]$DataObject.MemoryDevices) -SP_MemSlotUsed $dataObject.MemSlots`
				-SP_OS $DataObject.'Operating System' -SP_SP $DataObject.'Service Pack' -SP_OSCDKey $DataObject.'CD-Key' -SP_IPAddress $DataObject.'IP Addresses' -SP_MACAddress $DataObject.'MAC Addresses' -SP_HotFixes $DataObject.HotFixes`
				-SP_DiskDrives $DataObject.'Disk Drives' -SP_NetDrives $DataObject.'Network Disks' -SP_DiskSize $DataObject.'Disk Size' -SP_DiskFreeSpace $DataObject.'Disk Free Space' -SP_RDPStatus $DataObject.'Remote Desktop' `
				-SP_RAStatus $DataObject.'Remote Assistance' -SP_AUClient $DataObject.'Automatic Updates' -SP_AVServer $DataObject.'Anti-Virus Server' -SP_AVDefs $DataObject.'Anti-Virus Defs' -SP_ComputerTotalHealth $DataObject.'Computer Total Health'
				#endregion
			}
			if($outCSV)
			{
				$AllComputerInfo += $DataObject
				$DataObject = "" # Clear Data Object
			}
			
			# Clean up - Delete File
			Remove-Item -Path C:\CompDet.xml
		}
		else # Computer behind Firewall
		{
			#region Get Computer Main OU
				# Create command to run
				$cmdOU = "Cscript.exe -nologo '$ScriptLocation\SearchComputers-ReturnADSPath.vbs' $strComputer"
				$MainOU = Invoke-Expression $cmdOU
				If($MainOU.Contains(","))
				{
					$MainOU = $MainOU.Split(",")[-4].Replace("OU=","")
				}
			#endregion

			Write-Warning "$strComputer behind Firewall.`nNo Data was Collected"
			# Write the computer name in the Error Log File
			$strComputer | Out-File -Append $ErrorLogFile
			if($OutDB)
			{			
				StoredProc -SP_ComputerName $strComputer -SP_ComputerTotalHealth "FireWalled" -SP_MainOU $MainOU
			}
			if($outCSV)
			{
				$AllComputerInfo += $DataObject
				$DataObject = "" # Clear Data Object
			}
		}
	}
	else # No Ping to Computer
	{ 
		#region Get Computer Main OU
			# Create command to run
			$cmdOU = "Cscript.exe -nologo '$ScriptLocation\SearchComputers-ReturnADSPath.vbs' $strComputer"
			$MainOU = Invoke-Expression $cmdOU
			If($MainOU.Contains(","))
			{
				$MainOU = $MainOU.Split(",")[-4].Replace("OU=","")
			}
		#endregion
		
		Write-Warning "No Ping to $strComputer.`nNo Data was Collected"
		# Write the computer name in the Error Log File
		$strComputer | Out-File -Append $ErrorLogFile
		if($OutDB)
		{		
			StoredProc -SP_ComputerName $strComputer -SP_ComputerTotalHealth "No Ping" -SP_MainOU $MainOU
		}
		if($outCSV)
		{
			$AllComputerInfo += $DataObject
			$DataObject = "" # Clear Data Object
		}		
	}
	
	# Incerment Counter
	$counter = ($counter + 1 * $arrComputers.count) / 100
}

#region Finising
if($outCSV)
{
	# Export all the Data to the Log File
	$AllComputerInfo | Export-Csv -Encoding OEM -Path $LogFile
}	
if($OutDB)
{
	# Closing Connections
	$conn.Close
}
#endregion
