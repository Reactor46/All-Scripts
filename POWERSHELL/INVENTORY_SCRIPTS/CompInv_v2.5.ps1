# ==============================================================================================
# 
# Microsoft PowerShell Source File -- Created with SAPIEN Technologies PrimalScript 2007
# 
# NAME: Server/Workstation Inventory (CompInv_v2.ps1)
# 
# AUTHOR: Jesse Hamrick
# DATE  : 2/25/2009
# Web	: www.PowerShellPro.com
# COMMENT: Script Inventories Computers and sends results to an excel file.
# 
# ==============================================================================================

# ==============================================================================================
# Functions Section
# ==============================================================================================
# Function Name 'WMILookup' - Gathers info using WMI and places results in Excel
# ==============================================================================================
Function WMILookup {
foreach ($StrComputer in $colComputers){
		$GenItems1 = gwmi Win32_ComputerSystem -Comp $StrComputer
		$GenItems2 = gwmi Win32_OperatingSystem -Comp $StrComputer
		$SysItems1 = gwmi Win32_BIOS -Comp $StrComputer
		$ProcItems1 = gwmi Win32_Processor -Comp $StrComputer
		$MemItems1 = [math]::Round((gwmi -Class Win32_ComputerSystem  -Comp $StrComputer).TotalPhysicalMemory/1GB)
        #$DiskItems = gwmi Win32_LogicalDisk -Comp $StrComputer | where{$_.DeviceID -eq "C:"} 
		$DiskItems = Get-DiskFree -Comp $StrComputer -HumanReadable
        $NetItems = gwmi Win32_NetworkAdapterConfiguration -Comp $StrComputer |	where{$_.IPEnabled -eq "True"}
		}
#Populate General Sheet with information
	foreach ($objItem in $GenItems1){
		$Sheet1.Cells.Item($intRow, 1) = $StrComputer
		$Sheet1.Cells.Item($intRow, 2) = $objItem.Manufacturer
		$Sheet1.Cells.Item($intRow, 3) = $objItem.Model
		$Sheet1.Cells.Item($intRow, 9) = [math]::Round($objItem.TotalPhysicalMemory/1GB)
        $Sheet1.Cells.Item($intRow, 10) = $objItem.NumberOfProcessors		
        }
#Populate OS Information
	foreach ($objItem in $GenItems2){
		$Sheet1.Cells.Item($intRow, 6) = $objItem.Caption
		$Sheet1.Cells.Item($intRow, 7) = $objItem.csdversion
        $Sheet1.Cells.Item($intRow, 8) = $objItem.OSArchitecture
		}
#Populate CPU Information
    foreach ($objItem in $ProcItems1){
		$Sheet1.Cells.Item($intRow, 13) = $objItem.Name
        $Sheet1.Cells.Item($intRow, 11) = $objItem.NumberofCores
        $Sheet1.Cells.Item($intRow, 12) = $objItem.NumberofLogicalProcessors
        }
#Populate Bios Information
    foreach ($objItem in $SysItems1){
		$Sheet1.Cells.Item($intRow, 5) = $objItem.SMBIOSbiosVersion
		$Sheet1.Cells.Item($intRow, 4) = $objItem.SerialNumber
		}
#Populate Disk Information
	foreach ($objItem in $DiskItems){
		#$Sheet1.Cells.Item($intRow, 18) = $objItem.DeviceID
		#$Sheet1.Cells.Item($intRow, 19) = [math]::Round($objItem.Size+" GB"/1GB)
        #$Sheet1.Cells.Item($intRow, 20) = [math]::Round($objItem.FreeSpace+" GB"/1GB)
		#$Sheet1.Cells.Item($intRow, 21) = [math]::Round($objItem.FreeSpace/1GB),4 * 100
        $Sheet1.Cells.Item($intRow, 18) = $objItem.Vol
		$Sheet1.Cells.Item($intRow, 19) = $objItem.Size
        $Sheet1.Cells.Item($intRow, 20) = $objItem.Avail
		$Sheet1.Cells.Item($intRow, 21) = $objItem.'Use%'
		}
#Populate Network
	foreach ($objItem in $NetItems){
		$Sheet1.Cells.Item($intRow, 14) = $objItem.Description
		$Sheet1.Cells.Item($intRow, 15) = $objItem.IPAddress
		$Sheet1.Cells.Item($intRow, 16) = $objItem.IPSubnet
		$Sheet1.Cells.Item($intRow, 17) = $objItem.DefaultIPGateway
		}
		
$intRow = $intRow + 1
#$intRowCPU = $intRowCPU + 1
#$intRowMem = $intRowMem + 1
#$intRowDisk = $intRowDisk + 1
#$intRowNet = $intRowNet + 1
}


# =============================================================================================
# Function Name 'ListComputers' - Enumerates ALL computer objects in AD
# ==============================================================================================
Function ListComputers {
$strCategory = "computer"

$objDomain = New-Object System.DirectoryServices.DirectoryEntry

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.Filter = ("(objectCategory=$strCategory)")

$colProplist = "name"
foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}

$colResults = $objSearcher.FindAll()

foreach ($objResult in $colResults)
    {$objComputer = $objResult.Properties; $objComputer.name}
}


# ========================================================================
# Function Name 'ListTextFile' - Enumerates Computer Names in a text file
# Create a text file and enter the names of each computer. One computer
# name per line. Supply the path to the text file when prompted.
# ========================================================================
Function ListTextFile {
	$strText = Read-Host "Enter the path for the text file"
	$colComputers = Get-Content $strText
}

# ========================================================================
# Function Name 'SingleEntry' - Enumerates Computer from user input
# ========================================================================
Function ManualEntry {
	$colComputers = Read-Host "Enter Computer Name or IP" 
}

# ==============================================================================================
# Script Body
# ==============================================================================================
$erroractionpreference = "SilentlyContinue"


#Gather info from user.
Write-Host "********************************" 	-ForegroundColor Green
Write-Host "Computer Inventory Script" 			-ForegroundColor Green
Write-Host "By: Jesse Hamrick" 					-ForegroundColor Green
Write-Host "Created: 04/15/2009" 				-ForegroundColor Green
Write-Host "Contact: www.PowerShellPro.com" 	-ForegroundColor Green
Write-Host "********************************" 	-ForegroundColor Green
Write-Host " "
Write-Host "Admin rights are required to enumerate information." 	-ForegroundColor Green
Write-Host "Would you like to use an alternative credential?"		-ForegroundColor Green
$credResponse = Read-Host "[Y] Yes, [N] No"
	If($CredResponse -eq "y"){$cred = Get-Credential DOMAIN\USER}
Write-Host " "
Write-Host "Which computer resources would you like in the report?"	-ForegroundColor Green
$strResponse = Read-Host "[1] All Domain Computers, [2] All Domain Servers, [3] Computer names from a File, [4] Choose a Computer manually"
If($strResponse -eq "1"){$colComputers = ListComputers | Sort-Object}
	elseif($strResponse -eq "2"){$colComputers = ListServers | Sort-Object}
	elseif($strResponse -eq "3"){. ListTextFile}
	elseif($strResponse -eq "4"){. ManualEntry}
	else{Write-Host "You did not supply a correct response, Please run script again." -foregroundColor Red}				
Write-Progress -Activity "Getting Inventory" -status "Running..." -id 1

#New Excel Application
$Excel = New-Object -Com Excel.Application
$Excel.visible = $True

# Create  worksheets
$Excel = $Excel.Workbooks.Add()

# Assign each worksheet to a variable and
# name the worksheet.
$Sheet1 = $Excel.Worksheets.Item(1)
$Sheet1.Name = "General"


#Create Heading for General Sheet
$Sheet1.Cells.Item(1,1) = "PC Name"
$Sheet1.Cells.Item(1,2) = "Manufacturer"
$Sheet1.Cells.Item(1,3) = "Model"
$Sheet1.Cells.Item(1,4) = "ServiceTag"
$Sheet1.Cells.Item(1,5) = "BIOS_Version"
$Sheet1.Cells.Item(1,6) = "Operating_System"
$Sheet1.Cells.Item(1,7) = "SP_Level"
$Sheet1.Cells.Item(1,8) = "Arch"
$Sheet1.Cells.Item(1,9) = "Memory_GB"
$Sheet1.Cells.Item(1,10) = "CPU_Count"
$Sheet1.Cells.Item(1,11) = "Cores"
$Sheet1.Cells.Item(1,12) = "Threads"
$Sheet1.Cells.Item(1,13) = "Processor(s)"
$Sheet1.Cells.Item(1,14) = "Active NIC"
$Sheet1.Cells.Item(1,15) = "IP_Address"
$Sheet1.Cells.Item(1,16) = "Subnet"
$Sheet1.Cells.Item(1,17) = "GW_Address"
$Sheet1.Cells.Item(1,18) = "HardDrive"
$Sheet1.Cells.Item(1,19) = "Size"
$Sheet1.Cells.Item(1,20) = "FreeSpace"
$Sheet1.Cells.Item(1,21) = "Precent Free"


$colSheets = ($Sheet1) #, $Sheet2, $Sheet3, $Sheet4, $Sheet5, $Sheet6)
foreach ($colorItem in $colSheets){
$intRow = 2
#$intRowCPU = 2
#$intRowMem = 2
#$intRowDisk = 2
#$intRowNet = 2
$WorkBook = $colorItem.UsedRange
$WorkBook.Interior.ColorIndex = 20
$WorkBook.Font.ColorIndex = 11
$WorkBook.Font.Bold = $True
}

If($credResponse -eq "y"){WMILookupCred}
Else{WMILookup}

#Auto Fit all sheets in the Workbook
foreach ($colorItem in $colSheets){
$WorkBook = $colorItem.UsedRange															
$WorkBook.EntireColumn.AutoFit()
#clear
}
Write-Host "*******************************" -ForegroundColor Green
Write-Host "The Report has been completed."  -ForeGroundColor Green
Write-Host "*******************************" -ForegroundColor Green
# ========================================================================
# END of Script
# ========================================================================