# ==============================================================================================
# 
# Microsoft PowerShell Source File -- Created with SAPIEN Technologies PrimalScript 2007
# Web	: www.PowerShellPro.com
# COMMENT: Script Inventories Computers and sends results to an excel file.
# 
# UPDATED: Added calls to collect Windows Product information
# ==============================================================================================

# ==============================================================================================
# Functions Section
# ==============================================================================================
# Function Name 'WMILookup' - Gathers info using WMI and places results in Excel
# ==============================================================================================
Function WMILookup {
$sysCount = $colComputers.Count
$foo = 0
foreach ($StrComputer in $colComputers){
        $Sheet1.Cells.Item($intRow, 1) = $StrComputer
        $foo++
        $percent = ($foo / $sysCount) * 100
        Write-Progress -Activity $strComputer -percentComplete $percent

		$GenItems1 = gwmi Win32_ComputerSystem -Comp $StrComputer
		$GenItems2 = gwmi Win32_OperatingSystem -Comp $StrComputer
		$SysItems1 = gwmi Win32_BIOS -Comp $StrComputer
		$SysItems2 = gwmi Win32_TimeZone -Comp $StrComputer
#		$SysItems3 = gwmi Win32_WmiSetting -Comp $StrComputer
#		$ProcItems1 = gwmi Win32_Processor -Comp $StrComputer
#		$MemItems1 = gwmi Win32_PhysicalMemory -Comp $StrComputer
#		$memItems2 = gwmi Win32_PhysicalMemoryArray -Comp $StrComputer
#		$DiskItems = gwmi Win32_LogicalDisk -Comp $StrComputer
#		$NetItems = gwmi Win32_NetworkAdapterConfiguration -Comp $StrComputer |`
					where{$_.IPEnabled -eq "True"}

        # Now extract data to build the registered Product Key
        $hklm = 2147483650
        $regPath = "Software\Microsoft\Windows NT\CurrentVersion"
        $regValue = "DigitalProductId"
        
        
        $windowsKey = $null
        $win32os = $null
        $wmi = [WMIClass]"\\$strComputer\root\default:stdRegProv"
        $data = $wmi.GetBinaryValue($hklm,$regPath,$regValue)
        $binArray = ($data.uValue)[52..66]
        $charsArray = "B","C","D","F","G","H","J","K","M","P","Q","R","T","V","W","X","Y","2","3","4","6","7","8","9"
        ## decrypt base24 encoded binary data
        For ($i = 24; $i -ge 0; $i--) {
            $k = 0
            For ($j = 14; $j -ge 0; $j--) {
                $k = $k * 256 -bxor $binArray[$j]
                $binArray[$j] = [math]::truncate($k / 24)
                $k = $k % 24
            }
            $windowsKey = $charsArray[$k] + $windowsKey
            If (($i % 5 -eq 0) -and ($i -ne 0)) {
                $windowsKey = "-" + $windowsKey
            }
        }
		
				
# Populate General Sheet(1) with information
	foreach ($objItem in $GenItems1){
		#$Sheet1.Cells.Item($intRow, 1) = $StrComputer
		Switch($objItem.DomainRole)
			{
			0{$Sheet1.Cells.Item($intRow, 2) = "Stand Alone Workstation"}
			1{$Sheet1.Cells.Item($intRow, 2) = "Member Workstation"}
			2{$Sheet1.Cells.Item($intRow, 2) = "Stand Alone Server"}
			3{$Sheet1.Cells.Item($intRow, 2) = "Member Server"}
			4{$Sheet1.Cells.Item($intRow, 2) = "Back-up Domain Controller"}
			5{$Sheet1.Cells.Item($intRow, 2) = "Primary Domain Controller"}
			default{"Undetermined"}
			}
		$Sheet1.Cells.Item($intRow, 3) = $objItem.Manufacturer
		$Sheet1.Cells.Item($intRow, 4) = $objItem.Model
# Field 5 will be filled in later
		$Sheet1.Cells.Item($intRow, 6) = $objItem.SystemType
		$Sheet1.Cells.Item($intRow, 7) = $objItem.NumberOfProcessors
		$Sheet1.Cells.Item($intRow, 8) = $objItem.TotalPhysicalMemory / 1024 / 1024
		}
	foreach ($objItem in $GenItems2){
		$Sheet1.Cells.Item($intRow, 9) = $objItem.Caption
		$Sheet1.Cells.Item($intRow, 10) = $objItem.csdversion
        $Sheet1.Cells.Item($intRow, 11) = $objItem.OSArchitecture
        $Sheet1.Cells.Item($intRow, 12) = $windowsKey
		}
			
#Populate Systems Sheet
	foreach ($objItem in $SysItems1){

		$Sheet1.Cells.Item($intRow, 5) = $objItem.SerialNumber
		}

		
$intRow = $intRow + 1
$intRowCPU = $intRowCPU + 1
}
}

# ==============================================================================================
# Function Name 'WMILookupCred'-Uses Alternative Credential-Gathers info using WMI.
# ==============================================================================================
Function WMILookupCred {
foreach ($StrComputer in $colComputers){
		$GenItems1 = gwmi Win32_ComputerSystem -Comp $StrComputer -Credential $cred
		$GenItems2 = gwmi Win32_OperatingSystem -Comp $StrComputer -Credential $cred
		$SysItems1 = gwmi Win32_BIOS -Comp $StrComputer -Credential $cred
#		$SysItems2 = gwmi Win32_TimeZone -Comp $StrComputer -Credential $cred
#		$SysItems3 = gwmi Win32_WmiSetting -Comp $StrComputer -Credential $cred
#		$ProcItems1 = gwmi Win32_Processor -Comp $StrComputer -Credential $cred
#		$MemItems1 = gwmi Win32_PhysicalMemory -Comp $StrComputer -Credential $cred
#		$memItems2 = gwmi Win32_PhysicalMemoryArray -Comp $StrComputer -Credential $cred
#		$DiskItems = gwmi Win32_LogicalDisk -Comp $StrComputer -Credential $cred
#		$NetItems = gwmi Win32_NetworkAdapterConfiguration -Comp $StrComputer -Credential $cred |`
					where{$_.IPEnabled -eq "True"}
		# Now extract data to build the registered Product Key
        $hklm = 2147483650
        $regPath = "Software\Microsoft\Windows NT\CurrentVersion"
        $regValue = "DigitalProductId"
        
        
        $windowsKey = $null
        $win32os = $null
        $wmi = [WMIClass]"\\$strComputer\root\default:stdRegProv"
        $data = $wmi.GetBinaryValue($hklm,$regPath,$regValue)
        $binArray = ($data.uValue)[52..66]
        $charsArray = "B","C","D","F","G","H","J","K","M","P","Q","R","T","V","W","X","Y","2","3","4","6","7","8","9"
        ## decrypt base24 encoded binary data
        For ($i = 24; $i -ge 0; $i--) {
            $k = 0
            For ($j = 14; $j -ge 0; $j--) {
                $k = $k * 256 -bxor $binArray[$j]
                $binArray[$j] = [math]::truncate($k / 24)
                $k = $k % 24
            }
            $windowsKey = $charsArray[$k] + $windowsKey
            If (($i % 5 -eq 0) -and ($i -ne 0)) {
                $windowsKey = "-" + $windowsKey
            }
        }
				
# Populate General Sheet(1) with information
	foreach ($objItem in $GenItems1){
		#$Sheet1.Cells.Item($intRow, 1) = $StrComputer
		Switch($objItem.DomainRole)
			{
			0{$Sheet1.Cells.Item($intRow, 2) = "Stand Alone Workstation"}
			1{$Sheet1.Cells.Item($intRow, 2) = "Member Workstation"}
			2{$Sheet1.Cells.Item($intRow, 2) = "Stand Alone Server"}
			3{$Sheet1.Cells.Item($intRow, 2) = "Member Server"}
			4{$Sheet1.Cells.Item($intRow, 2) = "Back-up Domain Controller"}
			5{$Sheet1.Cells.Item($intRow, 2) = "Primary Domain Controller"}
			default{"Undetermined"}
			}
		$Sheet1.Cells.Item($intRow, 3) = $objItem.Manufacturer
		$Sheet1.Cells.Item($intRow, 4) = $objItem.Model
# Field 5 will be filled in later
		$Sheet1.Cells.Item($intRow, 6) = $objItem.SystemType
		$Sheet1.Cells.Item($intRow, 7) = $objItem.NumberOfProcessors
		$Sheet1.Cells.Item($intRow, 8) = $objItem.TotalPhysicalMemory / 1024 / 1024
		}
	foreach ($objItem in $GenItems2){
		$Sheet1.Cells.Item($intRow, 9) = $objItem.Caption
		$Sheet1.Cells.Item($intRow, 10) = $objItem.csdversion
        $Sheet1.Cells.Item($intRow, 11) = $objItem.OSArchitecture
        $Sheet1.Cells.Item($intRow, 12) = $windowsKey
		}
			
#Populate Systems Sheet
	foreach ($objItem in $SysItems1){
		$Sheet1.Cells.Item($intRow, 5) = $objItem.SerialNumber
		}

		
$intRow = $intRow + 1
$intRowCPU = $intRowCPU + 1

}
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

# ==============================================================================================
# Function Name 'ListServers' - Enumerates ALL Servers objects in AD
# ==============================================================================================
Function ListServers {
$strCategory = "computer"
$strOS = "Windows*Server*"

$objDomain = New-Object System.DirectoryServices.DirectoryEntry

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.Filter = ("(&(objectCategory=$strCategory)(OperatingSystem=$strOS))")

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
Write-Host "********************************" 	-ForegroundColor DarkGreen
Write-Host "Computer Inventory Script" 			-ForegroundColor Green
Write-Host "By: Jesse Hamrick" 					-ForegroundColor Green
Write-Host "Created: 04/15/2009" 				-ForegroundColor Green
Write-Host "Contact: www.PowerShellPro.com" 	-ForegroundColor Green
Write-Host "Updated: Jim Fornango - 9/2014"     -ForegroundColor Green
Write-Host "********************************" 	-ForegroundColor DarkGreen
Write-Host " "
Write-Host "Admin rights are required to enumerate information." 	-ForegroundColor Red
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
	else{Write-Host "You did not supply a correct response, `
	Please run script again." -foregroundColor Red}				

#New Excel Application
$Excel = New-Object -Com Excel.Application
$Excel.visible = $True

# Create a workbook with a single sheet
$Excel = $Excel.Workbooks.Add()
$S2 = 

# Assign each worksheet to a variable and
# name the worksheet.
$Sheet1 = $Excel.Worksheets.Item(1)

# The default workbook has three sheets, remove 2 
$Sheet2 = $Excel.Worksheets.Item(2)
$Sheet3 = $Excel.Worksheets.Item(3)
$Sheet2.delete() 
$Sheet3.delete()  

#Create Heading for General Sheet
$Sheet1.Cells.Item(1,1) = "Device_Name"
$Sheet1.Cells.Item(1,2) = "Role"
$Sheet1.Cells.Item(1,3) = "HW_Make"
$Sheet1.Cells.Item(1,4) = "HW_Model"
$Sheet1.Cells.Item(1,5) = "HW_Serial_#"
$Sheet1.Cells.Item(1,6) = "HW_Type"
$Sheet1.Cells.Item(1,7) = "CPU_Count"
$Sheet1.Cells.Item(1,8) = "Memory_MB"
$Sheet1.Cells.Item(1,9) = "Operating_System"
$Sheet1.Cells.Item(1,10) = "SP_Level"
$Sheet1.Cells.Item(1,11) = "OS_Architecture"
$Sheet1.Cells.Item(1,12) = "Windows Key"

$colSheets = ($Sheet1)
foreach ($colorItem in $colSheets){
$intRow = 2
$intRowCPU = 2
$intRowMem = 2
$intRowDisk = 2
$intRowNet = 2
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
clear
}
Write-Host "*******************************" -ForegroundColor Green
Write-Host "The Report has been completed."  -ForeGroundColor Green
Write-Host "*******************************" -ForegroundColor Green
# ========================================================================
# END of Script
# ========================================================================
