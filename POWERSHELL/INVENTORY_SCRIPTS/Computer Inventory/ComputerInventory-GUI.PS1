# ========================================================
#
# 	Script Information
#
#	Title:				Remote Computer Inventory GUI
#	Author:				Assaf Miron
#	Originally created:	21/06/2008
#	Last Updated:		20/01/2009 
#	Original path:		ComputerInventory-GUI.PS1
#	Description:		Collects Remote Computer Data Using WMI and Registry Access	
#						Outputs all information to a Data Grid Form and to a CSV Log File						
#	
# ========================================================

#region Constructor
# Import Assembly
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# Get Script Location (For Icons)
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

#endregion

#region Form Creation

#~~< Form1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Form1 = New-Object System.Windows.Forms.Form
$Form1.AutoSize = $TRUE
$Form1.ClientSize = New-Object System.Drawing.Size(522, 404)
$Form1.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$Form1.Text = "Computer Inventory"
$Form1.Icon = $ScriptLocation+".\Network.ico"
#~~< Panel1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Panel1 = new-object System.Windows.Forms.Panel
$Panel1.Dock = [System.Windows.Forms.DockStyle]::Fill
$Panel1.Location = new-object System.Drawing.Point(0, 24)
$Panel1.Size = new-object System.Drawing.Size(522, 380)
$Panel1.TabIndex = 20
$Panel1.Text = ""
#~~< btnRun >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Enabled = $FALSE
$btnRun.Location = New-Object System.Drawing.Point(431, 30)
$btnRun.Size = New-Object System.Drawing.Size(75, 23)
$btnRun.TabIndex = 2
$btnRun.Text = "Run"
$btnRun.UseVisualStyleBackColor = $TRUE
$btnRun.add_Click({ RunScript($btnRun) })
#~~< Label1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label1 = New-Object System.Windows.Forms.Label
$Label1.AutoSize = $TRUE
$Label1.Location = New-Object System.Drawing.Point(12, 33)
$Label1.Size = New-Object System.Drawing.Size(159, 13)
$Label1.TabIndex = 15
$Label1.Text = "File containing Computer Names"
#~~< TextBox1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TextBox1 = New-Object System.Windows.Forms.TextBox
$TextBox1.Location = New-Object System.Drawing.Point(177, 30)
$TextBox1.Size = New-Object System.Drawing.Size(161, 20)
$TextBox1.TabIndex = 0
$TextBox1.Text = ""
#~~< btnBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(347, 30)
$btnBrowse.Size = New-Object System.Drawing.Size(75, 23)
$btnBrowse.TabIndex = 1
$btnBrowse.Text = "Browse"
$btnBrowse.UseVisualStyleBackColor = $TRUE
$btnBrowse.add_Click({ BrowseFile($btnBrowse) })
#~~< DataGridView1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DataGridView1 = new-object System.Windows.Forms.DataGridView
$DataGridView1.Anchor = ([System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right))
$DataGridView1.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$DataGridView1.Location = New-Object System.Drawing.Point(12, 59)
$DataGridView1.Size = New-Object System.Drawing.Size(497, 280)
$DataGridView1.TabIndex = 4
$DataGridView1.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::Disable
$DataGridView1.Text = ""
#~~< ProgressBar1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ProgressBar1 = New-Object System.Windows.Forms.ProgressBar
$ProgressBar1.Anchor = ([System.Windows.Forms.AnchorStyles] ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right ))
$ProgressBar1.Location = New-Object System.Drawing.Point(12, 345)
$ProgressBar1.Size = New-Object System.Drawing.Size(410, 23)
$ProgressBar1.TabIndex = 5
$ProgressBar1.Text = ""
#~~< btnExit >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Anchor = ([System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right))
$btnExit.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$btnExit.Location = New-Object System.Drawing.Point(431, 345)
$btnExit.Size = New-Object System.Drawing.Size(78, 23)
$btnExit.TabIndex = 3
$btnExit.Text = "Exit"
$btnExit.UseVisualStyleBackColor = $TRUE
$btnExit.add_Click({ CloseForm($btnExit) })
#~~< MenuStrip1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MenuStrip1 = new-object System.Windows.Forms.MenuStrip
$MenuStrip1.Location = new-object System.Drawing.Point(0, 0)
$MenuStrip1.Size = new-object System.Drawing.Size(292, 24)
$MenuStrip1.TabIndex = 6
$MenuStrip1.Text = "MenuStrip1"
#~~< FileToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FileToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
$FileToolStripMenuItem.Size = new-object System.Drawing.Size(35, 20)
$FileToolStripMenuItem.Text = "File"
#~~< OpenLogFileToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenLogFileToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
$OpenLogFileToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
$OpenLogFileToolStripMenuItem.Text = "Open Log File"
$OpenLogFileToolStripMenuItem.add_Click({Open-file($LogFile)})
#~~< OpenErrorLogFileToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenErrorLogFileToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
$OpenErrorLogFileToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
$OpenErrorLogFileToolStripMenuItem.Text = "Open Error Log File"
$OpenErrorLogFileToolStripMenuItem.add_Click({Open-file($ErrorLogFile)})
#~~< ExitToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ExitToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
$ExitToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
$ExitToolStripMenuItem.Text = "Exit"
$ExitToolStripMenuItem.add_Click({CloseForm($ExitToolStripMenuItem)})
$FileToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@($OpenLogFileToolStripMenuItem, $OpenErrorLogFileToolStripMenuItem, $ExitToolStripMenuItem)))
#~~< HelpToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$HelpToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
$HelpToolStripMenuItem.Size = new-object System.Drawing.Size(40, 20)
$HelpToolStripMenuItem.Text = "Help"
#~~< AboutToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$AboutToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
$AboutToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
$AboutToolStripMenuItem.Text = "About"
$AboutToolStripMenuItem.add_Click({Show-About($AboutToolStripMenuItem)})
#~~< HowToToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$HowToToolStripMenuItem = new-object System.Windows.Forms.ToolStripMenuItem
$HowToToolStripMenuItem.Size = new-object System.Drawing.Size(152, 22)
$HowToToolStripMenuItem.Text = "How To?"
$HowToToolStripMenuItem.add_Click({Show-HowTo($HowToToolStripMenuItem)})
$HelpToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@($AboutToolStripMenuItem, $HowToToolStripMenuItem)))
$MenuStrip1.Items.AddRange([System.Windows.Forms.ToolStripItem[]](@($FileToolStripMenuItem, $HelpToolStripMenuItem)))
$Panel1.Controls.Add($MenuStrip1)
$Panel1.Controls.Add($btnRun)
$Panel1.Controls.Add($Label1)
$Panel1.Controls.Add($TextBox1)
$Panel1.Controls.Add($btnBrowse)
$Panel1.Controls.Add($ProgressBar1)
$Panel1.Controls.Add($btnExit)
$Panel1.Controls.Add($DataGridView1)
$Panel1.Controls.Add($Menu)
$Form1.Controls.Add($Panel1)
#~~< Ping1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Ping1 = New-Object System.Net.NetworkInformation.Ping
#~~< OpenFileDialog1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenFileDialog1 = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog1.Filter = "Text Files|*.txt|CSV Files|*.csv|All Files|*.*"
$OpenFileDialog1.InitialDirectory = "C:"
$OpenFileDialog1.Title = "Open Computers File"
#~~< objNotifyIcon >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon 
# Assign an Icon and Icon Type For the NotifyIcon Object
$objNotifyIcon.Icon = $ScriptLocation+".\Browse 2.ico"
$objNotifyIcon.BalloonTipIcon = "Info" 
#endregion

#region Functions

function out-DataTable
# Function Creates a New Data Table that will be set as the Data Source of the Data Grid View
# Thanks to /\/\o\/\/ http:\\thePowerShellGuy.com
{
	$dt = New-Object Data.datatable
	$First = $TRUE	

	foreach ($item in $Input)
	{
		$DR = $DT.NewRow()
		$Item.PsObject.get_properties() | foreach {
			if ($first)
			{
				$Col = New-Object Data.DataColumn
				$Col.ColumnName = $_.Name.ToString()
			$DT.Columns.Add($Col) }
			if ($_.value -eq $null -or $_.value -eq "")
			{
				$DR.Item($_.Name) = "[empty]"
			}
			elseif ($_.IsArray) {
				$DR.Item($_.Name) = [string]::Join($_.value, ";")
			}
			else
			{
				$DR.Item($_.Name) = $_.value
			}
		}
		$DT.Rows.Add($DR)
		$First = $FALSE
	}
		
	return @(, ( $dt ))
		
}



function Show-NotifyIcon {
# Function Controls the Notification Icon
# Changes its Title and Text
	param(
		$Title,
		$Text
	)
		# Change Notify Icon Title
		$objNotifyIcon.BalloonTipTitle = $Title
		# Change Notify Icon Text
		$objNotifyIcon.BalloonTipText = $Text
		# Show Notify Icon for 10 Secs
		$objNotifyIcon.Visible = $TRUE 
		$objNotifyIcon.ShowBalloonTip(10000)
}

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

function StoredProc {
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

#region Event Loop

function Main
# Main Function, Runs the Form
{
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run($Form1)
}

#endregion


#region Event Handlers

function BrowseFile($object)
# Function for Running the OpenFileDialog
# Used when Clicking on the Browse Button
{
	$OpenFileDialog1.showdialog()
	$TextBox1.Text = $OpenFileDialog1.FileName
	$btnRun.Enabled = $TRUE
}

function Open-File( $FileToOpen ){
# Function Open the Log File
	if(Test-Path $FileToOpen){
		Invoke-Item $FileToOpen
	}
}

function Show-PopUp
# Function for Showing Custom Pop up Forms
{
	param(
		$PopupTitle,
		$PopupText
		)
#~~< PopupForm >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PopupForm = New-Object System.Windows.Forms.Form
$PopupForm.ClientSize = New-Object System.Drawing.Size(381, 356)
$PopupForm.ControlBox = $false
$PopupForm.ShowInTaskbar = $false
$PopupForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
$PopupForm.Text = $PopupTitle
#~~< PopupColse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PopupColse = New-Object System.Windows.Forms.Button
$PopupColse.Anchor = ([System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right))
$PopupColse.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$PopupColse.Location = New-Object System.Drawing.Point(137, 321)
$PopupColse.Size = New-Object System.Drawing.Size(104, 23)
$PopupColse.TabIndex = 0
$PopupColse.Text = "Close"
$PopupColse.UseVisualStyleBackColor = $true
$PopupForm.AcceptButton = $PopupColse
$PopupForm.CancelButton = $PopupColse
#~~< PopupHeader >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PopupHeader = New-Object System.Windows.Forms.Label
$PopupHeader.Font = New-Object System.Drawing.Font("Calibri", 15.75, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$PopupHeader.Location = New-Object System.Drawing.Point(137, 9)
$PopupHeader.Size = New-Object System.Drawing.Size(104, 23)
$PopupHeader.TabIndex = 2
$PopupHeader.Text = $PopupTitle
$PopupHeader.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
#~~< PopUpTextArea >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PopUpTextArea = New-Object System.Windows.Forms.Label
$PopUpTextArea.Anchor = ([System.Windows.Forms.AnchorStyles]([System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right))
$PopUpTextArea.Location = New-Object System.Drawing.Point(12, 15)
$PopUpTextArea.Size = New-Object System.Drawing.Size(357, 265)
$PopUpTextArea.TabIndex = 1
$PopUpTextArea.Text = $PopupText
$PopUpTextArea.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$PopupForm.Controls.Add($PopupHeader)
$PopupForm.Controls.Add($PopUpTextArea)
$PopupForm.Controls.Add($PopupColse)

# Show Form
$PopupForm.Add_Shown({$PopupForm.Activate()})  
[void]$PopupForm.showdialog() 

}

function Show-About( $object ){
# Function Opens the About Page
$AboutText = @("
Script Title:	            Remote Computer Inventory`n
Script Author:              Assaf Miron`n
Script Description:         Collects Remote Computer Data Using WMI and Registry Access Outputs all information to a Data Grid Form and to a CSV Log File.`n

Log File Name:	$LogFile")

Show-Popup -PopupTitle "About" -PopupText $AboutText
}

function Show-HowTo( $object ){
# Function Opens the Help Page
$HowToText = @("
1. Click on the Browse Button and select a TXT or a CSV File Containing Computer Names`n
2. After File is Selected click on the Run Button.`n
3. You will see a Notify Icon with the Coresponding Text.`n
4. The Script has begon collecting Remote Computer Inventory!`n
`nWhen The script is Done you will see a Popup Message and all data will be presented in the DataGrid.`n
** Because Poweshell works only in MTA mode there is no Option Copying the Data off the DataGrid...`n
5. All Data will be Exported to a Log File Located Here: $LogFile")

Show-Popup -PopupTitle "How To?" -PopupText $HowToText
}

function CloseForm($object)
# Function End the Program
{
	$Form1.Close()
}

function RunScript($object)
# Function Runs the Program and starts collecting data
{

	# Open a new Connection and Create a new Command Type
	# Connect To DB
	$conn = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
	$conn.Open()
	
	# Create an Array of Computers Enterd in the Input File
	$arrComputers = Get-Content -path $textBox1.Text -encoding UTF8
	
	# Create an Empty Array to Contain all The Data of all Scaned Computers
	$AllComputerInfo = @()
	$ProgressBar1.Value = 0
	$DataGridView1.DataSource = ""
	# Init the Progress bar to it's Maximum Value
	if(($arrComputers -is [array]) -eq $FALSE) { $ProgressBar1.Maximum = 1 }
	else { $ProgressBar1.Maximum = $arrComputers.Count }
	$ProgressBar1.Step = 1 # Define the Progress bar Step value
	
	# Scan all Computers in the Array $arrComputers
	foreach ($strComputer in $arrComputers)
		{ 
			# Uses the Ping Command to check if the Computer is Alive
			$Alive=""
			$Alive = $Ping1.Send($strComputer).Status 
			if($Alive -eq "Success")
			{
				$cmdPortQry = "& '$ScriptLocation\portqry.exe' -n $strComputer -e 135 | find "+[char]34+"TCP port 135 (epmap service):"+[char]34
				Show-NotifyIcon -Title "Querying Computer Ports" -Text "Querying $strComputer For Opend Port (port 135)" 
				$PortQuery = Invoke-Expression $cmdPortQry
				If($PortQuery.split(":")[1].Trim() -eq "LISTENING") # Check Ports in dest computer
				{
					Show-NotifyIcon -Title "Retriving Computer Information" -Text "Scanning $strComputer For Hardware Data" 
					
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
					Show-NotifyIcon -Text "Exporting $strComputer Information" -Title "Exporting..."
			
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
					
					# Clean up - Delete File
					Remove-Item -Path C:\CompDet.xml
				}
				else # Computer behind Firewall
				{
					#region Get Computer Main OU
						# Create command to run
						$cmdOU = "Cscript.exe -nologo .\SearchComputers-ReturnADSPath.vbs $strComputer"
						$MainOU = Invoke-Expression $cmdOU
						If($MainOU.Contains(","))
						{
							$MainOU = $MainOU.Split(",")[-4].Replace("OU=","")
						}
					#endregion
					$objNotifyIcon.BalloonTipIcon = "Error" 
					Write-Warning "$strComputer behind Firewall.`nNo Data was Collected"
					Show-NotifyIcon -Title "$strComputer is not avialable" -Text "$strComputer behind Firewall.`n`tNo Data was Collected"
					# Write the computer name in the Error Log File
					$strComputer | Out-File -Append $ErrorLogFile
					$objNotifyIcon.BalloonTipIcon = "Info"
					StoredProc -SP_ComputerName $strComputer -SP_ComputerTotalHealth "FireWalled" -SP_MainOU $MainOU
				}
			}
			else # No Ping to Computer
			{ 
				#region Get Computer Main OU
					# Create command to run
					$cmdOU = "Cscript.exe -nologo .\SearchComputers-ReturnADSPath.vbs $strComputer"
					$MainOU = Invoke-Expression $cmdOU
					If($MainOU.Contains(","))
					{
						$MainOU = $MainOU.Split(",")[-4].Replace("OU=","")
					}
				#endregion
				$objNotifyIcon.BalloonTipIcon = "Error" 
				Write-Warning "No Ping to $strComputer.`nNo Data was Collected"
				Show-NotifyIcon -Title "$strComputer is not avialable" -Text "No Ping to $strComputer.`n`tNo Data was Collected"
				# Write the computer name in the Error Log File
				$strComputer | Out-File -Append $ErrorLogFile
				$objNotifyIcon.BalloonTipIcon = "Info" 
				StoredProc -SP_ComputerName $strComputer -SP_ComputerTotalHealth "No Ping" -SP_MainOU $MainOU
			}
			$AllComputerInfo += $DataObject
			$DataObject = "" # Clear Data Object
			$ProgressBar1.PerformStep()
			$Form1.Update()
			$Form1.Refresh()
		}
	
#region Finishing - Script is Done
	
	# Assign an Icon and Icon Type For the NotifyIcon Object
	$objNotifyIcon.Icon = $ScriptLocation+".\Publish.ico"
	$objNotifyIcon.BalloonTipIcon = "Info" 
	
	# Show Notify Icon with Finishing Text
	$objNotifyIcon.BalloonTipText = "Done!`nFile Saved in "+$LogFile
	$objNotifyIcon.Visible = $TRUE 
	$objNotifyIcon.ShowBalloonTip(10000)
	
	
	#region Exporting data - DataGrid and Log File
			
			# Export the DataObject to the DataGridView Using the out-DataTable
			$DataTable = $AllComputerInfo | out-dataTable
			
			# Define Data Grid's Data Source to the DataTable we Created
			$DataGridView1.DataSource = $DataTable.psObject.baseobject
			$DataGridView1.Refresh() # Refresh the Table View in order to View the new lines
			
			# Export all the Data to the Log File
			$AllComputerInfo | Export-Csv -Encoding OEM -Path $LogFile
	#endregion
	
	# Pop Up a Message box
	$MSGObject = new-object -comobject wscript.shell
	$MSGResult = $MSGObject.popup("Script Has Finished Running!",0,"I'm Done",0)
	
	# Set to False so that the Notify Icon will Disapear after the Script is Done
	$objNotifyIcon.Visible = $FALSE 
	$conn.Close
#endregion
}

Main # This call must remain below all other event functions

#endregion


