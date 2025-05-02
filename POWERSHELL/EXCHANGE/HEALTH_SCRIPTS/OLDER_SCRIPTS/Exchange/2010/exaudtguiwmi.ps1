[System.Reflection.Assembly]::LoadWithPartialName("System.Core")

$UserDNHash = @{ }
$FHash = @{ }
$RHash = @{ }
function FillDnHash(){
	$mbServers = get-mailboxserver
	$mbServers | foreach-object{
		get-mailbox -server $_.Name -ResultSize Unlimited | foreach-object{
			if ($_.LegacyExchangeDN.ToString() -ne ""){
				$UserDNHash.Add($_.LegacyExchangeDN.ToString(),$_)	
			}
		
		}
	}
	get-mailuser -ResultSize Unlimited | foreach-object{
			$UserDNHash.Add($_.LegacyExchangeDN.ToString(),$_)	
		
	}
}
FillDnHash

function showdetail(){
	
	if ($GnGroupbyDrop.SelectedItem -eq "UserName Used"){
		$rows = $fsTable.Select("AccessedBy = '" +  $f2Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "'")}
	else{
		$rows = $fsTable.Select("Mailbox = '" +  $f1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "'")}
	
	$frTable.clear()
	foreach ($row in $rows){
		$frTable.rows.add($row[0].ToString(),$row[1].ToString(),$row[2].ToString(),$row[3].ToString(),$row[4].ToString(),$row[5].ToString(),$row[6].ToString(),$row[7].ToString(),$row[8].ToString(),$row[9].ToString())
	}
	$dgDataGrid1.datasource = $frTable

}

[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

$form = new-object System.Windows.Forms.form 
# Add DataTable

$Dataset = New-Object System.Data.DataSet
$fsTable = New-Object System.Data.DataTable
$f1Table = New-Object System.Data.DataTable
$f2Table = New-Object System.Data.DataTable
$frTable = New-Object System.Data.DataTable

$f1Table.TableName = "Folder Access Forward"
$f1Table.Columns.Add("Mailbox")
$f1Table.Columns.Add("Inbox #")
$f1Table.Columns.Add("Calendar #")
$f1Table.Columns.Add("FreeBusy #")
$f1Table.Columns.Add("Other #")
$Dataset.tables.add($f1Table)


$f2Table.TableName = "Folder Access Reverse"
$f2Table.Columns.Add("User")
$f2Table.Columns.Add("Inbox #")
$f2Table.Columns.Add("Calendar #")
$f2Table.Columns.Add("FreeBusy #")
$f2Table.Columns.Add("Other #")
$Dataset.tables.add($f2Table)


$fsTable.TableName = "Folder Access Detail"
$fsTable.Columns.Add("RecordID")
$fsTable.Columns.Add("DateTime",[DATETIME])
$fsTable.Columns.Add("Mailbox")
$fsTable.Columns.Add("FolderName")
$fstable.Columns.Add("FolderPath")
$fsTable.Columns.Add("AccessedBy")
$fsTable.Columns.Add("IPAddress")
$fsTable.Columns.Add("MachineName")
$fsTable.Columns.Add("ProgramName")
$fsTable.Columns.Add("ApplicationID")
$Dataset.tables.add($fsTable)

$frtable.TableName = "Folder Access Detail Show"
$frTable.Columns.Add("RecordID")
$frtable.Columns.Add("DateTime",[DATETIME])
$frtable.Columns.Add("Mailbox")
$frtable.Columns.Add("FolderName")
$frtable.Columns.Add("FolderPath")
$frtable.Columns.Add("AccessedBy")
$frtable.Columns.Add("IPAddress")
$frtable.Columns.Add("MachineName")
$frtable.Columns.Add("ProgramName")
$frtable.Columns.Add("ApplicationID")
$Dataset.tables.add($frtable)

$sdetailButton = new-object System.Windows.Forms.Button
$sdetailButton.Location = new-object System.Drawing.Size(550,19)
$sdetailButton.Size = new-object System.Drawing.Size(120,23)
$sdetailButton.Text = "Show Folder Detail"
$sdetailButton.visible = $True
$sdetailButton.Add_Click({showDetail})
$form.Controls.Add($sdetailButton)

$gllogButton = new-object System.Windows.Forms.Button
$gllogButton.Location = new-object System.Drawing.Size(440,19)
$gllogButton.Size = new-object System.Drawing.Size(90,23)
$gllogButton.Text = "Get Logs"
$gllogButton.visible = $True
$gllogButton.Add_Click({getLogs})
$form.Controls.Add($gllogButton)

# Add Server DropLable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,20) 
$snServerNamelableBox.size = new-object System.Drawing.Size(80,20) 
$snServerNamelableBox.Text = "ServerName"
$form.Controls.Add($snServerNamelableBox) 

# Add Server Drop Down
$snServerNameDrop = new-object System.Windows.Forms.ComboBox
$snServerNameDrop.Location = new-object System.Drawing.Size(90,20)
$snServerNameDrop.Size = new-object System.Drawing.Size(100,30)
get-mailboxserver | ForEach-Object{$snServerNameDrop.Items.Add($_.Name)}
$form.Controls.Add($snServerNameDrop)


# Add DateTimePickers Button

$dpDatePickerFromlableBox = new-object System.Windows.Forms.Label
$dpDatePickerFromlableBox.Location = new-object System.Drawing.Size(10,50) 
$dpDatePickerFromlableBox.size = new-object System.Drawing.Size(90,20) 
$dpDatePickerFromlableBox.Text = "Logged Between"
$form.Controls.Add($dpDatePickerFromlableBox) 

$dpTimeFrom = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom.Location = new-object System.Drawing.Size(110,50)
$dpTimeFrom.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom)

$dpDatePickerFromlableBox1 = new-object System.Windows.Forms.Label
$dpDatePickerFromlableBox1.Location = new-object System.Drawing.Size(10,70) 
$dpDatePickerFromlableBox1.size = new-object System.Drawing.Size(50,20) 
$dpDatePickerFromlableBox1.Text = "and"
$form.Controls.Add($dpDatePickerFromlableBox1) 

$dpTimeFrom1 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom1.Location = new-object System.Drawing.Size(110,70)
$dpTimeFrom1.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom1)

$dpTimeFrom2 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom2.Format = "Time"
$dpTimeFrom2.value = [DateTime]::get_Now().AddHours(-1)
$dpTimeFrom2.ShowUpDown = $True
$dpTimeFrom2.Location = new-object System.Drawing.Size(310,50)
$dpTimeFrom2.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom2)

$dpTimeFrom3 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom3.Format = "Time"
$dpTimeFrom3.ShowUpDown = $True
$dpTimeFrom3.Location = new-object System.Drawing.Size(310,70)
$dpTimeFrom3.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom3)


# Add DropLable
$GnGroupbylableBox = new-object System.Windows.Forms.Label
$GnGroupbylableBox.Location = new-object System.Drawing.Size(200,20) 
$GnGroupbylableBox.size = new-object System.Drawing.Size(100,20) 
$GnGroupbylableBox.Text = "Group Results By"
$form.Controls.Add($GnGroupbylableBox) 


$GnGroupbyDrop = new-object System.Windows.Forms.ComboBox
$GnGroupbyDrop.Location = new-object System.Drawing.Size(310,20)
$GnGroupbyDrop.Size = new-object System.Drawing.Size(110,30)
$GnGroupbyDrop.Items.Add("Mailbox Accessed")
$GnGroupbyDrop.Items.Add("UserName Used")
$GnGroupbyDrop.Add_SelectedValueChanged({if ($GnGroupbyDrop.SelectedItem -eq "UserName Used"){$dgDataGrid.datasource = $f2Table}
				         else{$dgDataGrid.datasource = $f1Table}
}) 
$form.Controls.Add($GnGroupbyDrop)


# Add Export 1st Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(10,620)
$exButton1.Size = new-object System.Drawing.Size(125,20)
$exButton1.Text = "Export Summary Grid"
$exButton1.Add_Click({exportTable})
$form.Controls.Add($exButton1)

# Add Export 2sn Button

$exButton2 = new-object System.Windows.Forms.Button
$exButton2.Location = new-object System.Drawing.Size(550,620)
$exButton2.Size = new-object System.Drawing.Size(135,20)
$exButton2.Text = "Export Details Grid"
$exButton2.Add_Click({exportDetail})
$form.Controls.Add($exButton2)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,100) 
$dgDataGrid.size = new-object System.Drawing.Size(530,500)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)

$dgDataGrid1 = new-object System.windows.forms.DataGridView
$dgDataGrid1.Location = new-object System.Drawing.Size(550,100) 
$dgDataGrid1.size = new-object System.Drawing.Size(450,500)
$dgDataGrid1.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid1)


function getlogs{
$DateTo = New-Object System.DateTime $dpTimeFrom.value.year,$dpTimeFrom.value.month,$dpTimeFrom.value.day,$dpTimeFrom2.value.hour,$dpTimeFrom2.value.minute,$dpTimeFrom2.value.second
$DateFrom =  New-Object System.DateTime $dpTimeFrom1.value.year,$dpTimeFrom1.value.month,$dpTimeFrom1.value.day,$dpTimeFrom3.value.hour,$dpTimeFrom3.value.minute,$dpTimeFrom3.value.second

$f1Table.Clear()
$f2Table.Clear()
$fsTable.Clear()

$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($DateTo.ToUniversalTime())
$WmidtQueryDTf = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($DateFrom.ToUniversalTime())

$filterblock1 =  "timewritten  <= '" + $WmidtQueryDTf + "' and timewritten  >=  '" + $WmidtQueryDT + "'"

get-wmiobject -computer $snServerNameDrop.SelectedItem -query "select * from Win32_NTLogEvent where LogFile = 'Exchange Auditing' and $filterblock1" | foreach-object{
		$arry = $_.Message.ToString().Split("`n")
		$mbMailbox = ""
		$folderName = ""
		$folderpath = ""
		$AccessingUser = ""
		$processname = ""
		$machinename = ""
		$ApplicationId = ""
		$Address = ""
		$RecordID = ""
		$timewritten = [System.Management.ManagementDateTimeConverter]::ToDateTime($_.TimeWritten) 
		foreach($line in $arry){
			if ($line -match "Mailbox:"){
				$laLineArray = $line.Split(":")
				for($sc=$laLineArray[1].length-1;$sc -ne -1;$sc--){
					if ($laLineArray[1].substring($sc,1) -ne " "){
						$ec = $sc
						$sc = 0
					}
	
				}
				$mbMailbox = $laLineArray[1].Substring(1,$ec-2)
			}
			if ($line -match "Display Name:"){
				$laLineArray = $line.Split(":")
				$folderName = $laLineArray[1].Substring(1,($laLineArray[1].length-2))
			}
			if ($line -match "Accessing User:"){
				$laLineArray = $line.Split(":")
				for($sc=$laLineArray[1].length-1;$sc -ne -1;$sc--){
					if ($laLineArray[1].substring($sc,1) -ne " "){
						$ec = $sc
						$sc = 0
					}
	
				}
				$AccessingUser = $laLineArray[1].Substring(1,$ec-1)
			}
			if ($line -match "Process Name:"){
				$laLineArray = $line.Split(":")
				$processname = $laLineArray[1].Substring(1,($laLineArray[1].length-2))
			}
			if ($line -match "Machine Name:"){
				$laLineArray = $line.Split(":")
				$machinename = $laLineArray[1].Substring(1,($laLineArray[1].length-2))
			}
			if ($line -match "Application Id:"){
				$laLineArray = $line.Split(":")
				$ApplicationId = $laLineArray[1].Substring(1,($laLineArray[1].length-2))
			}
			if ($line -match "Address:"){
				$laLineArray = $line.Split(":")
				$Address = $laLineArray[1].Substring(1,($laLineArray[1].length-2))
			}
			if ($line -match "The folder "){
				$laLinenumbend = $line.indexof("in Mailbox")
				$folderpath = $line.Substring(11,($laLinenumbend-11))
			}
	        }
		if ($mbMailbox.length -gt 2 -band $AccessingUser.length -gt 2){
			$exAuditObject = "" | select RecordID,TimeCreated,FolderPath,FolderName,Mailbox,AccessingUser,MailboxLegacyExchangeDN,AccessingUserLegacyExchangeDN,MachineName,Address,ProcessName,ApplicationId
			$exAuditObject.RecordID = $_.RecordNumber.ToString()
			$exAuditObject.TimeCreated = $timewritten
			$exAuditObject.FolderPath = $folderpath
			$exAuditObject.FolderName = $folderName
			$exAuditObject.Mailbox = $mbMailbox
			$exAuditObject.AccessingUser = $AccessingUser
			$exAuditObject.AccessingUserLegacyExchangeDN = $AccessingUser
			$exAuditObject.MailboxLegacyExchangeDN = $mbMailbox
			$exAuditObject.MachineName = $machinename 
			$exAuditObject.Address = $Address
			$exAuditObject.ProcessName = $processname
			$exAuditObject.ApplicationId = $ApplicationId
		if ($mbMailbox -ne $AccessingUser){
		$fsTable.Rows.add($exAuditObject.RecordID,$exAuditObject.TimeCreated.ToString(),$UserDNHash[$mbMailbox].DisplayName,$exAuditObject.FolderName.ToString(),$exAuditObject.FolderPath.ToString(),$UserDNHash[$AccessingUser].DisplayName,$exAuditObject.Address.ToString(),$exAuditObject.MachineName.ToString(),$exAuditObject.ProcessName.ToString(),$exAuditObject.ApplicationId.ToString())
		
		if($Fhash.containskey($mbMailbox)){
			switch($exAuditObject.FolderName.ToString()){
				"Inbox" {$Fhash[$mbMailbox].InboxCount = $Fhash[$mbMailbox].InboxCount + 1}
				"Calendar" {$Fhash[$mbMailbox].CalendarCount = $Fhash[$mbMailbox].CalendarCount + 1}
				"FreeBusy Data" {$FHash[$mbMailbox].FreeBusyCount = $FHash[$mbMailbox].FreeBusyCount + 1}
				default {$Fhash[$mbMailbox].OtherCount = $Fhash[$mbMailbox].OtherCount + 1}
				}
		}
		else{
			$coCustObj = "" | select MailboxName,InboxCount,CalendarCount,FreeBusyCount,OtherCount
			$coCustObj.MailboxName = $mbMailbox
			$coCustObj.InboxCount = 0
			$coCustObj.CalendarCount = 0
			$coCustObj.FreeBusyCount = 0
			$coCustObj.OtherCount = 0
			$Fhash.add($mbMailbox,$coCustObj)	
			switch($folderName){
				"Inbox" {$Fhash[$mbMailbox].InboxCount = $Fhash[$mbMailbox].InboxCount + 1}
				"Calendar" {$Fhash[$mbMailbox].CalendarCount = $Fhash[$mbMailbox].CalendarCount + 1}
				"FreeBusy Data" {$FHash[$mbMailbox].FreeBusyCount = $FHash[$mbMailbox].FreeBusyCount + 1}
				default {$Fhash[$mbMailbox].OtherCount = $Fhash[$mbMailbox].OtherCount + 1}
			}
		}

		if($RHash.containskey($AccessingUser)){
				switch($exAuditObject.FolderName.ToString()){
					"Inbox" {$RHash[$AccessingUser].InboxCount = $RHash[$AccessingUser].InboxCount + 1}
					"Calendar" {$RHash[$AccessingUser].CalendarCount = $RHash[$AccessingUser].CalendarCount + 1}
					"FreeBusy Data" {$RHash[$AccessingUser].FreeBusyCount = $RHash[$AccessingUser].FreeBusyCount + 1}
					default {$RHash[$AccessingUser].OtherCount = $RHash[$AccessingUser].OtherCount + 1}

				}
		}
		else{
			$coCustObj = "" | select MailboxName,InboxCount,FreeBusyCount,CalendarCount,OtherCount
			$coCustObj.MailboxName = $AccessingUser
			$coCustObj.InboxCount = 0
			$coCustObj.CalendarCount = 0
			$coCustObj.FreeBusyCount = 0
			$coCustObj.OtherCount = 0
			$Rhash.add($AccessingUser,$coCustObj)	
			switch($exAuditObject.FolderName.ToString()){
				"Inbox" {$RHash[$AccessingUser].InboxCount = $RHash[$AccessingUser].InboxCount + 1}
				"Calendar" {$RHash[$AccessingUser].CalendarCount = $RHash[$AccessingUser].CalendarCount + 1}
				"FreeBusy Data" {$RHash[$AccessingUser].FreeBusyCount = $RHash[$AccessingUser].FreeBusyCount + 1}
				default {$RHash[$AccessingUser].OtherCount = $RHash[$AccessingUser].OtherCount + 1}
			}
		}}
	}
}
foreach($key in $Fhash.keys){
	if($key -ne " "){
		$f1Table.rows.add($UserDNHash[$Fhash[$key].MailboxName].DisplayName,$Fhash[$key].InboxCount,$Fhash[$key].CalendarCount,$Fhash[$key].FreeBusyCount,$Fhash[$key].OtherCount)		

}}
foreach($key in $Rhash.keys){
	if($key -ne " "){
	$f2Table.rows.add($UserDNHash[$Rhash[$key].MailboxName].DisplayName,$Rhash[$key].InboxCount,$Rhash[$key].CalendarCount,$Rhash[$key].FreeBusyCount,$Rhash[$key].OtherCount)		
}}

if ($GnGroupbyDrop.SelectedItem -eq "UserName Used"){$dgDataGrid.datasource = $f2Table}
else{$dgDataGrid.datasource = $f1Table}

}

function exportTable{
$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.Showhelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("MailboxName,InboxCount,FreeBusyCount,CalendarCount,OtherCount")
	$table = $dgDataGrid.datasource
	foreach($row in $table.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString()) 
	}
	$logfile.Close()
}

}

function exportDetail{
$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.Showhelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("RecordID,TimeCreated,FolderPath,FolderName,Mailbox,AccessingUser,MailboxLegacyExchangeDN,AccessingUserLegacyExchangeDN,MachineName,Address,ProcessName,ApplicationId")
	$table = $dgDataGrid1.datasource
	foreach($row in $table.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString() + "," + $row[5].ToString() + "," +  $row[6].ToString()+ "," + $row[7].ToString() + "," + $row[8].ToString() + "," + $row[9].ToString()) 
	}
	$logfile.Close()
}

}

$form.Text = "Exchange 2007 Mailbox Folder Access Audit Form"
$form.size = new-object System.Drawing.Size(1000,800) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()