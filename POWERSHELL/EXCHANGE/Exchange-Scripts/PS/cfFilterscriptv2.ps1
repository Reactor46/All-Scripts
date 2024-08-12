[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

function CheckExtraSettings{
$extrasettings = 0
if ($GroupbySCL.Checked -eq $true) {$extrasettings = 1}
if ($GroupByReciever.Checked -eq $true) {$extrasettings = 1}
if ($GroupBySender.Checked -eq $true) {$extrasettings = 1}
if ($GroupbySCLDomain.Checked -eq $true) {$extrasettings = 1}
if ($GroupByDate.Checked -eq $true) {$extrasettings = 1}
return $extrasettings
}

function AggResults([string]$saAddress){
	if ($gbhash1.ContainsKey($saAddress)){
	$tnum =  [int]$gbhash1[$saAddress]  + 1
	$gbhash1[$saAddress] = $tnum
	}
		else{
			$gbhash1.Add($saAddress,1)
	}

}

function adddata{

$adhash.Clear()
$ssTable.Clear()
$sclTable.Clear()
$dchash.Clear()
$gbhash1.Clear()
$gbhash2.Clear()
$sclhash.Clear()
$servername = $snServerNameTextBox.text
$extrasettings = CheckExtraSettings
$inIncludeit = 0
$recpstring = ""
$filter = ""
$filter = "$_.agent -eq ""Content Filter Agent"""
$dtQueryDT = New-Object System.DateTime $dpTimeFrom.value.year,$dpTimeFrom.value.month,$dpTimeFrom.value.day,$dpTimeFrom2.value.hour,$dpTimeFrom2.value.minute,$dpTimeFrom2.value.second
$dtQueryDTf =  New-Object System.DateTime $dpTimeFrom1.value.year,$dpTimeFrom1.value.month,$dpTimeFrom1.value.day,$dpTimeFrom3.value.hour,$dpTimeFrom3.value.minute,$dpTimeFrom3.value.second
if ($extrasettings -eq 0){
	get-agentlog -StartDate $dtQueryDT -EndDate $dtQueryDTf | where {$filter} | ForEach-Object {
	$exclude = 0
	if ($sclFilterboxCheck.Checked -eq $true -band $_.ReasonData -ne $sclFilterboxDrop.SelectedItem){$exclude = 1}
	$repstring = ""
	$incRec = $false
	$p2string = [string]::join(" , ", $_.P2FromAddresses)
	$repstring = [string]::join(" , ",$_.Recipients)
	if ($snSenderAddressTextBox.text -ne ""){
		if ($snSenderAddressTextBox.text.ToString().ToLower() -eq $_.P1FromAddress.ToString().ToLower()){
			$incRec = $true
		}
	}
	else {
		if ($snRecipientAddressTextBox.text.ToString().ToLower() -ne ""){
			if ($repstring -match $snRecipientAddressTextBox.text.ToString().ToLower()){
				$incRec = $true
			}
		}
		else{$incRec = $true}
	}
	if ($incRec -eq $true -band $exclude -eq 0){$ssTable.Rows.Add($_.Timestamp,$_.IPAddress,$_.P1FromAddress,$p2string,$repstring,$_.Action,$_.Reason,$_.ReasonData)}
	}
	$dgDataGrid.DataSource = $ssTable}
else{
	get-agentlog -StartDate $dtQueryDT -EndDate $dtQueryDTf | where {$filter} | ForEach-Object {
		if ($GroupbySCL.Checked -eq $true){
				[String]$sclival = "SCL " + $_.ReasonData
				if ($sclhash.ContainsKey($sclival)){
					$tsize = [int]$sclhash[$sclival] + 1
					$sclhash[$sclival] = $tsize
				}
				else{			
					$sclhash.add($sclival,1)
				}
				
		}
		if ($GroupByReciever.Checked -eq $true){
			foreach($recp in $_.Recipients){
				$sclagkey = $recp.ToString().replace("|","-") + "|" + $_.ReasonData
				AggResults($sclagkey)			
			}	
		
		}
		if ($GroupBySender.Checked -eq $true){
			$sclagkey = $_.P1FromAddress.ToString().replace("|","-") + "|" + $_.ReasonData
			AggResults($sclagkey)		
		}	
		if ($GroupByDate.Checked -eq $true){
			$sclagkey  = $_.Timestamp.toshortdatestring().replace("|","-") + "|" + $_.ReasonData
			AggResults($sclagkey)		
		}	
		
		}

		foreach($sclval in $sclhash.keys){
			$sclTable.rows.add("",$sclval,$sclhash[$sclval])
		}
		foreach($adr in $gbhash1.keys){
			$daDatarray = $adr.split("|")
			$sclTable.rows.add($daDatarray[0],$daDatarray[1],$gbhash1[$adr])
		}
		$dgDataGrid.DataSource = $sclTable


}
}


function Exportcsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("Date/Time,P1FromAddress,P2FromAddress,Recipients,Action,Reason,SCL")
	foreach($row in $ssTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`",`"" + $row[1].ToString() + "`",`"" + $row[2].ToString() + "`",`"" + $row[3].ToString() + "`",`"" + $row[4].ToString()  + "`",`"" +  $row[5].ToString()  + "`",`"" + $row[6].ToString() + "`"")
	}
	$logfile.Close()
}
}

$Dataset = New-Object System.Data.DataSet
$ssTable = New-Object System.Data.DataTable
$ssTable.TableName = "AgentLogs"
$ssTable.Columns.Add("Date/Time",[DateTime])
$ssTable.Columns.Add("IPAddress")
$ssTable.Columns.Add("P1FromAddress")
$ssTable.Columns.Add("P2FromAddress")
$ssTable.Columns.Add("Recipients")
$ssTable.Columns.Add("Action")
$ssTable.Columns.Add("Reason")
$ssTable.Columns.Add("SCL")
$Dataset.tables.add($ssTable)
$sclTable = New-Object System.Data.DataTable
$sclTable.TableName = "SCLGroupTable"
$sclTable.Columns.Add("Address")
$sclTable.Columns.Add("SCLValue")
$sclTable.Columns.Add("Number_Messages",[int])
$Dataset.tables.add($sclTable)

$dchash = @{ }
$gbhash1 = @{ }
$gbhash2 = @{ }
$adhash = @{ }
$sclhash = @{ }
$svSizeVal = 0

$form = new-object System.Windows.Forms.form 
$form.Text = "Agent Log SCL Tracker"

# Add Search Button

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(10,125)
$exButton.Size = new-object System.Drawing.Size(85,20)
$exButton.Text = "Search"
$exButton.Add_Click({adddata})
$form.Controls.Add($exButton)

# Add Export Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(100,125)
$exButton1.Size = new-object System.Drawing.Size(85,20)
$exButton1.Text = "Export"
$exButton1.Add_Click({Exportcsv})
$form.Controls.Add($exButton1)

# Add Sender Email-address Box
$snSenderAddressTextBox = new-object System.Windows.Forms.TextBox 
$snSenderAddressTextBox.Location = new-object System.Drawing.Size(100,30) 
$snSenderAddressTextBox.size = new-object System.Drawing.Size(200,20) 
$form.Controls.Add($snSenderAddressTextBox) 

# Add Sender Email-address Lable
$snSenderAddresslableBox = new-object System.Windows.Forms.Label
$snSenderAddresslableBox.Location = new-object System.Drawing.Size(10,30) 
$snSenderAddresslableBox.size = new-object System.Drawing.Size(100,20) 
$snSenderAddresslableBox.Text = "Senders Email"
$form.Controls.Add($snSenderAddresslableBox) 

# Add Recipient Email-address Box
$snRecipientAddressTextBox = new-object System.Windows.Forms.TextBox 
$snRecipientAddressTextBox.Location = new-object System.Drawing.Size(100,55) 
$snRecipientAddressTextBox.size = new-object System.Drawing.Size(200,20) 
$form.Controls.Add($snRecipientAddressTextBox) 

# Add Recipient Email-address Lable
$snRecipientAddresslableBox = new-object System.Windows.Forms.Label
$snRecipientAddresslableBox.Location = new-object System.Drawing.Size(10,55) 
$snRecipientAddresslableBox.size = new-object System.Drawing.Size(100,20) 
$snRecipientAddresslableBox.Text = "Recipients Email"
$form.Controls.Add($snRecipientAddresslableBox) 

# Add DateTimePickers Button

$dpDatePickerFromlableBox = new-object System.Windows.Forms.Label
$dpDatePickerFromlableBox.Location = new-object System.Drawing.Size(320,30) 
$dpDatePickerFromlableBox.size = new-object System.Drawing.Size(90,20) 
$dpDatePickerFromlableBox.Text = "Logged Between"
$form.Controls.Add($dpDatePickerFromlableBox) 

$dpTimeFrom = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom.Location = new-object System.Drawing.Size(410,30)
$dpTimeFrom.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom)

$dpDatePickerFromlableBox1 = new-object System.Windows.Forms.Label
$dpDatePickerFromlableBox1.Location = new-object System.Drawing.Size(350,50) 
$dpDatePickerFromlableBox1.size = new-object System.Drawing.Size(50,20) 
$dpDatePickerFromlableBox1.Text = "and"
$form.Controls.Add($dpDatePickerFromlableBox1) 

$dpTimeFrom1 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom1.Location = new-object System.Drawing.Size(410,50)
$dpTimeFrom1.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom1)

$dpTimeFrom2 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom2.Format = "Time"
$dpTimeFrom2.value = [DateTime]::get_Now().Date
$dpTimeFrom2.ShowUpDown = $True
$dpTimeFrom2.Location = new-object System.Drawing.Size(610,30)
$dpTimeFrom2.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom2)

$dpTimeFrom3 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom3.Format = "Time"
$dpTimeFrom3.ShowUpDown = $True
$dpTimeFrom3.Location = new-object System.Drawing.Size(610,50)
$dpTimeFrom3.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom3)

# Add SCLFilter

$sclFilterboxLable = new-object System.Windows.Forms.Label
$sclFilterboxLable.Location = new-object System.Drawing.Size(320,95) 
$sclFilterboxLable.size = new-object System.Drawing.Size(90,20) 
$sclFilterboxLable.Text = "SCL Value"
$form.Controls.Add($sclFilterboxLable)  

$sclFilterboxCheck =  new-object System.Windows.Forms.CheckBox
$sclFilterboxCheck.Location = new-object System.Drawing.Size(410,85)
$sclFilterboxCheck.Size = new-object System.Drawing.Size(30,25)
$sclFilterboxCheck.Add_Click({if ($sclFilterboxCheck.Checked -eq $true){$sclFilterboxDrop.Enabled = $true}
			else{$sclFilterboxDrop.Enabled = $false}})
$form.Controls.Add($sclFilterboxCheck)


$sclFilterboxDrop = new-object System.Windows.Forms.ComboBox
$sclFilterboxDrop.Location = new-object System.Drawing.Size(440,90)
$sclFilterboxDrop.Size = new-object System.Drawing.Size(70,30)
$sclFilterboxDrop.Enabled = $false
for ( $i = 0; $i -le 9; $i+=1 ) { $sclFilterboxDrop.Items.Add($i)}
$form.Controls.Add($sclFilterboxDrop)

# Add Extras

$GroupbySCL =  new-object System.Windows.Forms.CheckBox
$GroupbySCL.Location = new-object System.Drawing.Size(820,20)
$GroupbySCL.Size = new-object System.Drawing.Size(200,30)
$GroupbySCL.Text = "Group By SCL"
$GroupbySCL.Add_Click({
	$GroupbyReciever.Checked = $false
	$GroupBySender.Checked = $false
	$GroupbyDate.Checked = $false
	})
$form.Controls.Add($GroupbySCL)

$GroupByReciever =  new-object System.Windows.Forms.CheckBox
$GroupByReciever.Location = new-object System.Drawing.Size(820,42)
$GroupByReciever.Size = new-object System.Drawing.Size(200,30)
$GroupByReciever.Text = "Group By Reciever"
$GroupByReciever.Add_Click({
	$GroupbySCL.Checked = $false
	$GroupBySender.Checked = $false
	$GroupbyDate.Checked = $false
	})
$form.Controls.Add($GroupByReciever)

$GroupBySender =  new-object System.Windows.Forms.CheckBox
$GroupBySender.Location = new-object System.Drawing.Size(820,64)
$GroupBySender.Size = new-object System.Drawing.Size(205,28)
$GroupBySender.Text = "Group By Sender"
$GroupBySender.Add_Click({
	$GroupbySCL.Checked = $false
	$GroupbyReciever.Checked = $false
	$GroupbyDate.Checked = $false
})
$form.Controls.Add($GroupBySender)

$GroupByDate =  new-object System.Windows.Forms.CheckBox
$GroupByDate.Location = new-object System.Drawing.Size(820,86)
$GroupByDate.Size = new-object System.Drawing.Size(205,28)
$GroupByDate.Text = "Group By Date"
$GroupByDate.Add_Click({
	$GroupbySCL.Checked = $false
	$GroupbyReciever.Checked = $false
	$GroupBySender.Checked = $false
})
$form.Controls.Add($GroupByDate)

$Gbox =  new-object System.Windows.Forms.GroupBox
$Gbox.Location = new-object System.Drawing.Size(810,5)
$Gbox.Size = new-object System.Drawing.Size(220,135)
$Gbox.Text = "Extras"
$form.Controls.Add($Gbox)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,160) 
$dgDataGrid.size = new-object System.Drawing.Size(1024,700) 


$form.Controls.Add($dgDataGrid)

#populate DataGrid

$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
