[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

function CheckExtraSettings{
$extrasettings = 0
if ($GroupbySender.Checked -eq $true) {$extrasettings = 1}
if ($GroupByReciever.Checked -eq $true) {$extrasettings = 1}
if ($GroupByRecieverDomain.Checked -eq $true) {$extrasettings = 1}
if ($GroupBySenderDomain.Checked -eq $true) {$extrasettings = 1}
if ($GroupByDate.Checked -eq $true) {$extrasettings = 1}
return $extrasettings
}

Function GetLogs{
	$ssTable.clear()
	$DomainHash.clear()
	get-accepteddomain | ForEach-Object{
		if ($_.DomainType -eq "Authoritative"){
			$DomainHash.add($_.DomainName.SmtpDomain.ToString().ToLower(),1)
		}
	}
	$dtQueryDT = New-Object System.DateTime $dpTimeFrom.value.year,$dpTimeFrom.value.month,$dpTimeFrom.value.day,$dpTimeFrom2.value.hour,$dpTimeFrom2.value.minute,$dpTimeFrom2.value.second
	$dtQueryDTf =  New-Object System.DateTime $dpTimeFrom1.value.year,$dpTimeFrom1.value.month,$dpTimeFrom1.value.day,$dpTimeFrom3.value.hour,$dpTimeFrom3.value.minute,$dpTimeFrom3.value.second
	Get-TransportServer | Get-MessageTrackingLog -ResultSize Unlimited -Start $dtQueryDT -End $dtQueryDTf -recipient $snRecipientAddressTextBox.Text -EventId "Receive" | ForEach-Object{ 
		$recpients = ""
		foreach($recp in $_.Recipients){
			if ($recpients -eq ""){$recpients = $recp}
			else {$recpients = $recpients + ";" + $recp}
		}
		$ssTable.Rows.Add($_.TimeStamp,$_.EventId,$_.Sender,$recpients,$_.MessageSubject.ToString(),($_.TotalBytes/1024).ToString(0.00))
	}
	$dgDataGrid.DataSource = $ssTable
}

function Exportcsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("Time,Action,MessageID,SenderAddress,RecipientAddress,Subject,Size(KB)")
	foreach($row in $ssTable.Rows){
		$logfile.WriteLine($row[0].ToString() + "," + $row[1].ToString() + "," + $row[2].ToString() + ",`"" + $row[3].ToString() + "`"," + $row[4].ToString())
	}
	$logfile.Close()
}
}

$Dataset = New-Object System.Data.DataSet
$ssTable = New-Object System.Data.DataTable
$ssTable.TableName = "TrackingLogs"
$ssTable.Columns.Add("Time",[DateTime])
$ssTable.Columns.Add("Action")
$ssTable.Columns.Add("SenderAddress")
$ssTable.Columns.Add("RecipientAddress")
$ssTable.Columns.Add("Subject")
$ssTable.Columns.Add("Size (KB)",[int])
$Dataset.tables.add($ssTable)
$gsTable = New-Object System.Data.DataTable
$gsTable.TableName = "Grouped-TrackingLogs-Sender"
$gsTable.Columns.Add("EmailAddress")
$gsTable.Columns.Add("Domain")
$gsTable.Columns.Add("Number_Messages",[int])
$gsTable.Columns.Add("Size (KB)",[int])
$Dataset.tables.add($gsTable)

$DomainHash = @{ }
$dchash = @{ }
$gbhash1 = @{ }
$gbhash2 = @{ }
$adhash = @{ }
$svSizeVal = 0

$form = new-object System.Windows.Forms.form 
$form.Text = "Recipient Message Tracker Exchange 2007"


# Add Search Button

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(10,125)
$exButton.Size = new-object System.Drawing.Size(85,20)
$exButton.Text = "Search"
$exButton.Add_Click({GetLogs})
$form.Controls.Add($exButton)

# Add Export Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(100,125)
$exButton1.Size = new-object System.Drawing.Size(85,20)
$exButton1.Text = "Export"
$exButton1.Add_Click({Exportcsv})
$form.Controls.Add($exButton1)



# Add Recipient Email-address Box
$snRecipientAddressTextBox = new-object System.Windows.Forms.TextBox 
$snRecipientAddressTextBox.Location = new-object System.Drawing.Size(100,80) 
$snRecipientAddressTextBox.size = new-object System.Drawing.Size(200,20) 
$form.Controls.Add($snRecipientAddressTextBox) 

# Add Recipient Email-address Lable
$snRecipientAddresslableBox = new-object System.Windows.Forms.Label
$snRecipientAddresslableBox.Location = new-object System.Drawing.Size(10,80) 
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
$dpTimeFrom2.value = [DateTime]::get_Now().AddHours(-1)
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



# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,160) 
$dgDataGrid.size = new-object System.Drawing.Size(1024,700) 


$form.Controls.Add($dgDataGrid)

#populate DataGrid

$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
