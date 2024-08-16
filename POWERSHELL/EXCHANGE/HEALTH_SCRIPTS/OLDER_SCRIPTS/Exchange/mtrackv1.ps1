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

function AggResults([string]$saAddress){
	if ($gbhash1.ContainsKey($saAddress)){
	$tsize = [int]$gbhash2[$saAddress] + [int]$svSizeVal
	$tnum =  [int]$gbhash1[$saAddress]  + 1
	$gbhash1[$saAddress] = $tnum
	$gbhash2[$saAddress] = $tsize
	}
		else{
			$gbhash1.Add($saAddress,1)
			$gbhash2.Add($saAddress,$svSizeVal)
	}

}

function GetEmailDomains{

$reRootDse = New-Object System.DirectoryServices.DirectoryEntry("LDAP://rootDSE") 
$cfConfigroot = New-Object directoryservices.directoryentry("LDAP://" + $reRootDse.configurationnamingcontext) 
$dsSearcher = New-Object directoryservices.directorySearcher($cfConfigroot)
$dsSearcher.PropertiesToLoad.Add("gatewayProxy") 
$dsSearcher.filter = "(objectCategory=msExchRecipientPolicy)" 
$rsResults = $dsSearcher.findall()
foreach ($rsResult in $rsResults) {
	$rpProps = $rsResult.properties
	foreach ($eaEmailAddress in $rpProps.gatewayproxy){
		if ($eaEmailAddress.ToLower().indexofany("smtp:") -eq 0){
			$arEmailAddress = $eaEmailAddress.split("@")
			if ($adhash.ContainsKey($arEmailAddress[1])){}
			else {$adhash.Add($arEmailAddress[1],1)}
		}
	}
}
}

function adddata{

$adhash.Clear()
$ssTable.Clear()
$gsTable.Clear()
$dchash.Clear()
$gbhash1.Clear()
$gbhash2.Clear()
$servername = $snServerNameTextBox.text
$extrasettings = CheckExtraSettings
$inIncludeit = 0

$dtQueryDT = New-Object System.DateTime $dpTimeFrom.value.year,$dpTimeFrom.value.month,$dpTimeFrom.value.day,$dpTimeFrom2.value.hour,$dpTimeFrom2.value.minute,$dpTimeFrom2.value.second
$dtQueryDTf =  New-Object System.DateTime $dpTimeFrom1.value.year,$dpTimeFrom1.value.month,$dpTimeFrom1.value.day,$dpTimeFrom3.value.hour,$dpTimeFrom3.value.minute,$dpTimeFrom3.value.second
$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($dtQueryDT.ToUniversalTime())
$WmidtQueryDTf = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($dtQueryDTf.ToUniversalTime())
$WmiNamespace = "ROOT\MicrosoftExchangev2"

$filterblock1 =  "OriginationTime <= '" + $WmidtQueryDTf + "' and OriginationTime >=  '" + $WmidtQueryDT + "' and entrytype = '1028'"
$filterblock2 =  "OriginationTime <= '" + $WmidtQueryDTf + "' and OriginationTime >=  '" + $WmidtQueryDT + "' and entrytype = '1020'"

if ($seSizeCheck.Checked -eq $true){
	$schfor = $seSizeCheckNum.value*1024
	$filterblock1 = $filterblock1  + " and Size > '" + $schfor.ToString(0.00) + "'" 
        $filterblock2 = $filterblock2  + " and Size > '" + $schfor.ToString(0.00) + "'"
}

if ($snSenderAddressTextBox.text -ne ""){
      $filterblock1 = $filterblock1  + " and SenderAddress = '" + $snSenderAddressTextBox.text + "'" 
      $filterblock2 = $filterblock2  + " and SenderAddress = '" + $snSenderAddressTextBox.text + "'"
}

$filter = $filterblock1 + " or " + $filterblock2
if ($etTypeCheck.Checked -eq $true -bor $edTypeCheck.Checked -eq $true){GetEmailDomains}
if ($snRecipientAddressTextBox.text -eq ""){
	get-wmiobject -class Exchange_MessageTrackingEntry -Namespace $WmiNamespace -ComputerName $servername -filter $filter | ForEach-Object{ 
	if ($etTypeCheck.Checked -eq $true){
		$inInternal = 0
		$inIncludeit = 0
		$rmRecpMatch = 1
		$senddomarray1 = $_.SenderAddress.split("@")
		foreach($recp in $_.RecipientAddress){
			$recpdomarray1 =$recp.split("@")
			if ($adhash.ContainsKey($recpdomarray1[1])){}
			else {$rmRecpMatch = 0}
			}
			if ($senddomarray1.length -ne 1){ 
				if ($adhash.ContainsKey($senddomarray1[1]) -band $rmRecpMatch -eq 1){$inInternal = 1}}
			if ($etTypeCheckDrop.SelectedItem -eq "Internal" -band $inInternal -eq 1){
				$inIncludeit = 1
			}
			if ($etTypeCheckDrop.SelectedItem -eq "External" -band $inInternal -eq 0){
				$inIncludeit = 1
			}
		}
	else {$inIncludeit = 1}
	if ($edTypeCheck.Checked -eq $true -band $inIncludeit -eq 1){
		$issentex = 1
		$senddomarray1 = $_.SenderAddress.split("@")
		if ($senddomarray1.length -ne 1){ 
			if ($adhash.ContainsKey($senddomarray1[1])){
				$issentex = 0	
			}
		}
		if ($edTypeCheckDrop.SelectedItem -eq "Sent" -band $issentex -eq 1){$inIncludeit = 0}	
		if ($edTypeCheckDrop.SelectedItem -eq "Recieved" -band $issentex -eq 0){$inIncludeit = 0}
		
	}
	if ($inIncludeit -eq 1){
		$recpval = ""
		foreach($recp in $_.RecipientAddress){
			if ($recpval -eq ""){$recpval = $recp}
			else {$recpval = $recpval + ";" + $recp}}

		if($dchash.ContainsKey($_.MessageID)){}
		else{
			$dchash.Add($_.MessageID,1)
			$svSizeVal = $_.size/1024
			if ($extrasettings -eq 0){				
				$ssTable.Rows.Add([System.Management.ManagementDateTimeConverter]::ToDateTime($_.OriginationTime),$_.SenderAddress,$recpval,$_.subject,[int]$svSizeVal.ToString(0.00))
				}
			if ($GroupbySender.Checked -eq $true -bor $GroupBySenderDomain.Checked -eq $true ){
				if ($GroupbySender.Checked -eq $true){ 
					AggResults($_.SenderAddress)
				}
				else{
						$senddomarray = $_.SenderAddress.split("@")
						AggResults($senddomarray[1])
				}
				
			}
			if ($GroupByReciever.Checked -eq $true -bor $GroupByRecieverDomain.Checked -eq $true ){
				foreach($recp in $_.RecipientAddress){
					if ($GroupByReciever.Checked -eq $true){
						AggResults($recp)
					}
					else{
						$recpdomarray = $recp.split("@")
						AggResults($recpdomarray[1])
					}
				}
			}
			if ($GroupByDate.Checked -eq $true){
				$dateag = [System.Management.ManagementDateTimeConverter]::ToDateTime($_.OriginationTime)
				AggResults($dateag.ToShortDateString())
			}
		}
	}
}
}
else{
	get-wmiobject -class Exchange_MessageTrackingEntry -Namespace $WmiNamespace -ComputerName $servername -filter $filter | where-object  {$_.RecipientAddress -eq $snRecipientAddressTextBox.text} | ForEach-Object{ 
	if ($etTypeCheck.Checked -eq $true){
	$inInternal = 0
	$inIncludeit = 0
	$rmRecpMatch = 1
	$senddomarray1 = $_.SenderAddress.split("@")
	foreach($recp in $_.RecipientAddress){
		$recpdomarray1 =$recp.split("@")
		if ($adhash.ContainsKey($recpdomarray1[1])){}
		else {$rmRecpMatch = 0}
		}
		if ($senddomarray1.length -ne 1){ 
				if ($adhash.ContainsKey($senddomarray1[1]) -band $rmRecpMatch -eq 1){$inInternal = 1}}
		if ($etTypeCheckDrop.SelectedItem -eq "Internal" -band $inInternal -eq 1){
			$inIncludeit = 1
		}
		if ($etTypeCheckDrop.SelectedItem -eq "External" -band $inInternal -eq 0){
			$inIncludeit = 1
		}
	}
	else {$inIncludeit = 1}
	if ($edTypeCheck.Checked -eq $true -band $inIncludeit -eq 1){
		$issentex = 1
		$senddomarray1 = $_.SenderAddress.split("@")
		if ($senddomarray1.length -ne 1){ 
			if ($adhash.ContainsKey($senddomarray1[1])){
				$issentex = 0	
			}
		}
		if ($edTypeCheckDrop.SelectedItem -eq "Sent" -band $issentex -eq 1){$inIncludeit = 0}	
		if ($edTypeCheckDrop.SelectedItem -eq "Recieved" -band $issentex -eq 0){$inIncludeit = 0}
		
	}
	if ($inIncludeit -eq 1){
	$recpval = ""
	foreach($recp in $_.RecipientAddress){if ($recpval -eq ""){$recpval = $recp}
		else {$recpval = $recpval + ";" + $recp}}
	if($dchash.ContainsKey($_.MessageID)){}
	else{
		$dchash.Add($_.MessageID,1)
		if ($extrasettings -eq 0){
			$svSizeVal = $_.size/1024
			$ssTable.Rows.Add([System.Management.ManagementDateTimeConverter]::ToDateTime($_.OriginationTime),$_.SenderAddress,$recpval,$_.subject,[int]$svSizeVal.ToString(0.00))
			}
		if ($GroupbySender.Checked -eq $true -bor $GroupBySenderDomain.Checked -eq $true ){
			if ($GroupbySender.Checked -eq $true){ 
				AggResults($_.SenderAddress)
			}
			else{
					$senddomarray = $_.SenderAddress.split("@")
					AggResults($senddomarray[1])
			}
			
		}
		if ($GroupByReciever.Checked -eq $true -bor $GroupByRecieverDomain.Checked -eq $true ){
			foreach($recp in $_.RecipientAddress){
				if ($GroupByReciever.Checked -eq $true){
					AggResults($recp)
				}
				else{
					$recpdomarray = $recp.split("@")
					AggResults($recpdomarray[1])
				}
			}
		}
		if ($GroupByDate.Checked -eq $true){
			$dateag = [System.Management.ManagementDateTimeConverter]::ToDateTime($_.OriginationTime)
			AggResults($dateag.ToShortDateString())
		}
}}}}

if ($extrasettings -eq 0){$dgDataGrid.DataSource = $ssTable

}
else {	foreach ($htent in $gbhash1.keys){
		$spemarray = $htent.split("@")
		$gsTable.Rows.Add($htent,$spemarray[1],[int]$gbhash1[$htent],[int]$gbhash2[$htent])
	}
	$dgDataGrid.DataSource = $gsTable
 }}

function Exportcsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("OriginationTime,SenderAddress,RecipientAddress,Subject,Size")
	foreach($row in $ssTable.Rows){
		$logfile.WriteLine($row[0].ToString() + "," + $row[1].ToString() + "," + $row[2].ToString() + ",`"" + $row[3].ToString() + "`"," + $row[4].ToString())
	}
	$logfile.Close()
}
}

$Dataset = New-Object System.Data.DataSet
$ssTable = New-Object System.Data.DataTable
$ssTable.TableName = "TrackingLogs"
$ssTable.Columns.Add("Origination Time",[DateTime])
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

$dchash = @{ }
$gbhash1 = @{ }
$gbhash2 = @{ }
$adhash = @{ }
$svSizeVal = 0

$form = new-object System.Windows.Forms.form 
$form.Text = "BYO Message Tracker"

# Add Servername Box
$snServerNameTextBox = new-object System.Windows.Forms.TextBox 
$snServerNameTextBox.Location = new-object System.Drawing.Size(100,30) 
$snServerNameTextBox.size = new-object System.Drawing.Size(200,20) 
$form.Controls.Add($snServerNameTextBox) 

# Add Servername Lable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,30) 
$snServerNamelableBox.size = new-object System.Drawing.Size(100,20) 
$snServerNamelableBox.Text = "ServerName"
$form.Controls.Add($snServerNamelableBox) 

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
$snSenderAddressTextBox.Location = new-object System.Drawing.Size(100,55) 
$snSenderAddressTextBox.size = new-object System.Drawing.Size(200,20) 
$form.Controls.Add($snSenderAddressTextBox) 

# Add Sender Email-address Lable
$snSenderAddresslableBox = new-object System.Windows.Forms.Label
$snSenderAddresslableBox.Location = new-object System.Drawing.Size(10,55) 
$snSenderAddresslableBox.size = new-object System.Drawing.Size(100,20) 
$snSenderAddresslableBox.Text = "Senders Email"
$form.Controls.Add($snSenderAddresslableBox) 

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

# Add Size Clause

$esSizelableBox = new-object System.Windows.Forms.Label
$esSizelableBox.Location = new-object System.Drawing.Size(320,70) 
$esSizelableBox.size = new-object System.Drawing.Size(90,20) 
$esSizelableBox.Text = "Size Larger then"
$form.Controls.Add($esSizelableBox) 

$seSizeCheck =  new-object System.Windows.Forms.CheckBox
$seSizeCheck.Location = new-object System.Drawing.Size(410,65)
$seSizeCheck.Size = new-object System.Drawing.Size(30,25)
$seSizeCheck.Add_Click({if ($seSizeCheck.Checked -eq $true){$seSizeCheckNum.Enabled = $true}
			else{$seSizeCheckNum.Enabled = $false}})
$form.Controls.Add($seSizeCheck)

$seSizeCheckNum =  new-object System.Windows.Forms.numericUpDown
$seSizeCheckNum.Location = new-object System.Drawing.Size(440,70)
$seSizeCheckNum.Size = new-object System.Drawing.Size(70,30)
$seSizeCheckNum.Enabled = $false
$seSizeCheckNum.Maximum = 10000000000
$form.Controls.Add($seSizeCheckNum)

# Add Type

$etTypeBox = new-object System.Windows.Forms.Label
$etTypeBox.Location = new-object System.Drawing.Size(320,95) 
$etTypeBox.size = new-object System.Drawing.Size(90,20) 
$etTypeBox.Text = "Type of Email"
$form.Controls.Add($etTypeBox)  

$etTypeCheck =  new-object System.Windows.Forms.CheckBox
$etTypeCheck.Location = new-object System.Drawing.Size(410,85)
$etTypeCheck.Size = new-object System.Drawing.Size(30,25)
$etTypeCheck.Add_Click({if ($etTypeCheck.Checked -eq $true){$etTypeCheckDrop.Enabled = $true}
			else{$etTypeCheckDrop.Enabled = $false}})
$form.Controls.Add($etTypeCheck)


$etTypeCheckDrop = new-object System.Windows.Forms.ComboBox
$etTypeCheckDrop.Location = new-object System.Drawing.Size(440,90)
$etTypeCheckDrop.Size = new-object System.Drawing.Size(70,30)
$etTypeCheckDrop.Enabled = $false
$etTypeCheckDrop.Items.Add("Internal")
$etTypeCheckDrop.Items.Add("External")
$form.Controls.Add($etTypeCheckDrop)

# Add Direction

$edTypeBox = new-object System.Windows.Forms.Label
$edTypeBox.Location = new-object System.Drawing.Size(310,115) 
$edTypeBox.size = new-object System.Drawing.Size(100,20) 
$edTypeBox.Text = "Direction of Email"
$form.Controls.Add($edTypeBox)  

$edTypeCheck =  new-object System.Windows.Forms.CheckBox
$edTypeCheck.Location = new-object System.Drawing.Size(410,110)
$edTypeCheck.Size = new-object System.Drawing.Size(30,25)
$edTypeCheck.Add_Click({if ($edTypeCheck.Checked -eq $true){$edTypeCheckDrop.Enabled = $true}
			else{$edTypeCheckDrop.Enabled = $false}})
$form.Controls.Add($edTypeCheck)

$edTypeCheckDrop = new-object System.Windows.Forms.ComboBox
$edTypeCheckDrop.Location = new-object System.Drawing.Size(440,115)
$edTypeCheckDrop.Size = new-object System.Drawing.Size(70,30)
$edTypeCheckDrop.Enabled = $false
$edTypeCheckDrop.Items.Add("Sent")
$edTypeCheckDrop.Items.Add("Recieved")
$form.Controls.Add($edTypeCheckDrop)

# Add Extras

$GroupbySender =  new-object System.Windows.Forms.CheckBox
$GroupbySender.Location = new-object System.Drawing.Size(820,20)
$GroupbySender.Size = new-object System.Drawing.Size(200,30)
$GroupbySender.Text = "Group By Sender Size,Number"
$GroupBySender.Add_Click({
	$GroupbyReciever.Checked = $false
	$GroupbyRecieverDomain.Checked = $false
	$GroupbySenderDomain.Checked = $false
	$GroupbyDate.Checked = $false
	})
$form.Controls.Add($GroupBySender)

$GroupByReciever =  new-object System.Windows.Forms.CheckBox
$GroupByReciever.Location = new-object System.Drawing.Size(820,42)
$GroupByReciever.Size = new-object System.Drawing.Size(200,30)
$GroupByReciever.Text = "Group By Reciever Size,Number"
$GroupByReciever.Add_Click({
	$GroupbySender.Checked = $false
	$GroupbyRecieverDomain.Checked = $false
	$GroupbySenderDomain.Checked = $false
	$GroupbyDate.Checked = $false
	})
$form.Controls.Add($GroupByReciever)

$GroupByRecieverDomain =  new-object System.Windows.Forms.CheckBox
$GroupByRecieverDomain.Location = new-object System.Drawing.Size(820,64)
$GroupByRecieverDomain.Size = new-object System.Drawing.Size(205,28)
$GroupByRecieverDomain.Text = "Group By Recv Domain Size,Num"
$GroupByRecieverDomain.Add_Click({
	$GroupbySender.Checked = $false
	$GroupbyReciever.Checked = $false
	$GroupbySenderDomain.Checked = $false
	$GroupbyDate.Checked = $false
})
$form.Controls.Add($GroupByRecieverDomain)

$GroupBySenderDomain =  new-object System.Windows.Forms.CheckBox
$GroupBySenderDomain.Location = new-object System.Drawing.Size(820,86)
$GroupBySenderDomain.Size = new-object System.Drawing.Size(205,28)
$GroupBySenderDomain.Text = "Group By Send Domain Size,Num"
$GroupBySenderDomain.Add_Click({
	$GroupbySender.Checked = $false
	$GroupbyReciever.Checked = $false
	$GroupByRecieverDomain.Checked = $false
	$GroupbyDate.Checked = $false
})
$form.Controls.Add($GroupBySenderDomain)

$GroupByDate =  new-object System.Windows.Forms.CheckBox
$GroupByDate.Location = new-object System.Drawing.Size(820,108)
$GroupByDate.Size = new-object System.Drawing.Size(205,28)
$GroupByDate.Text = "Group By Date Size,Num"
$GroupByDate.Add_Click({
	$GroupbySender.Checked = $false
	$GroupbyReciever.Checked = $false
	$GroupbyRecieverDomain.Checked = $false
	$GroupbySenderDomain.Checked = $false
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
