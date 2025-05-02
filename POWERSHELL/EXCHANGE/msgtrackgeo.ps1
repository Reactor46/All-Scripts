[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

function createdb{
$aoADOXDb = new-object -com ADOX.Catalog
$aoADOXDb.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + $dbfilepath)  
$atADOXTable = new-object -com ADOX.Table
$atADOXTable.Name = "IPTOCOUNTRY"
$atADOXTable.Columns.Append("IPFrom", $adDouble)
$atADOXTable.Columns.Append("IPTo", $adDouble)
$atADOXTable.Columns.Append("Registry", $adVarWChar, 25)
$atADOXTable.Columns.Append("ASSIGNED", $adDouble)
$atADOXTable.Columns.Append("CTRY", $adVarWChar, 2)
$atADOXTable.Columns.Append("CNTRY", $adVarWChar, 3)
$atADOXTable.Columns.Append("COUNTRY", $adVarWChar, 100)
$atindex = new-object -com ADOX.index
$atindex.Name = "idxIPFrom"
$atindex.Columns.Append("IPFrom")
$atindex.Unique = $True
$atADOXTable.Indexes.Append($atindex)

$atindex1 = new-object -com ADOX.index
$atindex1.Name = "idxIPTo"
$atindex1.Columns.Append("IPTo")
$atindex1.Unique = $True
$atADOXTable.Indexes.Append($atindex1)

$aoADOXDb.Tables.Append($atADOXTable) 
$atADOXTablever = new-object -com ADOX.Table
$atADOXTablever.Name = "Version"
$atADOXTablever.Columns.Append("FileModifiedDate", $adVarWChar, 255)
$aoADOXDb.Tables.Append($atADOXTablever) 
# Cleanup
$aoADOXDb.ActiveConnection.close()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$atindex)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$atindex1)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$atADOXTable)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$atADOXTablever)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$aoADOXDb)
[system.gc]::Collect()
$atindex = $null
$atindex1 = $null
$atADOXTable = $null
$atADOXTablever = $null
$aoADOXDb = $null

}

function populatedb{

$file = get-item $importfilepath
$ocOdbcConnection.Open()
$dcOdBcommand = new-object System.Data.OleDb.OleDbCommand
$dcOdBcommand.connection = $ocOdbcConnection
$dcOdBcommand.commandtext = "Insert into Version values('" + $file.lastwritetime + "')"
$dcOdBcommand.ExecuteNonQuery()
$rcRowCount = 0 
"Filling Database this may take a few minutes"
get-content $importfilepath  | %{
	$linarr = $_.replace("'","``").split(",")
	if ($linarr[0].indexofany("#") -band $linarr.length -gt 1){
	$stSQLStatement = "Insert into IPTOCOUNTRY values('" +  $linarr[0].replace("`"","") + "','" `
	+  $linarr[1].replace("`"","") + "','"+  $linarr[2].replace("`"","") + "','"+ `
	 $linarr[3].replace("`"","") + "','" +  $linarr[4].replace("`"","") + "','" + `
	 $linarr[5].replace("`"","") + "','" +  $linarr[6].replace("`"","") + "')"
	$dcOdBcommand.commandtext = $stSQLStatement
	$inResult = $dcOdBcommand.ExecuteNonQuery()
	if ($inResult -ne 1){$inResult}
	}
	$rcRowCount = $rcRowCount + 1
	if ($rcRowCount -eq 10000){
		$rcRowCount = 0
		"10000 Rows Inserted"
		}
}
"Fill completed"
$dcOdBcommand = $null
$ocOdbcConnection.Close()
$ocOdbcConnection =  $null
$jrJroobj = new-object -com JRO.JetEngine
$dbSourceDB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + $dbfilepath
$dbDestinationDB = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + $tmpdbfilepath
$jrJroobj.CompactDatabase($dbSourceDB, $dbDestinationDB)
remove-item $dbfilepath
copy-item $tmpdbfilepath $dbfilepath
remove-item $tmpdbfilepath
}

function checkfileversion{
	$ocOdbcConnection.Open()
	$dcOdBcommand = new-object System.Data.OleDb.OleDbCommand
	$dcOdBcommand.connection = $ocOdbcConnection
	$dcOdBcommand.commandtext = "select FileModifiedDate from Version"
	$drDBreader = $dcOdBcommand.ExecuteReader()
	$retval = 0
	while ($drDBreader.read()){
		if ($drDBreader[0] -ne $file.lastwritetime){$retval = 1} 
		else{"DB up to date"}
	}
	$dcOdBcommand = $null
	$ocOdbcConnection.Close()
	return $retval
}

function CheckExtraSettings{
$extrasettings = 0
if ($GroupbyCountry.Checked -eq $true) {$extrasettings = 1}
if ($GroupBysidebyside.Checked -eq $true) {$extrasettings = 1}
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

function AggResultsSBS([string]$saAddress){
	if ($gbhash1.ContainsKey($saAddress)){
		if ($EntryType -eq 1019){
			$tsize = [int]$gbhash2[$saAddress] + [int]$svSizeVal
			$tnum =  [int]$gbhash1[$saAddress]  + 1
			$gbhash1[$saAddress] = $tnum
			$gbhash2[$saAddress] = $tsize
		}	
		else {
			$tsize = [int]$gbhash4[$saAddress] + [int]$svSizeVal
			$tnum =  [int]$gbhash3[$saAddress]  + 1
			$gbhash3[$saAddress] = $tnum
			$gbhash4[$saAddress] = $tsize		
		}
	}
	else{
		if ($EntryType -eq 1019){
			$gbhash1.Add($saAddress,1)
			$gbhash2.Add($saAddress,$svSizeVal)
			$gbhash3.Add($saAddress,0)
			$gbhash4.Add($saAddress,0)}
		else {
			$gbhash1.Add($saAddress,0)
			$gbhash2.Add($saAddress,0)
			$gbhash3.Add($saAddress,1)
			$gbhash4.Add($saAddress,$svSizeVal)
		}
	}
}

function getCountry($iptoQuery){
$ipa = $iptoQuery.split(".")
$iptoQuerynum = ([int]$ipa[0] *16777216) + ([int]$ipa[1] *65536) + ([int]$ipa[2] *256) + [int]$ipa[4]
$iptoQueryCache = ([int]$ipa[0] *16777216) + ([int]$ipa[1] *65536) + ([int]$ipa[2] *256)
if ($ipHash.ContainsKey($iptoQueryCache)){
	$cntry = $ipHash[$iptoQueryCache] 	
}
else{
$qsQueryString = "Select * FROM IPTOCOUNTRY where IPFrom <= " + $iptoQuerynum + " and IPTo >= " + $iptoQuerynum 
$cntry = "Not Defined"
$dcOdbcCommand = new-object System.Data.OleDb.OleDbCommand($qsQueryString,$ocOdbcConnection)
$drOdbcDataReader =  $dcOdbcCommand.ExecuteReader() 
while ($drOdbcDataReader.read()){
	if ([int64]$iptoQuerynum -ge [int64]$drOdbcDataReader[0] -band [int64]$iptoQuerynum -le [int64]$drOdbcDataReader[1]){
	$cntry =    $drOdbcDataReader[6] 
	$ipHash.Add($iptoQueryCache,$drOdbcDataReader[6])
	if ($ipHash.ContainsKey([int64]$drOdbcDataReader[0])){}else{$ipHash.Add([int64]$drOdbcDataReader[0],$drOdbcDataReader[6])}
	if ($ipHash.ContainsKey([int64]$drOdbcDataReader[1])){}else{$ipHash.Add([int64]$drOdbcDataReader[1],$drOdbcDataReader[6])}
	}
}
}
return $cntry 
}

function adddata{

$gbhash1.Clear()
$gbhash2.Clear()
$gbhash3.Clear()
$gbhash4.Clear()
$ipTable.Clear()
$cntTable.Clear()
$SBSTable.Clear()
$csConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" + $dbfilepath
$ocOdbcConnection = new-object System.Data.OleDb.OleDbConnection($csConnectionString)
$ocOdbcConnection.open()

$extrasettings = CheckExtraSettings
$servername = $snServerNameTextBox.text
$dchash = @{ }
$dtQueryDT = New-Object System.DateTime $dpTimeFrom.value.year,$dpTimeFrom.value.month,$dpTimeFrom.value.day,$dpTimeFrom2.value.hour,$dpTimeFrom2.value.minute,$dpTimeFrom2.value.second
$dtQueryDTf =  New-Object System.DateTime $dpTimeFrom1.value.year,$dpTimeFrom1.value.month,$dpTimeFrom1.value.day,$dpTimeFrom3.value.hour,$dpTimeFrom3.value.minute,$dpTimeFrom3.value.second
$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($dtQueryDT.ToUniversalTime())
$WmidtQueryDTf = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($dtQueryDTf.ToUniversalTime())
$WmiNamespace = "ROOT\MicrosoftExchangev2"
if ($etTypeCheckDrop.SelectedItem -eq "Spam"){$filter =  "OriginationTime <= '" + $WmidtQueryDTf + "' and OriginationTime >=  '" + $WmidtQueryDT + "' and entrytype = '1039' or OriginationTime <= '" + $WmidtQueryDTf + "' and OriginationTime >=  '" + $WmidtQueryDT + "' and entrytype = '1040'"}
else{$filter =  "OriginationTime <= '" + $WmidtQueryDTf + "' and OriginationTime >=  '" + $WmidtQueryDT + "' and entrytype = '1019'" 
}
if ($GroupBysidebyside.Checked -eq $true){$filter =  "OriginationTime <= '" + $WmidtQueryDTf + "' and OriginationTime >=  '" + $WmidtQueryDT + "' `
	and entrytype = '1039' or OriginationTime <= '" + $WmidtQueryDTf + "' and OriginationTime >=  '" + $WmidtQueryDT + "' and entrytype = '1040' or " + `
	 "OriginationTime <= '" + $WmidtQueryDTf + "' and OriginationTime >=  '" + $WmidtQueryDT + "' and entrytype = '1019'" 
}
get-wmiobject -class Exchange_MessageTrackingEntry -Namespace $WmiNamespace -ComputerName $servername -filter $filter | ForEach-Object{ 
$recpval = ""
foreach($recp in $_.RecipientAddress){if ($recpval -eq ""){$recpval = $recp}
	else {$recpval = $recpval + ";" + $recp}}
if($dchash.ContainsKey($_.MessageID)){}
else{
if ($_.ClientIP.length -gt 0){
	$svSizeVal = $_.size/1024
	$Country = getCountry($_.ClientIP)
	if ($extrasettings -eq 0){	
		$ipTable.Rows.Add([System.Management.ManagementDateTimeConverter]::ToDateTime($_.OriginationTime),$_.SenderAddress,$recpval,$_.subject,[int]$svSizeVal.ToString(0.00),$Country)}
	if ($GroupbyCountry.Checked -eq $true){
		AggResults($Country)
	}
	if ($GroupBysidebyside.Checked -eq $true){
		$EntryType = $_.EntryType
		AggResultsSBS($Country)
	}
}
$dchash.Add($_.MessageID,1)
}	
}

if ($extrasettings -eq 0){$dgDataGrid.DataSource = $ipTable}
else{
	if($GroupBysidebyside.Checked -eq $true){
		foreach ($htent in $gbhash1.keys){
		$sbsTable.Rows.Add($htent,[int]$gbhash1[$htent],[int]$gbhash2[$htent],[int]$gbhash3[$htent],[int]$gbhash4[$htent])
	}
	$dgDataGrid.DataSource = $sbsTable

	}
	else {
	foreach ($htent in $gbhash1.keys){
		$cntTable.Rows.Add($htent,[int]$gbhash1[$htent],[int]$gbhash2[$htent])
	}
	$dgDataGrid.DataSource = $cntTable
	}

}
$ocOdbcConnection.close

}

function Exportcsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.ShowDialog()
$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
$logfile.WriteLine("OriginationTime,SenderAddress,RecipientAddress,Subject,Size,Country")
foreach($row in $ipTable.Rows){
	$logfile.WriteLine($row)
}
$logfile.Close()

}

$form = new-object System.Windows.Forms.form 
$form.Text = "IP Location Message Tracker"
$ipHash = @{ }
$gbhash1 = @{ }
$gbhash2 = @{ }
$gbhash3 = @{ }
$gbhash4 = @{ }
$Dataset = New-Object System.Data.DataSet
$ipTable = New-Object System.Data.DataTable
$ipTable.TableName = "IPs"
$ipTable.Columns.Add("Origination Time")
$ipTable.Columns.Add("SenderAddress")
$ipTable.Columns.Add("RecipientAddress")
$ipTable.Columns.Add("Subject")
$ipTable.Columns.Add("Size")
$ipTable.Columns.Add("Country")
$Dataset.tables.add($ipTable)
$cntTable = New-Object System.Data.DataTable
$cntTable.TableName = "Contry-Table"
$cntTable.Columns.Add("Country")
$cntTable.Columns.Add("Number_Messages",[int])
$cntTable.Columns.Add("Size (KB)",[int])
$Dataset.tables.add($cntTable)
$sbsTable = New-Object System.Data.DataTable
$sbsTable.TableName = "Side by Side Contry-Table"
$sbsTable.Columns.Add("Country")
$sbsTable.Columns.Add("Normal Number_Messages",[int])
$sbsTable.Columns.Add("Normal Size (KB)",[int])
$sbsTable.Columns.Add("Spam Number_Messages",[int])
$sbsTable.Columns.Add("Spam Size (KB)",[int])
$Dataset.tables.add($sbsTable)


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
$exButton.Location = new-object System.Drawing.Size(10,75)
$exButton.Size = new-object System.Drawing.Size(85,20)
$exButton.Text = "Search"
$exButton.Add_Click({adddata})
$form.Controls.Add($exButton)

# Add Export Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(100,75)
$exButton1.Size = new-object System.Drawing.Size(85,20)
$exButton1.Text = "Export"
$exButton1.Add_Click({Exportcsv})
$form.Controls.Add($exButton1)

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

# Add Type Clause

$etTypelableBox = new-object System.Windows.Forms.Label
$etTypelableBox.Location = new-object System.Drawing.Size(320,70) 
$etTypelableBox.size = new-object System.Drawing.Size(90,20) 
$etTypelableBox.Text = "Email Type"
$form.Controls.Add($etTypelableBox) 

$etTypeCheckDrop = new-object System.Windows.Forms.ComboBox
$etTypeCheckDrop.Location = new-object System.Drawing.Size(410,70)
$etTypeCheckDrop.Size = new-object System.Drawing.Size(70,30)
$etTypeCheckDrop.Text = "Normal"
$etTypeCheckDrop.Items.Add("Normal")
$etTypeCheckDrop.Items.Add("Spam")
$form.Controls.Add($etTypeCheckDrop)

# Add Extras

$GroupbyCountry =  new-object System.Windows.Forms.CheckBox
$GroupbyCountry.Location = new-object System.Drawing.Size(820,20)
$GroupbyCountry.Size = new-object System.Drawing.Size(200,30)
$GroupbyCountry.Text = "Group By Country Size,Number"
$GroupbyCountry.Add_Click({
	$GroupBysidebyside.Checked = $false
	})
$form.Controls.Add($GroupbyCountry)

$GroupBysidebyside =  new-object System.Windows.Forms.CheckBox
$GroupBysidebyside.Location = new-object System.Drawing.Size(820,42)
$GroupBysidebyside.Size = new-object System.Drawing.Size(200,30)
$GroupBysidebyside.Text = "Group By Country Side By Side"
$GroupBysidebyside.Add_Click({
	$GroupbyCountry.Checked = $false
	})
$form.Controls.Add($GroupBysidebyside)

$Gbox =  new-object System.Windows.Forms.GroupBox
$Gbox.Location = new-object System.Drawing.Size(810,5)
$Gbox.Size = new-object System.Drawing.Size(220,85)
$Gbox.Text = "Extras"
$form.Controls.Add($Gbox)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGrid

$dgDataGrid.AllowSorting = $True
$dgDataGrid.Location = new-object System.Drawing.Size(10,100) 
$dgDataGrid.size = new-object System.Drawing.Size(1024,700) 

$form.Controls.Add($dgDataGrid)

#populate DataGrid

# MDB Database routines
$dbfilepath = "c:\temp\iptocountry.mdb"
$tmpdbfilepath = "c:\temp\iptocountrycomp.mdb"
$importfilepath = "c:\temp\iptocountry.csv"
$adDouble = 5
$adVarWChar = 202
$csConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;data source=" + $dbfilepath 
$ocOdbcConnection = new-object System.Data.OleDb.OleDbConnection($csConnectionString)
$file = get-item $importfilepath -ea silentlycontinue 
$dbfile = (resolve-path $dbfilepath  -ea silentlycontinue).path 
if ( ! $dbfile ) {  
	createdb
	populatedb
}
else{
	$vchk = checkfileversion
	$vchk
	if ($vchk -eq 1){
		$dcOdBcommand = $null
		$ocOdbcConnection = $null
		remove-item $dbfilepath 
		$ocOdbcConnection = new-object System.Data.OleDb.OleDbConnection($csConnectionString)
		createdb
		populatedb
	}
}
$dcOdBcommand = $null
$ocOdbcConnection = $null

$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
