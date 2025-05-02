[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 


function getMailboxSizes(){
$msTable.clear()

if ($mtTypeDrop.SelectedItem -ne $null){
	if ($mtTypeDrop.SelectedItem.ToString() -eq "Disconnected"){
		get-mailboxstatistics -Server $snServerNameDrop.SelectedItem.ToString() | Where {$_.DisconnectDate -ne $null} | ForEach-Object{
	        $icount = 0
		$tisize = 0
		$disize = 0
		if ($_.DisplayName -ne $null){$dname = $_.DisplayName}
		if ($_.ItemCount -ne $null){$icount = $_.ItemCount}
		if ($_.TotalItemSize.Value.ToMB() -ne $null){$tisize = $_.TotalItemSize.Value.ToMB()}
		if ($_.TotalDeletedItemSize.Value.ToKB() -ne $null){$disize = $_.TotalDeletedItemSize.Value.ToKB()}    
		$msTable.Rows.add($dname,$icount,$tisize,$disize)
		}
	}
	else{	get-mailboxstatistics -Server $snServerNameDrop.SelectedItem.ToString() | Where {$_.DisconnectDate -eq $null} | ForEach-Object{
	        $icount = 0
		$tisize = 0
		$disize = 0
		if ($_.DisplayName -ne $null){$dname = $_.DisplayName}
		if ($_.ItemCount -ne $null){$icount = $_.ItemCount}
		if ($_.TotalItemSize.Value.ToMB() -ne $null){$tisize = $_.TotalItemSize.Value.ToMB()}
		if ($_.TotalDeletedItemSize.Value.ToKB() -ne $null){$disize = $_.TotalDeletedItemSize.Value.ToKB()}    
		$msTable.Rows.add($dname,$icount,$tisize,$disize)
		}

	}
}
else{
		get-mailboxstatistics -Server $snServerNameDrop.SelectedItem.ToString() | ForEach-Object{
	        $icount = 0
		$tisize = 0
		$disize = 0
		if ($_.DisplayName -ne $null){$dname = $_.DisplayName}
		if ($_.ItemCount -ne $null){$icount = $_.ItemCount}
		if ($_.TotalItemSize.Value.ToMB() -ne $null){$tisize = $_.TotalItemSize.Value.ToMB()}
		if ($_.TotalDeletedItemSize.Value.ToKB() -ne $null){$disize = $_.TotalDeletedItemSize.Value.ToKB()}    
		$msTable.Rows.add($dname,$icount,$tisize,$disize)
	}

} 
write-host $fstring 

$dgDataGrid.DataSource = $msTable

}


function GetFolderSizes(){
$fsTable.clear()
$snServername = $snServerNameDrop.SelectedItem.ToString()
write-host $dgDataGrid.CurrentCell.RowIndex
$siSIDToSearch = get-user $msTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0]
write-host $siSIDToSearch.SamAccountName.ToString()
Get-MailboxFolderStatistics $siSIDToSearch.SamAccountName.ToString() | ForEach-Object{
	$ficount = 0
	$fisize = 0
	$fsisize = 0
	$fscount = 0
	$fname = $_.Name
	if ($_.FolderSize -ne $null){$fsisize = [math]::round(($_.FolderSize/1mb),2)}
	if ($_.ItemsInFolder -ne $null){$ficount = $_.ItemsInFolder}
	if ($_.ItemsInFolderAndSubfolders -ne $null){$fscount = $_.ItemsInFolderAndSubfolders} 
	if ($_.FolderAndSubfolderSize -ne $null){$fsisize = [math]::round(($_.FolderAndSubfolderSize/1mb),2)}      
	$fsTable.Rows.add($fname,$ficount,$fsisize,$fscount,$fsisize)
}
$dgDataGrid1.DataSource = $fsTable
}

function ExportMBcsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("UserName,# Items,MB Size(MB),DelItems(KB)")
	foreach($row in $msTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString()) 
	}
	$logfile.Close()
}
}

function ExportFScsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("DisplayName,# Items,Folder Size(MB),# Items + Sub,Folder Size + Sub(MB)")
	foreach($row in $fsTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString()) 
	}
	$logfile.Close()
}
}

$form = new-object System.Windows.Forms.form 
$global:LastFolder = ""
# Add DataTable

$Dataset = New-Object System.Data.DataSet
$fsTable = New-Object System.Data.DataTable
$fsTable.TableName = "Folder Sizes"
$fsTable.Columns.Add("DisplayName")
$fsTable.Columns.Add("# Items",[int64])
$fsTable.Columns.Add("Folder Size(MB)",[int64])
$fsTable.Columns.Add("# Items + Sub",[int64])
$fsTable.Columns.Add("Folder Size + Sub(MB)",[int64])
$Dataset.tables.add($fsTable)

$msTable = New-Object System.Data.DataTable
$msTable.TableName = "Mailbox Sizes"
$msTable.Columns.Add("UserName")
$msTable.Columns.Add("# Items")
$msTable.Columns.Add("MB Size(MB)",[int64])
$msTable.Columns.Add("DelItems(KB)",[int64])
$Dataset.tables.add($msTable)

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
$snServerNameDrop.Add_SelectedValueChanged({getMailboxSizes})  
$form.Controls.Add($snServerNameDrop)

# Add Mailbox Type DropLable
$mtTypeDroplableBox = new-object System.Windows.Forms.Label
$mtTypeDroplableBox.Location = new-object System.Drawing.Size(200,20) 
$mtTypeDroplableBox.size = new-object System.Drawing.Size(80,20) 
$mtTypeDroplableBox.Text = "MailboxType"
$form.Controls.Add($mtTypeDroplableBox) 

# Add Mailbox Type Drop Down
$mtTypeDrop = new-object System.Windows.Forms.ComboBox
$mtTypeDrop.Location = new-object System.Drawing.Size(290,20)
$mtTypeDrop.Size = new-object System.Drawing.Size(100,30)
$mtTypeDrop.Items.Add("Disconnected")
$mtTypeDrop.Items.Add("Connected")
$mtTypeDrop.Add_SelectedValueChanged({if ($snServerNameDrop.SelectedItem -ne $null){getMailboxSizes}})  
$form.Controls.Add($mtTypeDrop)

# Add Export MB Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(10,560)
$exButton1.Size = new-object System.Drawing.Size(125,20)
$exButton1.Text = "Export Mailbox Grid"
$exButton1.Add_Click({ExportMBcsv})
$form.Controls.Add($exButton1)

# Add Export FG Button

$exButton2 = new-object System.Windows.Forms.Button
$exButton2.Location = new-object System.Drawing.Size(500,560)
$exButton2.Size = new-object System.Drawing.Size(135,20)
$exButton2.Text = "Export FolderSize Grid"
$exButton2.Add_Click({ExportFScsv})
$form.Controls.Add($exButton2)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,50) 
$dgDataGrid.size = new-object System.Drawing.Size(450,500)
$form.Controls.Add($dgDataGrid)

$dgDataGrid1 = new-object System.windows.forms.DataGridView
$dgDataGrid1.Location = new-object System.Drawing.Size(500,50) 
$dgDataGrid1.size = new-object System.Drawing.Size(450,500)
$form.Controls.Add($dgDataGrid1)

# folder Size Button

$fsizeButton = new-object System.Windows.Forms.Button
$fsizeButton.Location = new-object System.Drawing.Size(500,19)
$fsizeButton.Size = new-object System.Drawing.Size(120,23)
$fsizeButton.Text = "Get Folder Size"
$fsizeButton.visible = $True
$fsizeButton.Add_Click({GetFolderSizes})
$form.Controls.Add($fsizeButton)



$form.Text = "Exchange 2007 Mailbox Size Form"
$form.size = new-object System.Drawing.Size(1000,620) 
$form.autoscroll = $true
$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
