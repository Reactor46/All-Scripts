[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
$form = new-object System.Windows.Forms.form 
function processmailboxes(){
$fsTable.clear()
get-mailbox -Server $snServerNameDrop.SelectedItem.ToString() -ResultSize Unlimited | ForEach-Object{
	$siSIDToSearch = get-user $_.DisplayName
	write-host $_.DisplayName
	Get-MailboxFolderStatistics $siSIDToSearch.SamAccountName.ToString() -IncludeOldestAndNewestItems | ForEach-Object{
		$ficount = 0
		$flastaccess = ""
		$fisize = 0
		$fsisize = 0
		$fscount = 0
		$fname = $_.Name
		if ($fnhash.ContainsKey($_.Name)){
			$fnhash[$_.Name] = [int]$fnhash[$_.Name] + 1
		} 
		else {
		    if ($_.Name -ne $null){ $fnhash.add($_.Name,1)}
		}
		if ($_.FolderSize -ne $null){$fsisize = [math]::round(($_.FolderSize/1mb),2)}
		if ($_.ItemsInFolder -ne $null){$ficount = $_.ItemsInFolder}
		if ($_.NewestItemReceivedDate -ne $null){$flastaccess = $_.NewestItemReceivedDate}
		if ($_.ItemsInFolderAndSubfolders -ne $null){$fscount = $_.ItemsInFolderAndSubfolders} 
		if ($_.FolderAndSubfolderSize -ne $null){$fsisize = [math]::round(($_.FolderAndSubfolderSize/1mb),2)}      
		$fsTable.Rows.add($siSIDToSearch.SamAccountName.ToString(),$siSIDToSearch.DisplayName.ToString(),$fname,$ficount,$fsisize,$fscount,$fsisize,$flastaccess)

	}
}
	$arylst = new-object System.Collections.ArrayList
	$Dataveiw.RowFilter =  "FolderName = 'Inbox'"
	$dgDataGrid.DataSource = $Dataveiw
	foreach ($key in $fnhash.keys){
		if ($fnhash[$key] -gt 5) {
			$arylst.Add($key)
			write-host $key	
	}}
	$arylst.Sort()
	foreach ($val in $arylst){
		$fnTypeDrop.Items.Add($val)
	}
	
}

function selectfolder(){
	$Dataveiw.RowFilter =  "FolderName = '" + $fnTypeDrop.SelectedItem.ToString() + "'"
	$dgDataGrid.DataSource = $Dataveiw

}

function ExportFScsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("SamAccountName,DisplayName,FolderName,# Items,Folder Size(MB),# Items + Sub,Folder Size + Sub(MB),Last Item Received")
	foreach($row in $dgDataGrid.Rows){
		if ($row.cells[0].Value -ne $null){
			$logfile.WriteLine("`"" + $row.cells[0].Value.ToString() + "`",`"" + $row.cells[1].Value.ToString() + "`",`"" + $row.cells[2].Value.ToString() + "`"," + $row.cells[3].Value.ToString() + "," + $row.cells[4].Value.ToString() + "," + $row.cells[5].Value.ToString() + "," + $row.cells[6].Value.ToString() + "," + $row.cells[7].Value.ToString())
		} 
	}
	$logfile.Close()
}
}

$fnhash = @{ }
$Dataset = New-Object System.Data.DataSet
$fsTable = New-Object System.Data.DataTable
$fsTable.TableName = "Folder Sizes"
$fsTable.Columns.Add("SamAccountName")
$fsTable.Columns.Add("DisplayName")
$fsTable.Columns.Add("FolderName")
$fsTable.Columns.Add("# Items",[int64])
$fsTable.Columns.Add("Folder Size(MB)",[int64])
$fsTable.Columns.Add("# Items + Sub",[int64])
$fsTable.Columns.Add("Folder Size + Sub(MB)",[int64])
$fsTable.Columns.Add("Last Item Received")
$Dataset.tables.add($fsTable)
$Dataveiw = New-Object System.Data.DataView($fsTable)

# Add Server DropLable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,20) 
$snServerNamelableBox.size = new-object System.Drawing.Size(80,20) 
$snServerNamelableBox.Text = "Server Name"
$form.Controls.Add($snServerNamelableBox) 

# Add Server Drop Down
$snServerNameDrop = new-object System.Windows.Forms.ComboBox
$snServerNameDrop.Location = new-object System.Drawing.Size(90,20)
$snServerNameDrop.Size = new-object System.Drawing.Size(100,30)
get-mailboxserver | ForEach-Object{$snServerNameDrop.Items.Add($_.Name)}
$form.Controls.Add($snServerNameDrop)

# folder Size Button

$fsizeButton = new-object System.Windows.Forms.Button
$fsizeButton.Location = new-object System.Drawing.Size(200,19)
$fsizeButton.Size = new-object System.Drawing.Size(120,23)
$fsizeButton.Text = "Get Folder Sizes"
$fsizeButton.visible = $True
$fsizeButton.Add_Click({processmailboxes})
$form.Controls.Add($fsizeButton)

# Add Folder Name Type Drop Down
$fnTypeDrop = new-object System.Windows.Forms.ComboBox
$fnTypeDrop.Location = new-object System.Drawing.Size(350,20)
$fnTypeDrop.Size = new-object System.Drawing.Size(150,30)
$fnTypeDrop.Add_SelectedValueChanged({if ($snServerNameDrop.SelectedItem -ne $null){SelectFolder}})  
$form.Controls.Add($fnTypeDrop)

# Add Export FolderSize Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(10,660)
$exButton1.Size = new-object System.Drawing.Size(125,20)
$exButton1.Text = "Export Folder Sizes"
$exButton1.Add_Click({ExportFScsv})
$form.Controls.Add($exButton1)


# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,50) 
$dgDataGrid.size = new-object System.Drawing.Size(900,600)
$form.Controls.Add($dgDataGrid)

$form.Text = "Exchange 2007 Mailbox Folder Compare Folder Size Form"
$form.size = new-object System.Drawing.Size(1000,620) 
$form.autoscroll = $true
$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
