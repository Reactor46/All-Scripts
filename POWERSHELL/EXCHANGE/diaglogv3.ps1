[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
Function getdiagvalues{
$daTable.clear()
$dgDataGrid.DataSource = $daTable
get-eventloglevel -Server $snServerNameDrop.SelectedItem.ToString() | ForEach-Object{
$daTable.Rows.Add($_.Identity,$_.EventLevel)
}
$dgDataGrid.DataSource = $daTable
}

Function UpdateLogLevel{
if ($dgDataGrid.SelectedRows.Count -eq 0){	 
	$idIdtoSet =  $dgDataGrid.Rows[$dgDataGrid.CurrentCell.RowIndex].Cells[0].Value 
	set-eventloglevel -Identity $idIdtoSet -Level $llLoglevelDrop.Text
	$dgDataGrid.Rows[$dgDataGrid.CurrentCell.RowIndex].Cells[1].Value = $llLoglevelDrop.Text
	
}
else{
	$msgbox = new-object -comobject wscript.shell
	$lcLoopCount = 0
	while ($lcLoopCount -le ($dgDataGrid.SelectedRows.Count-1)) {
	set-eventloglevel -Identity $dgDataGrid.SelectedRows[$lcLoopCount].Cells[0].Value -Level $llLoglevelDrop.SelectedItem
	$dgDataGrid.SelectedRows[$lcLoopCount].Cells[1].Value = $llLoglevelDrop.SelectedItem
	$lcLoopCount += 1
	}
}
}

$form = new-object System.Windows.Forms.form 

$llLableloc = 50
$VlLoc = 50

$Dataset = New-Object System.Data.DataSet
$daTable = New-Object System.Data.DataTable
$daTable.TableName = "Diag"
$daTable.Columns.Add("Identity")
$daTable.Columns.Add("Current Setting")


# Add Server DropLable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,20) 
$snServerNamelableBox.size = new-object System.Drawing.Size(100,20) 
$snServerNamelableBox.Text = "ServerName"
$form.Controls.Add($snServerNamelableBox) 

# Add Server Drop Down
$snServerNameDrop = new-object System.Windows.Forms.ComboBox
$snServerNameDrop.Location = new-object System.Drawing.Size(130,20)
$snServerNameDrop.Size = new-object System.Drawing.Size(130,30)
get-exchangeserver | ForEach-Object{$snServerNameDrop.Items.Add($_.Name)}
$snServerNameDrop.Add_SelectedValueChanged({getdiagvalues})  
$form.Controls.Add($snServerNameDrop)

# Add New Log Level Drop Down
$llLoglevelDrop = new-object System.Windows.Forms.ComboBox
$llLoglevelDrop.Location = new-object System.Drawing.Size(350,20)
$llLoglevelDrop.Size = new-object System.Drawing.Size(70,30)
$llLoglevelDrop.Items.Add("Lowest")
$llLoglevelDrop.Items.Add("Low")
$llLoglevelDrop.Items.Add("Medium")
$llLoglevelDrop.Items.Add("High")
$llLoglevelDrop.Items.Add("Expert")
$form.Controls.Add($llLoglevelDrop)

# Add Apply Button

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(430,20)
$exButton.Size = new-object System.Drawing.Size(60,20)
$exButton.Text = "Apply"
$exButton.Add_Click({UpdateLogLevel})
$form.Controls.Add($exButton)

# New setting Group Box

$OfGbox =  new-object System.Windows.Forms.GroupBox
$OfGbox.Location = new-object System.Drawing.Size(300,0)
$OfGbox.Size = new-object System.Drawing.Size(200,50)
$OfGbox.Text = "New Log Level Settings"
$form.Controls.Add($OfGbox)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,80) 
$dgDataGrid.size = new-object System.Drawing.Size(500,500) 
$dgDataGrid.AutoSizeColumnsMode = "AllCells"
$dgDataGrid.SelectionMode = "FullRowSelect"
$form.Controls.Add($dgDataGrid)


$form.Text = "Exchange 2007 Diagnostic Logging Form"
$form.size = new-object System.Drawing.Size(600,600) 
$form.autoscroll = $true
$form.topmost = $true

$form.ShowDialog() 





