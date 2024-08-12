[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

function UpdateBypassSenders(){
	$newcombCollection = @()
	foreach($Item in $dgDataGrid.Rows){
		if ($item.Cells[0].Value -ne $null){
			if ($item.Cells[0].Value -ne ""){
				$newcombCollection += $item.Cells[0].Value
			}
		}
	}
	Set-ContentFilterConfig -BypassedSenders $newcombCollection
	$fsTable.Clear()
	$contentFilterConfig = get-ContentFilterConfig
	foreach ($ent in $contentFilterConfig.BypassedSenders){
		$fsTable.Rows.add($ent.ToString())	
	}
	$dgDataGrid.DataSource = $fsTable
}

function UpdateBypassSenderDomains(){
	$newcombCollection = @()
	foreach($Item in $dgDataGrid1.Rows){
		if ($item.Cells[0].Value -ne $null){
			if ($item.Cells[0].Value -ne ""){
				$newcombCollection += $item.Cells[0].Value
			}
		}
	}
	Set-ContentFilterConfig -BypassedSenderDomains $newcombCollection
	$fsTable1.Clear()
	$contentFilterConfig = get-ContentFilterConfig
	foreach ($ent in $contentFilterConfig.BypassedSenderDomains){
		$fsTable1.Rows.add($ent.ToString())	
	}
	$dgDataGrid1.DataSource = $fsTable1
}

$form = new-object System.Windows.Forms.form 

$contentFilterConfig = get-ContentFilterConfig
$Dataset = New-Object System.Data.DataSet
$fsTable = New-Object System.Data.DataTable
$fsTable.TableName = "BypasssedSender"
$fsTable.Columns.Add("WhitelistEntry")
$BypassSenderArray = $contentFilterConfig.BypassedSenders.ToString().Split(',')

foreach ($ent in $contentFilterConfig.BypassedSenders){
	$fsTable.Rows.add($ent.ToString())	
}

$fsTable1 = New-Object System.Data.DataTable
$fsTable1.TableName = "BypassedSenderDomains"
$fsTable1.Columns.Add("WhitelistEntry")

foreach ($ent in $contentFilterConfig.BypassedSenderDomains){
	$fsTable1.Rows.add($ent.ToString())	
}

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(20,50) 
$dgDataGrid.size = new-object System.Drawing.Size(250,450)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)
$dgDataGrid.DataSource = $fsTable

$dgDataGrid1 = new-object System.windows.forms.DataGridView
$dgDataGrid1.Location = new-object System.Drawing.Size(320,50) 
$dgDataGrid1.size = new-object System.Drawing.Size(250,450)
$dgDataGrid1.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid1)
$dgDataGrid1.DataSource = $fsTable1




# Add Update Button

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(20,530)
$exButton.Size = new-object System.Drawing.Size(150,20)
$exButton.Text = "Update - Senders"
$exButton.Add_Click({UpdateBypassSenders})
$form.Controls.Add($exButton)

# Add Update  Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(320,530)
$exButton1.Size = new-object System.Drawing.Size(150,20)
$exButton1.Text = "Update - Domains"
$exButton1.Add_Click({UpdateBypassSenderDomains})
$form.Controls.Add($exButton1)

$Gbox =  new-object System.Windows.Forms.GroupBox
$Gbox.Location = new-object System.Drawing.Size(10,15)
$Gbox.Size = new-object System.Drawing.Size(270,550)
$Gbox.Text = "By-Passed Senders"
$form.Controls.Add($Gbox)

$Gbox1 =  new-object System.Windows.Forms.GroupBox
$Gbox1.Location = new-object System.Drawing.Size(310,15)
$Gbox1.Size = new-object System.Drawing.Size(270,550)
$Gbox1.Text = "By-Passed Sender Domains"
$form.Controls.Add($Gbox1)

$form.Text = "Exchange 2007 WhiteList Form"
$form.size = new-object System.Drawing.Size(600,600) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()