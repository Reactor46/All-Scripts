[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
$form = new-object System.Windows.Forms.form 

$usrinfo = @{ }
$ctab = @{ }
$agDepartment = @{ }
$agOffice  = @{ }
$htAg = @{ }
$ExcomCollection = @()
function Aggresults{
	$htAg.clear()
	if ($mtTypeDrop.SelectedItem -eq $null){
		$agMethod = "LastName"
	}
	else{
		$agMethod = $mtTypeDrop.SelectedItem.ToString()
	}
	foreach ($row in $usrinfo.Values){
		if ($row.$agMethod.ToString() -eq ""){
			if ($htAg.containskey("Not Set")){ 
				$htAg["Not Set"].Number = $htAg["Not Set"].Number + 1	
				$htAg["Not Set"].Size = $htAg["Not Set"].Size + [INT]$row.MailboxSize.ToString()	
			}
			else{
				$mbcnt = "" | select Name,Number,Size
				$mbcnt.Name = "Not Set"
				$mbcnt.Number = 1
				$mbcnt.Size = [INT]$row.MailboxSize.ToString()
				$htAg.Add("Not Set",$mbcnt)
			
			}
		}
		else{
			if ($htAg.containskey($row.$agMethod.ToString())){
				$htAg[$row.$agMethod.ToString()].Number = $htAg[$row.$agMethod.ToString()].Number + 1	
				$htAg[$row.$agMethod.ToString()].Size = $htAg[$row.$agMethod.ToString()].Size + [INT]$row.MailboxSize.ToString()	
			}
			else{
				$row.$agMethod.ToString()
				$mbcnt = "" | select Name,Number,Size
				$mbcnt.Name = $row.$agMethod.ToString()
				$mbcnt.Number = 1
				$mbcnt.Size = [INT]$row.MailboxSize.ToString()
				$htAg.Add($row.$agMethod.ToString(),$mbcnt)
			}
		}
	}
	$agtable.Clear()
	$valueBlock = ""
	$TitleBlock = ""
	$TitleBlock1 = ""
	foreach ($row in $htAg.Values){
		$agtable.rows.add($row.Name,$row.Number,$row.Size)
		if ($valueBlock -eq ""){$valueBlock = $row.Size.ToString()
			$lval = [INT]$row.Size
		}
		else {
			$valueBlock = $valueBlock + "," + $row.Size.ToString()
			if ($lval -lt [INT]$row.Size){$lval = [INT]$row.Size}
		}
		if ($TitleBlock -eq ""){$TitleBlock = $row.Name.ToString().Replace("&","-")
				       $TitleBlock1 = $row.Name.ToString().Replace("&","-")
		}
		else {$TitleBlock = $TitleBlock + "|" + $row.Name.ToString().Replace("&","-")
	  	    $TitleBlock1 = $row.Name.ToString().Replace("&","-") + "|" + $TitleBlock1
		}
	}
	$csString = "http://chart.apis.google.com/chart?cht=p3&chs=430x160&chds=0," + $lval + "&chd=t:" + $valueBlock + "&chl=" + $TitleBlock  + "&chco=0000ff,00ff00,ff0000,FFFFFF,000000"
	$pbox.ImageLocation = $csString
	$csString1 = "http://chart.apis.google.com/chart?cht=bhs&chs=530x300&chd=t:" + $valueBlock + "&chds=0," + ($lval+20)  + "&chxt=x,y&chbh=a&chxr=" + "&chxr=0,0," + ($lval+20) + "&chxl=1:|" + $TitleBlock1 + "&chco=0000ff,00ff00,ff0000,FFFFFF,000000"
	$pbox1.ImageLocation = $csString1
	$dgDataGrid.DataSource = $agtable
}

function getMailboxSizes(){

$mbcombCollection = @()
$usrinfo.clear()
$ctab.clear()
$agDepartment.clear()
$agOffice.clear()
$ExcomCollection = @()
$agtable.clear()

$serverName = $snServerNameDrop.SelectedItem.ToString()
if ($snServerNameDrop.SelectedItem.ToString() -eq "ALL Servers"){
	$mailboxes = get-mailbox -ResultSize Unlimited
}
else{
	$mailboxes = get-mailbox -server $snServerNameDrop.SelectedItem.ToString() -ResultSize Unlimited
} 
$mailboxes | foreach-object{
	$ctab.add($_.Guid.ToString(),$_.ExchangeGuid)
}
"Finished Get Mailbox"
Get-User -recipienttype UserMailbox -ResultSize Unlimited | foreach-object{
	if ($ctab.ContainsKey($_.Guid.ToString())){
		$usrobj = $_ | select *,MailboxSize,ExchangeGUID,OU,MailboxStore,StorageGroup
		$oustring = ""
		$idarry = $_.Identity.ToString().Split("/")
		for($i=1;$i-lt ($idarry.length-1);$i++){
			$oustring = $oustring + "\" + $idarry[$i] 
		}
		$usrobj.OU = $oustring
		$usrobj.ExchangeGUID = $ctab[$_.Guid.ToString()]
		$usrinfo.add($ctab[$_.Guid.ToString()].ToString(),$usrobj)	
	}
	
}
$mbServers = get-mailboxserver
if ($snServerNameDrop.SelectedItem.ToString() -eq "ALL Servers"){
	$mbServers | foreach-object{
	
		$mscombCollection += get-mailboxstatistics -server $_.Name  | Where {$_.DisconnectDate -eq $null}
			
		}
	}
	else{
		$mscombCollection += get-mailboxstatistics -server $snServerNameDrop.SelectedItem.ToString()  | Where {$_.DisconnectDate -eq $null}
	} 

$mscombCollection | ForEach-Object{
	$usrobj = $usrinfo[$_.MailboxGuid.ToString()]
	if ($usrinfo.ContainsKey($_.MailboxGuid.ToString())){	
		if ($_.TotalItemSize.Value -ne $null){
			$usrobj.MailboxSize = $_.TotalItemSize.Value.ToMB()
			$usrobj.MailboxStore = $_.DatabaseName
			$usrobj.StorageGroup = $_.StorageGroupName
		}
		else
		{
			$usrobj.MailboxSize = 0
			$usrobj.MailboxStore = $_.DatabaseName
			$usrobj.StorageGroup = $_.StorageGroupName
		}
	}
	
}
Aggresults	
}

function ShowMailboxes(){
	if ($mtTypeDrop.SelectedItem -eq $null){
		$agMethod = "LastName"
	}
	else{
		$agMethod = $mtTypeDrop.SelectedItem.ToString()
	}
	$global:mbtable = New-Object System.Data.DataTable
	$global:mbtable.TableName = "Mailboxes"
	$global:mbtable.Columns.Add("Display Name")
	$global:mbtable.Columns.Add($agMethod)
	$global:mbtable.Columns.Add("Mailbox Size(MB)",[int64])
	foreach ($row in $usrinfo.Values){
		if($row.$agMethod -eq $agtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0]){
				$global:mbtable.rows.add($row.DisplayName,$row.$agMethod,$row.MailboxSize)
		}
		else{
			if($row.$agMethod -eq ""){
				if("Not Set" -eq $agtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0]){
					$global:mbtable.rows.add($row.DisplayName,$row.$agMethod,$row.MailboxSize)
				}
			}
		}
	}
	$dgDataGrid1.DataSource = $global:mbtable
}

function ExportGrid1csv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("DisplayName,Number of Mailboxes,Mailbox Size(MB)")
	foreach($row in $agTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString()) 
	}
	$logfile.Close()
}
}

function ExportGrid2csv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	if ($mtTypeDrop.SelectedItem -eq $null){
		$agMethod = "LastName"
	}
	else{
		$agMethod = $mtTypeDrop.SelectedItem.ToString()
	}
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("DisplayName,$agMethod,Mailbox Size(MB)")
	foreach($row in $global:mbtable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString()) 
	}
	$logfile.Close()
}
}

$agtable = New-Object System.Data.DataTable
$agtable.TableName = "Mailbox Sizes"
$agtable.Columns.Add("Name")
$agtable.Columns.Add("# Mailboxes",[int64])
$agtable.Columns.Add("Mailbox Size(MB)",[int64])




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
$snServerNameDrop.Items.Add("ALL Servers")
$snServerNameDrop.Add_SelectedValueChanged({getMailboxSizes})  

$form.Controls.Add($snServerNameDrop)

# Add Agregate Type DropLable
$mtTypeDroplableBox = new-object System.Windows.Forms.Label
$mtTypeDroplableBox.Location = new-object System.Drawing.Size(200,20) 
$mtTypeDroplableBox.size = new-object System.Drawing.Size(70,20) 
$mtTypeDroplableBox.Text = "Group By"
$form.Controls.Add($mtTypeDroplableBox) 

# Add Mailbox Type Drop Down
$mtTypeDrop = new-object System.Windows.Forms.ComboBox
$mtTypeDrop.Location = new-object System.Drawing.Size(270,20)
$mtTypeDrop.Size = new-object System.Drawing.Size(135,30)
$properties = get-user -resultsize 1 | get-member -membertype Property 
foreach ($prop in $properties){
	$mtTypeDrop.Items.Add($prop.Name)
} 
$mtTypeDrop.Items.Add("OU")
$mtTypeDrop.Items.Add("MailboxStore")
$mtTypeDrop.Items.Add("StorageGroup")
$mtTypeDrop.Add_SelectedValueChanged({Aggresults})  
$form.Controls.Add($mtTypeDrop)

# Show Mailboxs Button

$shMailboxes = new-object System.Windows.Forms.Button
$shMailboxes.Location = new-object System.Drawing.Size(600,19)
$shMailboxes.Size = new-object System.Drawing.Size(120,23)
$shMailboxes.Text = "Show Mailboxes"
$shMailboxes.visible = $True
$shMailboxes.Add_Click({ShowMailboxes})
$form.Controls.Add($shMailboxes)

# Add Export Grid Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(430,19)
$exButton1.Size = new-object System.Drawing.Size(90,23)
$exButton1.Text = "Export Grid"
$exButton1.Add_Click({ExportGrid1csv})
$form.Controls.Add($exButton1)

# Add Export Grid Button 2

$exButton2 = new-object System.Windows.Forms.Button
$exButton2.Location = new-object System.Drawing.Size(750,19)
$exButton2.Size = new-object System.Drawing.Size(125,23)
$exButton2.Text = "Export Grid"
$exButton2.Add_Click({ExportGrid2csv})
$form.Controls.Add($exButton2)

#add Picture box

$pbox =  new-object System.Windows.Forms.PictureBox
$pbox.Location = new-object System.Drawing.Size(550,360)
$pbox.Size = new-object System.Drawing.Size(500,220)
$form.Controls.Add($pbox)

$pbox1 =  new-object System.Windows.Forms.PictureBox
$pbox1.Location = new-object System.Drawing.Size(10,360)
$pbox1.Size = new-object System.Drawing.Size(550,370)
$form.Controls.Add($pbox1)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,50) 
$dgDataGrid.size = new-object System.Drawing.Size(530,300)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)

$dgDataGrid1 = new-object System.windows.forms.DataGridView
$dgDataGrid1.Location = new-object System.Drawing.Size(550,50) 
$dgDataGrid1.size = new-object System.Drawing.Size(450,300)
$dgDataGrid1.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid1)

$form.Text = "Exchange 2007 Group by Mailbox Size Form"
$form.size = new-object System.Drawing.Size(1000,700) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()