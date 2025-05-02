[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
$form = new-object System.Windows.Forms.form 

$repeathashGroup = @{ }
$repeathashUser = @{ }

function Get-member($GroupName){
	$Grouppath = "LDAP://"  + $GroupName
	$groupObj =  [ADSI]$Grouppath
	foreach($member in $groupObj.Member){	
		$userPath = "LDAP://"  + $member
		$UserObj =  [ADSI]$userPath
		if($UserObj.groupType.Value -eq $null){
			if($repeathashUser.ContainsKey($UserObj.distinguishedName.ToString()) -eq $false){
				$repeathashUser.add($UserObj.distinguishedName.ToString(),1)
				$grTable.Rows.Add($UserObj.DisplayName.ToString(),$UserObj.extensionAttribute1.Value,`
				$UserObj.extensionAttribute2.Value,$UserObj.extensionAttribute3.Value,$UserObj.extensionAttribute4.Value,$UserObj.extensionAttribute5.Value,`
				$UserObj.extensionAttribute6.Value,$UserObj.extensionAttribute7.Value,$UserObj.extensionAttribute8.Value,$UserObj.extensionAttribute9.Value,`
				$UserObj.extensionAttribute10.Value,$UserObj.extensionAttribute11.Value,`
				$UserObj.extensionAttribute12.Value,$UserObj.extensionAttribute13.Value,$UserObj.extensionAttribute14.Value,$UserObj.extensionAttribute15.Value)
			}

		}
		else{
			if($repeathashGroup.ContainsKey($UserObj.distinguishedName.ToString()) -eq $false){
				$repeathashGroup.add($UserObj.distinguishedName.ToString(),1)
				Get-member($UserObj.distinguishedName)		
			}
		}
	}
	$dgDataGrid2.DataSource = $grTable
}

function Update-member($GroupName){
	#Warn first
	$buttons=[system.windows.forms.messageboxbuttons]::yesno
	$return = ""
	$return = [system.windows.forms.messagebox]::Show("This will set the selected Extended Properties for all members of the group and nested Groups do you want to proceed","",$buttons)
	if ($return -eq [Windows.Forms.DialogResult]::Yes){
		$Grouppath = "LDAP://"  + $GroupName
		$groupObj =  [ADSI]$Grouppath
		foreach($member in $groupObj.Member){	
			$userPath = "LDAP://"  + $member
			$UserObj =  [ADSI]$userPath
			if($UserObj.groupType.Value -eq $null){
				if($repeathashUser.ContainsKey($UserObj.distinguishedName.ToString()) -eq $false){
					$propname = $exPropDrop.SelectedItem.ToString()
					$repeathashUser.add($UserObj.distinguishedName.ToString(),1)
					$UserObj.$propname = $uvtext.Text
					$UserObj.SetInfo()
				}
	
			}
			else{
				if($repeathashGroup.ContainsKey($UserObj.distinguishedName.ToString()) -eq $false){
					$repeathashGroup.add($UserObj.distinguishedName.ToString(),1)
					Update-member($UserObj.distinguishedName)		
				}
			}
		}
	}
	else{
		[system.windows.forms.messagebox]::Show("Update Aborted")
	}
	$grTable.Clear()
	$repeathashUser.clear()
	$repeathashGroup.clear()
	Get-member($msTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][1])

}

function ExportGrid(){
$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("MemberName,extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute4,extensionAttribute5,extensionAttribute6,extensionAttribute7,extensionAttribute8,extensionAttribute9,extensionAttribute10,extensionAttribute11,extensionAttribute12,extensionAttribute13,extensionAttribute14,extensionAttribute15")
	foreach($row in $grTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString()+ "," + $row[5].ToString() + "," + $row[6].ToString() + "," + $row[7].ToString() + "," + $row[8].ToString()`
		+ "," + $row[9].ToString() + "," + $row[10].ToString() + "," + $row[11].ToString() + "," + $row[12].ToString() + "," + $row[13].ToString() + "," + $row[14].ToString() + "," + $row[15].ToString()) 
	}
	$logfile.Close()
}

}

function SearchfoGroup(){
		$msTable.clear()
		$root = [ADSI]'LDAP://RootDSE' 
		if ($rbSearchDWide.Checked -eq $true){
			$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
		}
		else{
			$dfDefaultRootPath = "LDAP://" + $ouOUNameDrop.SelectedItem.ToString()
		}
		$dfRoot = [ADSI]$dfDefaultRootPath
		$gfGALQueryFilter =  "(&(objectClass=group)(displayName=" + $pnProxyAddress.Text + "*))"
		$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
		if($ouOUCheckBox.Checked -eq $false -band $rbSearchDWide.Checked -eq $false){$dfsearcher.SearchScope = "OneLevel"}
		$dfsearcher.Filter = $gfGALQueryFilter
		$srSearchResult = $dfsearcher.FindAll()
		foreach ($emResult in $srSearchResult) {
			$unUserobject = New-Object System.DirectoryServices.directoryentry
			$unUserobject  = $emResult.GetDirectoryEntry()	
			$msTable.Rows.Add($unUserobject.DisplayName.ToString(),$unUserobject.distinguishedName.ToString())
			
		}
		$dgDataGrid.DataSource = $msTable

}

$msTable = New-Object System.Data.DataTable
$msTable.TableName = "GroupName"
$msTable.Columns.Add("GroupName")
$msTable.Columns.Add("DistinguishedName")

$grTable = New-Object System.Data.DataTable
$grTable.TableName = "Members"
$grTable.Columns.Add("MemberName")
$grTable.Columns.Add("extensionAttribute1")
$grTable.Columns.Add("extensionAttribute2")
$grTable.Columns.Add("extensionAttribute3")
$grTable.Columns.Add("extensionAttribute4")
$grTable.Columns.Add("extensionAttribute5")
$grTable.Columns.Add("extensionAttribute6")
$grTable.Columns.Add("extensionAttribute7")
$grTable.Columns.Add("extensionAttribute8")
$grTable.Columns.Add("extensionAttribute9")
$grTable.Columns.Add("extensionAttribute10")
$grTable.Columns.Add("extensionAttribute11")
$grTable.Columns.Add("extensionAttribute12")
$grTable.Columns.Add("extensionAttribute13")
$grTable.Columns.Add("extensionAttribute14")
$grTable.Columns.Add("extensionAttribute15")



# Add RadioButtons
$rbSearchDWide = new-object System.Windows.Forms.RadioButton
$rbSearchDWide.Location = new-object System.Drawing.Size(20,20)
$rbSearchDWide.size = new-object System.Drawing.Size(150,17)
$rbSearchDWide.Checked = $true
$rbSearchDWide.Text = "Search Domain Wide"
$rbSearchDWide.Add_Click({if ($rbSearchDWide.Checked -eq $true){$ouOUNameDrop.Enabled = $false}})
$form.Controls.Add($rbSearchDWide)

$rbSearchOUWide = new-object System.Windows.Forms.RadioButton
$rbSearchOUWide.Location = new-object System.Drawing.Size(20,60)
$rbSearchOUWide.size = new-object System.Drawing.Size(150,17)
$rbSearchOUWide.Checked = $false
$rbSearchOUWide.Add_Click({if ($rbSearchDWide.Checked -eq $false){$ouOUNameDrop.Enabled = $true}})
$rbSearchOUWide.Text = "Search within OU"
$form.Controls.Add($rbSearchOUWide)


$OulableBox = new-object System.Windows.Forms.Label
$OulableBox.Location = new-object System.Drawing.Size(220,60) 
$OulableBox.size = new-object System.Drawing.Size(120,20) 
$OulableBox.Text = "Select OU Name : "
$form.controls.Add($OulableBox) 

# Add OU Drop Down
$ouOUNameDrop = new-object System.Windows.Forms.ComboBox
$ouOUNameDrop.Location = new-object System.Drawing.Size(360,60)
$ouOUNameDrop.Size = new-object System.Drawing.Size(350,30)
$ouOUNameDrop.Enabled = $false
$root = [ADSI]'LDAP://RootDSE' 
$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
$dfRoot = [ADSI]$dfDefaultRootPath
$gfGALQueryFilter =  "(objectClass=organizationalUnit)"
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
$dfsearcher.Filter = $gfGALQueryFilter
$srSearchResult = $dfsearcher.FindAll()
foreach ($emResult in $srSearchResult) {
	$OUobject = New-Object System.DirectoryServices.directoryentry
	$OUobject  = $emResult.GetDirectoryEntry()
	$ouOUNameDrop.Items.Add($OUobject.distinguishedName.ToString())
}
$form.Controls.Add($ouOUNameDrop)

# Add Prop Drop Down
$exPropDrop = new-object System.Windows.Forms.ComboBox
$exPropDrop.Location = new-object System.Drawing.Size(240,580)
$exPropDrop.Size = new-object System.Drawing.Size(150,30)
$exPropDrop.Items.Add("extensionAttribute1")
$exPropDrop.Items.Add("extensionAttribute2")
$exPropDrop.Items.Add("extensionAttribute3")
$exPropDrop.Items.Add("extensionAttribute4")
$exPropDrop.Items.Add("extensionAttribute5")
$exPropDrop.Items.Add("extensionAttribute6")
$exPropDrop.Items.Add("extensionAttribute7")
$exPropDrop.Items.Add("extensionAttribute8")
$exPropDrop.Items.Add("extensionAttribute9")
$exPropDrop.Items.Add("extensionAttribute10")
$exPropDrop.Items.Add("extensionAttribute11")
$exPropDrop.Items.Add("extensionAttribute12")
$exPropDrop.Items.Add("extensionAttribute13")
$exPropDrop.Items.Add("extensionAttribute14")
$exPropDrop.Items.Add("extensionAttribute15")
$form.Controls.Add($exPropDrop)

$exlableBox = new-object System.Windows.Forms.Label
$exlableBox.Location = new-object System.Drawing.Size(20,580) 
$exlableBox.size = new-object System.Drawing.Size(240,20) 
$exlableBox.Text = "Extended Property to Set/Update : "
$form.controls.Add($exlableBox) 

$exlableBox2 = new-object System.Windows.Forms.Label
$exlableBox2.Location = new-object System.Drawing.Size(20,610) 
$exlableBox2.size = new-object System.Drawing.Size(150,20) 
$exlableBox2.Text = "Value : "
$form.controls.Add($exlableBox2) 


$uvtext = new-object System.Windows.Forms.TextBox 
$uvtext.Location = new-object System.Drawing.Size(240,610) 
$uvtext.size = new-object System.Drawing.Size(300,20) 
$form.controls.Add($uvtext) 


$ProxyAddresslableBox = new-object System.Windows.Forms.Label
$ProxyAddresslableBox.Location = new-object System.Drawing.Size(20,100) 
$ProxyAddresslableBox.size = new-object System.Drawing.Size(320,20) 
$ProxyAddresslableBox.Text = "Group to Search for"
$form.controls.Add($ProxyAddresslableBox) 

$ouOUCheckBox = new-object System.Windows.Forms.CheckBox
$ouOUCheckBox.Location = new-object System.Drawing.Size(750,60)
$ouOUCheckBox.Size = new-object System.Drawing.Size(200,20)
$ouOUCheckBox.Checked = $true
$ouOUCheckBox.Text = "Search in Sub OU's"
$form.Controls.Add($ouOUCheckBox)


# Add ProxyDomain Text Box
$pnProxyAddress = new-object System.Windows.Forms.TextBox 
$pnProxyAddress.Location = new-object System.Drawing.Size(350,100) 
$pnProxyAddress.size = new-object System.Drawing.Size(300,20) 
$form.controls.Add($pnProxyAddress) 


$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(700,100)
$exButton1.Size = new-object System.Drawing.Size(125,20)
$exButton1.Text = "Search"
$exButton1.Add_Click({SearchfoGroup})
$form.Controls.Add($exButton1)

$exButton3 = new-object System.Windows.Forms.Button
$exButton3.Location = new-object System.Drawing.Size(320,150)
$exButton3.Size = new-object System.Drawing.Size(105,20)
$exButton3.Text = "Show Members"
$exButton3.Add_Click({$grTable.clear()
	$repeathashUser.clear()
	$repeathashGroup.clear()
	Get-member($msTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][1])
	})
$form.Controls.Add($exButton3)

# Add Update Button

$exButton2 = new-object System.Windows.Forms.Button
$exButton2.Location = new-object System.Drawing.Size(10,640)
$exButton2.Size = new-object System.Drawing.Size(125,20)
$exButton2.Text = "Set/Update Value"
$exButton2.Add_Click({
	$repeathashUser.clear()
	$repeathashGroup.clear()
	update-member($msTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][1])
	})
$form.Controls.Add($exButton2)


# Add Export Grid Button

$exButton4 = new-object System.Windows.Forms.Button
$exButton4.Location = new-object System.Drawing.Size(700,550)
$exButton4.Size = new-object System.Drawing.Size(125,20)
$exButton4.Text = "Export Grid"
$exButton4.Add_Click({ExportGrid})
$form.Controls.Add($exButton4)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,145) 
$dgDataGrid.size = new-object System.Drawing.Size(300,400)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)

# Add DataGrid View

$dgDataGrid2 = new-object System.windows.forms.DataGridView
$dgDataGrid2.Location = new-object System.Drawing.Size(440,145) 
$dgDataGrid2.size = new-object System.Drawing.Size(400,400)
$dgDataGrid2.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid2)


$form.Text = "Group Extended Propery Update GUI"
$form.size = new-object System.Drawing.Size(1200,800) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
