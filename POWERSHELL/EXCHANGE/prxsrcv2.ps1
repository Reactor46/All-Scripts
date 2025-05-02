[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
$form = new-object System.Windows.Forms.form 

function SearchforProxy(){
	if ($pnProxyAddress.Text -match "@"){
		$msTable.clear()
		$root = [ADSI]'LDAP://RootDSE' 
		if ($rbSearchDWide.Checked -eq $true){
			$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
		}
		else{
			$dfDefaultRootPath = "LDAP://" + $ouOUNameDrop.SelectedItem.ToString()
		}
		$dfRoot = [ADSI]$dfDefaultRootPath
		$gfGALQueryFilter =  "(&(&(&(&(!mailnickname=systemmailbox*)(objectCategory=person)(objectClass=user)(proxyAddresses=smtp:*" + $pnProxyAddress.Text + ")))))"
		$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
		if($ouOUCheckBox.Checked -eq $false -band $rbSearchDWide.Checked -eq $false){$dfsearcher.SearchScope = "OneLevel"}
		$dfsearcher.Filter = $gfGALQueryFilter
		$srSearchResult = $dfsearcher.FindAll()
		foreach ($emResult in $srSearchResult) {
			$unUserobject = New-Object System.DirectoryServices.directoryentry
			$unUserobject  = $emResult.GetDirectoryEntry()	
			$address = ""
			$userName = ""
			$mail = ""
			$prxfound = ""
			foreach($prxaddress in $unUserobject.ProxyAddresses){
				$sstring = "*" + $pnProxyAddress.Text
				if ($prxaddress -like $sstring){
					if($prxfound -eq ""){$prxfound = $prxaddress}
					else{$prxfound = $prxfound + ";" + $prxaddress}

				}
				if ($address -eq ""){$address = $prxaddress}
				else{$address = $address + ";" + $prxaddress}
			}
			if ($unUserobject.Mail -ne $null){$mail = $unUserobject.Mail.ToString()}
			$msTable.Rows.Add($unUserobject.SamaccountName.ToString(),$mail,$prxfound,$unUserobject.psbase.Parent.Name.ToString(),$address,$unUserobject.distinguishedName.ToString())
			
		}
		$dgDataGrid.DataSource = $msTable
	}
	else{
		$msgstring = "Address format incorrect it must contain @ sign"
		$a = new-object -comobject wscript.shell
		$b = $a.popup($msgstring,0,"Warning",1)

	}
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
	$logfile.WriteLine("UserName,Primary-EmailAddress,ProxyAddres-Found,OU,ProxyAddresses,distinguishedName")
	foreach($row in $msTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString()+ $row[5].ToString()) 
	}
	$logfile.Close()
}

}

function DeletedSelected(){
if ($dgDataGrid.SelectedRows.Count -eq 0){	 
		$msgstring = "Now Rows Selected"
		$a = new-object -comobject wscript.shell
		$b = $a.popup($msgstring,0,"Warning",1)
}
else{
	$buttons=[system.windows.forms.messageboxbuttons]::yesno
	$return = ""
	$return = [system.windows.forms.messagebox]::Show("This will remove the Selected Proxy Addresses Click Yes to Procced","",$buttons)
	if ($return -eq [Windows.Forms.DialogResult]::Yes){
		$lcLoopCount = 0
		while ($lcLoopCount -le ($dgDataGrid.SelectedRows.Count-1)) {

			$UserDN = "LDAP://" + $dgDataGrid.SelectedRows[$lcLoopCount].Cells[5].Value
			$userObject = [ADSI]$UserDN 
			$sstring = "*" + $pnProxyAddress.Text
			foreach($prxaddress in $userObject.ProxyAddresses){
				if ($prxaddress -like $sstring){
					if($prxaddress -cmatch "SMTP:") {
						$Message = "Not Deleteing " + $prxaddress + " as this is the primary SMTP address"
						[system.windows.forms.messagebox]::Show($Message)
					}
					else{
						$Message = "Deleteing the following proxy Address's`r`n`r`n" + $prxaddress + "`r`n`r`nFrom " + $userObject.DisplayName.ToString()
						$mreturn = ""
						if ($cmfCheckBox.Checked -eq $true){
							$mreturn = [system.windows.forms.messagebox]::Show($Message,"",$buttons)
							if ($mreturn  -eq [Windows.Forms.DialogResult]::Yes){
								$userObject.PutEx(4, 'proxyAddresses', @("$prxaddress"))
								$userObject.setinfo()
							}
						}
						else{
							$userObject.PutEx(4, 'proxyAddresses', @("$prxaddress"))
							$userObject.setinfo()											
						}
					}
				}
			}
			$lcLoopCount += 1
		}
	}
	else {[system.windows.forms.messagebox]::Show("Not proceeding with delete") }
	SearchforProxy

}


}

$msTable = New-Object System.Data.DataTable
$msTable.TableName = "ProxyAddress"
$msTable.Columns.Add("UserName")
$msTable.Columns.Add("Primary-EmailAddress")
$msTable.Columns.Add("ProxyAddress-Found")
$msTable.Columns.Add("OU")
$msTable.Columns.Add("ProxyAddresses")
$msTable.Columns.Add("distinguishedName")


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

$ProxyAddresslableBox = new-object System.Windows.Forms.Label
$ProxyAddresslableBox.Location = new-object System.Drawing.Size(20,100) 
$ProxyAddresslableBox.size = new-object System.Drawing.Size(320,20) 
$ProxyAddresslableBox.Text = "Proxy Address to Search (eg @proxydomain.com)"
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
$exButton1.Add_Click({SearchforProxy})
$form.Controls.Add($exButton1)

# Add Export Grid Button

$exButton2 = new-object System.Windows.Forms.Button
$exButton2.Location = new-object System.Drawing.Size(10,760)
$exButton2.Size = new-object System.Drawing.Size(125,20)
$exButton2.Text = "Export Grid"
$exButton2.Add_Click({ExportGrid})
$form.Controls.Add($exButton2)

# Add Export Grid Button

$exButton3 = new-object System.Windows.Forms.Button
$exButton3.Location = new-object System.Drawing.Size(250,760)
$exButton3.Size = new-object System.Drawing.Size(300,20)
$exButton3.Text = "Deleted Select Proxy Addresses"
$exButton3.Add_Click({DeletedSelected})
$form.Controls.Add($exButton3)

$cmfCheckBox = new-object System.Windows.Forms.CheckBox
$cmfCheckBox.Location = new-object System.Drawing.Size(600,760)
$cmfCheckBox.Size = new-object System.Drawing.Size(200,20)
$cmfCheckBox.Checked = $true
$cmfCheckBox.Text = "Confirm Each removal"
$form.Controls.Add($cmfCheckBox)


# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,145) 
$dgDataGrid.size = new-object System.Drawing.Size(1000,600)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)

$form.Text = "Proxy Address Search and Remove Form"
$form.size = new-object System.Drawing.Size(1200,750) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()

