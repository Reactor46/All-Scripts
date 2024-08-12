[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

function enumMailboxperms() {
$root = [ADSI]'LDAP://RootDSE' 
$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
$dfRoot = [ADSI]$dfDefaultRootPath
$gfGALQueryFilter =  "(&(&(&(mailnickname=*)(objectCategory=person)(objectClass=user))))"
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
$dfsearcher.Filter = $gfGALQueryFilter
$dfsearcher.PageSize = 1000
$dfsearcher.PropertiesToLoad.Add("msExchMailboxSecurityDescriptor")
$srSearchResult = $dfsearcher.FindAll()
foreach ($emResult in $srSearchResult) {
	$uoUserobject = New-Object System.DirectoryServices.directoryentry
	$uoUserobject = $emResult.GetDirectoryEntry()
	$emProps = $emResult.Properties
	if ($emProps["msexchmailboxsecuritydescriptor"][0] -ne $null){
	[byte[]]$DaclByte = $emProps["msexchmailboxsecuritydescriptor"][0]
	$adDACL = new-object System.DirectoryServices.ActiveDirectorySecurity
	$adDACL.SetSecurityDescriptorBinaryForm($DaclByte)
	$mbRightsacls =$adDACL.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])
	"Processing Mailbox - " +  $uoUserobject.DisplayName
	$mbCount = 0
	foreach ($ace in $mbRightsacls){
		if($ace.IdentityReference.Value -ne "S-1-5-10" -band $ace.IdentityReference.Value -ne "S-1-5-18" -band $ace.IsInherited -ne $true){		
				$sidbind = "LDAP://<SID=" + $ace.IdentityReference.Value + ">"
				$AceName = $ace.IdentityReference.Value 
			        $aceuser = [ADSI]$sidbind
				if ($aceuser.name -ne $null){
					$AceName = $aceuser.samaccountname.ToString()
				}
				if ($rvMailboxPerms.Containskey($AceName)){
					$rvMailboxPerms[$AceName] = [int]$rvMailboxPerms[$AceName] +1			
				}
				else {
					$rvMailboxPerms.add($AceName,1)
				}
				If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::CreateChild){ 
					[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"MailboxRight",$aceName,"Full Mailbox Access",$ace.AccessControlType)
					$mbCount++}					
				If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::WriteOwner -ne 0){ 
					[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"MailboxRight",$aceName,"Take Ownership",$ace.AccessControlType)
					$mbCount++}
			        If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::WriteDacl){ 
					[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"MailboxRight",$aceName,"Modify User Attributes",$ace.AccessControlType)
					$mbCount++}
			        If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::ListChildren){ 
					[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"MailboxRight",$aceName,"Is mailbox primary owner of this object",$ace.AccessControlType)
					$mbCount++}		
				If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::Delete){ 
					[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"MailboxRight",$aceName,"Delete mailbox storage",$ace.AccessControlType)
					$mbCount++}
				If ($ace.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::ReadControl){ 
					[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"MailboxRight",$aceName,"Read permissions",$ace.AccessControlType)
					$mbCount++}

		}
	}
	$srCount = 0
	$Sendasacls = $uoUserobject.psbase.get_objectSecurity().getAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])|? {$_.ObjectType -eq 'ab721a54-1e2f-11d0-9819-00aa0040529b'}
	$Recieveasacls = $uoUserobject.psbase.get_objectSecurity().getAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])|? {$_.ObjectType -eq 'ab721a56-1e2f-11d0-9819-00aa0040529b'}
	if ($Sendasacls -ne $null){
		foreach ($ace in $Sendasacls)
		{
			if($ace.IdentityReference.Value -ne "S-1-5-10" -band $ace.IdentityReference.Value -ne "S-1-5-18" -band $ace.IsInherited -ne $true){
				$srCount++
				$sidbind = "LDAP://<SID=" + $ace.IdentityReference.Value + ">"
				$AceName = $ace.IdentityReference.Value 
				$aceuser = [ADSI]$sidbind
				if ($aceuser.name -ne $null){
					$AceName = $aceuser.samaccountname.ToString()
				}
				[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"SendAS-RecieveAS",$AceName,"Send As",$ace.AccessControlType)
				if ($rvSendRecieve.Containskey($AceName)){
					$rvSendRecieve[$AceName] = [int]$rvSendRecieve[$AceName] +1			
				}
				else {
					$rvSendRecieve.add($AceName,1)
				}
			}
	
		}
	}
	if ($Recieveasacls -ne $null){
		foreach ($ace in $Recieveasacls)
		{
			if($ace.IdentityReference.Value -ne "S-1-5-10" -band $ace.IdentityReference.Value -ne "S-1-5-18" -band $ace.IsInherited -ne $true){
				$srCount++
				$sidbind = "LDAP://<SID=" + $ace.IdentityReference.Value + ">"
				$AceName = $ace.IdentityReference.Value 
				$aceuser = [ADSI]$sidbind
				if ($aceuser.name -ne $null){
					$AceName = $aceuser.samaccountname.ToString()
				}
				[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"SendAS-RecieveAS",$AceName,"Recieve As",$ace.AccessControlType)
				if ($rvSendRecieve.Containskey($AceName)){
					$rvSendRecieve[$AceName] = [int]$rvSendRecieve[$AceName] +1			
				}
				else {
					$rvSendRecieve.add($AceName,1)
				}
			}
		}
	}
	$nmMailboxPerms.Add($uoUserobject.samaccountname.ToString(),$mbCount)
	$nmSendRecieve.Add($uoUserobject.samaccountname.ToString(),$srCount)
	}
}
foreach($key in $nmMailboxPerms.keys){
	$rvSRRights = 0
	$rvMbRights = 0
	if ($rvMailboxPerms.Containskey($key)){
		$rvMbRights = $rvMailboxPerms[$key]
	}
	if ($rvSendRecieve.Containskey($key)){
		$rvSRRights = $rvSendRecieve[$key]
	}

	$rs1Table.Rows.Add($key, $nmMailboxPerms[$key],$nmSendRecieve[$key],$rvMbRights,$rvSRRights)
	
}
# write-host $nmSendRecieve
$dgDataGrid.datasource = $rs1Table
}

function showACL{
if ($ObjTypeDrop.SelectedItem -eq $null -bor $ObjTypeDrop.SelectedItem -eq "Mailbox"){

	if ($AceTypeDrop.SelectedItem -ne $null){
		switch($AceTypeDrop.SelectedItem.ToString()){
			"Mailbox-Rights" {$rows = $rsTable.Select("MailboxName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType = 'MailboxRight'")}
			"SendAs/ReciveAS-Rights" {$rows = $rsTable.Select("MailboxName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType = 'SendAS-RecieveAS'")}
			"Reverse-Mailbox-Rights" {$rows = $rsTable.Select("UserName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType = 'MailboxRight'")}
			"Reverse-SendAs/ReciveAS-Rights" {$rows = $rsTable.Select("UserName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType = 'SendAS-RecieveAS'")}
			default {$rows = $rsTable.Select("UserName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' Or MailboxName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "'")}
		}
	}
	else{
		$rows = $rsTable.Select("UserName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' Or MailboxName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "'")
	}


	$frTable.clear()
	foreach ($row in $rows){
		$frTable.rows.add($row[0].ToString(),$row[1].ToString(),$row[2].ToString(),$row[3].ToString(),$row[4].ToString())
	}
	$dgDataGrid1.datasource = $frTable
}
else{
	$rows = $msrTable.Select("DistinguishedName='" + $msr1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][3] + "'")
	$fr1Table.clear()
	foreach ($row in $rows){
		$fr1Table.rows.add($row[0].ToString(),$row[1].ToString(),$row[2].ToString(),$row[3].ToString(),$row[4].ToString())
	}
	$dgDataGrid1.datasource = $fr1Table
}
}

function EnumMailStorePerms(){

$dse  = [adsi]("LDAP://Rootdse")
$ERtbl  = @{}
$ext  = [adsi]("LDAP://cn=Extended-rights,"+$dse.configurationNamingContext)
$ext.psbase.children |% { 
	if ($ERtbl.containskey($_.rightsGuid.ToString()) -eq $false){
		$ERtbl.Add($_.rightsGuid.ToString(),$_.displayName.toString()) 
	}
}

$root = [ADSI]'LDAP://RootDSE' 
$cfConfigRootPath = "LDAP://" + $root.configurationNamingContext.tostring()
$cfRoot = [ADSI]$cfConfigRootPath
$sQueryFilter =  "(objectCategory=msExchPrivateMDB)"
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($cfRoot)
$dfsearcher.Filter = $sQueryFilter
$srSearchResult = $dfsearcher.FindAll()
foreach ($emResult in $srSearchResult) {
	$soStoreobject = New-Object System.DirectoryServices.directoryentry
	$soStoreobject = $emResult.GetDirectoryEntry()
	$Storeacls = $soStoreobject.psbase.get_objectSecurity().getAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])
	foreach($ace in $Storeacls){
		if ($ace.IdentityReference.Value -ne "S-1-5-7" -band $ace.IdentityReference.Value -ne "S-1-1-0" -band $ace.IsInherited -ne $true){
			$sidbind = "LDAP://<SID=" + $ace.IdentityReference.Value + ">"
			$AceName = $ace.IdentityReference.Value 
			$aceuser = [ADSI]$sidbind
			if ($aceuser.name -ne $null){
				$AceName = $aceuser.samaccountname.ToString()
			}	
			if ($ERtbl.containskey($ace.ObjectType.ToString())){
				$rights = $ERtbl[$ace.ObjectType.ToString()] + " " + $ace.ObjectType.ToString()
			}
			else {
				$rights = $ace.activeDirectoryRights.toString()
			}
			$msrTable.rows.add($soStoreobject.Name.ToString(),$soStoreobject.DistinguishedName.ToString(),$AceName,$rights,$ace.AccessControlType)

	}
	}
	$soServer = [ADSI]("LDAP://" + $soStoreobject.msExchOwningServer)
	"Processing MailStore - " + $soServer
	$sgStorageGroup = $soStoreobject.psbase.Parent
	$msr1Table.rows.add($soServer.Name.ToString(),$sgStorageGroup.Name.ToString(),$soStoreobject.Name.ToString(),$soStoreobject.DistinguishedName.ToString())
}
}

function exportTable{
$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.Showhelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("MailboxName,ACLType,UserName,Rights,Status")
	foreach($row in $rsTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString()) 
	}
	$logfile.Close()
}

}

function enDelExchangeRight(){

$root = [ADSI]'LDAP://RootDSE'
$exRootPath = "LDAP://CN=Microsoft Exchange,CN=Services," + $root.configurationNamingContext.tostring()
$exRoot = [ADSI]$exRootPath
$Storeacls = $exRoot.psbase.get_objectSecurity().getAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])
	foreach($ace in $Storeacls){
		if ($ace.IdentityReference.Value -ne "S-1-5-7" -band $ace.IdentityReference.Value -ne "S-1-1-0" -band $ace.IsInherited -ne $true){
			$sidbind = "LDAP://<SID=" + $ace.IdentityReference.Value + ">"
			$AceName = $ace.IdentityReference.Value 
			$aceuser = [ADSI]$sidbind
			if ($aceuser.name -ne $null){
				$AceName = $aceuser.samaccountname.ToString()
			}	
			switch ($ace.activeDirectoryRights.GetHashCode()){
				983551 {$drTable1.rows.add($AceName,"Exchange Full Administration")}
				197119 {$drTable1.rows.add($AceName,"Exchange Administration")}
				131220 {$drTable1.rows.add($AceName,"Exchange View Only Administrator")}
			}
			

	}
}

}


$nmMailboxPerms = @{ }
$nmSendRecieve = @{ }
$rvMailboxPerms = @{ }
$rvSendRecieve = @{ }

$Dataset = New-Object System.Data.DataSet
$rsTable = New-Object System.Data.DataTable
$rsTable.TableName = "Mailbox Rights"
$rsTable.Columns.Add("MailboxName")
$rsTable.Columns.Add("ACLType")
$rsTable.Columns.Add("UserName")
$rsTable.Columns.Add("Rights")
$rsTable.Columns.Add("Status")
$Dataset.tables.add($rsTable)

$Dataveiw = New-Object System.Data.DataView($rsTable1)
$rs1Table = New-Object System.Data.DataTable
$rs1Table.TableName = "ACL-Numbers"
$rs1Table.Columns.Add("AccountName")
$rs1Table.Columns.Add("Mailbox")
$rs1Table.Columns.Add("Send/RecieveAS")
$rs1Table.Columns.Add("Revese_Mailbox")
$rs1Table.Columns.Add("Revese_Send/RecieveAS")
$Dataset.tables.add($rs1Table)

$Dataveiw = New-Object System.Data.DataView($drTable1)
$drTable1 = New-Object System.Data.DataTable
$drTable1.Columns.Add("AccountName")
$drTable1.Columns.Add("DelegatedRights")
$Dataset.tables.add($drTable1)

$frTable = New-Object System.Data.DataTable
$frTable.TableName = "Filtered Mailbox Rights"
$frTable.Columns.Add("MailboxName")
$frTable.Columns.Add("ACLType")
$frTable.Columns.Add("UserName")
$frTable.Columns.Add("Rights")
$frTable.Columns.Add("Status")
$Dataset.tables.add($frTable)


$fr1Table = New-Object System.Data.DataTable
$fr1Table.TableName = "Mailbox Store Rights"
$fr1Table.Columns.Add("MailStore")
$fr1Table.Columns.Add("DistinguishedName")
$fr1Table.Columns.Add("UserName")
$fr1Table.Columns.Add("Rights")
$fr1Table.Columns.Add("Status")
$Dataset.tables.add($fr1Table)

$msrTable = New-Object System.Data.DataTable
$msrTable.TableName = "Filtered Mailbox Store Rights"
$msrTable.Columns.Add("MailStore")
$msrTable.Columns.Add("DistinguishedName")
$msrTable.Columns.Add("UserName")
$msrTable.Columns.Add("Rights")
$msrTable.Columns.Add("Status")
$Dataset.tables.add($msrTable)

$msr1Table = New-Object System.Data.DataTable
$msr1Table.TableName = "Mailbox Store Table"

$msr1Table.Columns.Add("ServerName")
$msr1Table.Columns.Add("StorageGroupName")
$msr1Table.Columns.Add("MailStoreName")
$msr1Table.Columns.Add("DistinguishedName")
$Dataset.tables.add($msr1Table)

$form = new-object System.Windows.Forms.form 
$form.Text = "Exchange Permissions Gui"


# Add Show ACE Button

$shaces = new-object System.Windows.Forms.Button
$shaces.Location = new-object System.Drawing.Size(560,19)
$shaces.Size = new-object System.Drawing.Size(90,23)
$shaces.Text = "Show ACE's"
$shaces.Add_Click({showACL})
$form.Controls.Add($shaces)

# Add Object Type Drop Down
$ObjTypeDrop = new-object System.Windows.Forms.ComboBox
$ObjTypeDrop.Location = new-object System.Drawing.Size(160,20)
$ObjTypeDrop.Size = new-object System.Drawing.Size(200,30)
$ObjTypeDrop.Items.Add("Mailbox")
$ObjTypeDrop.Items.Add("Mailbox-Store")
$ObjTypeDrop.Items.Add("Delegated Exchange Admin")
$ObjTypeDrop.Add_SelectedValueChanged({
	switch($ObjTypeDrop.SelectedItem.ToString()){
		"Mailbox" {$dgDataGrid.datasource = $rs1Table}
		"Mailbox-Store" {$dgDataGrid.datasource = $msr1Table}
		"Delegated Exchange Admin" {$dgDataGrid.datasource = $drTable1}	
	}
})
$form.Controls.Add($ObjTypeDrop)

# Add Object Type DropLable
$ObjTypelableBox = new-object System.Windows.Forms.Label
$ObjTypelableBox.Location = new-object System.Drawing.Size(20,20) 
$ObjTypelableBox.size = new-object System.Drawing.Size(150,20) 
$ObjTypelableBox.Text = "Select Permission Type"
$form.Controls.Add($ObjTypelableBox) 

# Add Ace Type DropLable
$AceTypelableBox = new-object System.Windows.Forms.Label
$AceTypelableBox.Location = new-object System.Drawing.Size(660,20) 
$AceTypelableBox.size = new-object System.Drawing.Size(80,20) 
$AceTypelableBox.Text = "ACE Type"
$form.Controls.Add($AceTypelableBox) 

# Add Ace Type Drop Down
$AceTypeDrop = new-object System.Windows.Forms.ComboBox
$AceTypeDrop.Location = new-object System.Drawing.Size(740,20)
$AceTypeDrop.Size = new-object System.Drawing.Size(200,30)
$AceTypeDrop.Items.Add("Mailbox-Rights")
$AceTypeDrop.Items.Add("SendAs/ReciveAS-Rights")
$AceTypeDrop.Items.Add("Reverse-Mailbox-Rights")
$AceTypeDrop.Items.Add("Reverse-SendAs/ReciveAS-Rights")
$form.Controls.Add($AceTypeDrop)

# Select Target Group Box

$OfGbox1 =  new-object System.Windows.Forms.GroupBox
$OfGbox1.Location = new-object System.Drawing.Size(12,0)
$OfGbox1.Size = new-object System.Drawing.Size(520,75)
$OfGbox1.Text = "Select Permission Object Type"
$form.Controls.Add($OfGbox1)

# DACL Content Group Box

$OfGbox =  new-object System.Windows.Forms.GroupBox
$OfGbox.Location = new-object System.Drawing.Size(550,0)
$OfGbox.Size = new-object System.Drawing.Size(450,75)
$OfGbox.Text = "Show DACL Contents"
$form.Controls.Add($OfGbox)

# Add Export Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(10,585)
$exButton1.Size = new-object System.Drawing.Size(150,20)
$exButton1.Text = "Export All Permissions"
$exButton1.Add_Click({exportTable})
$form.controls.Add($exButton1)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,80) 
$dgDataGrid.size = new-object System.Drawing.Size(520,500)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)

$dgDataGrid1 = new-object System.windows.forms.DataGridView
$dgDataGrid1.Location = new-object System.Drawing.Size(550,80) 
$dgDataGrid1.size = new-object System.Drawing.Size(450,500)
$dgDataGrid1.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid1)


enumMailboxperms
EnumMailStorePerms
enDelExchangeRight

$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()




