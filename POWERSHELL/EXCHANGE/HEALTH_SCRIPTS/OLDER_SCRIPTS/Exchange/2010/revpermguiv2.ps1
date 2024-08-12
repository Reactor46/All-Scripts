[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
[void][Reflection.Assembly]::LoadFile("c:\temp\EWSUtil.dll")

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
	[byte[]]$DaclByte = $emProps["msexchmailboxsecuritydescriptor"][0]
	$adDACL = new-object System.DirectoryServices.ActiveDirectorySecurity
	$adDACL.SetSecurityDescriptorBinaryForm($DaclByte)
	$mbRightsacls =$adDACL.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])
	"Processing Mailbox - " +  $uoUserobject.DisplayName
	$mbCount = 0
	$fpCount = 0
	$dgCount = 0
	$dfCount = 0
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
	$mbMailboxEmail = $uoUserobject.mail.ToString()
	write-host "Doing Mailbox Permissions with EWS"
	if ($AutoDiscoveryCheck.Checked -eq $true){ $casurl = $null}
	else { $casurl= $unCASUrlTextBox.text}
	$useImp = $false
	if ($seImpersonationCheck.Checked -eq $true) {
		$useImp = $true
	}
	if ($seAuthCheck.Checked -eq $false) {
		$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$useImp, $unUserNameTextBox.Text, $unPasswordTextBox.Text, $unDomainTextBox.Text,$casUrl)
	}
	else{
		$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$useImp, "", "", "",$casUrl)
	}
	$ewc.esb.CookieContainer = New-Object System.Net.CookieContainer
	$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 6
	for ($fcint=0;$fcint -lt 6;$fcint++){
		$dTypeFld = new-object EWSUtil.EWS.DistinguishedFolderIdType
	
		switch ($fcint){
			0 {$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox}
			1 {$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::calendar}
			2 {$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::contacts}
			3 {$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::tasks}
			4 {$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::journal}
			5 {$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::msgfolderroot}
		}
		$mbMailbox = new-object EWSUtil.EWS.EmailAddressType
		$mbMailbox.EmailAddress = $mbMailboxEmail
		$dTypeFld.Mailbox = $mbMailbox
		$fldarry[$fcint] = $dTypeFld
	}
	$Folders = $ewc.GetFolder($fldarry)
	If ($Folders.Count -ne 0) {
		 ForEach ($Folder in $Folders) {
			if ($Folder.GetType() -eq  [EWSUtil.EWS.CalendarFolderType]){
				ForEach ($Permissions in $Folder.PermissionSet.CalendarPermissions){
					if ($Permissions.UserId.DistinguishedUserSpecified -eq $false){
						$sidbind = "LDAP://<SID=" + $Permissions.UserId.SID.ToString() + ">"
						$AceName = $ace.IdentityReference.Value 
						$aceuser = [ADSI]$sidbind
						[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),("FolderRight-" + $Folder.DisplayName),$aceuser.samaccountname.ToString(),$ewc.enumOutlookRole($Permissions),"Allow")
						$fpCount++
						if ($rvFolderPerms.Containskey($aceuser.samaccountname.ToString())){
							$rvFolderPerms[$aceuser.samaccountname.ToString()] = [int]$rvFolderPerms[$aceuser.samaccountname.ToString()] +1			
						}
						else {
							$rvFolderPerms.add($aceuser.samaccountname.ToString(),1)
						}
					}
					else{
						if ($Permissions.UserId.DistinguishedUser -eq [EWSUtil.EWS.DistinguishedUserType]::Default){
							if ($Permissions.CalendarPermissionLevel -ne [EWSUtil.EWS.CalendarPermissionLevelType]::None){
								[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),("FolderRight-" + $Folder.DisplayName),"Default",$ewc.enumOutlookRole($Permissions),"Allow")
								$dfCount++
							 }
						}
					}
				}


			}
			else {
				ForEach ($Permissions in $Folder.PermissionSet.Permissions){
					if ($Permissions.UserId.DistinguishedUserSpecified -eq $false){
						$sidbind = "LDAP://<SID=" + $Permissions.UserId.SID.ToString() + ">"
						$AceName = $ace.IdentityReference.Value 
						$aceuser = [ADSI]$sidbind
						[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),("FolderRight-" + $Folder.DisplayName),$aceuser.samaccountname.ToString(),$Permissions.PermissionLevel.ToString(),"Allow")
						$fpCount++
						if ($rvFolderPerms.Containskey($aceuser.samaccountname.ToString())){
							$rvFolderPerms[$aceuser.samaccountname.ToString()] = [int]$rvFolderPerms[$aceuser.samaccountname.ToString()] +1			
						}
						else {
							$rvFolderPerms.add($aceuser.samaccountname.ToString(),1)
						}
					}
					else{
						if ($Permissions.UserId.DistinguishedUser -eq [EWSUtil.EWS.DistinguishedUserType]::Default){
							if ($Permissions.PermissionLevel -ne [EWSUtil.EWS.PermissionLevelType]::None){
								[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),("FolderRight-" + $Folder.DisplayName),"Default",$Permissions.PermissionLevel.ToString(),"Allow")
								$dfCount++
							 }
						}
					}
				}
			}
		}
	}
	$delegates = $ewc.getdeletgates($mbMailboxEmail)
	foreach ($delegate in $delegates){
		$sidbind = "LDAP://<SID=" + $delegate.DelegateUser.UserId.SID.ToString() + ">"
		$AceName = $ace.IdentityReference.Value 
		$aceuser = [ADSI]$sidbind
		$dgCount++
		[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"DelegateRight",$aceuser.samaccountname.ToString(),"Mailbox Delegated","Allow")
		if ($delegate.DelegateUser.ReceiveCopiesOfMeetingMessagesSpecified -eq $true){
			if ($delegate.DelegateUser.ReceiveCopiesOfMeetingMessages-eq $true){
				$dgCount++
				[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"DelegateRight",$aceuser.samaccountname.ToString(),"Recieve Meetings","Allow")
				if ($rvDelegatePerms.Containskey($aceuser.samaccountname.ToString())){
					$rvDelegatePerms[$aceuser.samaccountname.ToString()] = [int]$rvDelegatePerms[$aceuser.samaccountname.ToString()] +1			
				}
				else {
				$rvDelegatePerms.add($aceuser.samaccountname.ToString(),1)
				}
			}
		}

		if ($delegate.DelegateUser.ViewPrivateItemsSpecified -eq $true){
			if ($delegate.DelegateUser.ViewPrivateItems-eq $true){
				$dgCount++
				[VOID]$rsTable.rows.add($uoUserobject.samaccountname.ToString(),"DelegateRight",$aceuser.samaccountname.ToString(),"View Private Items","Allow")
				if ($rvDelegatePerms.Containskey($aceuser.samaccountname.ToString())){
					$rvDelegatePerms[$aceuser.samaccountname.ToString()] = [int]$rvDelegatePerms[$aceuser.samaccountname.ToString()] +1			
				}
				else {
					$rvDelegatePerms.add($aceuser.samaccountname.ToString(),1)
				}
			}
		}
		if ($rvDelegatePerms.Containskey($aceuser.samaccountname.ToString())){
			$rvDelegatePerms[$aceuser.samaccountname.ToString()] = [int]$rvDelegatePerms[$aceuser.samaccountname.ToString()] +1			
		}
		else {
			$rvDelegatePerms.add($aceuser.samaccountname.ToString(),1)
		}

	}
	$nmMailboxPerms.Add($uoUserobject.samaccountname.ToString(),$mbCount)
	$nmSendRecieve.Add($uoUserobject.samaccountname.ToString(),$srCount)
	$fpFolderPerms.Add($uoUserobject.samaccountname.ToString(),$fpCount)
	$nmDelegatePerms.Add($uoUserobject.samaccountname.ToString(),$dgCount)
	$nmDefualtPerms.Add($uoUserobject.samaccountname.ToString(),$dfCount)
}
foreach($key in $nmMailboxPerms.keys){
	$rvSRRights = 0
	$rvMbRights = 0
	$rvFpRights = 0
	$rvDRRights = 0
	if ($rvMailboxPerms.Containskey($key)){
		$rvMbRights = $rvMailboxPerms[$key]
	}
	if ($rvSendRecieve.Containskey($key)){
		$rvSRRights = $rvSendRecieve[$key]
	}
	if ($rvFolderPerms.Containskey($key)){
		$rvFpRights = $rvFolderPerms[$key]
	}
	if ($rvDelegatePerms.Containskey($key)){
		$rvDRRights = $rvDelegatePerms[$key]
	}
	$rs1Table.Rows.Add($key, $nmMailboxPerms[$key],$fpFolderPerms[$key],$nmDelegatePerms[$key],$nmDefualtPerms[$key],$nmSendRecieve[$key],$rvMbRights,$rvFpRights,$rvDRRights,$rvSRRights)
	
}
# write-host $nmSendRecieve
$dgDataGrid.datasource = $rs1Table
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

function showACL{
if ($ObjTypeDrop.SelectedItem -eq $null -bor $ObjTypeDrop.SelectedItem -eq "Mailbox"){

	if ($AceTypeDrop.SelectedItem -ne $null){
		switch($AceTypeDrop.SelectedItem.ToString()){
			"Mailbox-Rights" {$rows = $rsTable.Select("MailboxName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType = 'MailboxRight'")}
			"Folder-Rights" {$rows = $rsTable.Select("MailboxName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType like 'Folder%'")}				
			"Delegate-Rights" {$rows = $rsTable.Select("MailboxName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType = 'DelegateRight'")}				
			"SendAs/ReciveAS-Rights" {$rows = $rsTable.Select("MailboxName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType ='SendAS-RecieveAS'")}
			"Reverse-Mailbox-Rights" {$rows = $rsTable.Select("UserName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType = 'MailboxRight'")}
			"Reverse-Folder-Rights" {$rows = $rsTable.Select("UserName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType like 'Folder%'")}
			"Reverse-Delegate-Rights" {$rows = $rsTable.Select("UserName = '" +  $rs1Table.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] + "' And ACLType = 'DelegateRight'")}
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

$formconf = new-object System.Windows.Forms.form 
$formconf.Text = "2007 Audit Tracker EWS config"

$AutoDiscoverylableBox = new-object System.Windows.Forms.Label
$AutoDiscoverylableBox.Location = new-object System.Drawing.Size(10,30) 
$AutoDiscoverylableBox.Size = new-object System.Drawing.Size(130,20) 
$AutoDiscoverylableBox.Text = "Use AutoDiscover"
$formconf.controls.Add($AutoDiscoverylableBox) 

$AutoDiscoveryCheck =  new-object System.Windows.Forms.CheckBox
$AutoDiscoveryCheck.Location = new-object System.Drawing.Size(140,25)
$AutoDiscoveryCheck.Size = new-object System.Drawing.Size(30,25)
$AutoDiscoveryCheck.Checked = $true
$AutoDiscoveryCheck.Add_Click({if ($AutoDiscoveryCheck.Checked -eq $false){
	$unCASUrlTextBox.enabled = $true}
	else{$unCASUrlTextBox.enabled = $false}})
$formconf.controls.Add($AutoDiscoveryCheck)

# Add Impersonation Clause

$esImpersonationlableBox = new-object System.Windows.Forms.Label
$esImpersonationlableBox.Location = new-object System.Drawing.Size(10,55) 
$esImpersonationlableBox.Size = new-object System.Drawing.Size(130,20) 
$esImpersonationlableBox.Text = "Use EWS Impersonation"
$formconf.controls.Add($esImpersonationlableBox) 

$seImpersonationCheck =  new-object System.Windows.Forms.CheckBox
$seImpersonationCheck.Location = new-object System.Drawing.Size(140,50)
$seImpersonationCheck.Size = new-object System.Drawing.Size(30,25)
$formconf.controls.Add($seImpersonationCheck)

# Add Auth Clause

$esAuthlableBox = new-object System.Windows.Forms.Label
$esAuthlableBox.Location = new-object System.Drawing.Size(10,80) 
$esAuthlableBox.Size = new-object System.Drawing.Size(130,20) 
$esAuthlableBox.Text = "Use Current Credentials"
$formconf.controls.Add($esAuthlableBox) 

$seAuthCheck =  new-object System.Windows.Forms.CheckBox
$seAuthCheck.Location = new-object System.Drawing.Size(140,80)
$seAuthCheck.Size = new-object System.Drawing.Size(30,25)
$seAuthCheck.Checked = $true
$seAuthCheck.Add_Click({if ($seAuthCheck.Checked -eq $false){
			$unUserNameTextBox.Enabled = $true
			$unPasswordTextBox.Enabled = $true
			$unDomainTextBox.Enabled = $true
			}
			else{
				$unUserNameTextBox.Enabled = $false
				$unPasswordTextBox.Enabled = $false
				$unDomainTextBox.Enabled = $false}})
$formconf.controls.Add($seAuthCheck)

# Add UserName Box
$unUserNameTextBox = new-object System.Windows.Forms.TextBox 
$unUserNameTextBox.Location = new-object System.Drawing.Size(230,85) 
$unUserNameTextBox.size = new-object System.Drawing.Size(100,20) 
$formconf.controls.Add($unUserNameTextBox) 

# Add UserName Lable
$unUserNamelableBox = new-object System.Windows.Forms.Label
$unUserNamelableBox.Location = new-object System.Drawing.Size(170,85) 
$unUserNamelableBox.size = new-object System.Drawing.Size(60,20) 
$unUserNamelableBox.Text = "UserName"
$unUserNameTextBox.Enabled = $false
$formconf.controls.Add($unUserNamelableBox) 

# Add Password Box
$unPasswordTextBox = new-object System.Windows.Forms.TextBox 
$unPasswordTextBox.PasswordChar = "*"
$unPasswordTextBox.Location = new-object System.Drawing.Size(400,85) 
$unPasswordTextBox.size = new-object System.Drawing.Size(100,20) 
$formconf.controls.Add($unPasswordTextBox) 

# Add Password Lable
$unPasswordlableBox = new-object System.Windows.Forms.Label
$unPasswordlableBox.Location = new-object System.Drawing.Size(340,85) 
$unPasswordlableBox.size = new-object System.Drawing.Size(60,20) 
$unPasswordlableBox.Text = "Password"
$unPasswordTextBox.Enabled = $false
$formconf.controls.Add($unPasswordlableBox) 

# Add Domain Box
$unDomainTextBox = new-object System.Windows.Forms.TextBox 
$unDomainTextBox.Location = new-object System.Drawing.Size(550,85) 
$unDomainTextBox.size = new-object System.Drawing.Size(100,20) 
$formconf.controls.Add($unDomainTextBox) 

# Add Domain Lable
$unDomainlableBox = new-object System.Windows.Forms.Label
$unDomainlableBox.Location = new-object System.Drawing.Size(510,85) 
$unDomainlableBox.size = new-object System.Drawing.Size(50,20) 
$unDomainlableBox.Text = "Domain"
$unDomainTextBox.Enabled = $false
$formconf.controls.Add($unDomainlableBox) 


# Add CASUrl Box
$unCASUrlTextBox = new-object System.Windows.Forms.TextBox 
$unCASUrlTextBox.Location = new-object System.Drawing.Size(240,30) 
$unCASUrlTextBox.size = new-object System.Drawing.Size(400,20) 
$unCASUrlTextBox.text = ""
$unCASUrlTextBox.enabled = $false
$formconf.Controls.Add($unCASUrlTextBox) 

# Add CASUrl Lable
$unCASUrllableBox = new-object System.Windows.Forms.Label
$unCASUrllableBox.Location = new-object System.Drawing.Size(180,30) 
$unCASUrllableBox.size = new-object System.Drawing.Size(50,20) 
$unCASUrllableBox.Text = "CASUrl"
$formconf.Controls.Add($unCASUrllableBox) 


# Add Execute Button

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(10,130)
$exButton.Size = new-object System.Drawing.Size(110,20)
$exButton.Text = "Execute"
$exButton.Add_Click({$formconf.Close()})
$formconf.controls.Add($exButton)

$formconf.size = new-object System.Drawing.Size(800,200) 
$formconf.Add_Shown({$formconf.Activate()})
$formconf.autoscroll = $true
$formconf.ShowDialog()


$nmMailboxPerms = @{ }
$nmSendRecieve = @{ }
$fpFolderPerms = @{ }
$duFolderPerms = @{ }
$rvMailboxPerms = @{ }
$rvSendRecieve = @{ }
$rvFolderPerms = @{ }
$nmDelegatePerms = @{ }
$rvDelegatePerms = @{ }
$nmDefualtPerms = @{ }

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
$rs1Table.Columns.Add("Folder")
$rs1Table.Columns.Add("Delegate")
$rs1Table.Columns.Add("Default")
$rs1Table.Columns.Add("Send/RecieveAS")
$rs1Table.Columns.Add("Revese_Mailbox")
$rs1Table.Columns.Add("Revese_Folder")
$rs1Table.Columns.Add("Revese_Delegate")
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


$dfrTable = New-Object System.Data.DataTable
$dfrTable.TableName = "Delegate Folder Rights"
$dfrTable.Columns.Add("MailboxName")
$dfrTable.Columns.Add("ACLType")
$dfrTable.Columns.Add("UserName")
$dfrTable.Columns.Add("Rights")
$dfrTable.Columns.Add("Status")
$Dataset.tables.add($dfrTable)


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
$shaces.Location = new-object System.Drawing.Size(20,530)
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
$AceTypelableBox.Location = new-object System.Drawing.Size(120,535) 
$AceTypelableBox.size = new-object System.Drawing.Size(60,20) 
$AceTypelableBox.Text = "ACE Type"
$form.Controls.Add($AceTypelableBox) 

# Add Ace Type Drop Down
$AceTypeDrop = new-object System.Windows.Forms.ComboBox
$AceTypeDrop.Location = new-object System.Drawing.Size(190,530)
$AceTypeDrop.Size = new-object System.Drawing.Size(200,30)
$AceTypeDrop.Items.Add("Mailbox-Rights")
$AceTypeDrop.Items.Add("Folder-Rights")
$AceTypeDrop.Items.Add("Delegate-Rights")
$AceTypeDrop.Items.Add("SendAs/ReciveAS-Rights")
$AceTypeDrop.Items.Add("Reverse-Mailbox-Rights")
$AceTypeDrop.Items.Add("Reverse-Folder-Rights")
$AceTypeDrop.Items.Add("Reverse-Delegate-Rights")
$AceTypeDrop.Items.Add("Reverse-SendAs/ReciveAS-Rights")
$form.Controls.Add($AceTypeDrop)

# Select Target Group Box

$OfGbox1 =  new-object System.Windows.Forms.GroupBox
$OfGbox1.Location = new-object System.Drawing.Size(12,0)
$OfGbox1.Size = new-object System.Drawing.Size(520,50)
$OfGbox1.Text = "Select Permission Object Type"
$form.Controls.Add($OfGbox1)

# DACL Content Group Box

$OfGbox =  new-object System.Windows.Forms.GroupBox
$OfGbox.Location = new-object System.Drawing.Size(12,510)
$OfGbox.Size = new-object System.Drawing.Size(450,50)
$OfGbox.Text = "Show DACL Contents"
$form.Controls.Add($OfGbox)

# Add Export Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(490,530)
$exButton1.Size = new-object System.Drawing.Size(150,20)
$exButton1.Text = "Export All Permissions"
$exButton1.Add_Click({exportTable})
$form.controls.Add($exButton1)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,55) 
$dgDataGrid.size = new-object System.Drawing.Size(1000,450)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)

$dgDataGrid1 = new-object System.windows.forms.DataGridView
$dgDataGrid1.Location = new-object System.Drawing.Size(10,570) 
$dgDataGrid1.size = new-object System.Drawing.Size(1000,230)
$dgDataGrid1.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid1)


enumMailboxperms
EnumMailStorePerms
enDelExchangeRight

$form.Add_Shown({$form.Activate()})
$form.ShowDialog()




