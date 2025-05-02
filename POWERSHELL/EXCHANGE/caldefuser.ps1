[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
[void][Reflection.Assembly]::LoadFile("c:\temp\EWSUtil.dll")
[array]$calurls = Get-WebServicesVirtualDirectory
$strRootURI = $calurls[0].InternalUrl.AbsoluteUri
$strRootURI

function GetPerms(){
$logTable.clear()


get-mailbox -server $snServerNameDrop.SelectedItem.ToString() -ResultSize Unlimited | foreach-object{$colnum
	$defperm = ""
	$unUsername = ""
	$pnPpassword = ""
	$dnDomainName = ""
	$inInpersonate = $false
	if ($seAuthCheck.Checked -eq $true){
		$unUsername = $unUserNameTextBox.Text
		$pnPpassword = $unPasswordTextBox.Text
		$dnDomainName = $unDomainTextBox.Text
	}
	if ($seImpersonationCheck.Checked -eq $true){
		$uoUser = [ADSI]("LDAP://" + $_.DistinguishedName.ToString())	
		if ($uoUser.PSBase.InvokeGet("AccountDisabled") -ne $true){$inInpersonate = $true }
			
	}
	$calutil = new-object EWSUtil.CalendarUtil($_.WindowsEmailAddress,$inInpersonate,$unUsername,$pnPpassword,$dnDomainName,$unCASUrlTextBox.text)
	for ($cpint=0;$cpint -lt $calutil.CalendarDACL.Count; $cpint++){
		if ($calutil.CalendarDACL[$cpint].UserId.DistinguishedUserSpecified -eq $true){
			if ($calutil.CalendarDACL[$cpint].UserId.DistinguishedUser -eq [EWSUtil.EWS.DistinguishedUserType]::Default){
				write-host "Processing : " + $_.WindowsEmailAddress
				$defperm = $calutil.enumOutlookRole($calutil.CalendarDACL[$cpint])
			}
				
		}
	}
	$logTable.rows.add($_.DisplayName,$_.WindowsEmailAddress,$defperm)
}
$dgDataGrid.DataSource = $logTable

}


Function UpdatePerms{
$unUsername = ""
$pnPpassword = ""
$dnDomainName = ""
$inInpersonate = $false
if ($seAuthCheck.Checked -eq $true){
	$unUsername = $unUserNameTextBox.Text
	$pnPpassword = $unPasswordTextBox.Text
	$dnDomainName = $unDomainTextBox.Text
}
	if ($seImpersonationCheck.Checked -eq $true){
		$mbchk = get-mailbox $dgDataGrid.Rows[$dgDataGrid.CurrentCell.RowIndex].Cells[1].Value
		$uoUser = [ADSI]("LDAP://" + $mbchk.DistinguishedName.ToString())	
		if ($uoUser.PSBase.InvokeGet("AccountDisabled") -ne $true){$inInpersonate = $true }
			
	}
if ($dgDataGrid.SelectedRows.Count -eq 0){
	$mbtoSet =  $dgDataGrid.Rows[$dgDataGrid.CurrentCell.RowIndex].Cells[1].Value 
	$calutil = new-object EWSUtil.CalendarUtil($mbtoSet,$inInpersonate,$unUsername,$pnPpassword,$dnDomainName,$unCASUrlTextBox.text)
	switch ($npNewpermDrop.Text){
		"None" {$calutil.CalendarDACL.Add($calutil.NonePermissions("default"))
			$calutil.FreeBusyDACL.Add($calutil.FolderEmpty("default"))}
		"FreeBusyTimeOnly" {$calutil.CalendarDACL.Add($calutil.FreeBusyTimeOnly("default"))
				  $calutil.FreeBusyDACL.Add($calutil.FolderEmpty("default"))}									
		"FreeBusyTimeAndSubjectAndLocation" {$calutil.CalendarDACL.Add($calutil.FreeBusyTimeAndSubjectAndLocation("default"))
				  $calutil.FreeBusyDACL.Add($calutil.FolderEmpty("default"))}
		"Reviewer" {$calutil.CalendarDACL.Add($calutil.Reviewer("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderReviewer("default"))}
		"Contributer" {$calutil.CalendarDACL.Add($calutil.Contributer("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderContributer("default"))}
		"Author" {$calutil.CalendarDACL.Add($calutil.AuthorPermissions("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderAuthorPermissions("default"))}
		"NonEditingAuthor" {$calutil.CalendarDACL.Add($calutil.NonEditingAuthorPermissions("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderNonEditingAuthorPermissions("default"))}
		"PublishingAuthor" {$calutil.CalendarDACL.Add($calutil.PublishingAuthorPermissions("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderPublishingAuthorPermissions("default"))}
		"Author" {$calutil.CalendarDACL.Add($calutil.AuthorPermissions("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderAuthorPermissions("default"))}
		"Editor" {$calutil.CalendarDACL.Add($calutil.EditorPermissions("default"))
				  $calutil.FreeBusyDACL.Add($calutil.FolderEditorPermissions("default"))}
		"PublishingEditor"{$calutil.CalendarDACL.Add($calutil.PublishingEditorPermissions("default"))
				  $calutil.FreeBusyDACL.Add($calutil.FolderPublishingEditorPermissions("default"))}	 
	}
	$calutil.update()
	write-host "Permission updated" + $npNewpermDrop.Text
}
else{
	$lcLoopCount = 0
	while ($lcLoopCount -le ($dgDataGrid.SelectedRows.Count-1)) {
		$mbtoSet = $dgDataGrid.SelectedRows[$lcLoopCount].Cells[1].Value
		$calutil = new-object EWSUtil.CalendarUtil($mbtoSet,$inInpersonate,$unUsername,$pnPpassword,$dnDomainName,$unCASUrlTextBox.text)
		switch ($npNewpermDrop.Text){
			"None" {$calutil.CalendarDACL.Add($calutil.NonePermissions("default"))
				$calutil.FreeBusyDACL.Add($calutil.FolderEmpty("default"))}
			"FreeBusyTimeOnly" {$calutil.CalendarDACL.Add($calutil.FreeBusyTimeOnly("default"))
				  $calutil.FreeBusyDACL.Add($calutil.FolderEmpty("default"))}									
			"FreeBusyTimeAndSubjectAndLocation" {$calutil.CalendarDACL.Add($calutil.FreeBusyTimeAndSubjectAndLocation("default"))
				  $calutil.FreeBusyDACL.Add($calutil.FolderEmpty("default"))}
			"Reviewer" {$calutil.CalendarDACL.Add($calutil.Reviewer("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderReviewer("default"))}
			"Contributer" {$calutil.CalendarDACL.Add($calutil.Contributer("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderContributer("default"))}
			"Author" {$calutil.CalendarDACL.Add($calutil.AuthorPermissions("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderAuthorPermissions("default"))}
			"NonEditingAuthor" {$calutil.CalendarDACL.Add($calutil.NonEditingAuthorPermissions("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderNonEditingAuthorPermissions("default"))}
			"PublishingAuthor" {$calutil.CalendarDACL.Add($calutil.PublishingAuthorPermissions("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderPublishingAuthorPermissions("default"))}
			"Author" {$calutil.CalendarDACL.Add($calutil.AuthorPermissions("default"))
				   $calutil.FreeBusyDACL.Add($calutil.FolderAuthorPermissions("default"))}
			"Editor" {$calutil.CalendarDACL.Add($calutil.EditorPermissions("default"))
				  $calutil.FreeBusyDACL.Add($calutil.FolderEditorPermissions("default"))}
			"PublishingEditor"{$calutil.CalendarDACL.Add($calutil.PublishingEditorPermissions("default"))
				  $calutil.FreeBusyDACL.Add($calutil.FolderPublishingEditorPermissions("default"))}	 
		}
		$calutil.update()
		write-host "Permission updated" + $npNewpermDrop.Text
		$lcLoopCount += 1	
	}
}
write-host "end PermUpdate"
write-host "Refresh Perms"
GetPerms
}


$form = new-object System.Windows.Forms.form 
$form.Text = "Calender Permission Enum Tool"
$Dataset = New-Object System.Data.DataSet
$logTable = New-Object System.Data.DataTable
$logTable.TableName = "ActiveSyncLogs"
$logTable.Columns.Add("DisplayName");
$logTable.Columns.Add("EmailAddress");
$logTable.Columns.Add("Default-Permissions");



# Add Server DropLable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,20) 
$snServerNamelableBox.size = new-object System.Drawing.Size(70,20) 
$snServerNamelableBox.Text = "ServerName"
$form.Controls.Add($snServerNamelableBox) 

# Add Server Drop Down
$snServerNameDrop = new-object System.Windows.Forms.ComboBox
$snServerNameDrop.Location = new-object System.Drawing.Size(80,20)
$snServerNameDrop.Size = new-object System.Drawing.Size(130,30)
get-exchangeserver | ForEach-Object{$snServerNameDrop.Items.Add($_.Name)}
$form.Controls.Add($snServerNameDrop)


# Add Get Perms Button

$gpgetperms = new-object System.Windows.Forms.Button
$gpgetperms.Location = new-object System.Drawing.Size(220,20)
$gpgetperms.Size = new-object System.Drawing.Size(85,23)
$gpgetperms.Text = "Enum Perms"
$gpgetperms.Add_Click({GetPerms})
$form.Controls.Add($gpgetperms)

# Add New Permission Drop Down
$npNewpermDrop = new-object System.Windows.Forms.ComboBox
$npNewpermDrop.Location = new-object System.Drawing.Size(350,20)
$npNewpermDrop.Size = new-object System.Drawing.Size(190,30)
$npNewpermDrop.Items.Add("None")
$npNewpermDrop.Items.Add("FreeBusyTimeOnly")
$npNewpermDrop.Items.Add("FreeBusyTimeAndSubjectAndLocation")
$npNewpermDrop.Items.Add("Reviewer")
$npNewpermDrop.Items.Add("Contributer")
$npNewpermDrop.Items.Add("Author")
$npNewpermDrop.Items.Add("NonEditingAuthor")
$npNewpermDrop.Items.Add("PublishingAuthor")
$npNewpermDrop.Items.Add("Editor")
$npNewpermDrop.Items.Add("PublishingEditor")
$form.Controls.Add($npNewpermDrop)

# Add Apply Button

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(550,20)
$exButton.Size = new-object System.Drawing.Size(60,20)
$exButton.Text = "Apply"
$exButton.Add_Click({UpdatePerms})
$form.Controls.Add($exButton)

# New setting Group Box

$OfGbox =  new-object System.Windows.Forms.GroupBox
$OfGbox.Location = new-object System.Drawing.Size(320,0)
$OfGbox.Size = new-object System.Drawing.Size(300,50)
$OfGbox.Text = "New Permission Settings"
$form.Controls.Add($OfGbox)

# Add Impersonation Clause

$esImpersonationlableBox = new-object System.Windows.Forms.Label
$esImpersonationlableBox.Location = new-object System.Drawing.Size(10,50) 
$esImpersonationlableBox.Size = new-object System.Drawing.Size(130,20) 
$esImpersonationlableBox.Text = "Use EWS Impersonation"
$form.Controls.Add($esImpersonationlableBox) 

$seImpersonationCheck =  new-object System.Windows.Forms.CheckBox
$seImpersonationCheck.Location = new-object System.Drawing.Size(140,45)
$seImpersonationCheck.Size = new-object System.Drawing.Size(30,25)
$form.Controls.Add($seImpersonationCheck)

# Add Auth Clause

$esAuthlableBox = new-object System.Windows.Forms.Label
$esAuthlableBox.Location = new-object System.Drawing.Size(10,70) 
$esAuthlableBox.Size = new-object System.Drawing.Size(130,20) 
$esAuthlableBox.Text = "Specify Credentials"
$form.Controls.Add($esAuthlableBox) 

$seAuthCheck =  new-object System.Windows.Forms.CheckBox
$seAuthCheck.Location = new-object System.Drawing.Size(140,65)
$seAuthCheck.Size = new-object System.Drawing.Size(30,25)
$seAuthCheck.Add_Click({if ($seAuthCheck.Checked -eq $true){
			$unUserNameTextBox.Enabled = $true
			$unPasswordTextBox.Enabled = $true
			$unDomainTextBox.Enabled = $true
			}
			else{
				$unUserNameTextBox.Enabled = $false
				$unPasswordTextBox.Enabled = $false
				$unDomainTextBox.Enabled = $false}})
$form.Controls.Add($seAuthCheck)

# Add UserName Box
$unUserNameTextBox = new-object System.Windows.Forms.TextBox 
$unUserNameTextBox.Location = new-object System.Drawing.Size(230,70) 
$unUserNameTextBox.size = new-object System.Drawing.Size(100,20) 
$form.Controls.Add($unUserNameTextBox) 

# Add UserName Lable
$unUserNamelableBox = new-object System.Windows.Forms.Label
$unUserNamelableBox.Location = new-object System.Drawing.Size(170,70) 
$unUserNamelableBox.size = new-object System.Drawing.Size(60,20) 
$unUserNamelableBox.Text = "UserName"
$unUserNameTextBox.Enabled = $false
$form.Controls.Add($unUserNamelableBox) 

# Add Password Box
$unPasswordTextBox = new-object System.Windows.Forms.TextBox 
$unPasswordTextBox.PasswordChar = "*"
$unPasswordTextBox.Location = new-object System.Drawing.Size(400,70) 
$unPasswordTextBox.size = new-object System.Drawing.Size(100,20) 
$form.Controls.Add($unPasswordTextBox) 

# Add Password Lable
$unPasswordlableBox = new-object System.Windows.Forms.Label
$unPasswordlableBox.Location = new-object System.Drawing.Size(340,70) 
$unPasswordlableBox.size = new-object System.Drawing.Size(60,20) 
$unPasswordlableBox.Text = "Password"
$unPasswordTextBox.Enabled = $false
$form.Controls.Add($unPasswordlableBox) 

# Add Domain Box
$unDomainTextBox = new-object System.Windows.Forms.TextBox 
$unDomainTextBox.Location = new-object System.Drawing.Size(550,70) 
$unDomainTextBox.size = new-object System.Drawing.Size(100,20) 
$form.Controls.Add($unDomainTextBox) 

# Add Domain Lable
$unDomainlableBox = new-object System.Windows.Forms.Label
$unDomainlableBox.Location = new-object System.Drawing.Size(510,70) 
$unDomainlableBox.size = new-object System.Drawing.Size(50,20) 
$unDomainlableBox.Text = "Domain"
$unDomainTextBox.Enabled = $false
$form.Controls.Add($unDomainlableBox) 

# Add CASUrl Box
$unCASUrlTextBox = new-object System.Windows.Forms.TextBox 
$unCASUrlTextBox.Location = new-object System.Drawing.Size(70,100) 
$unCASUrlTextBox.size = new-object System.Drawing.Size(500,20) 
$unCASUrlTextBox.text = $strRootURI
$form.Controls.Add($unCASUrlTextBox) 

# Add CASUrl Lable
$unCASUrllableBox = new-object System.Windows.Forms.Label
$unCASUrllableBox.Location = new-object System.Drawing.Size(10,100) 
$unCASUrllableBox.size = new-object System.Drawing.Size(60,20) 
$unCASUrllableBox.Text = "CASUrl"
$form.Controls.Add($unCASUrllableBox) 


# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,130) 
$dgDataGrid.size = new-object System.Drawing.Size(650,550) 
$dgDataGrid.AutoSizeColumnsMode = "AllCells"
$dgDataGrid.SelectionMode = "FullRowSelect"
$form.Controls.Add($dgDataGrid)


$form.Text = "Exchange 2007 Default Calendar Permissions Form"
$form.size = new-object System.Drawing.Size(700,730) 

$form.autoscroll = $true
$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
