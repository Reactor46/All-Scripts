[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 


function GetPerms(){
$logTable.clear()

If ($seAuthCheck.Checked -eq $true){
	$mailboxes = get-mailbox -server $snServerNameTextBox.Text -ResultSize Unlimited 
}
else{
	$mailboxes = get-mailbox -ResultSize Unlimited 
}
$mailboxes | foreach-object{
	$alias = $_.alias + ":\Calendar"
	$displayName = $_.DisplayName
	write-host $alias
 	$permissions = Get-MailboxFolderPermission $alias | Where-Object {$_.Identity.ToString() -eq "Default"}
	if($permissions -ne $null){
		Add-Member -InputObject $permissions -MemberType NoteProperty -Name "Alias" -Value $alias -Force
		$stringPerms = ""
		foreach($perms in $permissions.AccessRights){$stringPerms = $stringPerms + $perms + " "}
		$logTable.rows.add($displayName,$permissions.alias,$stringPerms)
	}
	
}
$dgDataGrid.DataSource = $logTable

}

Function UpdatePerms{
if ($dgDataGrid.SelectedRows.Count -eq 0){
	$mbtoSet =  $dgDataGrid.Rows[$dgDataGrid.CurrentCell.RowIndex].Cells[1].Value 
	$newperm = ""
	switch ($npNewpermDrop.Text){
		"None" {$newperm = "None"}
		"FreeBusyTimeOnly" {$newperm = "AvailabilityOnly"}									
		"FreeBusyTimeAndSubjectAndLocation" {$newperm = "LimitedDetails"}	
		"Reviewer" {$newperm = "Reviewer"}	
		"Contributer" {$newperm = "Contributer"}	
		"Author" {$newperm = "Author"}	
		"NonEditingAuthor" {$newperm = "NonEditingAuthor"}	
		"PublishingAuthor"{$newperm = "PublishingAuthor"}	
		"Author" {$newperm = "Author"}	
		"Editor" {$newperm = "Editor"}	
		"PublishingEditor"{$newperm = "PublishingEditor"}	
	}
	Set-MailboxFolderPermission -id $mbtoSet -User Default -AccessRights $newperm
	write-host "Permission updated" + $npNewpermDrop.Text
}
else{
	$lcLoopCount = 0
	while ($lcLoopCount -le ($dgDataGrid.SelectedRows.Count-1)) {
		$mbtoSet =  $dgDataGrid.SelectedRows[$lcLoopCount].Cells[1].Value
		
		switch ($npNewpermDrop.Text){
			"None" {$newperm = "None"}
			"FreeBusyTimeOnly" {$newperm = "AvailabilityOnly"}									
			"FreeBusyTimeAndSubjectAndLocation" {$newperm = "LimitedDetails"}	
			"Reviewer" {$newperm = "Reviewer"}	
			"Contributer" {$newperm = "Contributer"}	
			"Author" {$newperm = "Author"}	
			"NonEditingAuthor" {$newperm = "NonEditingAuthor"}	
			"PublishingAuthor"{$newperm = "PublishingAuthor"}	
			"Author" {$newperm = "Author"}	
			"Editor" {$newperm = "Editor"}	
			"PublishingEditor"{$newperm = "PublishingEditor"}	
		}
		Set-MailboxFolderPermission -id $mbtoSet -User Default -AccessRights $newperm
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
$logTable.Columns.Add("MailboxFolderId");
$logTable.Columns.Add("Default-Permissions");



# Add Server DropLable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,60) 
$snServerNamelableBox.size = new-object System.Drawing.Size(70,20) 
$snServerNamelableBox.Text = "ServerName"
$form.Controls.Add($snServerNamelableBox) 

# Add ServerNameText
$snServerNameTextBox = new-object System.Windows.Forms.TextBox 
$snServerNameTextBox.Location = new-object System.Drawing.Size(90,60) 
$snServerNameTextBox.size = new-object System.Drawing.Size(150,20) 
$snServerNameTextBox.Enabled = $false
$form.Controls.Add($snServerNameTextBox) 

$seAuthCheck =  new-object System.Windows.Forms.CheckBox
$seAuthCheck.Location = new-object System.Drawing.Size(250,60)
$seAuthCheck.Size = new-object System.Drawing.Size(130,25)
$seAuthCheck.Text = "Filter by"
$seAuthCheck.Add_Click({if ($seAuthCheck.Checked -eq $true){
			$snServerNameTextBox.Enabled = $true
			}
			else{
			$snServerNameTextBox.Enabled = $false}})
$form.Controls.Add($seAuthCheck)

# Add Get Perms Button

$gpgetperms = new-object System.Windows.Forms.Button
$gpgetperms.Location = new-object System.Drawing.Size(10,20)
$gpgetperms.Size = new-object System.Drawing.Size(140,23)
$gpgetperms.Text = "Enumerate Permissions"
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

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,130) 
$dgDataGrid.size = new-object System.Drawing.Size(750,550) 
$dgDataGrid.AutoSizeColumnsMode = "AllCells"
$dgDataGrid.SelectionMode = "FullRowSelect"
$form.Controls.Add($dgDataGrid)


$form.Text = "Exchange 2010 Default Calendar Permissions Form"
$form.size = new-object System.Drawing.Size(800,730) 

$form.autoscroll = $true
$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
