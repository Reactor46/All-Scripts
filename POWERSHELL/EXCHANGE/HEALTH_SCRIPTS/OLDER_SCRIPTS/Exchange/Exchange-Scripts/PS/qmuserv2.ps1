[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

Function Enableuser{
	$result = Enable-MailUser -Identity $emIdentityTextBox.Text  -Alias $AliasNameText.Text -ExternalEmailAddress ("SMTP:" + $exExternalMail.Text)
	if ($result -ne $null){write-host "User Mail enabled"}
	else{write-host "Error Mail enabling user check command Line for Details"}
	$unUserNameDropBox.Items.Clear()
	get-user  -sortby "LastName" -OrganizationalUnit $OUhash1[$ouOuNameDrop.SelectedItem.ToString()] | where { $_.RecipientType -eq "User" } | foreach-object{
	# $uname = $_.LastName + " " + $_.FirstName
	$uname = $_.DisplayName
	$unUserNameDropBox.Items.Add($uname)}
}


$OUhash1 = @{ }
$MBhash1 = @{ }

$form = new-object System.Windows.Forms.form 
$form.Text = "Exchange 2007 Quick User Mail User Form"
$form.size = new-object System.Drawing.Size(600,400) 

# Add OU DropLable
$ouOuNamelableBox = new-object System.Windows.Forms.Label
$ouOuNamelableBox.Location = new-object System.Drawing.Size(10,10) 
$ouOuNamelableBox.size = new-object System.Drawing.Size(70,20) 
$ouOuNamelableBox.Text = "OU Name"
$form.Controls.Add($ouOuNamelableBox) 

# Add OU Drop Down
$ouOuNameDrop = new-object System.Windows.Forms.ComboBox
$ouOuNameDrop.Location = new-object System.Drawing.Size(100,10)
$ouOuNameDrop.Size = new-object System.Drawing.Size(230,30)
$ouOuNameDrop.Items.Add("/Users")
$OUhash1.Add("/Users","Users")
$root = [ADSI]''
$searcher = new-object System.DirectoryServices.DirectorySearcher($root)
$searcher.Filter = '(objectClass=organizationalUnit)'
$searcher.PropertiesToLoad.Add("canonicalName")
$searcher.PropertiesToLoad.Add("Name")
$searcher1 = $searcher.FindAll()
foreach ($person in $searcher1){ 
[string]$ent = $person.Properties.canonicalname
$OUhash1.Add($ent.substring($ent.indexof("/"),$ent.length-$ent.indexof("/")),$ent)
$ouOuNameDrop.Items.Add($ent.substring($ent.indexof("/"),$ent.length-$ent.indexof("/")))
}
$ouOuNameDrop.Add_SelectedValueChanged({
	$unUserNameDropBox.Items.Clear()
	get-user  -sortby "LastName" -OrganizationalUnit $OUhash1[$ouOuNameDrop.SelectedItem.ToString()] | where { $_.RecipientType -eq "User" } | foreach-object{
	$uname = $_.DisplayName
	$unUserNameDropBox.Items.Add($uname)	
}


})

$form.Controls.Add($ouOuNameDrop)


# Add UserName Box
$unUserNameDropBox = new-object System.Windows.Forms.ComboBox
$unUserNameDropBox.Location = new-object System.Drawing.Size(100,40) 
$unUserNameDropBox.size = new-object System.Drawing.Size(330,20) 
$unUserNameDropBox.Add_SelectedValueChanged({
	$user = get-user $unUserNameDropBox.SelectedItem.ToString()
	$emIdentityTextBox.text = $user.Identity
	$unFirstNameTextBox.text = $User.FirstName
	$lnLastNameTextBox.text = $user.LastName
	$dsDisplayNameTextBox.text = $user.DisplayName
	$AliasNameText.text = $user.SamAccountName
	$exExternalMail.text = $user.WindowsEmailAddress
	$pscmd = "Enable-MailUser -Identity '" + $user.Identity + "' -Alias '" + $user.SamAccountName + "' -ExternalEmailAddress 'SMTP:" + $user.WindowsEmailAddress + "'"
	$msCmd.text = $pscmd 
})
$form.Controls.Add($unUserNameDropBox) 


# Add Username Lable
$unUserNamelableBox = new-object System.Windows.Forms.Label
$unUserNamelableBox.Location = new-object System.Drawing.Size(10,40) 
$unUserNamelableBox.size = new-object System.Drawing.Size(100,20) 
$unUserNamelableBox.Text = "Username UPN"
$form.Controls.Add($unUserNamelableBox) 

# Add Identity Box
$emIdentityTextBox = new-object System.Windows.Forms.TextBox 
$emIdentityTextBox.Location = new-object System.Drawing.Size(100,65) 
$emIdentityTextBox.size = new-object System.Drawing.Size(330,20) 
$emIdentityTextBox.Add_TextChanged({
	$pscmd = "Enable-MailUser -Identity '" + $emIdentityTextBox.Text + "' -Alias '" + $AliasNameText.Text + "' -ExternalEmailAddress 'SMTP:" + $exExternalMail.Text + "'"
	$msCmd.text = $pscmd 
})
$form.Controls.Add($emIdentityTextBox) 

# Add Identity Lable
$emIdentitylableBox = new-object System.Windows.Forms.Label
$emIdentitylableBox.Location = new-object System.Drawing.Size(10,65) 
$emIdentitylableBox.size = new-object System.Drawing.Size(100,20) 
$emIdentitylableBox.Text = "Identity"
$form.Controls.Add($emIdentitylableBox) 

# Add FirstName Box
$unFirstNameTextBox = new-object System.Windows.Forms.TextBox 
$unFirstNameTextBox.Location = new-object System.Drawing.Size(100,90) 
$unFirstNameTextBox.size = new-object System.Drawing.Size(130,20) 
$form.Controls.Add($unFirstNameTextBox) 

# Add FirstName Lable
$unFirstNamelableBox = new-object System.Windows.Forms.Label
$unFirstNamelableBox.Location = new-object System.Drawing.Size(10,90) 
$unFirstNamelableBox.size = new-object System.Drawing.Size(60,20) 
$unFirstNamelableBox.Text = "First Name"
$form.Controls.Add($unFirstNamelableBox) 

# Add LastName Box
$lnLastNameTextBox = new-object System.Windows.Forms.TextBox 
$lnLastNameTextBox.Location = new-object System.Drawing.Size(100,120) 
$lnLastNameTextBox.size = new-object System.Drawing.Size(130,20) 
$form.Controls.Add($lnLastNameTextBox) 

# Add LastName Lable
$lnLastNamelableBox = new-object System.Windows.Forms.Label
$lnLastNamelableBox.Location = new-object System.Drawing.Size(10,120) 
$lnLastNamelableBox.size = new-object System.Drawing.Size(60,20) 
$lnLastNamelableBox.Text = "Last Name"
$form.Controls.Add($lnLastNamelableBox) 

# Add DisplayName Box
$dsDisplayNameTextBox = new-object System.Windows.Forms.TextBox 
$dsDisplayNameTextBox.Location = new-object System.Drawing.Size(100,150) 
$dsDisplayNameTextBox.size = new-object System.Drawing.Size(130,20) 
$form.Controls.Add($dsDisplayNameTextBox) 

# Add DisplayName Lable
$dsDisplayNamelableBox = new-object System.Windows.Forms.Label
$dsDisplayNamelableBox.Location = new-object System.Drawing.Size(10,150) 
$dsDisplayNamelableBox.size = new-object System.Drawing.Size(100,20) 
$dsDisplayNamelableBox.Text = "Display Name"
$form.Controls.Add($dsDisplayNamelableBox) 

# Add Alias Text
$AliasNameText = new-object System.Windows.Forms.TextBox
$AliasNameText.Location = new-object System.Drawing.Size(100,180)
$AliasNameText.Size = new-object System.Drawing.Size(230,30)
$AliasNameText.Add_TextChanged({
	$pscmd = "Enable-MailUser -Identity '" + $emIdentityTextBox.Text + "' -Alias '" + $AliasNameText.Text + "' -ExternalEmailAddress 'SMTP:" + $exExternalMail.Text + "'"
	$msCmd.text = $pscmd 
})
$form.Controls.Add($AliasNameText)

# Add Alias TextBox Lable
$AliasNameTextlableBox = new-object System.Windows.Forms.Label
$AliasNameTextlableBox.Location = new-object System.Drawing.Size(10,180) 
$AliasNameTextlableBox.size = new-object System.Drawing.Size(100,20) 
$AliasNameTextlableBox.Text = "Alias"
$form.Controls.Add($AliasNameTextlableBox) 

# Add External Email Address
$exExternalMail = new-object System.Windows.Forms.TextBox
$exExternalMail.Location = new-object System.Drawing.Size(100,210)
$exExternalMail.Size = new-object System.Drawing.Size(300,30)
$exExternalMail.Add_TextChanged({
	$pscmd = "Enable-MailUser -Identity '" + $emIdentityTextBox.Text + "' -Alias '" + $AliasNameText.Text + "' -ExternalEmailAddress 'SMTP:" + $exExternalMail.Text + "'"
	$msCmd.text = $pscmd 
})
$form.Controls.Add($exExternalMail)

# Add External DropLable
$exExternalMaillableBox = new-object System.Windows.Forms.Label
$exExternalMaillableBox.Location = new-object System.Drawing.Size(10,210) 
$exExternalMaillableBox.size = new-object System.Drawing.Size(100,20) 
$exExternalMaillableBox.Text = "External Email"
$form.Controls.Add($exExternalMaillableBox) 

# Add cmdbox 
$msCmd = new-object System.Windows.Forms.RichTextBox
$msCmd.Location = new-object System.Drawing.Size(100,240)
$msCmd.Size = new-object System.Drawing.Size(400,75)
$msCmd.readonly = $true
$form.Controls.Add($msCmd)

# Add CMd DropLable
$msCmdlableBox = new-object System.Windows.Forms.Label
$msCmdlableBox.Location = new-object System.Drawing.Size(10,240) 
$msCmdlableBox.size = new-object System.Drawing.Size(100,20) 
$msCmdlableBox.Text = "PowerShell CMD"
$form.Controls.Add($msCmdlableBox) 


# Add Mail Enabled Button

$crButton = new-object System.Windows.Forms.Button
$crButton.Location = new-object System.Drawing.Size(110,330)
$crButton.Size = new-object System.Drawing.Size(100,23)
$crButton.Text = "Enabled User"
$crButton.Add_Click({Enableuser})
$form.Controls.Add($crButton)

$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
