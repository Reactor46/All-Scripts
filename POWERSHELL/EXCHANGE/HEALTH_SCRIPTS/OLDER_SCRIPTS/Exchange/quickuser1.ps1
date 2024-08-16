[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

Function createmailbox{
$psSecurePasswordString = new-object System.Security.SecureString
foreach($char in $pwPassWordTextBox.Text.ToCharArray())
   {
      $psSecurePasswordString.AppendChar($char)
   }
$result = New-mailbox -UserPrincipalName $unUserNameTextBox.text  -alias $emAliasTextBox.text -database $MBhash1[$msMailStoreDrop.SelectedItem.ToString()] `
-Name $dsDisplayNameTextBox.text  -OrganizationalUnit $OUhash1[$ouOuNameDrop.SelectedItem.ToString()] -password $psSecurePasswordString `
-FirstName $unFirstNameTextBox.text -LastName $lnLastNameTextBox.text -DisplayName $dsDisplayNameTextBox.text
$msgbox = new-object -comobject wscript.shell
if ($result -ne $null){write-host "Mailbox Created Sucessfully"}
else{write-host "Error Creating Mailbox check command Line for Details"}

}


$OUhash1 = @{ }
$MBhash1 = @{ }

$form = new-object System.Windows.Forms.form 
$form.Text = "Exchange 2007 Quick User Create Form"
$form.size = new-object System.Drawing.Size(430,400) 

# Add UserName Box
$unUserNameTextBox = new-object System.Windows.Forms.TextBox 
$unUserNameTextBox.Location = new-object System.Drawing.Size(100,30) 
$unUserNameTextBox.size = new-object System.Drawing.Size(130,20) 
$form.Controls.Add($unUserNameTextBox) 

# Add Username Lable
$unUserNamelableBox = new-object System.Windows.Forms.Label
$unUserNamelableBox.Location = new-object System.Drawing.Size(10,30) 
$unUserNamelableBox.size = new-object System.Drawing.Size(100,20) 
$unUserNamelableBox.Text = "Username UPN"
$form.Controls.Add($unUserNamelableBox) 

# Add Alias Box
$emAliasTextBox = new-object System.Windows.Forms.TextBox 
$emAliasTextBox.Location = new-object System.Drawing.Size(100,60) 
$emAliasTextBox.size = new-object System.Drawing.Size(130,20) 
$form.Controls.Add($emAliasTextBox) 

# Add Alias Lable
$emAliaslableBox = new-object System.Windows.Forms.Label
$emAliaslableBox.Location = new-object System.Drawing.Size(10,60) 
$emAliaslableBox.size = new-object System.Drawing.Size(100,20) 
$emAliaslableBox.Text = "Alias"
$form.Controls.Add($emAliaslableBox) 

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

# Add OU Drop Down
$ouOuNameDrop = new-object System.Windows.Forms.ComboBox
$ouOuNameDrop.Location = new-object System.Drawing.Size(100,180)
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
$form.Controls.Add($ouOuNameDrop)

# Add OU DropLable
$ouOuNamelableBox = new-object System.Windows.Forms.Label
$ouOuNamelableBox.Location = new-object System.Drawing.Size(10,180) 
$ouOuNamelableBox.size = new-object System.Drawing.Size(100,20) 
$ouOuNamelableBox.Text = "OU Name"
$form.Controls.Add($ouOuNamelableBox) 

# Add Server Drop Down
$snServerNameDrop = new-object System.Windows.Forms.ComboBox
$snServerNameDrop.Location = new-object System.Drawing.Size(100,210)
$snServerNameDrop.Size = new-object System.Drawing.Size(130,30)
get-mailboxserver | ForEach-Object{$snServerNameDrop.Items.Add($_.Name)}
$snServerNameDrop.Add_SelectedValueChanged({
	$msMailStoreDrop.Items.Clear()
	get-mailboxdatabase -Server $snServerNameDrop.SelectedItem.ToString()| ForEach-Object{$msMailStoreDrop.Items.Add($_.Name)
	$MBhash1.add($_.Name,$_.ServerName + "\" + $_.StorageGroup.Name + "\" + $_.Name) 	
	}
})  
$form.Controls.Add($snServerNameDrop)

# Add Server DropLable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,210) 
$snServerNamelableBox.size = new-object System.Drawing.Size(100,20) 
$snServerNamelableBox.Text = "ServerName"
$form.Controls.Add($snServerNamelableBox) 

# Add MailStore Drop Down
$msMailStoreDrop = new-object System.Windows.Forms.ComboBox
$msMailStoreDrop.Location = new-object System.Drawing.Size(100,240)
$msMailStoreDrop.Size = new-object System.Drawing.Size(130,30)
$form.Controls.Add($msMailStoreDrop)

# Add MailStore DropLable
$msMailStorelableBox = new-object System.Windows.Forms.Label
$msMailStorelableBox.Location = new-object System.Drawing.Size(10,240) 
$msMailStorelableBox.size = new-object System.Drawing.Size(100,20) 
$msMailStorelableBox.Text = "Mail-Store"
$form.Controls.Add($msMailStorelableBox) 

# Add Password Box
$pwPassWordTextBox = new-object System.Windows.Forms.TextBox 
$pwPassWordTextBox.Location = new-object System.Drawing.Size(100,270) 
$pwPassWordTextBox.size = new-object System.Drawing.Size(130,20) 
$pwPasswordTextBox.UseSystemPasswordChar = $true
$form.Controls.Add($pwPassWordTextBox) 

# Add Password Lable
$pwPassWordlableBox = new-object System.Windows.Forms.Label
$pwPassWordlableBox.Location = new-object System.Drawing.Size(10,270) 
$pwPassWordlableBox.size = new-object System.Drawing.Size(60,20) 
$pwPassWordlableBox.Text = "Password"
$form.Controls.Add($pwPassWordlableBox) 

# Add Create Button

$crButton = new-object System.Windows.Forms.Button
$crButton.Location = new-object System.Drawing.Size(110,310)
$crButton.Size = new-object System.Drawing.Size(100,23)
$crButton.Text = "Create Mailbox"
$crButton.Add_Click({CreateMailbox})
$form.Controls.Add($crButton)

$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
