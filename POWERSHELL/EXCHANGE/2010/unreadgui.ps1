[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
$form = new-object System.Windows.Forms.form 
[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")
$mbcombCollection = @()
$Emailhash = @{ }

function ExportGrid{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.Showhelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("UserName,EmailAddress,Last_Logon,MailboxSize,Inbox_Number_Unread,Inbox_Unread_LastRecieved,Sent_Items_LastSent")
	foreach($row in $msTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString() + "," + $row[5].ToString() + "," + $row[6].ToString()) 
	}
	$logfile.Close()
}

}

function Getinfo(){

$mbcombCollection = @()
$Emailhash.clear()
$msTable.clear()

Get-Mailbox -server $snServerNameDrop.SelectedItem.ToString() -ResultSize Unlimited | foreach-object{
	if ($Emailhash.containskey($_.ExchangeGuid) -eq $false){
		$Emailhash.add($_.ExchangeGuid.ToString(),$_.windowsEmailAddress.ToString())
	}
}

get-mailboxstatistics  -server $snServerNameDrop.SelectedItem.ToString()  | foreach-object{
	$mbcomb = "" | select DisplayName,EmailAddress,Last_Logon,MailboxSize,Inbox_Number_Unread,Inbox_Unread_LastRecieved,Sent_Items_LastSent
	$mbcomb.DisplayName = $_.DisplayName.ToString()
	if ($Emailhash.ContainsKey($_.MailboxGUID.ToString())){
		$mbcomb.EmailAddress = $Emailhash[$_.MailboxGUID.ToString()]
	}
	if ($_.LastLogonTime -ne $null){
		$mbcomb.Last_Logon = $_.LastLogonTime.ToString()
	}
	$mbcomb.MailboxSize = $_.TotalItemSize.Value.ToMB()

	"Mailbox : " + $mbcomb.EmailAddress
	if ($mbcomb.EmailAddress -ne $null){
	$mbMailboxEmail = $mbcomb.EmailAddress
	if ($unCASUrlTextBox.text -eq ""){ $casurl = $null}
	else { $casurl= $unCASUrlTextBox.text}
	$useImp = $false
	if ($seImpersonationCheck.Checked -eq $true) {
		$useImp = $true
	}
	if ($seAuthCheck.Checked -eq $true) {
		$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$useImp, $unUserNameTextBox.Text, $unPasswordTextBox.Text, $unDomainTextBox.Text,$casUrl)
	}
	else{
		$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$useImp, "", "", "",$casUrl)
	}


	$drDuration = new-object EWSUtil.EWS.Duration
	$drDuration.StartTime = [DateTime]::UtcNow.AddDays(-365)
	$drDuration.EndTime = [DateTime]::UtcNow

	$dTypeFld = new-object EWSUtil.EWS.DistinguishedFolderIdType
	$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox
	$dTypeFld2 = new-object EWSUtil.EWS.DistinguishedFolderIdType
	$dTypeFld2.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::sentitems

	$mbMailbox = new-object EWSUtil.EWS.EmailAddressType
	$mbMailbox.EmailAddress = $mbMailboxEmail
	$dTypeFld.Mailbox = $mbMailbox
	$dTypeFld2.Mailbox = $mbMailbox

	$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
	$fldarry[0] = $dTypeFld
	$msgList = $ewc.FindUnread($fldarry, $drDuration, $null, "")
	$mbcomb.Inbox_Number_Unread = $msgList.Count
	if ($msgList.Count -ne 0){
	        $mbcomb.Inbox_Unread_LastRecieved = $msgList[0].DateTimeSent.ToLocalTime().ToString()
	}

	$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
	$fldarry[0] = $dTypeFld2
	$msgList = $ewc.FindItems($fldarry, $drDuration, $null, "")
	if ($msgList.Count -ne 0){
	     $mbcomb.Sent_Items_LastSent = $msgList[0].DateTimeSent.ToLocalTime().ToString()
	}
	$msTable.Rows.add($mbcomb.DisplayName,$mbcomb.EmailAddress,$mbcomb.Last_Logon,$mbcomb.MailboxSize,$mbcomb.Inbox_Number_Unread,$mbcomb.Inbox_Unread_LastRecieved,$mbcomb.Sent_Items_LastSent)
	$mbcombCollection += $mbcomb}
}
$dgDataGrid.DataSource = $msTable
}

$msTable = New-Object System.Data.DataTable
$msTable.TableName = "Mailbox Info"
$msTable.Columns.Add("UserName")
$msTable.Columns.Add("EmailAddress")
$msTable.Columns.Add("Last Logon Time",[DateTime])
$msTable.Columns.Add("Mailbox Size(MB)",[int64])
$msTable.Columns.Add("Inbox Number Unread",[int64])
$msTable.Columns.Add("Inbox unread Last Recieved",[DateTime])
$msTable.Columns.Add("Sent Item Last Sent",[DateTime])



# Add Server DropLable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,20) 
$snServerNamelableBox.size = new-object System.Drawing.Size(80,20) 
$snServerNamelableBox.Text = "ServerName"
$form.Controls.Add($snServerNamelableBox) 

# Add Server Drop Down
$snServerNameDrop = new-object System.Windows.Forms.ComboBox
$snServerNameDrop.Location = new-object System.Drawing.Size(90,20)
$snServerNameDrop.Size = new-object System.Drawing.Size(150,30)
get-mailboxserver | ForEach-Object{$snServerNameDrop.Items.Add($_.Name)}
$form.Controls.Add($snServerNameDrop)

# Add Export Grid Button

$exButton2 = new-object System.Windows.Forms.Button
$exButton2.Location = new-object System.Drawing.Size(250,20)
$exButton2.Size = new-object System.Drawing.Size(125,20)
$exButton2.Text = "Execute"
$exButton2.Add_Click({Getinfo})
$form.Controls.Add($exButton2)


# Add Export Grid Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(10,610)
$exButton1.Size = new-object System.Drawing.Size(125,20)
$exButton1.Text = "Export Grid"
$exButton1.Add_Click({ExportGrid})
$form.Controls.Add($exButton1)

# Add Impersonation Clause

$esImpersonationlableBox = new-object System.Windows.Forms.Label
$esImpersonationlableBox.Location = new-object System.Drawing.Size(10,75) 
$esImpersonationlableBox.Size = new-object System.Drawing.Size(130,20) 
$esImpersonationlableBox.Text = "Use EWS Impersonation"
$form.controls.Add($esImpersonationlableBox) 

$seImpersonationCheck =  new-object System.Windows.Forms.CheckBox
$seImpersonationCheck.Location = new-object System.Drawing.Size(150,70)
$seImpersonationCheck.Size = new-object System.Drawing.Size(30,25)
$form.controls.Add($seImpersonationCheck)

# Add Auth Clause

$esAuthlableBox = new-object System.Windows.Forms.Label
$esAuthlableBox.Location = new-object System.Drawing.Size(10,105) 
$esAuthlableBox.Size = new-object System.Drawing.Size(130,20) 
$esAuthlableBox.Text = "Specify Credentials"
$form.controls.Add($esAuthlableBox) 

$seAuthCheck =  new-object System.Windows.Forms.CheckBox
$seAuthCheck.Location = new-object System.Drawing.Size(140,100)
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
$form.controls.Add($seAuthCheck)

# Add UserName Box
$unUserNameTextBox = new-object System.Windows.Forms.TextBox 
$unUserNameTextBox.Location = new-object System.Drawing.Size(230,100) 
$unUserNameTextBox.size = new-object System.Drawing.Size(100,20) 
$form.controls.Add($unUserNameTextBox) 

# Add UserName Lable
$unUserNamelableBox = new-object System.Windows.Forms.Label
$unUserNamelableBox.Location = new-object System.Drawing.Size(170,105) 
$unUserNamelableBox.size = new-object System.Drawing.Size(60,20) 
$unUserNamelableBox.Text = "UserName"
$unUserNameTextBox.Enabled = $false
$form.controls.Add($unUserNamelableBox) 

# Add Password Box
$unPasswordTextBox = new-object System.Windows.Forms.TextBox 
$unPasswordTextBox.PasswordChar = "*"
$unPasswordTextBox.Location = new-object System.Drawing.Size(400,100) 
$unPasswordTextBox.size = new-object System.Drawing.Size(100,20) 
$form.controls.Add($unPasswordTextBox) 

# Add Password Lable
$unPasswordlableBox = new-object System.Windows.Forms.Label
$unPasswordlableBox.Location = new-object System.Drawing.Size(340,105) 
$unPasswordlableBox.size = new-object System.Drawing.Size(60,20) 
$unPasswordlableBox.Text = "Password"
$unPasswordTextBox.Enabled = $false
$form.controls.Add($unPasswordlableBox) 

# Add Domain Box
$unDomainTextBox = new-object System.Windows.Forms.TextBox 
$unDomainTextBox.Location = new-object System.Drawing.Size(550,100) 
$unDomainTextBox.size = new-object System.Drawing.Size(100,20) 
$form.controls.Add($unDomainTextBox) 

# Add Domain Lable
$unDomainlableBox = new-object System.Windows.Forms.Label
$unDomainlableBox.Location = new-object System.Drawing.Size(510,105) 
$unDomainlableBox.size = new-object System.Drawing.Size(50,20) 
$unDomainlableBox.Text = "Domain"
$unDomainTextBox.Enabled = $false
$form.controls.Add($unDomainlableBox) 


# Add CASUrl Box
$unCASUrlTextBox = new-object System.Windows.Forms.TextBox 
$unCASUrlTextBox.Location = new-object System.Drawing.Size(280,75) 
$unCASUrlTextBox.size = new-object System.Drawing.Size(400,20) 
$unCASUrlTextBox.text = $strRootURI
$form.Controls.Add($unCASUrlTextBox) 

# Add CASUrl Lable
$unCASUrllableBox = new-object System.Windows.Forms.Label
$unCASUrllableBox.Location = new-object System.Drawing.Size(200,75) 
$unCASUrllableBox.size = new-object System.Drawing.Size(50,20) 
$unCASUrllableBox.Text = "CASUrl"
$form.Controls.Add($unCASUrllableBox) 

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,145) 
$dgDataGrid.size = new-object System.Drawing.Size(1000,450)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)



$form.Text = "Exchange 2007 Unused Mailbox Form"
$form.size = new-object System.Drawing.Size(1200,700) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()

