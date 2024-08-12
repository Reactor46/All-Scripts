[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

## end code from http://poshcode.org/624

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
$form = new-object System.Windows.Forms.form 
$service.TraceEnabled = $true
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$Treeinfo = @{ }

function OpenMailbox(){
	
	$tvTreView.Nodes.Clear()
	$Treeinfo.Clear()
	$TNRoot = new-object System.Windows.Forms.TreeNode("Root")
	$TNRoot.Name = "Mailbox"
	$TNRoot.Text = "Mailbox - " + $emEmailAddressTextBox.Text.ToString()
	[void]$tvTreView.Nodes.Add($TNRoot) 
	$rfRootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$emEmailAddressTextBox.Text.ToString())
	if ($seImpersonationCheck.Checked -eq $true){
		$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$emEmailAddressTextBox.Text.ToString())
	}
	if ($seAuthCheck.Checked -eq $false){
		$service.Credentials = New-Object System.Net.NetworkCredential($unUserNameTextBox.Text,$unPasswordTextBox.Text,$unDomainTextBox.Text)
	}
	else{
		$service.UseDefaultCredentials = $true
	}
	if ($adAutoDiscoCheck.Checked -eq $true){
		$service.AutodiscoverUrl($emEmailAddressTextBox.Text.ToString())
	}
	else{
		$uri=[system.URI]$unCASUrlTextBox.text
		$service.Url = $uri
	}
	$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$rfRootFolderID);
	$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000);
	$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
	$ffResponse = $rfRootFolder.FindFolders($fvFolderView);
	foreach ($ffFolder in $ffResponse.Folders)            {
		$TNChild = new-object System.Windows.Forms.TreeNode($ffFolder.DisplayName.ToString())
		$TNChild.Name = $ffFolder.DisplayName.ToString()
		$TNChild.Text = $ffFolder.DisplayName.ToString()
		$TNChild.tag = $ffFolder.Id.UniqueId.ToString()
		if ($ffFolder.ParentFolderId.UniqueId -eq $rfRootFolder.Id.UniqueId ){
			$ffFolder.DisplayName
			[void]$TNRoot.Nodes.Add($TNChild) 
			$Treeinfo.Add($ffFolder.Id.UniqueId.ToString(),$TNChild)
		}
		else{
			$pfFolder = $Treeinfo[$ffFolder.ParentFolderId.UniqueId.ToString()]
			[void]$pfFolder.Nodes.Add($TNChild)
			if ($Treeinfo.ContainsKey($ffFolder.Id.UniqueId.ToString()) -eq $false){
				$Treeinfo.Add($ffFolder.Id.UniqueId.ToString(),$TNChild)
			}
		}
	}

}

function GetFolderItems(){
	$mbtable.Clear()
	$folderID = $Global:lfFolderID 
	write-host $folderID
	$cfCurrentFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId($folderID)
	$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($neResultCheckNum.Value)
	if ($seSearchCheck.Checked -eq $true){
		switch($snSearchPropDrop.SelectedItem.ToString()){
				"Subject" {$sfilter = [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject}
				"Body" {$sfilter = [Microsoft.Exchange.WebServices.Data.ItemSchema]::Body}
				"From" {$sfilter = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender}
			}
		$SearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring($sfilter,$sbSearchTextBox.Text.ToString())
		 
	}
	$findResults = $service.FindItems($cfCurrentFolderID,$SearchFilter,$view)
	foreach($mail in  $findResults.Items){
		if ($mail.From.Name -ne $null){$fnFromName = $mail.From.Name.ToString()}
		else{$fnFromName = "N/A"}
		if ($mail.Subject -ne $null){$sbSubject = $mail.Subject.ToString()}
		else{$sbSubject = "N/A"}
		$mbtable.rows.add($fnFromName,$sbSubject,$mail.DateTimeSent.ToString(),$mail.Size,$mail.id.UniqueID.ToString())
	}
	$dgDataGrid.DataSource = $mbtable
}

function newMessage($reply){
	$global:newmsgform = new-object System.Windows.Forms.form 
	$global:newmsgform.Text = $global:msMessage.Subject
	$global:newmsgform.size = new-object System.Drawing.Size(1000,800) 

	# Add Message To Lable
	$miMessageTolableBox = new-object System.Windows.Forms.Label
	$miMessageTolableBox.Location = new-object System.Drawing.Size(20,20) 
	$miMessageTolableBox.size = new-object System.Drawing.Size(80,20) 
	$miMessageTolableBox.Text = "To"
	$global:newmsgform.controls.Add($miMessageTolableBox) 

	# Add Message Subject Lable
	$miMessageSubjectlableBox = new-object System.Windows.Forms.Label
	$miMessageSubjectlableBox.Location = new-object System.Drawing.Size(20,65) 
	$miMessageSubjectlableBox.size = new-object System.Drawing.Size(80,20) 
	$miMessageSubjectlableBox.Text = "Subject"
	$global:newmsgform.controls.Add($miMessageSubjectlableBox) 

	# Add Message To
	$miMessageTotextlabelBox = new-object System.Windows.Forms.TextBox
	$miMessageTotextlabelBox.Location = new-object System.Drawing.Size(100,20) 
	$miMessageTotextlabelBox.size = new-object System.Drawing.Size(400,20) 
	$global:newmsgform.controls.Add($miMessageTotextlabelBox) 

	# Add Message Subject 
	$miMessageSubjecttextlabelBox = new-object System.Windows.Forms.TextBox
	$miMessageSubjecttextlabelBox.Location = new-object System.Drawing.Size(100,65) 
	$miMessageSubjecttextlabelBox.size = new-object System.Drawing.Size(600,20) 
	$global:newmsgform.controls.Add($miMessageSubjecttextlabelBox) 

	# Add Message body 
	$miMessageBodytextlabelBox = new-object System.Windows.Forms.RichTextBox
	$miMessageBodytextlabelBox.Location = new-object System.Drawing.Size(100,100) 
	$miMessageBodytextlabelBox.size = new-object System.Drawing.Size(600,350) 
	$global:newmsgform.controls.Add($miMessageBodytextlabelBox) 


	$exButton7 = new-object System.Windows.Forms.Button
	$exButton7.Location = new-object System.Drawing.Size(95,460)
	$exButton7.Size = new-object System.Drawing.Size(125,20)
	$exButton7.Text = "Send Message"
	$exButton7.Add_Click({SendMessage})
	$global:newmsgform.Controls.Add($exButton7)


	$global:newmsgform.autoscroll = $true
	$global:newmsgform.Add_Shown({$form.Activate()})
	$global:newmsgform.ShowDialog()

}

function SendMessage(){
	$nmNewMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
	$nmNewMessage.Subject = $miMessageSubjecttextlabelBox.Text
	$nmNewMessage.Body = $miMessageBodytextlabelBox.Text
	$nmNewMessage.ToRecipients.Add($miMessageTotextlabelBox.Text)
	$nmNewMessage.Send.Invoke()
	$global:newmsgform.close()

}

function ExportMessage{
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
	$MessageID = $mbtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][4]
	$global:msMessage = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service,$MessageID,$psPropset);
	write-host $MessageID
	$exFileName = new-object System.Windows.Forms.saveFileDialog
	$exFileName.DefaultExt = "eml"
	$exFileName.Filter = "emlFiles files (*.eml)|*.eml"
	$exFileName.InitialDirectory = "c:\temp"
	$exFileName.ShowHelp = $true
	$exFileName.ShowDialog()
	if ($exFileName.FileName -ne ""){
		$fiFile = new-object System.IO.FileStream($exFileName.FileName, [System.IO.FileMode]::Create)
                $fiFile.Write($global:msMessage.MimeContent.Content, 0, $global:msMessage.MimeContent.Content.Length)
                $fiFile.Close()

	}

}

function showMessage($MessageID){
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties);
        $psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

	$MessageID = $mbtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][4]
	$global:msMessage = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service,$MessageID,$psPropset);
	write-host $MessageID
	$msgform = new-object System.Windows.Forms.form 
	$msgform.Text = $global:msMessage.Subject
	$msgform.size = new-object System.Drawing.Size(1000,800) 
	

	# Add Message From Lable
	$miMessageTolableBox = new-object System.Windows.Forms.Label
	$miMessageTolableBox.Location = new-object System.Drawing.Size(20,20) 
	$miMessageTolableBox.size = new-object System.Drawing.Size(80,20) 
	$miMessageTolableBox.Text = "To"
	$msgform.controls.Add($miMessageTolableBox) 

	# Add MessageID Lable
	$miMessageSentlableBox = new-object System.Windows.Forms.Label
	$miMessageSentlableBox.Location = new-object System.Drawing.Size(20,40) 
	$miMessageSentlableBox.size = new-object System.Drawing.Size(80,20) 
	$miMessageSentlableBox.Text = "From"
	$msgform.controls.Add($miMessageSentlableBox) 

	# Add Message Subject Lable
	$miMessageSubjectlableBox = new-object System.Windows.Forms.Label
	$miMessageSubjectlableBox.Location = new-object System.Drawing.Size(20,60) 
	$miMessageSubjectlableBox.size = new-object System.Drawing.Size(80,20) 
	$miMessageSubjectlableBox.Text = "Subject"
	$msgform.controls.Add($miMessageSubjectlableBox) 

	# Add Message To
	$miMessageTotextlabelBox = new-object System.Windows.Forms.Label
	$miMessageTotextlabelBox.Location = new-object System.Drawing.Size(100,20) 
	$miMessageTotextlabelBox.size = new-object System.Drawing.Size(400,20) 
	$msgform.controls.Add($miMessageTotextlabelBox) 
	$miMessageTotextlabelBox.Text = $global:msMessage.DisplayTo.ToString()

	# Add Message From
	$miMessageSenttextlabelBox = new-object System.Windows.Forms.Label
	$miMessageSenttextlabelBox.Location = new-object System.Drawing.Size(100,40) 
	$miMessageSenttextlabelBox.size = new-object System.Drawing.Size(600,20) 
	$msgform.controls.Add($miMessageSenttextlabelBox) 
	$miMessageSenttextlabelBox.Text = $global:msMessage.From.Name.ToString() + " (" + $global:msMessage.From.Address.ToString() + ")" 

	# Add Message Subject 
	$miMessageSubjecttextlabelBox = new-object System.Windows.Forms.Label
	$miMessageSubjecttextlabelBox.Location = new-object System.Drawing.Size(100,60) 
	$miMessageSubjecttextlabelBox.size = new-object System.Drawing.Size(600,20) 
	$msgform.controls.Add($miMessageSubjecttextlabelBox) 
	$miMessageSubjecttextlabelBox.Text  = $global:msMessage.Subject.ToString()

	# Add Message body 
	$miMessageBodytextlabelBox = new-object System.Windows.Forms.RichTextBox
	$miMessageBodytextlabelBox.Location = new-object System.Drawing.Size(100,80) 
	$miMessageBodytextlabelBox.size = new-object System.Drawing.Size(600,350) 
	$miMessageBodytextlabelBox.text = $global:msMessage.Body.Text.ToString()
	$msgform.controls.Add($miMessageBodytextlabelBox) 

	# Add Message Attachments Lable
	$miMessageAttachmentslableBox = new-object System.Windows.Forms.Label
	$miMessageAttachmentslableBox.Location = new-object System.Drawing.Size(20,445) 
	$miMessageAttachmentslableBox.size = new-object System.Drawing.Size(80,20) 
	$miMessageAttachmentslableBox.Text = "Attachments"
	$msgform.controls.Add($miMessageAttachmentslableBox) 

	$miMessageAttachmentslableBox1 = new-object System.Windows.Forms.Label
	$miMessageAttachmentslableBox1.Location = new-object System.Drawing.Size(100,445) 
	$miMessageAttachmentslableBox1.size = new-object System.Drawing.Size(600,20) 
	$miMessageAttachmentslableBox1.Text = ""
	$msgform.Controls.Add($miMessageAttachmentslableBox1) 

	
	$exButton4 = new-object System.Windows.Forms.Button
	$exButton4.Location = new-object System.Drawing.Size(10,465)
	$exButton4.Size = new-object System.Drawing.Size(150,20)
	$exButton4.Text = "Download Attachments"
	$exButton4.Enabled = $false
	$exButton4.Add_Click({DownloadAttachments})
	$msgform.Controls.Add($exButton4)
	
	$attname = ""
	if ($global:msMessage.hasattachments){
		write-host "Attachment"
		$exButton4.Enabled = $true
		foreach($attach in $global:msMessage.Attachments)
		{	
		
			$attname = $attname + $attach.Name.ToString() + "; "
		}
	}
	$miMessageAttachmentslableBox1.Text = $attname
	# Add Download Button

	$msgform.autoscroll = $true
	$msgform.Add_Shown({$form.Activate()})
	$msgform.ShowDialog()

}

function downloadattachments{
	$dlfolder = new-object -com shell.application 
	$dlfolderpath = $dlfolder.BrowseForFolder(0,"Download attachments to",0) 
	
	write-host
	foreach($attach in $global:msMessage.Attachments){
		$attach.Load()
		$fiFile = new-object System.IO.FileStream(($dlfolderpath.Self.Path + "\" + $attach.Name.ToString()), [System.IO.FileMode]::Create)
                $fiFile.Write($attach.Content, 0, $attach.Content.Length)
                $fiFile.Close()
		write-host  "Downloaded Attachment : " +  ($dlfolderpath.Self.Path + "\" + $attach.Name.ToString())
	}

}

function ShowHeader{
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties);
        $psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

	$MessageID = $mbtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][4]
	$global:msMessage = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service,$MessageID,$psPropset);
	write-host $MessageID
	$hdrform = new-object System.Windows.Forms.form 
	$hdrform.Text = $global:msMessage.Subject
	$hdrform.size = new-object System.Drawing.Size(800,600) 
	foreach ($ihead in $global:msMessage.InternetMessageHeaders){
        	$headertext =  $headertext + $ihead.Name.ToString() + ":" +  $ihead.Value.ToString() + "`r`n"
                
        }
	# Add Message header
	$miMessageHeadertextlabelBox = new-object System.Windows.Forms.RichTextBox
	$miMessageHeadertextlabelBox.Location = new-object System.Drawing.Size(10,10) 
	$miMessageHeadertextlabelBox.size = new-object System.Drawing.Size(800,600) 
	$miMessageHeadertextlabelBox.text = $headertext
	$hdrform.controls.Add($miMessageHeadertextlabelBox) 
	$hdrform.autoscroll = $true
	$hdrform.Add_Shown({$form.Activate()})
	$hdrform.ShowDialog()


}
$mbtable = New-Object System.Data.DataTable
$mbtable.TableName = "Folder Item"
$mbtable.Columns.Add("From")
$mbtable.Columns.Add("Subject")
$mbtable.Columns.Add("Recieved",[DATETIME])
$mbtable.Columns.Add("Size",[INT64])
$mbtable.Columns.Add("ID")

# Add Email Address 
$emEmailAddressTextBox = new-object System.Windows.Forms.TextBox 
$emEmailAddressTextBox.Location = new-object System.Drawing.Size(130,20) 
$emEmailAddressTextBox.size = new-object System.Drawing.Size(300,20) 
$emEmailAddressTextBox.Enabled = $true
$emEmailAddressTextBox.text = $aceuser.mail.ToString()
$form.controls.Add($emEmailAddressTextBox) 

# Add  Email Address  Lable
$emEmailAddresslableBox = new-object System.Windows.Forms.Label
$emEmailAddresslableBox.Location = new-object System.Drawing.Size(10,20) 
$emEmailAddresslableBox.size = new-object System.Drawing.Size(120,20) 
$emEmailAddresslableBox.Text = "Email Address"
$form.controls.Add($emEmailAddresslableBox) 

# AutoDisco Check

$adAutoDiscolableBox = new-object System.Windows.Forms.Label
$adAutoDiscolableBox.Location = new-object System.Drawing.Size(10,55) 
$adAutoDiscolableBox.Size = new-object System.Drawing.Size(170,20) 
$adAutoDiscolableBox.Text = "Use AutoDiscover"
$form.controls.Add($adAutoDiscolableBox) 

$adAutoDiscoCheck =  new-object System.Windows.Forms.CheckBox
$adAutoDiscoCheck.Location = new-object System.Drawing.Size(180,55)
$adAutoDiscoCheck.Size = new-object System.Drawing.Size(30,25)
$adAutoDiscoCheck.Checked = $true
$adAutoDiscoCheck.Add_Click({if ($adAutoDiscoCheck.Checked -eq $false){$unCASUrlTextBox.Enabled = $true}
	else{$unCASUrlTextBox.Enabled = $false}
})
$form.controls.Add($adAutoDiscoCheck)



# Add CASUrl Box
$unCASUrlTextBox = new-object System.Windows.Forms.TextBox 
$unCASUrlTextBox.Location = new-object System.Drawing.Size(280,55) 
$unCASUrlTextBox.size = new-object System.Drawing.Size(400,20) 
$unCASUrlTextBox.text = $strRootURI
$unCASUrlTextBox.Enabled = $false
$form.Controls.Add($unCASUrlTextBox) 

# Add CASUrl Lable
$unCASUrllableBox = new-object System.Windows.Forms.Label
$unCASUrllableBox.Location = new-object System.Drawing.Size(220,55) 
$unCASUrllableBox.size = new-object System.Drawing.Size(50,20) 
$unCASUrllableBox.Text = "CASUrl"
$form.Controls.Add($unCASUrllableBox) 

# Add Impersonation Clause

$esImpersonationlableBox = new-object System.Windows.Forms.Label
$esImpersonationlableBox.Location = new-object System.Drawing.Size(10,80) 
$esImpersonationlableBox.Size = new-object System.Drawing.Size(170,20) 
$esImpersonationlableBox.Text = "Use EWS Impersonation"
$form.controls.Add($esImpersonationlableBox) 

$seImpersonationCheck =  new-object System.Windows.Forms.CheckBox
$seImpersonationCheck.Location = new-object System.Drawing.Size(180,80)
$seImpersonationCheck.Size = new-object System.Drawing.Size(30,25)
$form.controls.Add($seImpersonationCheck)

# Add Auth Clause

$dfDefautcredlableBox = new-object System.Windows.Forms.Label
$dfDefautcredlableBox.Location = new-object System.Drawing.Size(10,80) 
$dfDefautcredlableBox.Size = new-object System.Drawing.Size(170,20) 
$dfDefautcredlableBox.Text = "Use Default Credentials"
$form.controls.Add($dfDefautcredlableBox) 


# Add Auth Clause

$dfDefautcredlableBox = new-object System.Windows.Forms.Label
$dfDefautcredlableBox.Location = new-object System.Drawing.Size(10,105) 
$dfDefautcredlableBox.Size = new-object System.Drawing.Size(170,20) 
$dfDefautcredlableBox.Text = "Use Default Credentials"
$form.controls.Add($dfDefautcredlableBox) 


# Add UserName Lable
$unUserNamelableBox = new-object System.Windows.Forms.Label
$unUserNamelableBox.Location = new-object System.Drawing.Size(220,105) 
$unUserNamelableBox.size = new-object System.Drawing.Size(70,20) 
$unUserNamelableBox.Text = "UserName"
$form.controls.Add($unUserNamelableBox) 

# Add UserName Box
$unUserNameTextBox = new-object System.Windows.Forms.TextBox 
$unUserNameTextBox.Location = new-object System.Drawing.Size(300,105) 
$unUserNameTextBox.size = new-object System.Drawing.Size(140,20) 
$unUserNameTextBox.Enabled = $false
$form.controls.Add($unUserNameTextBox) 

# Add Password Box
$unPasswordTextBox = new-object System.Windows.Forms.TextBox 
$unPasswordTextBox.PasswordChar = "*"
$unPasswordTextBox.Location = new-object System.Drawing.Size(520,105) 
$unPasswordTextBox.size = new-object System.Drawing.Size(140,20) 
$form.controls.Add($unPasswordTextBox) 

# Add Password Lable
$unPasswordlableBox = new-object System.Windows.Forms.Label
$unPasswordlableBox.Location = new-object System.Drawing.Size(450,105) 
$unPasswordlableBox.size = new-object System.Drawing.Size(80,20) 
$unPasswordlableBox.Text = "Password"
$unPasswordTextBox.Enabled = $false
$form.controls.Add($unPasswordlableBox) 

# Add Domain Box
$unDomainTextBox = new-object System.Windows.Forms.TextBox 
$unDomainTextBox.Location = new-object System.Drawing.Size(720,105) 
$unDomainTextBox.size = new-object System.Drawing.Size(100,20) 
$form.controls.Add($unDomainTextBox) 

# Add Domain Lable
$unDomainlableBox = new-object System.Windows.Forms.Label
$unDomainlableBox.Location = new-object System.Drawing.Size(670,105) 
$unDomainlableBox.size = new-object System.Drawing.Size(60,20) 
$unDomainlableBox.Text = "Domain"
$unDomainTextBox.Enabled = $false
$form.controls.Add($unDomainlableBox) 

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(10,130)
$exButton1.Size = new-object System.Drawing.Size(125,20)
$exButton1.Text = "Open Mailbox"
$exButton1.Add_Click({OpenMailbox})
$form.Controls.Add($exButton1)

# Add Numeric Results

$neResultCheckNum =  new-object System.Windows.Forms.numericUpDown
$neResultCheckNum.Location = new-object System.Drawing.Size(250,130)
$neResultCheckNum.Size = new-object System.Drawing.Size(70,30)
$neResultCheckNum.Enabled = $true
$neResultCheckNum.Value = 100
$neResultCheckNum.Maximum = 10000000000
$form.Controls.Add($neResultCheckNum)

$exButton2 = new-object System.Windows.Forms.Button
$exButton2.Location = new-object System.Drawing.Size(330,130)
$exButton2.Size = new-object System.Drawing.Size(125,25)
$exButton2.Text = "Show Message"
$exButton2.Add_Click({ShowMessage})
$form.Controls.Add($exButton2)

$exButton5 = new-object System.Windows.Forms.Button
$exButton5.Location = new-object System.Drawing.Size(455,130)
$exButton5.Size = new-object System.Drawing.Size(125,25)
$exButton5.Text = "Show Header"
$exButton5.Add_Click({ShowHeader})
$form.Controls.Add($exButton5)

$exButton6 = new-object System.Windows.Forms.Button
$exButton6.Location = new-object System.Drawing.Size(330,155)
$exButton6.Size = new-object System.Drawing.Size(125,25)
$exButton6.Text = "New Message"
$exButton6.Add_Click({NewMessage})
$form.Controls.Add($exButton6)

$exButton7 = new-object System.Windows.Forms.Button
$exButton7.Location = new-object System.Drawing.Size(960,165)
$exButton7.Size = new-object System.Drawing.Size(90,25)
$exButton7.Text = "Update"
$exButton7.Add_Click({GetFolderItems})
$form.Controls.Add($exButton7)

$exButton8 = new-object System.Windows.Forms.Button
$exButton8.Location = new-object System.Drawing.Size(455,155)
$exButton8.Size = new-object System.Drawing.Size(125,25)
$exButton8.Text = "Export Message"
$exButton8.Add_Click({ExportMessage})
$form.Controls.Add($exButton8)

# Add Search Lable

$saSeachBoxLable = new-object System.Windows.Forms.Label
$saSeachBoxLable.Location = new-object System.Drawing.Size(600,135) 
$saSeachBoxLable.Size = new-object System.Drawing.Size(170,20) 
$saSeachBoxLable.Text = "Search by Property"
$form.controls.Add($saSeachBoxLable) 

$saNumItemsBoxLable = new-object System.Windows.Forms.Label
$saNumItemsBoxLable.Location = new-object System.Drawing.Size(160,135) 
$saNumItemsBoxLable.Size = new-object System.Drawing.Size(170,20) 
$saNumItemsBoxLable.Text = "Number of Items"
$form.controls.Add($saNumItemsBoxLable) 

$seSearchCheck =  new-object System.Windows.Forms.CheckBox
$seSearchCheck.Location = new-object System.Drawing.Size(585,130)
$seSearchCheck.Size = new-object System.Drawing.Size(30,25)
$seSearchCheck.Add_Click({if ($seSearchCheck.Checked -eq $false){
	$sbSearchTextBox.Enabled = $false
	$snSearchPropDrop.Enabled = $false
	}
	else{
		$sbSearchTextBox.Enabled = $true
		$snSearchPropDrop.Enabled = $true
	}
})
$form.controls.Add($seSearchCheck)

#Add Search box
$snSearchPropDrop = new-object System.Windows.Forms.ComboBox
$snSearchPropDrop.Location = new-object System.Drawing.Size(585,165)
$snSearchPropDrop.Size = new-object System.Drawing.Size(150,30)
$snSearchPropDrop.Items.Add("Subject")
$snSearchPropDrop.Items.Add("Body")
$snSearchPropDrop.Items.Add("From")
$snSearchPropDrop.Enabled = $false
$form.Controls.Add($snSearchPropDrop)

# Add Search TextBox
$sbSearchTextBox = new-object System.Windows.Forms.TextBox 
$sbSearchTextBox.Location = new-object System.Drawing.Size(750,165) 
$sbSearchTextBox.size = new-object System.Drawing.Size(200,20) 
$sbSearchTextBox.Enabled = $false
$form.controls.Add($sbSearchTextBox) 

$seAuthCheck =  new-object System.Windows.Forms.CheckBox
$seAuthCheck.Location = new-object System.Drawing.Size(180,105)
$seAuthCheck.Checked = $true
$seAuthCheck.Size = new-object System.Drawing.Size(30,25)
$seAuthCheck.Add_Click({if ($seAuthCheck.Checked -eq $false){
			$unUserNameTextBox.Enabled = $true
			$unPasswordTextBox.Enabled = $true
			$unDomainTextBox.Enabled = $true
			}
			else{
				$unUserNameTextBox.Enabled = $false
				$unPasswordTextBox.Enabled = $false
				$unDomainTextBox.Enabled = $false}})
$form.controls.Add($seAuthCheck)

$tvTreView = new-object System.Windows.Forms.TreeView
$tvTreView.Location = new-object System.Drawing.Size(10,155)  
$tvTreView.size = new-object System.Drawing.Size(216,400) 
$tvTreView.Anchor = "Top,left,Bottom"
$tvTreView.add_AfterSelect({
	$Global:lfFolderID = $this.SelectedNode.tag
	GetFolderItems
	
})
$form.Controls.Add($tvTreView)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(250,200) 
$dgDataGrid.size = new-object System.Drawing.Size(800,600)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$form.Controls.Add($dgDataGrid)

$form.Text = "Simple Exchange Mailbox Client"
$form.size = new-object System.Drawing.Size(1200,700) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
