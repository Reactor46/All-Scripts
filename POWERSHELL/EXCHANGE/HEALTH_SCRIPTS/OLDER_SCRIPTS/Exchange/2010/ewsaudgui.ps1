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

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$global:service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$form = new-object System.Windows.Forms.form 
$service.TraceEnabled = $true
$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$mbtable = New-Object System.Data.DataTable
$mbtable.TableName = "Folder Item"
$mbtable.Columns.Add("From")
$mbtable.Columns.Add("Subject")
$mbtable.Columns.Add("Recieved",[DATETIME])
$mbtable.Columns.Add("ItemClass")
$mbtable.Columns.Add("Size",[INT64])
$mbtable.Columns.Add("ID")

$prmtable = New-Object System.Data.DataTable

$prmtable.TableName = "PRM"
$prmtable.Columns.Add("Name")
$prmtable.Columns.Add("Value")

$mprmtable = New-Object System.Data.DataTable
$mprmtable.TableName = "PRM2"
$mprmtable.Columns.Add("Name")
$mprmtable.Columns.Add("OldValue")
$mprmtable.Columns.Add("NewValue")

$msTable = New-Object System.Data.DataTable

$msTable.TableName = "Event"
$msTable.Columns.Add("RunDate")
$msTable.Columns.Add("Caller")
$msTable.Columns.Add("Cmdlet")
$msTable.Columns.Add("ObjectModified")
$msTable.Columns.Add("Succeeded")
$msTable.Columns.Add("Error")
$msTable.Columns.Add("OriginatingServer")
$msTable.Columns.Add("EntryNumb")

$gbTable = New-Object System.Data.DataTable

$gbTable.TableName = "GroupByUser"
$gbTable.Columns.Add("UserName")
$gbTable.Columns.Add("#Entries",[INT64])

$gbTable1 = New-Object System.Data.DataTable

$gbTable1.TableName = "GroupByUser1"
$gbTable1.Columns.Add("Cmd")
$gbTable1.Columns.Add("#Entries",[INT64])

$userhash = @{ }
$cmdlethash = @{ }
$enthash = @{ }
function showreport(){
	$userhash.clear()
	$cmdlethash.clear()
	$enthash.clear()
	$entnum = 0
	$msgid = $dgDataGrid.SelectedRows[0].Cells[5].Value
	$Item = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service,$msgid)
	$Item.Load()
	[XML]$sResult = $null
   	foreach($attach in $Item.Attachments){
		$fname = $env:temp + "\" + [GUID]::newguid() + ".tmp"
		Write-Host $fname
		$attach.Load()
		$fiFile = new-object System.IO.FileStream($fname.ToString(), [System.IO.FileMode]::Create)
        $fiFile.Write($attach.Content, 0, $attach.Content.Length)
        $fiFile.Close()	
		[XML]$sResult = get-content $fname.ToString()
	}  
	$sResult.SearchResults.Event | foreach-object{
	$mbcomb = "" | select Caller,Cmdlet,ObjectModified,RunDate,Succeeded,Error,OriginatingServer,CmdletParameters,ModifiedProperties,EntryNumb
	$mbcomb.Caller = $_.Caller
	$mbcomb.Cmdlet = $_.Cmdlet
	$mbcomb.ObjectModified = $_.ObjectModified
	$mbcomb.RunDate = $_.RunDate
	$mbcomb.Succeeded = $_.Succeeded
	$mbcomb.Error = $_.Error
	$mbcomb.OriginatingServer = $_.OriginatingServer
	$mbcomb.CmdletParameters = $_.CmdletParameters.Parameter 
	$mbcomb.ModifiedProperties = $_.ModifiedProperties.Property
	$mbcomb.EntryNumb = $entnum   
	$entnum++
	$enthash.add($entnum.ToString(),$mbcomb)
	if($userhash.containsKey($mbcomb.Caller)){
		$userhash[$mbcomb.Caller] += $mbcomb
	}
	else{
		$mbcombCollection = @()
		$mbcombCollection += $mbcomb 
		$userhash.add($mbcomb.Caller,$mbcombCollection)
	}
	if($cmdlethash.containsKey($mbcomb.Cmdlet)){
		$cmdlethash[$mbcomb.Cmdlet] += $mbcomb
	}
	else{
		$mbcombCollection = @()
		$mbcombCollection += $mbcomb 
		$cmdlethash.add($mbcomb.Cmdlet,$mbcombCollection)
	}
	}


	
	ReGroup
	$repform.Add_Shown({$repform.Activate()})
	$repform.ShowDialog()
	

}
function ReGroup(){
	$gbTable.clear()
	if($gbNameDrop.SelectedItem.ToString() -eq "Cmdlet"){
		$cmdlethash.GetEnumerator() | foreach-object {
			$gbTable1.rows.add($_.Key,$_.Value.Count)
		}
		$dgDataGrid1.datasource = $gbTable1
	}
	else{
		$userhash.GetEnumerator() | foreach-object {
			$gbTable.rows.add($_.Key,$_.Value.Count)
		}
		$dgDataGrid1.datasource = $gbTable
	}
}
function GetAuditEmails(){
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
		$service.AutodiscoverUrl($emEmailAddressTextBox.Text.ToString(),{return $true})
	}
	else{
		$uri=[system.URI]$unCASUrlTextBox.text
		$service.Url = $uri
	}

	$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$emEmailAddressTextBox.text)
	$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
	$sfSearchFilter = "System.Message.AttachmentContents:SearchResult.xml"
	$fiItems = $null
	do{
		$fiItems = $service.findItems($folderid,$sfSearchFilter,$ivItemView)
		foreach($mail in $fiItems.Items){
			if ($mail.From.Name -ne $null){$fnFromName = $mail.From.Name.ToString()}
			else{$fnFromName = "N/A"}
			if ($mail.Subject -ne $null){$sbSubject = $mail.Subject.ToString()}
			else{$sbSubject = "N/A"}
			if ($mail.DateTimeSent -ne $null){$recTime = $mail.DateTimeSent.ToString()}
			else{$recTime = $mail.DateTimeCreated.ToString()}
			$mbtable.rows.add($fnFromName,$sbSubject,$recTime,$mail.ItemClass,$mail.Size,$mail.id.UniqueID.ToString())
		}		
		$ivItemView.Offset += $fiItems.Items.Count
	}
	while($fiItems.MoreAvailable -eq $true)
	$dgDataGrid.datasource = $mbtable
}
function downloadAttachasXML(){
	$exFileName = new-object System.Windows.Forms.saveFileDialog
	$exFileName.DefaultExt = "xml"
	$exFileName.Filter = "xml files (*.xml)|*.xml"
	$exFileName.InitialDirectory = "c:\temp"
	$exFileName.ShowHelp = $true
	$exFileName.ShowDialog()
	$msgid = $dgDataGrid.SelectedRows[0].Cells[5].Value
	$Item = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service,$msgid)
	$Item.Load()
	
	write-host $Item.Subject
	foreach($attach in $Item.Attachments){
		$attach.Load()
		$fiFile = new-object System.IO.FileStream($exFileName.FileName, [System.IO.FileMode]::Create)
                $fiFile.Write($attach.Content, 0, $attach.Content.Length)
                $fiFile.Close()
		write-host  "Downloaded Attachment : " +  ($dlfolderpath.Self.Path + "\" + $attach.Name.ToString())
	}
	
}
function downloadCSV(){
	$exFileName = new-object System.Windows.Forms.saveFileDialog
	$exFileName.DefaultExt = "csv"
	$exFileName.Filter = "csv files (*.csv)|*.csv"
	$exFileName.InitialDirectory = "c:\temp"
	$exFileName.ShowHelp = $true
	$exFileName.ShowDialog()
	$msgid = $dgDataGrid.SelectedRows[0].Cells[5].Value
	$Item = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service,$msgid)
	$Item.Load()
   	foreach($attach in $Item.Attachments){
		$fname = $env:temp + "\" + [GUID]::newguid() + ".tmp"
		Write-Host $fname
		$attach.Load()
		$fiFile = new-object System.IO.FileStream($fname.ToString(), [System.IO.FileMode]::Create)
        $fiFile.Write($attach.Content, 0, $attach.Content.Length)
        $fiFile.Close()	
		[XML]$sResult = get-content $fname.ToString()
		$sResult.SearchResults.Event | export-csv -notype $exFileName.FileName
		Remove-Item $fname
	}  
		



}
function showdetails(){
	$ent = $dgDataGrid1.SelectedRows[0].Cells[0].Value
	$msTable.clear()
	if($gbNameDrop.SelectedItem.ToString() -eq "Cmdlet"){
		$cmdlethash[$ent] | foreach-object{
			$msTable.rows.add($_.RunDate,$_.Caller,$_.Cmdlet,$_.ObjectModified,$_.Succeeded,$_.Error,$_.OriginatingServer,$_.EntryNumb)
		}
		$dgDataGrid2.datasource = $msTable
	}
	else{
		$userhash[$ent] | foreach-object{
			$msTable.rows.add($_.RunDate,$_.Caller,$_.Cmdlet,$_.ObjectModified,$_.Succeeded,$_.Error,$_.OriginatingServer,$_.EntryNumb)
		}
		$dgDataGrid2.datasource = $msTable
	}
}
function ShowParameters($ptype){
	$prmform = new-object System.Windows.Forms.form
	$prmform.Text = "Parameter Form"
	$prmform.size = new-object System.Drawing.Size(800,400)	
	$prmform.topmost = $true
	$dgDataGrid3 = new-object System.windows.forms.DataGridView
	$dgDataGrid3.Location = new-object System.Drawing.Size(10,10)
	$dgDataGrid3.size = new-object System.Drawing.Size(400,400)
	$dgDataGrid3.AutoSizeRowsMode = "AllHeaders"
	$dgDataGrid3.SelectionMode = "FullRowSelect"
	$prmform.Controls.Add($dgDataGrid3)
	$mprmtable.clear()
	$prmtable.clear()
	$ent = $dgDataGrid2.SelectedRows[0].Cells[6].Value
	if($ptype -eq "cmdlet"){
		$enthash[$ent].CmdletParameters | foreach-object{
			$prmtable.rows.add($_.Name,$_.Value)
		}
		$dgDataGrid3.datasource = $prmtable
	}
	else{
		$enthash[$ent].ModifiedProperties | foreach-object{
			$mprmtable.rows.add($_.Name,$_.OldValue,$_.NewValue)
		}
		$dgDataGrid3.datasource = $mprmtable
	}
	$prmform.Add_Shown({$repform.Activate()})
	$prmform.ShowDialog()

}
function ExportGrid1{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("Caller,Cmdlet,ObjectModified,RunDate,Succeeded,Error,OriginatingServer")
	foreach($row in $msTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString() + "," + $row[5].ToString()) 
	}
	$logfile.Close()
}
}
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
$exButton1.Add_Click({GetAuditEmails})
$form.Controls.Add($exButton1)

$exButton7 = new-object System.Windows.Forms.Button
$exButton7.Location = new-object System.Drawing.Size(10,160)
$exButton7.Size = new-object System.Drawing.Size(180,25)
$exButton7.Text = "Export Attachment as XML"
$exButton7.Add_Click({downloadAttachasXML})
$form.Controls.Add($exButton7)

$exButton8 = new-object System.Windows.Forms.Button
$exButton8.Location = new-object System.Drawing.Size(10,190)
$exButton8.Size = new-object System.Drawing.Size(180,40)
$exButton8.Text = "Export Attachment converted to CSV"
$exButton8.Add_Click({downloadCSV})
$form.Controls.Add($exButton8)

$exButton9 = new-object System.Windows.Forms.Button
$exButton9.Location = new-object System.Drawing.Size(10,240)
$exButton9.Size = new-object System.Drawing.Size(180,25)
$exButton9.Text = "Show Report in Report Form"
$exButton9.Add_Click({showreport})
$form.Controls.Add($exButton9)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(200,150) 
$dgDataGrid.size = new-object System.Drawing.Size(800,600)
$dgDataGrid.AutoSizeRowsMode = "AllHeaders"
$dgDataGrid.SelectionMode = "FullRowSelect"
$form.Controls.Add($dgDataGrid)


$form.Text = "Simple Exchange Mailbox Audit Log Veiwer"
$form.size = new-object System.Drawing.Size(1200,700) 
$form.autoscroll = $true
$form.Add_Shown({$form.Activate()})

$repform = new-object System.Windows.Forms.form
	$repform.Text = "Exchange AuditReport Form"
	$repform.size = new-object System.Drawing.Size(800,800)



	# Add DataGrid View

	$dgDataGrid1 = new-object System.windows.forms.DataGridView
	$dgDataGrid1.Location = new-object System.Drawing.Size(10,60)
	$dgDataGrid1.size = new-object System.Drawing.Size(750,200)
	$dgDataGrid1.AutoSizeRowsMode = "AllHeaders"
	$dgDataGrid1.SelectionMode = "FullRowSelect"
	$repform.Controls.Add($dgDataGrid1)

	# Add DataGrid View

	$dgDataGrid2 = new-object System.windows.forms.DataGridView
	$dgDataGrid2.Location = new-object System.Drawing.Size(10,300)
	$dgDataGrid2.size = new-object System.Drawing.Size(750,305)
	$dgDataGrid2.AutoSizeRowsMode = "AllHeaders"
	$dgDataGrid2.SelectionMode = "FullRowSelect"
	$repform.Controls.Add($dgDataGrid2)


	# Add Groupby Drop Down
	$gbNameDrop = new-object System.Windows.Forms.ComboBox
	$gbNameDrop.Location = new-object System.Drawing.Size(100,20)
	$gbNameDrop.Size = new-object System.Drawing.Size(230,30)
	$gbNameDrop.Items.Add("Users")
	$gbNameDrop.Items.Add("Cmdlet")
	$gbNameDrop.Add_SelectedValueChanged({ReGroup})
	$repform.Controls.Add($gbNameDrop)


	# Add Groupby DropLable
	$ouOuNamelableBox = new-object System.Windows.Forms.Label
	$ouOuNamelableBox.Location = new-object System.Drawing.Size(10,20)
	$ouOuNamelableBox.size = new-object System.Drawing.Size(100,20)
	$ouOuNamelableBox.Text = "Group By"
	$repform.Controls.Add($ouOuNamelableBox)



	# Add Show Details Button

	$sdButton = new-object System.Windows.Forms.Button
	$sdButton.Location = new-object System.Drawing.Size(10,270)
	$sdButton.Size = new-object System.Drawing.Size(150,23)
	$sdButton.Text = "Show Details"
	$sdButton.Add_Click({showdetails})
	$repform.Controls.Add($sdButton)

	# Add Show parameters Button

	$sdprms = new-object System.Windows.Forms.Button
	$sdprms.Location = new-object System.Drawing.Size(170,270)
	$sdprms.Size = new-object System.Drawing.Size(150,23)
	$sdprms.Text = "Show Cmdlet Parameters"
	$sdprms.Add_Click({ShowParameters("cmdlet")})
	$repform.Controls.Add($sdprms)

	# Add Show parameters Button

	$mdprms = new-object System.Windows.Forms.Button
	$mdprms.Location = new-object System.Drawing.Size(330,270)
	$mdprms.Size = new-object System.Drawing.Size(150,23)
	$mdprms.Text = "Show Modified Parameters"
	$mdprms.Add_Click({ShowParameters("Modified")})

	$repform.Controls.Add($mdprms)
	# Export Grid


	$expGrid1 = new-object System.Windows.Forms.Button
	$expGrid1.Location = new-object System.Drawing.Size(600,610)
	$expGrid1.Size = new-object System.Drawing.Size(150,23)
	$expGrid1.Text = "Export Grid"
	$expGrid1.Add_Click({ExportGrid1})

	$repform.Controls.Add($expGrid1)

	$gbNameDrop.SelectedItem = "Users"

$form.ShowDialog()

