[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")
$global:Message = ""
$global:ewc = ""

function addtoCollection($trackingEntryAddress){
	$sdtime = $mdate.ToString("yyddMMHH")
	if ($emMailAgArray.containsKey($sdtime)){
		$hrArray = $emMailAgArray[$sdtime]
		if ($hrArray.containsKey($trackingEntryAddress.ToString())){
			$valuesobj = $hrArray[$trackingEntryAddress.ToString()]
		} 
		else{
			$valuesobj = "" |  Select-Object Email,InternalRecievedNumber,InternalRecievedSize,InternalSentNumber,InternalSentSize,ExternalRecievedNumber,ExternalRecievedSize,ExternalSentNumber,ExternalSentSize
			$valuesobj.Email = $trackingEntryAddress.ToString()
			$hrArray.Add($trackingEntryAddress.ToString(),$valuesobj)
		}
		
		
	}
	else{
		$valuesobj = "" |  Select-Object Email,InternalRecievedNumber,InternalRecievedSize,InternalSentNumber,InternalSentSize,ExternalRecievedNumber,ExternalRecievedSize,ExternalSentNumber,ExternalSentSize
		$valuesobj.Email = $trackingEntryAddress
		$emMailAddressAgArray = @{ }
		$emMailAddressAgArray.Add($trackingEntryAddress,$valuesobj)
		$emMailAgArray.Add($sdtime,$emMailAddressAgArray)
	}
	return $valuesobj
}

$emMailAgArray = @{ }
$msgIDArray = @{}
$emTotalAg = @{}

$hcRecdHourCount = @{ }
$hcSentHourCount = @{ }
$titleBlockarry = @{}
$DomainHash = @{ }

$Dataset = New-Object System.Data.DataSet
$ssTable = New-Object System.Data.DataTable
$ssTable.TableName = "TrackingLogs"
$ssTable.Columns.Add("Time",[DateTime])
$ssTable.Columns.Add("Action")
$ssTable.Columns.Add("MessageID")
$ssTable.Columns.Add("SenderAddress")
$ssTable.Columns.Add("RecipientAddress")
$ssTable.Columns.Add("Subject")
$ssTable.Columns.Add("Size (KB)",[int])
$Dataset.tables.add($ssTable)


$mdtable = New-Object System.Data.DataTable
$mdtable.TableName = "TrackingLogs-Detail"
$mdtable.Columns.Add("Time",[DateTime])
$mdtable.Columns.Add("Action")
$mdtable.Columns.Add("MessageID")
$mdtable.Columns.Add("SenderAddress")
$mdtable.Columns.Add("RecipientAddress")
$mdtable.Columns.Add("Subject")
$mdtable.Columns.Add("Size (KB)",[int])
$Dataset.tables.add($mdtable)

$agTable = New-Object System.Data.DataTable
$agTable.TableName = "Summary Tracking Logs"
$agTable.Columns.Add("Email Address   ")
$agTable.Columns.Add("Int Rcv",[int])
$agTable.Columns.Add("Int Rcv(KB)",[int])
$agTable.Columns.Add("Int Sent",[int])
$agTable.Columns.Add("Int Sent(KB)",[int])
$agTable.Columns.Add("Ext Rcv",[int])
$agTable.Columns.Add("Ext Rcv(KB)",[int])
$agTable.Columns.Add("Ext Sent",[int])
$agTable.Columns.Add("Ext Sent(KB)",[int])
$agTable.Columns.Add("Total #",[int])
$agTable.Columns.Add("Total Size(KB)",[int])

get-accepteddomain | ForEach-Object{
	if ($_.DomainType -eq "Authoritative"){
		$DomainHash.add($_.DomainName.SmtpDomain.ToString().ToLower(),1)
	}

}

$Dataset.tables.add($agTable)


function getData{

$ssTable.Clear()
$agTable.Clear()
$emMailAgArray.Clear() 
$msgIDArray.Clear() 
$emTotalAg.Clear() 
$hcRecdHourCount.Clear() 
$hcSentHourCount.Clear() 
$titleBlockarry.Clear() 
$InternalCount = 0
$ExternalCount = 0
$ExternalCount = 0
$sendcount = 0
$recievecount = 0 


$dtQueryDT = New-Object System.DateTime $dpTimeFrom.value.year,$dpTimeFrom.value.month,$dpTimeFrom.value.day,$dpTimeFrom2.value.hour,$dpTimeFrom2.value.minute,$dpTimeFrom2.value.second
$dtQueryDTf =  New-Object System.DateTime $dpTimeFrom1.value.year,$dpTimeFrom1.value.month,$dpTimeFrom1.value.day,$dpTimeFrom3.value.hour,$dpTimeFrom3.value.minute,$dpTimeFrom3.value.second


Get-MessageTrackingLog -Server $snServerNameDrop.SelectedItem.ToString() -ResultSize Unlimited -Start $dtQueryDT -End $dtQueryDTf  | ForEach-Object{ 
	$etTypeCheckDrop.SelectedItem = "All"
	$mdate = $_.TimeStamp
	$sdtime = $mdate.ToString("yyddMMHH")
	if($titleBlockArry.ContainsKey($sdtime) -eq $false){$titleBlockArry.Add($sdtime,1)}
	if ($_.EventID.ToString() -eq "SEND" -bor $_.EventID.ToString() -eq "RECEIVE"){
		foreach($recp in $_.recipients){
		if($recp.ToString() -ne ""){
			$unkey = $recp.ToString() + $_.Sender.ToString() + $_.MessageId.ToString()
			if ($msgIDArray.ContainsKey($unkey) -eq $false){
				$msgIDArray.Add($unkey,1)
				[VOID]$ssTable.Rows.Add($_.TimeStamp,$_.EventId,$_.MessageId.ToString(),$_.Sender,$recp,$_.MessageSubject.ToString(),($_.TotalBytes/1024).ToString(0.00))
				$recparray = $recp.split("@")
				$sndArray = $_.Sender.split("@")
				if ($DomainHash.ContainsKey($recparray[1])){
					if ($DomainHash.ContainsKey($sndArray[1])){
						$valuesobj = addtoCollection($recp.ToString())
						$valuesobj.InternalRecievedNumber = $valuesobj.InternalRecievedNumber + 1
						$valuesobj.InternalRecievedSize = $valuesobj.InternalRecievedSize + $_.TotalBytes/1024	
						$valuesobj = $null
						$valuesobj = addtoCollection($_.Sender.ToString())
						$valuesobj.InternalSentNumber = $valuesobj.InternalSentNumber + 1
						$valuesobj.InternalSentSize = $valuesobj.InternalSentSize + $_.TotalBytes/1024	
					}
					else{
						$valuesobj = addtoCollection($recp.ToString())
						$valuesobj.ExternalRecievedNumber = $valuesobj.ExternalRecievedNumber + 1
						$valuesobj.ExternalRecievedSize = $valuesobj.ExternalRecievedSize + $_.TotalBytes/1024	
					}			
				}
				else{
					if ($DomainHash.ContainsKey($sndArray[1])){
						$valuesobj = addtoCollection($_.Sender.ToString())
						$valuesobj.ExternalSentNumber = $valuesobj.ExternalSentNumber + 1
						$valuesobj.ExternalSentSize = $valuesobj.ExternalSentSize + $_.TotalBytes/1024	
					}				
				}
			}
			
		}
		}     
	}

}
$valueBlockSent = ""
$valueBlockRcvd = ""
$TitleBlock = ""
$lval = 0
$plval = 0
$plval1 = 0
$oSix = 0
$emMailAgArray.GetEnumerator() | sort name -descending | foreach-object {
	$sentnum = 0
	$recvnum = 0
	$valuearry = $_.value
	$valuearry.GetEnumerator() | sort name -descending | foreach-object {
		if ($emTotalAg.containskey($_.value.Email.ToString()) -eq $false){
			$valuesobj = "" |  Select-Object Email,InternalRecievedNumber,InternalRecievedSize,InternalSentNumber,InternalSentSize,ExternalRecievedNumber,ExternalRecievedSize,ExternalSentNumber,ExternalSentSize
			$valuesobj.Email = $_.value.Email.ToString()
			$emTotalAg.Add($_.value.Email.ToString(),$valuesobj)
		}
		$emTotalAg[$_.value.Email.ToString()].InternalSentNumber =  $emTotalAg[$_.value.Email.ToString()].InternalSentNumber + $_.value.InternalSentNumber
		$emTotalAg[$_.value.Email.ToString()].InternalSentSize =  $emTotalAg[$_.value.Email.ToString()].InternalSentSize + $_.value.InternalSentSize
		$emTotalAg[$_.value.Email.ToString()].ExternalSentNumber =  $emTotalAg[$_.value.Email.ToString()].ExternalSentNumber + $_.value.ExternalSentNumber
		$emTotalAg[$_.value.Email.ToString()].ExternalSentSize =  $emTotalAg[$_.value.Email.ToString()].ExternalSentSize + $_.value.ExternalSentSize
		$emTotalAg[$_.value.Email.ToString()].InternalRecievedNumber =  $emTotalAg[$_.value.Email.ToString()].InternalRecievedNumber + $_.value.InternalRecievedNumber
		$emTotalAg[$_.value.Email.ToString()].InternalRecievedSize =  $emTotalAg[$_.value.Email.ToString()].InternalRecievedSize + $_.value.InternalRecievedSize
		$emTotalAg[$_.value.Email.ToString()].ExternalRecievedNumber =  $emTotalAg[$_.value.Email.ToString()].ExternalRecievedNumber + $_.value.ExternalRecievedNumber
		$emTotalAg[$_.value.Email.ToString()].ExternalRecievedSize =  $emTotalAg[$_.value.Email.ToString()].ExternalRecievedSize + $_.value.ExternalRecievedSize
		$sentnum = $sentnum +  $_.value.InternalSentNumber +  $_.value.ExternalSentNumber
		$recvnum = $recvnum +  $_.value.InternalRecievedNumber +  $_.value.ExternalRecievedNumber
		$InternalCount = $InternalCount + $_.value.InternalRecievedNumber + $_.value.InternalSentNumber
		$ExternalCount = $ExternalCount + $_.value.ExternalRecievedNumber +  $_.value.ExternalSentNumber
		$sendcount = $sendcount + $_.value.InternalSentNumber + $_.value.ExternalSentNumber
		$recievecount = $recievecount + $_.value.InternalRecievedNumber + $_.value.ExternalRecievedNumber
	}
	if ($oSix -lt 7){ 
		if ($sentnum -eq $null){$sentnum = 0}
		if ($recvnum -eq $null){$recvnum = 0}
		if ($lval -lt $sentnum){$lval = $sentnum}
		if ($ValueBlockSent -eq "") {$ValueBlockSent = $sentnum.ToString() }
		else {$ValueBlockSent =   $ValueBlockSent + "," + $sentnum.ToString()}
		if ($lval -lt $recvnum){$lval = $recvnum}
		if ($valueBlockRcvd -eq "") {$valueBlockRcvd = $recvnum.ToString()}
		else {$valueBlockRcvd =   $valueBlockRcvd + "," + $recvnum.ToString()}
		if ($TitleBlock -eq ""){$TitleBlock = $_.key.ToString().SubString(6,2) + ":00"}
		else {$TitleBlock =  $_.key.ToString().SubString(6,2) + ":00" + "|" + $TitleBlock}
		$oSix++
	}

}
$intTotalSizeSent = 0
$intTotalNumberSent = 0
$intTotalSizercvd = 0
$intTotalNumberrcvd = 0
$extTotalSizeSent = 0
$extTotalNumberSent = 0
$extTotalSizercvd = 0
$extTotalNumberrcvd = 0

$emTotalAg.GetEnumerator() | sort name -descending | foreach-object {
	$totalSize = 0
	$totalNumber = 0
	$totalNumber = $_.Value.InternalRecievedNumber + $_.Value.InternalSentNumber + $_.Value.ExternalRecievedNumber + $_.Value.ExternalSentNumber
	$totalSize = $_.Value.InternalRecievedSize + $_.Value.InternalSentSize + $_.Value.ExternalRecievedSize + $_.Value.ExternalSentSize
	[VOID]$agTable.Rows.Add($_.Value.Email.ToString(),$_.Value.InternalRecievedNumber,$_.Value.InternalRecievedSize,$_.Value.InternalSentNumber,$_.Value.InternalSentSize,$_.Value.ExternalRecievedNumber,$_.Value.ExternalRecievedSize,$_.Value.ExternalSentNumber,$_.Value.ExternalSentSize,$totalNumber,$totalSize )
	$intTotalSizeSent = $intTotalSizeSent + $_.Value.InternalSentSize
	$intTotalNumberSent = $intTotalNumberSent + $_.Value.InternalSentNumber
	$intTotalSizercvd = $intTotalSizercvd + $_.Value.InternalRecievedSize
	$intTotalNumberrcvd = $intTotalNumberrcvd + $_.Value.InternalRecievedNumber
	$extTotalSizeSent = $extTotalSizeSent + $_.Value.ExternalSentSize 
	$extTotalNumberSent = $extTotalNumberSent + $_.Value.ExternalSentNumber
	$extTotalSizercvd = $extTotalSizercvd + $_.Value.ExternalRecievedSize
	$extTotalNumberrcvd = $extTotalNumberrcvd + $_.Value.ExternalRecievedNumber
}

# last hours date
$lasthourNum = $null
$lasthourNum = @()
$lbListView.Clear()
$lbListView.Columns.Add("Last hour top Sender/Recievers",200)
$lbListView.Columns.Add("# Total Items",80)
$lbListView.Columns.Add("Total Size",80)
$datenow = get-date
$lasthour = $emMailAgArray[$datenow.ToString("yyddMMHH")]
$lasthour.GetEnumerator() | sort name -descending | foreach-object {
	$totalSize = 0
	$totalNumber = 0
	$totalNumber = $_.Value.InternalRecievedNumber + $_.Value.InternalSentNumber + $_.Value.ExternalRecievedNumber + $_.Value.ExternalSentNumber
	$totalSize = $_.Value.InternalRecievedSize + $_.Value.InternalSentSize + $_.Value.ExternalRecievedSize + $_.Value.ExternalSentSize
	$lasthourobj = "" | select Email,TotalNumber,TotalSize
	$lasthourobj.Email = $_.Key.ToString()
	$lasthourobj.TotalNumber = $totalNumber
	$lasthourobj.TotalSize = $totalSize.ToString(0.00)
	$lasthourNum += $lasthourobj
}
$icCount = 0
$lasthourNum.GetEnumerator() | sort TotalNumber -descending | foreach-object { 
	if ($icCount -lt 5){
		write-host $_.Email
		$item1  = new-object System.Windows.Forms.ListViewItem($_.Email.ToString())
		$item1.SubItems.Add($_.TotalNumber.ToString())
		$item1.SubItems.Add($_.TotalSize.ToString())
		$lbListView.items.add($item1)
	}
	$icCount++
}
$tab1.Controls.Add($lbListView)

$lbListView1.Clear()
$lbListView1.Columns.Add("Organization Totals",200)
$lbListView1.Columns.Add("# Total Items",80)
$lbListView1.Columns.Add("Total Size",80)

$item1  = new-object System.Windows.Forms.ListViewItem("Internal Recieved Mail")
$item1.SubItems.Add($intTotalNumberrcvd.ToString())
$item1.SubItems.Add($intTotalSizercvd.ToString(0.00))
$lbListView1.items.add($item1)

$item1  = new-object System.Windows.Forms.ListViewItem("Internal Sent Mail")
$item1.SubItems.Add($intTotalNumberSent.ToString())
$item1.SubItems.Add($intTotalSizeSent.ToString(0.00))
$lbListView1.items.add($item1)

$item1  = new-object System.Windows.Forms.ListViewItem("External Recieved Mail")
$item1.SubItems.Add($extTotalNumberrcvd.ToString())
$item1.SubItems.Add($extTotalSizercvd.ToString(0.00))
$lbListView1.items.add($item1)

$item1  = new-object System.Windows.Forms.ListViewItem("External Sent Mail")
$item1.SubItems.Add($extTotalNumberSent.ToString())
$item1.SubItems.Add($extTotalSizeSent.ToString(0.00))
$lbListView1.items.add($item1)





if ($plval -lt $InternalCount){$plval = $InternalCount} 
if ($plval -lt $ExternalCount){$plval = $ExternalCount} 
if ($plval1 -lt $sendcount){$plval1 = $sendcount} 
if ($plval1 -lt $recievecount){$plval1 = $recievecount} 

$csString = "http://chart.apis.google.com/chart?cht=bhs&chs=400x250&chd=t:" + $ValueBlockSent + "|" + $valueBlockRcvd + "&chds=0," + ($lval+20)  + "&chxt=x,y&chxr=" + "&chxr=0,0," + ($lval+20) + "&chxl=1:|" + $TitleBlock + "&chdl=Sent|Recieved&chco=4d89f9,ff0000&chtt=Message+Volume++Last+6+Hours"
$csString1 = "http://chart.apis.google.com/chart?cht=p3&chs=370x120&chds=0," + $plval + "&chd=t:" + $InternalCount + "," + $ExternalCount  + "&chl=Internal|External&chco=4d89f9,ff0000&"
$csString2 = "http://chart.apis.google.com/chart?cht=p3&chs=370x120&chds=0," + $plval1 + "&chd=t:" + $sendcount + "," + $recievecount  + "&chl=Sent|Recieved&chco=4d89f9,ff0000&"
$pbox.ImageLocation = $csString
$pbox1.ImageLocation = $csString1
$pbox2.ImageLocation = $csString2
$tabControl.SelectedTab = $tab1


}

function ShowHeaders{
	$headerform = new-object System.Windows.Forms.form 
	$headerform.Text = "Internet Headers"
	$headertext = ""
	foreach ($ihead in $global:Message.InternetMessageHeaders){
        	$headertext =  $headertext + $ihead.HeaderName.ToString() + ":" +  $ihead.Value.ToString() + "`r`n"
                
        }
	# Add Messageheader box 
	$mhMessageBodytextlabelBox = new-object System.Windows.Forms.RichTextBox
	$mhMessageBodytextlabelBox.Location = new-object System.Drawing.Size(0,10) 
	$mhMessageBodytextlabelBox.size = new-object System.Drawing.Size(600,300) 
	$mhMessageBodytextlabelBox.text = $headertext
	$headerform.controls.Add($mhMessageBodytextlabelBox) 
	$headerform.size = new-object System.Drawing.Size(620,420) 
	$headerform.Add_Shown({$headerform.Activate()})
	$headerform.autoscroll = $true
	$headerform.ShowDialog()	

}

function getMessages{
$exbutton4.Enabled = $true
$mdtable.Clear()
foreach($row in $ssTable.Rows){
	$reciparry1 =$row[4].split("@")
	$sndr = $row[3].split("@")
	switch ($etTypeCheckDrop.SelectedItem.ToString()){
		"Internal Recieved" {
			if ($DomainHash.ContainsKey($reciparry1[1]) -band $DomainHash.ContainsKey($sndr[1]) -band $row[4] -eq  $agTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0]){
				$mdtable.rows.add($row[0],$row[1],$row[2],$row[3],$row[4],$row[5],$row[6])
			}
		}
		"Internal Sent" {
			if ($DomainHash.ContainsKey($reciparry1[1]) -band $DomainHash.ContainsKey($sndr[1]) -band $row[3] -eq  $agTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0]){
				$mdtable.rows.add($row[0],$row[1],$row[2],$row[3],$row[4],$row[5],$row[6])
			}
		}
		"External Recieved" {
			if ($DomainHash.ContainsKey($reciparry1[1]) -band $DomainHash.ContainsKey($sndr[1]) -eq $false -band $row[4] -eq  $agTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0]){
				$mdtable.rows.add($row[0],$row[1],$row[2],$row[3],$row[4],$row[5],$row[6])
			}
		}
		"External Sent" {
			if ($DomainHash.ContainsKey($reciparry1[1]) -eq $false -band $DomainHash.ContainsKey($sndr[1]) -band $row[3] -eq  $agTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0]){
				$mdtable.rows.add($row[0],$row[1],$row[2],$row[3],$row[4],$row[5],$row[6])
			}
		}
		"All" {
			if ($row[4] -eq  $agTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] -bor $row[3] -eq  $agTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0]){
				$mdtable.rows.add($row[0],$row[1],$row[2],$row[3],$row[4],$row[5],$row[6])
			}
		}
	}	
		
}
$dgDataGrid2.DataSource = $mdtable
}

function showMessage(){

$exButton2.Enabled = $false
$exButton3.Enabled = $false
$exButton1.Enabled = $false
$flFolderDrop.SelectedItem = "All Folders"
$exButton4.Enabled = $true
$miMessageAttachmentslableBox1.Text = ""
$miMessageSubjecttextlabelBox.Text = ""
$miMessageBodytextlabelBox.Text = ""
$miMessageTotextlabelBox.Text = ""
$miMessageSenttextlabelBox.Text = ""
$miMessageIDTextBox.text = $mdtable.DefaultView[$dgDataGrid2.CurrentCell.RowIndex][2]
$miMailboxTextBox.text = $agtable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0]
$tabControl.SelectedTab = $tab4

}
function getMessage(){
if ($unCASUrlTextBox.text -eq ""){ $casurl = $null}
else { $casurl= $unCASUrlTextBox.text}
$useImp = $false
$exButton2.Enabled = $true
$exButton3.Enabled = $true
$exButton1.Enabled = $true
if ($seImpersonationCheck.Checked -eq $true) {
	$useImp = $true
}
if ($seAuthCheck.Checked -eq $true) {
	$ewc = new-object EWSUtil.EWSConnection($miMailboxTextBox.Text.ToString(),$useImp, $unUserNameTextBox.Text, $unPasswordTextBox.Text, $unDomainTextBox.Text,$casUrl)
}
else{
	$ewc = new-object EWSUtil.EWSConnection($miMailboxTextBox.Text.ToString(),$useImp, "", "", "",$casUrl)
}
$global:ewc = $ewc
$dType = new-object EWSUtil.EWS.DistinguishedFolderIdType
switch ($flFolderDrop.SelectedItem.ToString()){
	"Inbox" {$dType.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox}
	"Sent Items" {$dType.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::sentitems}
	"All Folders" {$dType.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::msgfolderroot}

}

$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
$mbMailbox = new-object EWSUtil.EWS.EmailAddressType
$mbMailbox.EmailAddress = $miMailboxTextBox.Text.ToString()
$dType.Mailbox = $mbMailbox
$fldarry[0] = $dType
$randNumber = New-Object system.random
$prop = new-object EWSUtil.EWS.PathToUnindexedFieldType
$prop.FieldURI = [EWSUtil.EWS.UnindexedFieldURIType]::messageInternetMessageId
if ($flFolderDrop.SelectedItem.ToString() -eq "All Folders"){
	$messagelist = $ewc.RecurseFolder($fldarry,$prop,$miMessageIDTextBox.Text,$false)
}
else {
	$messagelist = $ewc.FindMessage($fldarry,$prop,$miMessageIDTextBox.Text,$false)
}
foreach ($message in $messagelist){
	$global:Message = $message
	$miMessageSubjecttextlabelBox.Text = $message.Subject
	$miMessageSenttextlabelBox.Text = $message.Sender.Item.EmailAddress.ToString()
	$miMessageBodytextlabelBox.Text = $message.Body.Value.ToString()
	$recpval = ""
	foreach ($Recp in $message.ToRecipients){
		$recpval = $recpval + $Recp.EmailAddress.ToString() + ";"
	}
	$miMessageTotextlabelBox.Text = $recpval
	$siStart = 555
	$attname = ""
	if ($message.hasattachments){
		foreach($attach in $message.Attachments)
		{
			$attname = $attname + $attach.Name.ToString() + "; "
		}
	}
	$miMessageAttachmentslableBox1.Text = $attname

}
}

function downloadattachments{
	$dlfolder = new-object -com shell.application 
	$dlfolderpath = $dlfolder.BrowseForFolder(0,"Download attachments to",0) 
	
	write-host $global:Message
	foreach($attach in $global:Message.Attachments){
      	           $global:ewc.DownloadAttachment(($dlfolderpath.Self.Path + "\" + $attach.Name.ToString()),$attach.AttachmentId);
      	            write-host  "Downloaded Attachment : " +  ($dlfolderpath.Self.Path + "\" + $attach.Name.ToString())
	}

}

function ExportMessage{
	$exFileName = new-object System.Windows.Forms.saveFileDialog
	$exFileName.DefaultExt = "eml"
	$exFileName.Filter = "emlFiles files (*.eml)|*.eml"
	$exFileName.InitialDirectory = "c:\temp"
	$exFileName.ShowHelp = $true
	$exFileName.ShowDialog()
	if ($exFileName.FileName -ne ""){
		$ascii = new-object System.Text.ASCIIEncoding
		$baByteArray = [Convert]::FromBase64String($global:Message.MimeContent.Value)
		$emlMessage = $ascii.GetString($baByteArray)
		$emlfile = new-object IO.StreamWriter($exFileName.FileName,$true)
		$emlfile.WriteLine($emlMessage)
		$emlfile.Close()
	}

}

function ExportAgGrid{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.Showhelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("EmailAddress,Int Rcv,Int Rcv(KB),Int Sent,Int Sent(KB),Ext Rcv,Ext Rcv(KB),Ext Sent,Ext Sent(KB),Total#,Total Size (KB)")
	foreach($row in $agTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString() + "," + $row[5].ToString()+ "," + $row[6].ToString() + "," + $row[7].ToString() + "," + $row[8].ToString() + "," + $row[9].ToString() + "," + $row[10].ToString()) 
	}
	$logfile.Close()
}

}


function ExportMessagedetailGrid{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.Showhelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("Time,Action,MessageID,SenderAddress,RecipientAddress,Subject,Size(KB)")
	foreach($row in $mdtable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString() + "," + $row[5].ToString()+ "," + $row[6].ToString()) 
	}
	$logfile.Close()
}

}

$form = new-object System.Windows.Forms.form 
$form.Text = "WizBang 2007 Message Tracker"


# Add Tabs

$tabControl = new-object System.windows.forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill

$tab0 = new-object System.windows.forms.TabPage
$tab0.Dock = [System.Windows.Forms.DockStyle]::Fill
$tab0.Text= "Query Settings"

$tab1 = new-object System.windows.forms.TabPage
$tab1.Dock = [System.Windows.Forms.DockStyle]::Fill
$tab1.Text= "DashBoard"


$tab2 = new-object System.windows.forms.TabPage
$tab2.Dock = [System.Windows.Forms.DockStyle]::Fill
$tab2.Text= "Email Summaries"

$tab3 = new-object System.windows.forms.TabPage
$tab3.Dock = [System.Windows.Forms.DockStyle]::Fill
$tab3.Text= "Tracking Data Raw"

$tab4 = new-object System.windows.forms.TabPage
$tab4.size = new-object System.Drawing.Size(1000,800) 
$tab4.Text= "Message Find"

# Add Servername Box
$snServerNameDrop = new-object System.Windows.Forms.ComboBox
get-Exchangeserver | ForEach-Object{$snServerNameDrop.Items.Add($_.Name)}
$snServerNameDrop.Location = new-object System.Drawing.Size(100,30) 
$snServerNameDrop.size = new-object System.Drawing.Size(200,20) 
$tab0.Controls.Add($snServerNameDrop) 

# Add Servername Lable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,30) 
$snServerNamelableBox.size = new-object System.Drawing.Size(100,20) 
$snServerNamelableBox.Text = "ServerName"
$tab0.Controls.Add($snServerNamelableBox) 

# Add Search Button

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(10,80)
$exButton.Size = new-object System.Drawing.Size(85,20)
$exButton.Text = "Get Logs"
$exButton.Add_Click({getData})
$tab0.Controls.Add($exButton)


# Add DateTimePickers Button

$dpDatePickerFromlableBox = new-object System.Windows.Forms.Label
$dpDatePickerFromlableBox.Location = new-object System.Drawing.Size(320,30) 
$dpDatePickerFromlableBox.size = new-object System.Drawing.Size(90,20) 
$dpDatePickerFromlableBox.Text = "Logged Between"
$tab0.Controls.Add($dpDatePickerFromlableBox) 

$dpTimeFrom = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom.Location = new-object System.Drawing.Size(410,30)
$dpTimeFrom.Size = new-object System.Drawing.Size(190,20)
$tab0.Controls.Add($dpTimeFrom)

$dpDatePickerFromlableBox1 = new-object System.Windows.Forms.Label
$dpDatePickerFromlableBox1.Location = new-object System.Drawing.Size(350,50) 
$dpDatePickerFromlableBox1.size = new-object System.Drawing.Size(50,20) 
$dpDatePickerFromlableBox1.Text = "and"
$tab0.Controls.Add($dpDatePickerFromlableBox1) 

$dpTimeFrom1 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom1.Location = new-object System.Drawing.Size(410,50)
$dpTimeFrom1.Size = new-object System.Drawing.Size(190,20)
$tab0.Controls.Add($dpTimeFrom1)

$dpTimeFrom2 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom2.Format = "Time"
$dpTimeFrom2.value = [DateTime]::get_Now().AddHours(-1)
$dpTimeFrom2.ShowUpDown = $True
$dpTimeFrom2.Location = new-object System.Drawing.Size(610,30)
$dpTimeFrom2.Size = new-object System.Drawing.Size(190,20)
$tab0.Controls.Add($dpTimeFrom2)

$dpTimeFrom3 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom3.Format = "Time"
$dpTimeFrom3.ShowUpDown = $True
$dpTimeFrom3.Location = new-object System.Drawing.Size(610,50)
$dpTimeFrom3.Size = new-object System.Drawing.Size(190,20)
$tab0.Controls.Add($dpTimeFrom3)

# Add Get Messages Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(10,315)
$exButton1.Size = new-object System.Drawing.Size(85,20)
$exButton1.Text = "Get Messages"
$exButton1.Add_Click({getMessages})
$tab2.Controls.Add($exButton1)

$exButton4 = new-object System.Windows.Forms.Button
$exButton4.Location = new-object System.Drawing.Size(260,315)
$exButton4.Size = new-object System.Drawing.Size(130,20)
$exButton4.Text = "Show Message"
$exButton4.Enabled = $false
$exButton4.Add_Click({showMessage})
$tab2.Controls.Add($exButton4)


$exButton5 = new-object System.Windows.Forms.Button
$exButton5.Location = new-object System.Drawing.Size(400,315)
$exButton5.Size = new-object System.Drawing.Size(130,20)
$exButton5.Text = "Export Top Grid"
$exButton5.Enabled = $true
$exButton5.Add_Click({ExportAgGrid})
$tab2.Controls.Add($exButton5)


$exButton6 = new-object System.Windows.Forms.Button
$exButton6.Location = new-object System.Drawing.Size(540,315)
$exButton6.Size = new-object System.Drawing.Size(130,20)
$exButton6.Text = "Export Bottom Grid"
$exButton6.Enabled = $true
$exButton6.Add_Click({ExportMessagedetailGrid})
$tab2.Controls.Add($exButton6)

$etTypeCheckDrop = new-object System.Windows.Forms.ComboBox
$etTypeCheckDrop.Location = new-object System.Drawing.Size(100,315)
$etTypeCheckDrop.Size = new-object System.Drawing.Size(140,30)
$etTypeCheckDrop.Enabled = $true
$etTypeCheckDrop.Items.Add("Internal Recieved")
$etTypeCheckDrop.Items.Add("External Recieved")
$etTypeCheckDrop.Items.Add("Internal Sent")
$etTypeCheckDrop.Items.Add("External Sent")
$etTypeCheckDrop.Items.Add("All")
$tab2.Controls.Add($etTypeCheckDrop)

#add Picture box
$pbox =  new-object System.Windows.Forms.PictureBox
$pbox.Location = new-object System.Drawing.Size(10,10)
$pbox.Size = new-object System.Drawing.Size(400,250)
$pbox.ImageLocation = $csString

$pbox1 =  new-object System.Windows.Forms.PictureBox
$pbox1.Location = new-object System.Drawing.Size(420,10)
$pbox1.Size = new-object System.Drawing.Size(400,120)
$pbox1.ImageLocation = $csString1

$pbox2 =  new-object System.Windows.Forms.PictureBox
$pbox2.Location = new-object System.Drawing.Size(420,140)
$pbox2.Size = new-object System.Drawing.Size(400,130)
$pbox2.ImageLocation = $csString2

$tab1.Controls.Add($pbox)
$tab1.Controls.Add($pbox1)
$tab1.Controls.Add($pbox2)

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,10) 
$dgDataGrid.size = new-object System.Drawing.Size(800,300)



$tab2.Controls.Add($dgDataGrid)

$dgDataGrid1 = new-object System.windows.forms.DataGridView
$dgDataGrid1.Location = new-object System.Drawing.Size(10,10) 
$dgDataGrid1.size = new-object System.Drawing.Size(800,600)
$tab3.Controls.Add($dgDataGrid1)

$dgDataGrid2 = new-object System.windows.forms.DataGridView
$dgDataGrid2.Location = new-object System.Drawing.Size(10,350) 
$dgDataGrid2.size = new-object System.Drawing.Size(800,300)
$tab2.Controls.Add($dgDataGrid2)

$lbListView = new-object System.Windows.Forms.ListView
$lbListView.Location = new-object System.Drawing.Size(10,280) 
$lbListView.size = new-object System.Drawing.Size(400,90)
$lbListView.LabelEdit = $True
$lbListView.AllowColumnReorder = $True
$lbListView.CheckBoxes = $False
$lbListView.FullRowSelect = $True
$lbListView.GridLines = $True
$lbListView.View = "Details"
$lbListView.Sorting = "Ascending"

$lbListView1 = new-object System.Windows.Forms.ListView
$lbListView1.Location = new-object System.Drawing.Size(410,280) 
$lbListView1.size = new-object System.Drawing.Size(400,90)
$lbListView1.LabelEdit = $True
$lbListView1.AllowColumnReorder = $True
$lbListView1.CheckBoxes = $False
$lbListView1.FullRowSelect = $True
$lbListView1.GridLines = $True
$lbListView1.View = "Details"
$lbListView1.Sorting = "Ascending"
$tab1.Controls.Add($lbListView1)


$dgDataGrid.DataSource = $agTable
$dgDataGrid1.DataSource = $ssTable

# Add MessageID Box
$miMessageIDTextBox = new-object System.Windows.Forms.TextBox 
$miMessageIDTextBox.Location = new-object System.Drawing.Size(110,20) 
$miMessageIDTextBox.size = new-object System.Drawing.Size(500,20) 
$tab4.controls.Add($miMessageIDTextBox) 

# Add MessageID Lable
$miMessageIDlableBox = new-object System.Windows.Forms.Label
$miMessageIDlableBox.Location = new-object System.Drawing.Size(20,20) 
$miMessageIDlableBox.size = new-object System.Drawing.Size(100,20) 
$miMessageIDlableBox.Text = "Message ID"
$tab4.controls.Add($miMessageIDlableBox) 

# Add Mailbox
$miMailboxTextBox = new-object System.Windows.Forms.TextBox 
$miMailboxTextBox.Location = new-object System.Drawing.Size(110,45) 
$miMailboxTextBox.size = new-object System.Drawing.Size(300,20) 
$tab4.controls.Add($miMailboxTextBox) 

# Add Mailbox Lable

$miMailboxlableBox = new-object System.Windows.Forms.Label
$miMailboxlableBox.Location = new-object System.Drawing.Size(20,45) 
$miMailboxlableBox.size = new-object System.Drawing.Size(100,20) 
$miMailboxlableBox.Text = "Mailbox"
$tab4.controls.Add($miMailboxlableBox) 

# Add Folder DropLable
$flFolderlableBox = new-object System.Windows.Forms.Label
$flFolderlableBox.Location = new-object System.Drawing.Size(20,70) 
$flFolderlableBox.size = new-object System.Drawing.Size(90,20) 
$flFolderlableBox.Text = "Folder to Search"
$tab4.controls.Add($flFolderlableBox) 

# Add Folder Drop Down
$flFolderDrop = new-object System.Windows.Forms.ComboBox
$flFolderDrop.Location = new-object System.Drawing.Size(110,70)
$flFolderDrop.Size = new-object System.Drawing.Size(100,30)
$flFolderDrop.Items.Add("Inbox")
$flFolderDrop.Items.Add("Sent Items")
$flFolderDrop.Items.Add("All Folders")

$tab4.controls.Add($flFolderDrop)

# Add Impersonation Clause

$esImpersonationlableBox = new-object System.Windows.Forms.Label
$esImpersonationlableBox.Location = new-object System.Drawing.Size(210,75) 
$esImpersonationlableBox.Size = new-object System.Drawing.Size(130,20) 
$esImpersonationlableBox.Text = "Use EWS Impersonation"
$tab4.controls.Add($esImpersonationlableBox) 

$seImpersonationCheck =  new-object System.Windows.Forms.CheckBox
$seImpersonationCheck.Location = new-object System.Drawing.Size(350,70)
$seImpersonationCheck.Size = new-object System.Drawing.Size(30,25)
$tab4.controls.Add($seImpersonationCheck)

# Add Auth Clause

$esAuthlableBox = new-object System.Windows.Forms.Label
$esAuthlableBox.Location = new-object System.Drawing.Size(10,105) 
$esAuthlableBox.Size = new-object System.Drawing.Size(130,20) 
$esAuthlableBox.Text = "Specify Credentials"
$tab4.controls.Add($esAuthlableBox) 

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
$tab4.controls.Add($seAuthCheck)

# Add UserName Box
$unUserNameTextBox = new-object System.Windows.Forms.TextBox 
$unUserNameTextBox.Location = new-object System.Drawing.Size(230,100) 
$unUserNameTextBox.size = new-object System.Drawing.Size(100,20) 
$tab4.controls.Add($unUserNameTextBox) 

# Add UserName Lable
$unUserNamelableBox = new-object System.Windows.Forms.Label
$unUserNamelableBox.Location = new-object System.Drawing.Size(170,105) 
$unUserNamelableBox.size = new-object System.Drawing.Size(60,20) 
$unUserNamelableBox.Text = "UserName"
$unUserNameTextBox.Enabled = $false
$tab4.controls.Add($unUserNamelableBox) 

# Add Password Box
$unPasswordTextBox = new-object System.Windows.Forms.TextBox 
$unPasswordTextBox.PasswordChar = "*"
$unPasswordTextBox.Location = new-object System.Drawing.Size(400,100) 
$unPasswordTextBox.size = new-object System.Drawing.Size(100,20) 
$tab4.controls.Add($unPasswordTextBox) 

# Add Password Lable
$unPasswordlableBox = new-object System.Windows.Forms.Label
$unPasswordlableBox.Location = new-object System.Drawing.Size(340,105) 
$unPasswordlableBox.size = new-object System.Drawing.Size(60,20) 
$unPasswordlableBox.Text = "Password"
$unPasswordTextBox.Enabled = $false
$tab4.controls.Add($unPasswordlableBox) 

# Add Domain Box
$unDomainTextBox = new-object System.Windows.Forms.TextBox 
$unDomainTextBox.Location = new-object System.Drawing.Size(550,100) 
$unDomainTextBox.size = new-object System.Drawing.Size(100,20) 
$tab4.controls.Add($unDomainTextBox) 

# Add Domain Lable
$unDomainlableBox = new-object System.Windows.Forms.Label
$unDomainlableBox.Location = new-object System.Drawing.Size(510,105) 
$unDomainlableBox.size = new-object System.Drawing.Size(50,20) 
$unDomainlableBox.Text = "Domain"
$unDomainTextBox.Enabled = $false
$tab4.controls.Add($unDomainlableBox) 


# Add CASUrl Box
$unCASUrlTextBox = new-object System.Windows.Forms.TextBox 
$unCASUrlTextBox.Location = new-object System.Drawing.Size(440,75) 
$unCASUrlTextBox.size = new-object System.Drawing.Size(400,20) 
$unCASUrlTextBox.text = $strRootURI
$tab4.Controls.Add($unCASUrlTextBox) 

# Add CASUrl Lable
$unCASUrllableBox = new-object System.Windows.Forms.Label
$unCASUrllableBox.Location = new-object System.Drawing.Size(380,75) 
$unCASUrllableBox.size = new-object System.Drawing.Size(50,20) 
$unCASUrllableBox.Text = "CASUrl"
$tab4.Controls.Add($unCASUrllableBox) 


# Add Search Button

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(10,130)
$exButton.Size = new-object System.Drawing.Size(110,20)
$exButton.Text = "Search Mailbox"
$exButton.Add_Click({getMessage})
$tab4.controls.Add($exButton)

# Add Download Button

$exButton1 = new-object System.Windows.Forms.Button
$exButton1.Location = new-object System.Drawing.Size(260,130)
$exButton1.Size = new-object System.Drawing.Size(150,20)
$exButton1.Text = "Download Attachments"
$exButton1.Enabled = $false
$exButton1.Add_Click({DownloadAttachments})
$tab4.Controls.Add($exButton1)

# Add Export Button
$exButton2 = new-object System.Windows.Forms.Button
$exButton2.Location = new-object System.Drawing.Size(130,130)
$exButton2.Size = new-object System.Drawing.Size(120,20)
$exButton2.Text = "Export Message"
$exButton2.Enabled = $false
$exButton2.Add_Click({ExportMessage})
$tab4.controls.Add($exButton2)

# Add Headers Button

$exButton3 = new-object System.Windows.Forms.Button
$exButton3.Location = new-object System.Drawing.Size(420,130)
$exButton3.Size = new-object System.Drawing.Size(150,20)
$exButton3.Text = "Show Headers"
$exButton3.Enabled = $false
$exButton3.Add_Click({ShowHeaders})
$tab4.Controls.Add($exButton3)

# Add Message From Lable
$miMessageTolableBox = new-object System.Windows.Forms.Label
$miMessageTolableBox.Location = new-object System.Drawing.Size(20,165) 
$miMessageTolableBox.size = new-object System.Drawing.Size(80,20) 
$miMessageTolableBox.Text = "To"
$tab4.controls.Add($miMessageTolableBox) 

# Add MessageID Lable
$miMessageSentlableBox = new-object System.Windows.Forms.Label
$miMessageSentlableBox.Location = new-object System.Drawing.Size(20,185) 
$miMessageSentlableBox.size = new-object System.Drawing.Size(80,20) 
$miMessageSentlableBox.Text = "From"
$tab4.controls.Add($miMessageSentlableBox) 

# Add Message Subject Lable
$miMessageSubjectlableBox = new-object System.Windows.Forms.Label
$miMessageSubjectlableBox.Location = new-object System.Drawing.Size(20,205) 
$miMessageSubjectlableBox.size = new-object System.Drawing.Size(80,20) 
$miMessageSubjectlableBox.Text = "Subject"
$tab4.controls.Add($miMessageSubjectlableBox) 

# Add Message To
$miMessageTotextlabelBox = new-object System.Windows.Forms.Label
$miMessageTotextlabelBox.Location = new-object System.Drawing.Size(100,165) 
$miMessageTotextlabelBox.size = new-object System.Drawing.Size(400,20) 
$tab4.controls.Add($miMessageTotextlabelBox) 

# Add Message From
$miMessageSenttextlabelBox = new-object System.Windows.Forms.Label
$miMessageSenttextlabelBox.Location = new-object System.Drawing.Size(100,185) 
$miMessageSenttextlabelBox.size = new-object System.Drawing.Size(400,20) 
$tab4.controls.Add($miMessageSenttextlabelBox) 

# Add Message Subject 
$miMessageSubjecttextlabelBox = new-object System.Windows.Forms.Label
$miMessageSubjecttextlabelBox.Location = new-object System.Drawing.Size(100,205) 
$miMessageSubjecttextlabelBox.size = new-object System.Drawing.Size(400,20) 
$tab4.controls.Add($miMessageSubjecttextlabelBox) 

# Add Message body 
$miMessageBodytextlabelBox = new-object System.Windows.Forms.RichTextBox
$miMessageBodytextlabelBox.Location = new-object System.Drawing.Size(100,225) 
$miMessageBodytextlabelBox.size = new-object System.Drawing.Size(600,200) 
$tab4.controls.Add($miMessageBodytextlabelBox) 

# Add Message Attachments Lable
$miMessageAttachmentslableBox = new-object System.Windows.Forms.Label
$miMessageAttachmentslableBox.Location = new-object System.Drawing.Size(20,435) 
$miMessageAttachmentslableBox.size = new-object System.Drawing.Size(80,20) 
$miMessageAttachmentslableBox.Text = "Attachments"
$tab4.controls.Add($miMessageAttachmentslableBox) 

$miMessageAttachmentslableBox1 = new-object System.Windows.Forms.Label
$miMessageAttachmentslableBox1.Location = new-object System.Drawing.Size(100,445) 
$miMessageAttachmentslableBox1.size = new-object System.Drawing.Size(600,20) 
$miMessageAttachmentslableBox1.Text = ""
$tab4.Controls.Add($miMessageAttachmentslableBox1) 
$siStart = $siStart + 20


$tabControl.Tabpages.add($tab0)
$tabControl.Tabpages.add($tab1)
$tabControl.Tabpages.add($tab2)
$tabControl.Tabpages.add($tab3)
$tabControl.Tabpages.add($tab4)


$form.Controls.Add($tabControl)
$form.size = new-object System.Drawing.Size(1000,800) 
$form.Add_Shown({$form.Activate()})
$form.autoscroll = $true
$form.ShowDialog()