[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

$unUserName = "username"
$psPassword = "password$"
$dnDomainName = "domain"
$cdUsrCredentials = new-object System.Net.NetworkCredential($unUserName , $psPassword , $dnDomainName)

function getMailboxSizes(){

$lbListView.clear()
$lbListView.Columns.Add("UserName",150)
$lbListView.Columns.Add("# Items",70)
$lbListView.Columns.Add("MB Size(MB)",80)
$lbListView.Columns.Add("DelItems (KB)",90)

get-mailboxstatistics -Server $snServerNameDrop.SelectedItem.ToString() | ForEach-Object{
$item1  = new-object System.Windows.Forms.ListViewItem($_.DisplayName)
$item1.SubItems.Add($_.ItemCount)
$item1.SubItems.Add($_.TotalItemSize.Value.ToMB() )
$item1.SubItems.Add($_.TotalDeletedItemSize.Value.ToKB())

$lbListView.items.add($item1)
}

$form.Controls.Add($lbListView)
}

function enumFolderSizes($fsFolderIDtoSearch){
$private:enfsFilterString = "ParentID = '" + $fsFolderIDtoSearch + "'"
$private:enSubFolders = $fsTable.select($enfsFilterString)
for($fcount1 = 0;$fcount1 -le $enSubFolders.GetUpperBound(0); $fcount1++){
	 $global:fldSize  =  $global:fldSize  + $enSubFolders[$fcount1][3] 
	 $global:itemCount = $global:itemCount  + $enSubFolders[$fcount1][5] 
	 if ($enSubFolders[$fcount1][4] -ne 0){
		enumFolderSizes($enSubFolders[$fcount1][1])
	}

}
}

function BackFolder(){
$private:bfFilterString = "FolderID = '" + $global:LastFolder + "'"
$private:bfFolder = $fsTable.select($bfFilterString)
GetSubFolderSizes($bfFolder[0][2]) 

}

function GetSubFolderSizes($fiFIDToSearch){
$global:LastFolder = $fiFIDToSearch
$upButton.visible = $true
$lbFldListView.clear()
$lbFldListView.Columns.Add("Folder Name",150)
$lbFldListView.Columns.Add("# Items",80)
$lbFldListView.Columns.Add("Size(MB)",80)
$lbFldListView.Columns.Add("Has Sub",80)
$lbFldListView.Columns.Add("FID",0)
$subfsFilterString = "ParentID = '" + $fiFIDToSearch + "'"
$subFolders = $fstable.select($subfsFilterString)
for($fcount2 = 0;$fcount2 -le $subFolders.GetUpperBound(0); $fcount2++){
	$global:fldSize = $subFolders[$fcount2][3]
	$global:itemCount = $subFolders[$fcount2][5]
	if ($subFolders[$fcount2][4] -ne 0){
		enumFolderSizes($subFolders[$fcount2][1])
	}
	$item1  = new-object System.Windows.Forms.ListViewItem($subFolders[$fcount2][0])
	$item1.SubItems.Add($global:itemCount)
	$item1.SubItems.Add([math]::round(($fldsize/1mb),2))
	if ($subFolders[$fcount2][4] -ne 0){
		$item1.SubItems.Add("Yes")
	}
	else { 
		$item1.SubItems.Add("No")
	}
	$item1.SubItems.Add($subFolders[$fcount2][1])
	$lbFldListView.items.add($item1)      
}


}


function GetFolderSizes($siSIDToSearch){
$fsTable.clear()
$lbFldListView.clear()
$lbFldListView.Columns.Add("Folder Name",150)
$lbFldListView.Columns.Add("# Items",80)
$lbFldListView.Columns.Add("Size(MB)",80)
$lbFldListView.Columns.Add("Has Sub",80)
$lbFldListView.Columns.Add("FID",0)
$snServername = $snServerNameDrop.SelectedItem.ToString()
$siSIDToSearch = get-user $siSIDToSearch 

$smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" `
+ "<soap:Envelope xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`" " `
+ " xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`"" `
+ " xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`" >" `
+ "<soap:Header>" `
+ "<t:ExchangeImpersonation>" `
+ "<t:ConnectingSID>" `
+ "<t:SID>" + $siSIDToSearch.SID + "</t:SID>" `
+ "</t:ConnectingSID>" `
+ "</t:ExchangeImpersonation>" `
+ "</soap:Header>" `
+ "<soap:Body>" `
+ "<FindFolder xmlns=`"http://schemas.microsoft.com/exchange/services/2006/messages`" " `
+ "xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`" Traversal=`"Deep`"> " `
+ "<FolderShape>" `
+ "<t:BaseShape>AllProperties</t:BaseShape>" `
+ "<AdditionalProperties xmlns=""http://schemas.microsoft.com/exchange/services/2006/types"">" `
+ "<ExtendedFieldURI PropertyTag=""0x0e08"" PropertyType=""Integer"" />" `
+ "</AdditionalProperties>" `
+ "</FolderShape>" `
+ "<ParentFolderIds>" `
+ "<t:DistinguishedFolderId Id=`"root`"/>" `
+ "</ParentFolderIds>" `
+ "</FindFolder>" `
+ "</soap:Body></soap:Envelope>"

$strRootURI = "https://" + $snServername + "/ews/Exchange.asmx"
$WDRequest = [System.Net.WebRequest]::Create($strRootURI)
$WDRequest.ContentType = "text/xml"
$WDRequest.Headers.Add("Translate", "F")
$WDRequest.Method = "Post"
$WDRequest.Credentials = $cdUsrCredentials
$bytes = [System.Text.Encoding]::UTF8.GetBytes($smSoapMessage)
$WDRequest.ContentLength = $bytes.Length
$RequestStream = $WDRequest.GetRequestStream()
$RequestStream.Write($bytes, 0, $bytes.Length)
$RequestStream.Close()
$WDResponse = $WDRequest.GetResponse()
$ResponseStream = $WDResponse.GetResponseStream()
$ResponseXmlDoc = new-object System.Xml.XmlDocument
$ResponseXmlDoc.Load($ResponseStream)
$DisplayNameNodes = @($ResponseXmlDoc.getElementsByTagName("t:DisplayName"))
$ExtenedPropertyField = @($ResponseXmlDoc.getElementsByTagName("t:Value"))
$FolderIdNodes = @($ResponseXmlDoc.getElementsByTagName("t:FolderId"))
$ParentFolderIdNodes = @($ResponseXmlDoc.getElementsByTagName("t:ParentFolderId"))
$ChildFolderCountNodes = @($ResponseXmlDoc.getElementsByTagName("t:ChildFolderCount"))
$TotalItemCountNodes = @($ResponseXmlDoc.getElementsByTagName("t:TotalCount"))
for($i=0;$i -lt $DisplayNameNodes.Count;$i++){
	if ($DisplayNameNodes[$i].'#text' -eq "Top of Information Store"){$rootFolderID = $FolderIdNodes[$i].GetAttributeNode("Id").'#text'}  
	$fiFolderID = $FolderIdNodes[$i].GetAttributeNode("Id")
        $pfParentFolderID = $ParentFolderIdNodes[$i].GetAttributeNode("Id")	
	$fsTable.Rows.Add($DisplayNameNodes[$i].'#text',$fiFolderID.'#text',$pfParentFolderID.'#text',$ExtenedPropertyField[$i].'#text',$ChildFolderCountNodes[$i].'#text',$TotalItemCountNodes[$i].'#text')
}
$fsFilterString = "ParentID = '" + $rootFolderID + "'"
$rrRootFolders = $fstable.select($fsFilterString)
for($fcount = 0;$fcount -le $rrRootFolders.GetUpperBound(0); $fcount++){
if ($rrRootFolders[$fcount][0] -ne "Top of Information Store"){  
	$global:fldSize = $rrRootFolders[$fcount][3]
	$global:itemCount = $rrRootFolders[$fcount][5]
	if ($rrRootFolders[$fcount][4] -ne 0){
		enumFolderSizes($rrRootFolders[$fcount][1])
	}
	$item1  = new-object System.Windows.Forms.ListViewItem($rrRootFolders[$fcount][0])
	$item1.SubItems.Add($global:itemCount)
	$item1.SubItems.Add([math]::round(($fldsize/1mb),2))
	if ($rrRootFolders[$fcount][4] -ne 0){
		$item1.SubItems.Add("Yes")
	}
	else { 
		$item1.SubItems.Add("No")
	}
	$item1.SubItems.Add($rrRootFolders[$fcount][1])
	$lbFldListView.items.add($item1)      
}
}

$form.Controls.Add($lbFldListView)
}
$form = new-object System.Windows.Forms.form 
$global:LastFolder = ""
# Add DataTable

$Dataset = New-Object System.Data.DataSet
$fsTable = New-Object System.Data.DataTable
$fsTable.TableName = "Folder Sizes"
$fsTable.Columns.Add("DisplayName")
$fsTable.Columns.Add("FolderID")
$fsTable.Columns.Add("ParentID")
$fsTable.Columns.Add("Size)",[int])
$fsTable.Columns.Add("ChildFolderCount",[int])
$fsTable.Columns.Add("TotalCount",[int])
$Dataset.tables.add($fsTable)

# Add Server DropLable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,20) 
$snServerNamelableBox.size = new-object System.Drawing.Size(100,20) 
$snServerNamelableBox.Text = "ServerName"
$form.Controls.Add($snServerNamelableBox) 

# Add Server Drop Down
$snServerNameDrop = new-object System.Windows.Forms.ComboBox
$snServerNameDrop.Location = new-object System.Drawing.Size(130,20)
$snServerNameDrop.Size = new-object System.Drawing.Size(130,30)
get-mailboxserver | ForEach-Object{$snServerNameDrop.Items.Add($_.Name)}
$snServerNameDrop.Add_SelectedValueChanged({getMailboxSizes})  
$form.Controls.Add($snServerNameDrop)

# Add List Box to DisplayMailboxs


$lbListView = new-object System.Windows.Forms.ListView
$lbListView.Location = new-object System.Drawing.Size(10,50) 
$lbListView.size = new-object System.Drawing.Size(400,500)
$lbListView.LabelEdit = $True
$lbListView.AllowColumnReorder = $True
$lbListView.CheckBoxes = $False
$lbListView.FullRowSelect = $True
$lbListView.GridLines = $True
$lbListView.View = "Details"
$lbListView.Sorting = "Ascending"
$lbListView.add_click({GetFolderSizes($this.SelectedItems.item(0).text)}); 


# Add List Box to Display FolderSizes


$lbFldListView = new-object System.Windows.Forms.ListView
$lbFldListView.Location = new-object System.Drawing.Size(500,50) 
$lbFldListView.size = new-object System.Drawing.Size(400,500)
$lbFldListView.LabelEdit = $True
$lbFldListView.AllowColumnReorder = $True
$lbFldListView.FullRowSelect = $True
$lbFldListView.GridLines = $True
$lbFldListView.View = "Details"
$lbFldListView.Sorting = "Ascending"
$lbFldListView.add_click({GetSubFolderSizes($this.SelectedItems.item(0).subitems[4].text)}); 

# UP folder Button

$upButton = new-object System.Windows.Forms.Button
$upButton.Location = new-object System.Drawing.Size(500,19)
$upButton.Size = new-object System.Drawing.Size(120,23)
$upButton.Text = "Back Folder level"
$upButton.visible = $false
$upButton.Add_Click({BackFolder})
$form.Controls.Add($upButton)

$form.Text = "Exchange 2007 Mailbox Size Form"
$form.size = new-object System.Drawing.Size(1000,600) 
$form.autoscroll = $true
$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
