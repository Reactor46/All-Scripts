[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

function getBusyStatus($bstat){
switch($bstat)
  {
    0 {$bret = "Free"}
    1 {$bret = "Tentative"}
    2 {$bret = "Busy"}
    3 {$bret = "Out of Office"}
  }
return $bret
}

function QueryMailbox($mbURI){
$datetimetoquery = get-date
write-host $mbURI
$wdWebDAVQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" " `
+ " xmlns:b=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"">" `
+ "<D:sql>SELECT  ""DAV:displayname"",  ""urn:schemas:httpmail:subject"", ""DAV:getlastmodified"", "`
+ " ""http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8205"" As BusyStatus, "`
+ " ""http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x825E"" As NewClients, "`
+ " ""urn:schemas:httpmail:fromname"",  ""urn:schemas:calendar:dtstart"", ""urn:schemas:calendar:dtend"", " `
+ " ""urn:schemas:calendar:instancetype"", ""urn:schemas:calendar:location"" FROM scope('shallow traversal of """`
+  $mbURI + """') Where ""DAV:ishidden"" = False AND ""DAV:contentclass"" " `
+ "= 'urn:content-classes:appointment' AND " `
+ " NOT ""urn:schemas:calendar:instancetype"" = 1 AND" `
+ """urn:schemas:calendar:dtstart"" &lt;= CAST(""" + $dpTimeFrom1.Value.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ")  + """ as 'dateTime') AND " `
+ """urn:schemas:calendar:dtend"" &gt;= CAST(""" + $dpTimeFrom.Value.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ") + """ as 'dateTime')</D:sql></D:searchrequest>"

$WDRequest = [System.Net.WebRequest]::Create($mbURI)
$WDRequest.ContentType = "text/xml"
$WDRequest.Headers.Add("Translate", "F")
$WDRequest.Method = "SEARCH"
$WDRequest.UseDefaultCredentials = $True
$bytes = [System.Text.Encoding]::UTF8.GetBytes($wdWebDAVQuery)
$WDRequest.ContentLength = $bytes.Length
$RequestStream = $WDRequest.GetRequestStream()
$RequestStream.Write($bytes, 0, $bytes.Length)
$RequestStream.Close()
$WDResponse = $WDRequest.GetResponse()
$ResponseStream = $WDResponse.GetResponseStream()
$ResponseXmlDoc = new-object System.Xml.XmlDocument
$ResponseXmlDoc.Load($ResponseStream)
$subjectnodes = @($ResponseXmlDoc.getElementsByTagName("d:subject"))
$LastModifiedNodes = @($ResponseXmlDoc.getElementsByTagName("a:getlastmodified"))
$Organizer = @($ResponseXmlDoc.getElementsByTagName("d:fromname"))
$dfDateFrom = @($ResponseXmlDoc.getElementsByTagName("e:dtstart"))
$dtDateTo = @($ResponseXmlDoc.getElementsByTagName("e:dtend"))
$location = @($ResponseXmlDoc.getElementsByTagName("e:location"))
$busyStatusnodes = @($ResponseXmlDoc.getElementsByTagName("BusyStatus"))
$newclientsnodes = @($ResponseXmlDoc.getElementsByTagName("NewClients"))
for($i=0;$i -lt $subjectnodes.Count;$i++){
	$startTime = [System.Convert]::ToDateTime($dfDateFrom[$i].'#text'.ToString())
	$EndTime = [System.Convert]::ToDateTime($dtDateTo[$i].'#text'.ToString())
	$LastModified = [System.Convert]::ToDateTime($LastModifiedNodes[$i].'#text'.ToString())
	$mbhash[$uaUserAccount.displayname.toString()] = $mbhash[$uaUserAccount.displayname.toString()] + 1
	$BusyStatus = getBusyStatus($busyStatusnodes[$i].'#text')
	if ($newclientsnodes[$i].'#text' -eq $null){
		$newclients = "False"	
	}
	else{
		$newclients = "True"
	}
	$fsTable.Rows.Add($uaUserAccount.displayname.toString(),$subjectnodes[$i].'#text',$startTime,$EndTime,$location[$i].'#text',$Organizer[$i].'#text',$LastModified,$BusyStatus,$newclients)
}

}

function DisplayMailboxes(){
$filters = 0
$sqlFilter.clear()
if ($mdFilterCheck.Checked -eq $true){$filters = 1}
if ($fbFilterCheck.Checked -eq $true){$filters = 1}
if ($ncFilterCheck.Checked -eq $true){$filters = 1}
if ($filters -eq 1){
	$mbhashfilter.clear()
	$mdDatetoCheck = New-Object System.DateTime $dpcdTime.value.year,$dpcdTime.value.month,$dpcdTime.value.day,$dpcdTime1.value.hour,$dpcdTime1.value.minute,$dpcdTime1.value.second
	if ($mdFilterCheck.Checked -eq $true){
		if ($rbafterRadioButton.Checked -eq $true){
			$sqlFilter.add(1,"ModifiedDate >= #" + $mdDatetoCheck  + "#")}
		else {
			$sqlFilter.add(1,"ModifiedDate <= #" + $mdDatetoCheck  + "#")}
	}
	if ($fbFilterCheck.Checked -eq $true){
		if ($fbDrop.SelectedItem.ToString() -eq "Free"){
			$sqlFilter.add(2,"BusyStatus = 'Free'")}
		else {
			$sqlFilter.add(2,"BusyStatus <> 'Free'")}
	}
	if ($ncFilterCheck.Checked -eq $true){
			if ($rbncRadioButton.Checked -eq $true){
			$sqlFilter.add(3,"NewClients = 'True'")}
		else {
			$sqlFilter.add(3,"NewClients = 'False'")}
	}
	$useand = 0 
	foreach($sqlstatment in $sqlFilter.keys){
		if ($useand -eq 0){
			$subfsFilterString = $sqlFilter[$sqlstatment]
			$useand = 1}
		else {
			$subfsFilterString = $subfsFilterString + " and " + $sqlFilter[$sqlstatment]
		     }
	}
	$ftappointments = $fstable.select($subfsFilterString)
	foreach ($mbox in $mbhash.keys){
		$mbhashfilter.add($mbox,0)
	}
	for($fcount2 = 0;$fcount2 -le $ftappointments.GetUpperBound(0); $fcount2++){
		$mbhashfilter[$ftappointments[$fcount2][0]] = $mbhashfilter[$ftappointments[$fcount2][0]] + 1
	}
	foreach ($mbox in $mbhashfilter.keys){
		$item1  = new-object System.Windows.Forms.ListViewItem($mbox)
		$item1.SubItems.Add($mbhashfilter[$mbox])
		$lbListView.items.add($item1)
	}
}
else{
	foreach ($mbox in $mbhash.keys){
		$item1  = new-object System.Windows.Forms.ListViewItem($mbox)
		$item1.SubItems.Add($mbhash[$mbox])
		$lbListView.items.add($item1)
	}
}
$form.Controls.Add($lbListView)
$form.Controls.Add($lbAptListView)
}

function GetUsers(){
$mbhash.clear()
$fsTable.clear()
$lbListView.clear()
$lbAptListView.clear()
$lbListView.Columns.Add("Mailbox",150)
$lbListView.Columns.Add("# Appointments",100)
$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
$configRoot = [ADSI]$cfConfigRootpath 
$dfRoot = [ADSI]$dfDefaultRootPath
$searcher = new-object System.DirectoryServices.DirectorySearcher($configRoot)
$searcher.Filter = '(&(objectCategory=msExchExchangeServer)(cn=' + $snServerNameDrop.SelectedItem.ToString()  + '))'
$searcher.PropertiesToLoad.Add("cn")
$searcher.PropertiesToLoad.Add("gatewayProxy")
$searcher.PropertiesToLoad.Add("legacyExchangeDN")
$searcher1 = $searcher.FindAll()
foreach ($server in $searcher1){ 
	$snServerEntry = New-Object System.DirectoryServices.directoryentry 
        $snServerEntry = $server.GetDirectoryEntry() 
	$snServerName = $snServerEntry.cn
	$snExchangeDN = $snServerEntry.legacyExchangeDN
}
$searcher.Filter = '(&(objectCategory=msExchRecipientPolicy)(cn=Default Policy))'
$searcher1 = $searcher.FindAll()
foreach ($recppolicies in $searcher1){ 
     	$gwaddrrs = New-Object System.DirectoryServices.directoryentry 
        $gwaddrrs = $recppolicies.GetDirectoryEntry() 
	foreach ($address in $gwaddrrs.gatewayProxy){
		if($address.Substring(0,5) -ceq "SMTP:"){$dfAddress = $address.Replace("SMTP:@","")}
	}	
	
}
$arMbRoot = "http://" + $snServerNameDrop.SelectedItem.ToString() + "/exadmin/admin/" + $dfAddress + "/mbx/"
$gfGALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)" `
+ "(objectClass=user)(msExchHomeServerName=" + $snExchangeDN + ")) )))))"
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
$dfsearcher.Filter = $gfGALQueryFilter
$searcher2 = $dfsearcher.FindAll()
foreach ($uaUsers in $searcher2){ 
     	$uaUserAccount = New-Object System.DirectoryServices.directoryentry 
        $uaUserAccount = $uaUsers.GetDirectoryEntry() 
	foreach ($address in $uaUserAccount.proxyaddresses){
		if($address.Substring(0,5) -ceq "SMTP:"){$uaAddress = $address.Replace("SMTP:","")}
	}
	$mbhash.add($uaUserAccount.displayname.toString(),0)
	QueryMailbox($arMbRoot + $uaAddress + "/calendar") 
}
DisplayMailboxes
}

function ShowAppts($mbMailbox){

$lbAptListView.clear()
if ($mdFilterCheck.Checked -eq $true){
	$lbAptListView.Columns.Add("Modified Time",135)
	$lbAptListView.Columns.Add("Start Time",135)}
else{
	$lbAptListView.Columns.Add("Start Time",135)
	$lbAptListView.Columns.Add("End Time",135)
}
$lbAptListView.Columns.Add("Subject",200)
$lbAptListView.Columns.Add("Location",100)
if($fbFilterCheck.Checked -eq $true){
	$lbAptListView.Columns.Add("Oganizer",80)
	$lbAptListView.Columns.Add("FreeBusy",40)
}
else
{
	$lbAptListView.Columns.Add("Oganizer",120)
}

$filters = 0
$sqlFilter.clear()
if ($mdFilterCheck.Checked -eq $true){$filters = 1}
if ($fbFilterCheck.Checked -eq $true){$filters = 1}
if ($ncFilterCheck.Checked -eq $true){$filters = 1}
if ($filters -eq 1){
	$mdDatetoCheck = New-Object System.DateTime $dpcdTime.value.year,$dpcdTime.value.month,$dpcdTime.value.day,$dpcdTime1.value.hour,$dpcdTime1.value.minute,$dpcdTime1.value.second
	if ($mdFilterCheck.Checked -eq $true){
		if ($rbafterRadioButton.Checked -eq $true){
			$sqlFilter.add(1,"ModifiedDate >= #" + $mdDatetoCheck  + "#")}
		else {
			$sqlFilter.add(1,"ModifiedDate <= #" + $mdDatetoCheck  + "#")}
	}
	if ($fbFilterCheck.Checked -eq $true){
		if ($fbDrop.SelectedItem.ToString() -eq "Free"){
			$sqlFilter.add(2,"BusyStatus = 'Free'")}
		else {
			$sqlFilter.add(2,"BusyStatus <> 'Free'")}
	}
	if ($ncFilterCheck.Checked -eq $true){
			if ($rbncRadioButton.Checked -eq $true){
			$sqlFilter.add(3,"NewClients = 'True'")}
		else {
			$sqlFilter.add(3,"NewClients = 'False'")}
	}
	$subfsFilterString = "Mailbox = '" + $mbMailbox + "'"
	foreach($sqlstatment in $sqlFilter.keys){
			$subfsFilterString = $subfsFilterString + " and " + $sqlFilter[$sqlstatment]
	}
}

else{
	$subfsFilterString = "Mailbox = '" + $mbMailbox + "'"
}
$apAppointments = $fstable.select($subfsFilterString)
for($fcount2 = 0;$fcount2 -le $apAppointments.GetUpperBound(0); $fcount2++){
	if ($mdFilterCheck.Checked -eq $true){
		$item1  = new-object System.Windows.Forms.ListViewItem($apAppointments[$fcount2][6])
		$item1.SubItems.Add($apAppointments[$fcount2][2])
	}
	else{
		$item1  = new-object System.Windows.Forms.ListViewItem($apAppointments[$fcount2][2])
		$item1.SubItems.Add($apAppointments[$fcount2][3])	
	}
	$item1.SubItems.Add($apAppointments[$fcount2][1])
	$item1.SubItems.Add($apAppointments[$fcount2][4])
	$item1.SubItems.Add($apAppointments[$fcount2][5])
	if($fbFilterCheck.Checked -eq $true){
		$item1.SubItems.Add($apAppointments[$fcount2][7])
	}	
	$lbAptListView.items.add($item1)      
}
$form.Controls.Add($lbAptListView)
}

function cdFilterButtons(){
if ($mdFilterCheck.Checked -eq $true){
	$rbbeforeRadioButton.enabled = $true
	$rbafterRadioButton.Enabled = $true
	$dpcdTime.Enabled = $true
	$dpcdTime1.Enabled = $true}
else{
	$rbbeforeRadioButton.enabled = $false
	$rbafterRadioButton.Enabled = $false
	$dpcdTime.Enabled = $false
	$dpcdTime1.Enabled = $false}

}

function FilterData(){
$lbListView.clear()
$lbAptListView.clear()
$lbListView.Columns.Add("Mailbox",150)
$lbListView.Columns.Add("# Appointments",100)
DisplayMailboxes

}
function NcFilter(){
if ($ncFilterCheck.Checked -eq $true){
	$rbncRadioButton.enabled = $true
	$rbnonncRadioButton.Enabled = $true}
else{
	$rbncRadioButton.enabled = $false
	$rbnonncRadioButton.Enabled = $false
}
}
$form = new-object System.Windows.Forms.form 

# Add DataTable

$Dataset = New-Object System.Data.DataSet
$fsTable = New-Object System.Data.DataTable
$mbhash = @{ }
$mbhashfilter = @{ }
$sqlFilter = @{ }
$fsTable.TableName = "Appoinments"
$fstable.Columns.Add("Mailbox")
$fsTable.Columns.Add("Subject")
$fsTable.Columns.Add("StartDate")
$fsTable.Columns.Add("EndDate")
$fsTable.Columns.Add("Location")
$fsTable.Columns.Add("Oganizer")
$fsTable.Columns.Add("ModifiedDate",[datetime])
$fsTable.Columns.Add("BusyStatus")
$fsTable.Columns.Add("NewClients")
$Dataset.tables.add($fsTable)

# Add Server DropLable
$snServerNamelableBox = new-object System.Windows.Forms.Label
$snServerNamelableBox.Location = new-object System.Drawing.Size(10,60) 
$snServerNamelableBox.size = new-object System.Drawing.Size(100,20) 
$snServerNamelableBox.Text = "ServerName"
$form.Controls.Add($snServerNamelableBox) 

# Add DateTimePickers Button

$dpDatePickerFromlableBox = new-object System.Windows.Forms.Label
$dpDatePickerFromlableBox.Location = new-object System.Drawing.Size(10,20) 
$dpDatePickerFromlableBox.size = new-object System.Drawing.Size(90,20) 
$dpDatePickerFromlableBox.Text = "Start-Date"
$form.Controls.Add($dpDatePickerFromlableBox) 

$dpTimeFrom = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom.Location = new-object System.Drawing.Size(130,20)
$dpTimeFrom.Value =  New-Object System.DateTime 2007,03,11,00,00,00
$dpTimeFrom.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom)

$dpDatePickerFromlableBox1 = new-object System.Windows.Forms.Label
$dpDatePickerFromlableBox1.Location = new-object System.Drawing.Size(10,40) 
$dpDatePickerFromlableBox1.size = new-object System.Drawing.Size(90,20) 
$dpDatePickerFromlableBox1.Text = "End-Date"
$form.Controls.Add($dpDatePickerFromlableBox1) 

$dpTimeFrom1 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom1.Location = new-object System.Drawing.Size(130,40)
$dpTimeFrom1.Value =  New-Object System.DateTime 2007,04,01,00,00,00
$dpTimeFrom1.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom1)



# Add Server Drop Down
$snServerNameDrop = new-object System.Windows.Forms.ComboBox
$snServerNameDrop.Location = new-object System.Drawing.Size(130,60)
$snServerNameDrop.Size = new-object System.Drawing.Size(130,30)
$root = [ADSI]'LDAP://RootDSE' 
$cfConfigRootpath = "LDAP://" + $root.ConfigurationNamingContext.tostring()
$configRoot = [ADSI]$cfConfigRootpath 
$searcher = new-object System.DirectoryServices.DirectorySearcher($configRoot)
$searcher.Filter = '(objectCategory=msExchExchangeServer)'
$searcher.PropertiesToLoad.Add("cn")
$searcher.PropertiesToLoad.Add("Name")
$searcher1 = $searcher.FindAll()
foreach ($server in $searcher1){ 
	$snServerNameDrop.Items.Add([String]$server.Properties.cn)
}
$form.Controls.Add($snServerNameDrop)



# Add Get-Appointments Button

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(10,90)
$exButton.Size = new-object System.Drawing.Size(135,20)
$exButton.Text = "Get-Appoitments"
$exButton.Add_Click({GetUsers})
$form.Controls.Add($exButton)

# Add List Box to DisplayMailboxs

$lbListView = new-object System.Windows.Forms.ListView
$lbListView.Location = new-object System.Drawing.Size(10,120) 
$lbListView.size = new-object System.Drawing.Size(250,500)
$lbListView.LabelEdit = $True
$lbListView.AllowColumnReorder = $True
$lbListView.CheckBoxes = $False
$lbListView.FullRowSelect = $True
$lbListView.GridLines = $True
$lbListView.View = "Details"
$lbListView.Sorting = "Ascending"
$lbListView.add_click({ShowAppts($this.SelectedItems.item(0).text)}); 

# Add Modified Time Filter
$mdFilterCheck =  new-object System.Windows.Forms.CheckBox
$mdFilterCheck.Location = new-object System.Drawing.Size(330,25)
$mdFilterCheck.Text = "Filter By Last Modified Date"
$mdFilterCheck.Size = new-object System.Drawing.Size(100,25)
$mdFilterCheck.Add_Click({cdFilterButtons})
$form.Controls.Add($mdFilterCheck)

# Add Free Busy Filter
$fbFilterCheck =  new-object System.Windows.Forms.CheckBox
$fbFilterCheck.Location = new-object System.Drawing.Size(330,60)
$fbFilterCheck.Text = "Filter By Free/Busy"
$fbFilterCheck.Size = new-object System.Drawing.Size(100,25)
$fbFilterCheck.Add_Click({if ($fbFilterCheck.Checked -eq $true){$fbDrop.Enabled = $true}
else{$fbDrop.Enabled = $false}})
$form.Controls.Add($fbFilterCheck)

# Add Free Busy DropDown
$fbDrop = new-object System.Windows.Forms.ComboBox
$fbDrop.Location = new-object System.Drawing.Size(430,60)
$fbDrop.Size = new-object System.Drawing.Size(70,30)
$fbDrop.Enabled = $false
$fbDrop.Items.Add("Free")
$fbDrop.Items.Add("Busy")
$form.Controls.Add($fbDrop)

# Add New Clients Check
$ncFilterCheck =  new-object System.Windows.Forms.CheckBox
$ncFilterCheck.Location = new-object System.Drawing.Size(520,60)
$ncFilterCheck.Text = "Filter By New Clients"
$ncFilterCheck.Size = new-object System.Drawing.Size(130,25)
$ncFilterCheck.Add_Click({NcFilter})
$form.Controls.Add($ncFilterCheck)

$Panel =  new-object System.Windows.Forms.Panel
$Panel.Location = new-object System.Drawing.Size(645,50)
$Panel.Size = new-object System.Drawing.Size(150,55)

# Add newclients RadioButtons
$rbncRadioButton = new-object System.Windows.Forms.RadioButton
$rbncRadioButton.Location = new-object System.Drawing.Size(5,10)
$rbncRadioButton.size = new-object System.Drawing.Size(100,17) 
$rbncRadioButton.Checked = $true
$rbncRadioButton.Text = "New Clients"
$rbncRadioButton.Enabled = $false
$Panel.Controls.Add($rbncRadioButton) 

# Add newclients RadioButtons
$rbnonncRadioButton = new-object System.Windows.Forms.RadioButton
$rbnonncRadioButton.Location = new-object System.Drawing.Size(5,30)
$rbnonncRadioButton.size = new-object System.Drawing.Size(130,17) 
$rbnonncRadioButton.Checked = $false
$rbnonncRadioButton.Text = "Non New Clients"
$rbnonncRadioButton.Enabled = $false
$Panel.Controls.Add($rbnonncRadioButton) 
$form.Controls.Add($panel)

# Add before RadioButtons
$rbbeforeRadioButton = new-object System.Windows.Forms.RadioButton
$rbbeforeRadioButton.Location = new-object System.Drawing.Size(430,20)
$rbbeforeRadioButton.size = new-object System.Drawing.Size(69,17) 
$rbbeforeRadioButton.Checked = $true
$rbbeforeRadioButton.Enabled = $false
$rbbeforeRadioButton.Text = "Before"
# $rbbeforeRadioButton.Add_Click({})
$form.Controls.Add($rbbeforeRadioButton) 


# Add after RadioButtons
$rbafterRadioButton = new-object System.Windows.Forms.RadioButton
$rbafterRadioButton.Location = new-object System.Drawing.Size(430,39)
$rbafterRadioButton.size = new-object System.Drawing.Size(69,17) 
$rbafterRadioButton.Checked = $false
$rbafterRadioButton.Text = "After"
$rbafterRadioButton.Enabled = $false
# $rbafterRadioButton.Add_Click({})
$form.Controls.Add($rbafterRadioButton) 

$dpcdTime = new-object System.Windows.Forms.DateTimePicker
$dpcdTime.Location = new-object System.Drawing.Size(500,30)
$dpcdTime.Size = new-object System.Drawing.Size(190,20)
$dpcdTime.Enabled = $false
$form.Controls.Add($dpcdTime)

$dpcdTime1 = new-object System.Windows.Forms.DateTimePicker
$dpcdTime1.Format = "Time"
$dpcdTime1.Enabled = $false
$dpcdTime1.value = [DateTime]::get_Now()
$dpcdTime1.ShowUpDown = $True
$dpcdTime1.Location = new-object System.Drawing.Size(690,30)
$dpcdTime1.Size = new-object System.Drawing.Size(100,20)
$form.Controls.Add($dpcdTime1)

# Add filter Button

$fiButton = new-object System.Windows.Forms.Button
$fiButton.Location = new-object System.Drawing.Size(800,30)
$fiButton.Size = new-object System.Drawing.Size(85,20)
$fiButton.Text = "Filter"
$fiButton.Add_Click({FilterData})
$form.Controls.Add($fiButton)

# Add List box for appointments
$lbAptListView = new-object System.Windows.Forms.ListView
$lbAptListView.Location = new-object System.Drawing.Size(280,120) 
$lbAptListView.size = new-object System.Drawing.Size(690,500)
$lbAptListView.LabelEdit = $True
$lbAptListView.AllowColumnReorder = $True
$lbAptListView.FullRowSelect = $True
$lbAptListView.GridLines = $True
$lbAptListView.View = "Details"
$lbAptListView.Sorting = "Ascending"

$form.Text = "Exchange Appointment Audit Form"
$form.size = new-object System.Drawing.Size(1000,800) 
$form.autoscroll = $true
$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
