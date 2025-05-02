[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

function openLog{
$exFileName = new-object System.Windows.Forms.openFileDialog
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
$fnFileNamelableBox.Text = $exFileName.FileName
Populatetable
}

#  
function Populatetable{

$logTable.clear()
$asyncsumTable.clear()
$unHash.clear()
$lfLogFieldHash.clear()
$tcountline = -1
$lnum = 0
$rnma = 1000
if ($rbVeiwAllOnlyRadioButton.Checked -eq $true){$tcountline = $lnLogfileLineNum.value}
get-content $fnFileNamelableBox.Text | Where-Object -FilterScript { $_ -ilike “*Microsoft-Server-ActiveSync*” } | %{ 
	$lnum ++
	if ($lnum -eq $rnma){ Write-Progress -Activity "Read Lines" -Status $lnum
		$rnma = $rnma + 1000
	}
	$linarr = $_.split(" ")
	$lfDate = ""
	$async = 0
	$lfTime = ""
	$lfSourceIP = ""
	$lfMethod = ""
	$lfAction = ""
	$lfUserName = ""
	$devid = ""
	$lfPPCversion = ""
	if ($linarr[0].substring(0, 1) -ne "#"){
		 if ($linarr.Length -gt 0){$lfDate = $linarr[0]}
		 if ($linarr.Length -gt 1){$lfTime = $linarr[1]}
		 if ($linarr.Length -gt 8){$lfSourceIP= $linarr[9]}
		 if ($linarr.Length -gt 3){$lfMethod  = $linarr[4]}
		 if ($linarr.Length -gt 5){$lfAction = $linarr[6]}
		 if ($linarr.Length -gt 7){$lfUserName  = $linarr[8]}
		 if ($linarr.Length -gt 10){$lfPPCversion = $linarr[11]}
		 if ($linarr.Length -gt 15){$bcSent = $linarr[16]}
		 if ($linarr.Length -gt 16){$bcRecieved = $linarr[17]}
     	         if ($lfMethod -ne "Options"){
				$cmdint = $lfAction.IndexOf("&Cmd=")+5
				$aysnccmd = $lfAction.Substring($cmdint,($lfAction.IndexOfAny("&",$cmdint)-$cmdint))
		 }
		else {$aysnccmd = "Options"}
		$ldLogDatecombine = $lfDate + " " + $lfTime
		if ($ltimeconv.Checked -eq $true){
			$LogDate = Get-Date($ldLogDatecombine)
			$LocalLogDate = $LogDate.tolocaltime() 
			$ldLogDatecombine = $LocalLogDate.ToString("yyyy-MM-dd HH:mm:ss")
		}
		$devidint = $lfAction.indexof("&DeviceId")+10
		$devid = $lfAction.Substring($devidint,($lfAction.IndexOfAny("&",$devidint)-$devidint))
		if ($lfLogFieldHash.ContainsKey -eq $false){$lfLogFieldHash.Add($aysnccmd,$lfLogFieldHashnum)
		$lfLogFieldHashnum = $lfLogFieldHashnum + 1 }
		if ($unHash.ContainsKey($devid)){
				$usrHash = $unHash[$devid]
				$usrHash["BytesSent"] = [Int]$usrHash["BytesSent"] + [Int]$bcSent
				$usrHash["BytesRecieved"] = [Int]$usrHash["BytesRecieved"] + [Int]$bcRecieved
				$usrHash["LastSeen"] = $ldLogDatecombine
				$lfPPCversion = $usrHash["MobileVersion"]
				if ($usrHash.ContainsKey($aysnccmd)){
					$usrHash[$aysnccmd] = $usrHash[$aysnccmd] +1
				}	
				else{
					$usrHash.Add($aysnccmd,1)
				}
			}
		else{
				$lfPPCversion =  getppcversion($lfPPCversion)
				$usrHash = @{ }
				$usrHash.Add("UserName",$lfUserName)
				$usrHash.Add("BytesSent",[Int]$bcSent)
				$usrHash.Add("BytesRecieved",[Int]$bcRecieved)
				$usrHash.Add("MobileVersion",$lfPPCversion)
				$usrHash.Add("LastSeen",$ldLogDatecombine)
				$usrHash.Add($aysnccmd,1)
				$unHash.Add($devid,$usrHash)
			}
	        $logTable.Rows.Add($ldLogDatecombine,$lfSourceIP,$lfMethod,$lfAction,$aysnccmd,$devid,$lfUserName,$lfPPCversion,$bcSent,$bcRecieved)
 
	}
}


	foreach($devkey in $unHash.keys){
		$devid = $devkey
		$devhash = $unHash[$devkey]
		$userName = $devhash["UserName"]
		$MobileVersion = $devhash["MobileVersion"]
		$LastSeen = $devhash["LastSeen"]
		$pings = 0
		if ($devhash.Contains("Ping")){
			$pings = $devhash["Ping"]
		}
		$Sync = 0
		if ($devhash.Contains("Sync")){
			$Sync = $devhash["Sync"]
		}
		$FolderSync = 0
		if ($devhash.Contains("FolderSync")){
			$FolderSync = $devhash["FolderSync"]
		}
		$SendMail = 0
		if ($devhash.Contains("SendMail")){
			$SendMail = $devhash["SendMail"]
		}
		$SmartReply = 0
		if ($devhash.Contains("SmartReply")){
			$SmartReply = $devhash["SmartReply"]
		}
		$SmartForward = 0
		if ($devhash.Contains("SmartForward")){
			$SmartForward = $devhash["SmartForward"]
		}
		$gattach = 0
		if ($devhash.Contains("GetAttachment")){
			$gattach= $devhash["GetAttachment"]
		}
		$bsent = 0
		if ($devhash.Contains("BytesSent")){
			$bsent = [math]::round($devhash["BytesSent"]/1024,2)
		}
		$brecv = 0
		if ($devhash.Contains("BytesSent")){
			$brecv = [math]::round($devhash["BytesRecieved"]/1024,2)
		}
		$asyncsumTable.Rows.Add($devid,$userName,$pings,$Sync,$FolderSync,$SendMail,$SmartReply,$SmartForward,$gattach,$MobileVersion,$bsent,$brecv,$LastSeen)
		
	}

if ($SsumBox.Checked -eq $true){ 	
	$dgDataGrid.DataSource = $asyncsumTable
}
else {
	$dgDataGrid.DataSource = $logTable
}
write-progress "complete" "100" -completed
}

function getppcversion($ppcVersion){
	$ppcret = $ppcVersion
	$verarry = $ppcVersion.split("/")
	$versegarry = $verarry[1].split(".")
	if ($verarry[0] -eq "Microsoft-PocketPC"){
		if ($versegarry[0] -eq "3"){
			$ppcret = "PocketPC 2003"
		}
		if ($versegarry[0] -eq "2"){
			$ppcret = "PocketPC 2002"
		}
	}
	else {
		if ($versegarry[0] -eq "5"){
			if ($versegarry[1] -eq "1"){
				if ([Int]$versegarry[2] -ge 2300) {$ppcret = "SmartPhone 2003"}
				else {$ppcret = "Mobile 5"}
		
			}
			if ($versegarry[1] -eq "2"){
				$ppcret = "Mobile 6"
			}
	
		}
	}
	if ($ppcVersion -eq "Microsoft-PocketPC/3.0") {}
	return $ppcret
}


$unHash = @{ }
$lfLogFieldHash = @{ }
$lfLogFieldHashnum = 2
$form = new-object System.Windows.Forms.form 
$form.Text = "ActiveSync Log Tool"
$Dataset = New-Object System.Data.DataSet
$logTable = New-Object System.Data.DataTable
$logTable.TableName = "ActiveSyncLogs"
$logTable.Columns.Add("Time");
$logTable.Columns.Add("SourceIPAddress");
$logTable.Columns.Add("Method");
$logTable.Columns.Add("Action");
$logTable.Columns.Add("ActiveSyncCommand");
$logTable.Columns.Add("DeviceID");
$logTable.Columns.Add("UserName");
$logTable.Columns.Add("PPCVersion");
$logTable.Columns.Add("Sent");
$logTable.Columns.Add("Recieved");

$asyncsumTable = New-Object System.Data.DataTable
$asyncsumTable.TableName = "ActiveSyncSummary"
$asyncsumTable.Columns.Add("DeviceID")
$asyncsumTable.Columns.Add("UserName")
$asyncsumTable.Columns.Add("Ping")
$asyncsumTable.Columns.Add("Sync")
$asyncsumTable.Columns.Add("FolderSync")
$asyncsumTable.Columns.Add("SendMail")
$asyncsumTable.Columns.Add("SmartReply")
$asyncsumTable.Columns.Add("SmartForward")
$asyncsumTable.Columns.Add("GetAttachment")
$asyncsumTable.Columns.Add("MobileVersion")
$asyncsumTable.Columns.Add("BytesSent(KB)")
$asyncsumTable.Columns.Add("BytesRecieved(KB)")
$asyncsumTable.Columns.Add("LastSeen")





# Content
$cmClickMenu = new-object System.Windows.Forms.ContextMenuStrip
$cmClickMenu.Items.add("test122")

# Add Open Log file Button

$olButton = new-object System.Windows.Forms.Button
$olButton.Location = new-object System.Drawing.Size(20,19)
$olButton.Size = new-object System.Drawing.Size(75,23)
$olButton.Text = "Select file"
$olButton.Add_Click({openLog})
$form.Controls.Add($olButton)



# Add FileName Lable
$fnFileNamelableBox = new-object System.Windows.Forms.Label
$fnFileNamelableBox.Location = new-object System.Drawing.Size(110,25)
$fnFileNamelableBox.forecolor = "MenuHighlight"
$fnFileNamelableBox.size = new-object System.Drawing.Size(200,20) 
$form.Controls.Add($fnFileNamelableBox) 

# Add Refresh Log file Button

$refreshButton = new-object System.Windows.Forms.Button
$refreshButton.Location = new-object System.Drawing.Size(390,29)
$refreshButton.Size = new-object System.Drawing.Size(75,23)
$refreshButton.Text = "Refresh"
$refreshButton.Add_Click({Populatetable})
$form.Controls.Add($refreshButton)

# Add Veiw RadioButtons
$rbVeiwAllRadioButton = new-object System.Windows.Forms.RadioButton
$rbVeiwAllRadioButton.Location = new-object System.Drawing.Size(20,52)
$rbVeiwAllRadioButton.size = new-object System.Drawing.Size(69,17) 
$rbVeiwAllRadioButton.Checked = $true
$rbVeiwAllRadioButton.Text = "View All"
$rbVeiwAllRadioButton.Add_Click({if ($rbVeiwAllRadioButton.Checked -eq $true){$lnLogfileLineNum.Enabled = $false}})
$form.Controls.Add($rbVeiwAllRadioButton) 

$rbVeiwAllOnlyRadioButton = new-object System.Windows.Forms.RadioButton
$rbVeiwAllOnlyRadioButton.Location = new-object System.Drawing.Size(110,52)
$rbVeiwAllOnlyRadioButton.size = new-object System.Drawing.Size(89,17) 
$rbVeiwAllOnlyRadioButton.Text = "View Only #"
$rbVeiwAllOnlyRadioButton.Add_Click({if ($rbVeiwAllOnlyRadioButton.Checked -eq $true){$lnLogfileLineNum.Enabled = $true}})
$form.Controls.Add($rbVeiwAllOnlyRadioButton) 

# Local Time Converstion CheckBox

$ltimeconv =  new-object System.Windows.Forms.CheckBox
$ltimeconv.Location = new-object System.Drawing.Size(310,7)
$ltimeconv.Size = new-object System.Drawing.Size(150,20)
$ltimeconv.Checked = $true
$ltimeconv.Text = "Convert to Local Time"
$form.Controls.Add($ltimeconv)

# Show Summary CheckBox

$SsumBox =  new-object System.Windows.Forms.CheckBox
$SsumBox.Location = new-object System.Drawing.Size(510,7)
$SsumBox.Size = new-object System.Drawing.Size(200,20)
$SsumBox.Checked = $true
$SsumBox.Add_Click({if ($SsumBox.Checked -eq $true){$dgDataGrid.DataSource = $asyncsumTable}
		    else {$dgDataGrid.DataSource = $logTable}})

$SsumBox.Text = "Show Summary"
$form.Controls.Add($SsumBox)


# Add Numeric log line number 
$lnLogfileLineNum =  new-object System.Windows.Forms.numericUpDown
$lnLogfileLineNum.Location = new-object System.Drawing.Size(201,52)
$lnLogfileLineNum.Size = new-object System.Drawing.Size(69,20)
$lnLogfileLineNum.Enabled = $false
$lnLogfileLineNum.Maximum = 10000000000
$form.Controls.Add($lnLogfileLineNum)

# File setting Group Box

$OfGbox =  new-object System.Windows.Forms.GroupBox
$OfGbox.Location = new-object System.Drawing.Size(12,0)
$OfGbox.Size = new-object System.Drawing.Size(464,75)
$OfGbox.Text = "Log File Settings"
$form.Controls.Add($OfGbox)



# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGrid
$dgDataGrid.AllowSorting = $True
$dgDataGrid.Location = new-object System.Drawing.Size(12,81) 
$dgDataGrid.size = new-object System.Drawing.Size(1024,750) 
$form.Controls.Add($dgDataGrid)


$form.Add_Shown({$form.Activate()})
$form.ShowDialog()


