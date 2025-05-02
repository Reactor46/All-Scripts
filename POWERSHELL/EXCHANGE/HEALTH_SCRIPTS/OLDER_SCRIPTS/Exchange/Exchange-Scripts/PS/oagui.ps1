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
$logTable.Clear()
$sumTable.Clear()
$fname = $fnFileNamelableBox.Text
$mbcombCollection = @()
$FldHash = @{}
$usHash = @{} 
$sumHash = @{} 
$fieldsline = (Get-Content $fname)[3]
$fldarray = $fieldsline.Split(" ")
$fnum = -1
foreach ($fld in $fldarray){
	$FldHash.add($fld,$fnum)
	$fnum++
}

get-content $fname | Where-Object -FilterScript { $_ -ilike “*MSRPC*” } | %{ 
	$lnum ++
	if ($lnum -eq $rnma){ Write-Progress -Activity "Read Lines" -Status $lnum
		$rnma = $rnma + 1000
	}
	$linarr = $_.split(" ")
	$uid = $linarr[$FldHash["cs-username"]] + $linarr[$FldHash["c-ip"]]
	if ($linarr[$FldHash["cs-username"]].length -gt 2){
		if ($usHash.Containskey($uid) -eq $false){
			if ($linarr.Length -gt 0){$lfDate = $linarr[$FldHash["date"]]}
			if ($linarr.Length -gt 1){$lfTime = $linarr[$FldHash["time"]]}
			$ldLogDatecombine = $lfDate + " " + $lfTime
			if ($ltimeconv.Checked -eq $true){
				$LogDate = Get-Date($ldLogDatecombine)
				$LocalLogDate = $LogDate.tolocaltime() 
				$ldLogDatecombine = $LocalLogDate.ToString("yyyy-MM-dd HH:mm:ss")
			}
			$usrobj = "" | select UserName,IpAddress,LogonTime,LogOffTime,Duration,NumberofRequests,BytesSent,BytesRecieved
			$usrobj.UserName = $linarr[$FldHash["cs-username"]]
			$usrobj.IpAddress = $linarr[$FldHash["c-ip"]]
			$usrobj.LogonTime = $ldLogDatecombine
			$usrobj.LogOffTime = $ldLogDatecombine
			$usrobj.Duration = 0
			$usrobj.NumberofRequests = 0
			$usrobj.BytesSent = $linarr[$FldHash["sc-bytes"]]
			$usrobj.BytesRecieved = $linarr[$FldHash["cs-bytes"]]
			$usHash.add($uid,$usrobj)
			
		}
		else{
			if ($linarr.Length -gt 0){$lfDate = $linarr[$FldHash["date"]]}
			if ($linarr.Length -gt 1){$lfTime = $linarr[$FldHash["time"]]}
			$ldLogDatecombine = $lfDate + " " + $lfTime
			if ($ltimeconv.Checked -eq $true){
				$LogDate = Get-Date($ldLogDatecombine)
				$LocalLogDate = $LogDate.tolocaltime() 
				$ldLogDatecombine = $LocalLogDate.ToString("yyyy-MM-dd HH:mm:ss")
			}
			$duration = New-TimeSpan $usHash[$uid].LogOffTime $ldLogDatecombine
			if ([INT]$duration.Totalminutes -gt 30){
				$mbcombCollection += $usHash[$uid]
				$usHash.remove($uid)
				$usrobj = "" | select UserName,IpAddress,LogonTime,LogOffTime,Duration,NumberofRequests,BytesSent,BytesRecieved
				$usrobj.UserName = $linarr[$FldHash["cs-username"]]
				$usrobj.IpAddress = $linarr[$FldHash["c-ip"]]
				$usrobj.LogonTime = $ldLogDatecombine
				$usrobj.LogOffTime = $ldLogDatecombine
				$usrobj.BytesSent = $linarr[$FldHash["sc-bytes"]]
				$usrobj.BytesRecieved = $linarr[$FldHash["cs-bytes"]]
				$usrobj.Duration = 0
				$usrobj.NumberofRequests = 0
				$usHash.add($uid,$usrobj)			
			}
			else{
				$usHash[$uid].LogOffTime = $ldLogDatecombine
				$lgduration = New-TimeSpan $usHash[$uid].LogonTime $ldLogDatecombine
				$usHash[$uid].Duration = [INT]$lgduration.Totalminutes 
				$usrobj.NumberofRequests = [INT]$usrobj.NumberofRequests + 1
				$usrobj.BytesSent = [INT]$usrobj.BytesSent + [INT]$linarr[$FldHash["sc-bytes"]]
				$usrobj.BytesRecieved = [INT]$usrobj.BytesRecieved + [INT]$linarr[$FldHash["cs-bytes"]]
			}

		
		}
	}
}

$usHash.GetEnumerator() | sort LogonTime -descending | foreach-object {
	$logTable.Rows.Add($_.value.UserName,$_.value.IpAddress,$_.value.LogonTime,$_.value.LogOffTime,$_.value.Duration,$_.value.NumberofRequests,$_.value.BytesSent,$_.value.BytesRecieved)
	if($sumHash.contains($_.value.UserName)){
		$sumHash[$_.value.UserName].NumberofLogons = $sumHash[$_.value.UserName].NumberofLogons + 1
		$sumHash[$_.value.UserName].Duration = [INT]$sumHash[$_.value.UserName].Duration + [INT]$_.value.Duration
		$sumHash[$_.value.UserName].BytesSent = [INT]$sumHash[$_.value.UserName].BytesSent + [INT]$_.value.BytesSent
		$sumHash[$_.value.UserName].BytesRecieved = [INT]$sumHash[$_.value.UserName].BytesRecieved + [INT]$_.value.BytesRecieved
	}
	else{
		$usrobj = "" | select UserName,NumberofLogons,Duration,BytesSent,BytesRecieved
		$usrobj.UserName = $_.value.UserName
		$usrobj.NumberofLogons = 1 
		$usrobj.Duration = $_.value.Duration
		$usrobj.BytesSent = $_.value.BytesSent
		$usrobj.BytesRecieved = $_.value.BytesRecieved	
		$sumHash.add($_.value.UserName,$usrobj)
	}
}
$sumHash.GetEnumerator() | sort LogonTime -descending | foreach-object {
	$sumTable.Rows.Add($_.value.UserName,$_.value.NumberofLogons,$_.value.Duration,$_.value.BytesSent,$_.value.BytesRecieved)
}
$dgDataGrid.DataSource = $sumTable
}

$Dataset = New-Object System.Data.DataSet
$logTable = New-Object System.Data.DataTable
$logTable.TableName = "RCPLogons"
$logTable.Columns.Add("UserName");
$logTable.Columns.Add("IpAddress");
$logTable.Columns.Add("LogonTime");
$logTable.Columns.Add("LogOffTime");
$logTable.Columns.Add("Duration",[INT64]);
$logTable.Columns.Add("NumberofRequests",[INT64]);
$logTable.Columns.Add("Sent",[INT64]);
$logTable.Columns.Add("Recieved",[INT64]);

$sumTable = New-Object System.Data.DataTable
$sumTable.TableName = "RpcSummary"
$sumTable.Columns.Add("UserName");
$sumTable.Columns.Add("NumberofLogons",[INT64]);
$sumTable.Columns.Add("Duration",[INT64]);
$sumTable.Columns.Add("Sent",[INT64]);
$sumTable.Columns.Add("Recieved",[INT64]);

$form = new-object System.Windows.Forms.form 
$form.Text = "Outlook Anywhere Log Tool"


# Add DataGrid View

# Local Time Converstion CheckBox

$ltimeconv =  new-object System.Windows.Forms.CheckBox
$ltimeconv.Location = new-object System.Drawing.Size(310,7)
$ltimeconv.Size = new-object System.Drawing.Size(150,20)
$ltimeconv.Checked = $true
$ltimeconv.Text = "Convert to Local Time"
$form.Controls.Add($ltimeconv)

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

# Show Summary CheckBox

$SsumBox =  new-object System.Windows.Forms.CheckBox
$SsumBox.Location = new-object System.Drawing.Size(510,7)
$SsumBox.Size = new-object System.Drawing.Size(200,20)
$SsumBox.Checked = $true
$SsumBox.Add_Click({if ($SsumBox.Checked -eq $true){$dgDataGrid.DataSource = $sumTable}
		    else {$dgDataGrid.DataSource = $logTable}})

$SsumBox.Text = "Show Summary"
$form.Controls.Add($SsumBox)

$dgDataGrid = new-object System.windows.forms.DataGrid
$dgDataGrid.AllowSorting = $True
$dgDataGrid.Location = new-object System.Drawing.Size(12,81) 
$dgDataGrid.size = new-object System.Drawing.Size(1024,750) 
$form.Controls.Add($dgDataGrid)



$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()