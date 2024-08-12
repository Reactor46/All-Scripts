[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
$DomainHash = @{ }
$Senders = @{ }
$Recipeints = @{ }
$RecipeintsSubject = @{ }
$SendersSubject = @{ }

function GetLogs(){
$dtQueryDT = New-Object System.DateTime $dpTimeFrom.value.year,$dpTimeFrom.value.month,$dpTimeFrom.value.day,$dpTimeFrom2.value.hour,$dpTimeFrom2.value.minute,$dpTimeFrom2.value.second
$dtQueryDTf =  New-Object System.DateTime $dpTimeFrom1.value.year,$dpTimeFrom1.value.month,$dpTimeFrom1.value.day,$dpTimeFrom3.value.hour,$dpTimeFrom3.value.minute,$dpTimeFrom3.value.second
$ltTable.clear()
$Recipeints.clear()
$Senders.clear()
$DomainHash.clear()
$SendersSubject.clear()
$RecipeintsSubject.clear()
get-accepteddomain | ForEach-Object{
	if ($_.DomainType -eq "Authoritative"){
		$DomainHash.add($_.DomainName.SmtpDomain.ToString().ToLower(),1)
	}

}
get-messagetrackinglog -Server $snServerNameDrop.SelectedItem.ToString()  -Start $dtQueryDT -End $dtQueryDTf -ResultSize 10000 | ForEach-Object{
	if ($_.EventId -eq "Send" -bor $_.EventId -eq "Receive"){
	$saSndArray = $_.Sender.ToString().Split("@")
	if ($saSndArray.length -gt 0){
		if ($DomainHash.ContainsKey($saSndArray[1].ToLower())){
			if ($Senders.ContainsKey($_.Sender.ToString().ToLower()) -eq $false){
				$Senders.Add($_.Sender.ToString().ToLower(),$_.Timestamp.ToString())
				$SendersSubject.Add($_.Sender.ToString().ToLower(),$_.MessageSubject.ToString())
			}
			else{	
				if ([DateTime]::Parse($Senders[$_.Sender.ToString().ToLower()]) -lt $_.Timestamp){
					$Senders[$_.Sender.ToString().ToLower()] = $_.Timestamp.ToString()
					$SendersSubject[$_.Sender.ToString().ToLower()] = $_.MessageSubject.ToString()
				}
			}
		}
	}
	foreach ($recp in $_.recipients){
		if ($recp -ne ""){
			if ($Recipeints.ContainsKey($recp.ToLower()) -eq $false){
				$Recipeints.Add($recp.ToLower(),$_.Timestamp.ToString())
				$RecipeintsSubject.Add($recp.ToLower(),$_.MessageSubject.ToString())
			}
			else{	
				if ([DateTime]::Parse($Recipeints[$recp.ToLower()]) -lt $_.Timestamp){
					$Recipeints[$recp.ToLower()] = $_.Timestamp.ToString()
					$RecipeintsSubject[$recp.ToLower()] = $_.MessageSubject.ToString()
				}
			}		
		}
	}
	

}}
foreach ($key in $Recipeints.keys){
	$Recvtimespan = New-TimeSpan -start ($Recipeints[$key]) -end $(Get-Date)
	$user = get-user $key
	if ($user -ne $null){
		if ($Senders.ContainsKey($key)){
			$SentTimeSpan =  New-TimeSpan -start ($Senders[$key]) -end $(Get-Date)
			$ltTable.Rows.Add($user.displayName,$Senders[$key],[math]::round($SentTimeSpan.TotalMinutes,0),$Recipeints[$key],[math]::round($Recvtimespan.TotalMinutes),$RecipeintsSubject[$key],$SendersSubject[$key])}
		else {$ltTable.Rows.Add($user.displayName,"n/a","n/a",$Recipeints[$key],[math]::round($Recvtimespan.TotalMinutes,0),$RecipeintsSubject[$key],"n/a")}
}}
foreach ($key in $Senders.keys){
	$user = get-user $key
	if ($user -ne $null){
		if ($Recipeints.ContainsKey($key) -eq $false){
			$SentTimeSpan =  New-TimeSpan -start ($Senders[$key]) -end $(Get-Date)
			$ltTable.Rows.Add($user.displayName,$Senders[$key],[math]::round($SentTimeSpan.TotalMinutes,0),"n/a","n/a","n/a",$SendersSubject[$key])

	}}

}
$dgDataGrid.DataSource = $ltTable
}

$ltTable = New-Object System.Data.DataTable
$ltTable.TableName = "SentAndRecv"
$ltTable.Columns.Add("Mailbox")
$ltTable.Columns.Add("Last Mail Sent")
$ltTable.Columns.Add("Minutes since Last Mail Sent")
$ltTable.Columns.Add("Last Recieved Mail")
$ltTable.Columns.Add("Minutes since Last Mail Recieved")
$ltTable.Columns.Add("Subject of Last Mail Recieved")
$ltTable.Columns.Add("Subject of Last Mail Sent")
$form = new-object System.Windows.Forms.form 
$form.Text = "Exchange Last Sent and Recieved Form"


# Add Search Button

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(270,20)
$exButton.Size = new-object System.Drawing.Size(85,20)
$exButton.Text = "Search"
$exButton.Add_Click({GetLogs})
$form.Controls.Add($exButton)

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
get-Exchangeserver | ForEach-Object{$snServerNameDrop.Items.Add($_.Name)}
# $snServerNameDrop.Add_SelectedValueChanged({GetLogs})  
$form.Controls.Add($snServerNameDrop)
# Add DateTimePickers Button

$dpDatePickerFromlableBox = new-object System.Windows.Forms.Label
$dpDatePickerFromlableBox.Location = new-object System.Drawing.Size(10,50) 
$dpDatePickerFromlableBox.size = new-object System.Drawing.Size(90,20) 
$dpDatePickerFromlableBox.Text = "Logged Between"
$form.Controls.Add($dpDatePickerFromlableBox) 

$dpTimeFrom = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom.Location = new-object System.Drawing.Size(120,50)
$dpTimeFrom.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom)

$dpDatePickerFromlableBox1 = new-object System.Windows.Forms.Label
$dpDatePickerFromlableBox1.Location = new-object System.Drawing.Size(10,70) 
$dpDatePickerFromlableBox1.size = new-object System.Drawing.Size(50,20) 
$dpDatePickerFromlableBox1.Text = "and"
$form.Controls.Add($dpDatePickerFromlableBox1) 

$dpTimeFrom1 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom1.Location = new-object System.Drawing.Size(120,70)
$dpTimeFrom1.Size = new-object System.Drawing.Size(190,20)
$form.Controls.Add($dpTimeFrom1)

$dpTimeFrom2 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom2.Format = "Time"
$dpTimeFrom2.value = [DateTime]::get_Now().AddHours(-1)
$dpTimeFrom2.ShowUpDown = $True
$dpTimeFrom2.Location = new-object System.Drawing.Size(315,50)
$dpTimeFrom2.Size = new-object System.Drawing.Size(90,20)
$form.Controls.Add($dpTimeFrom2)

$dpTimeFrom3 = new-object System.Windows.Forms.DateTimePicker
$dpTimeFrom3.Format = "Time"
$dpTimeFrom3.ShowUpDown = $True
$dpTimeFrom3.Location = new-object System.Drawing.Size(315,70)
$dpTimeFrom3.Size = new-object System.Drawing.Size(90,20)
$form.Controls.Add($dpTimeFrom3)



# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,100) 
$dgDataGrid.size = new-object System.Drawing.Size(1000,600) 
$dgDataGrid.AutoSizeColumnsMode = "AllCells"

$form.Controls.Add($dgDataGrid)

#populate DataGrid

$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
