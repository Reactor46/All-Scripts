[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 

$servername = "servername"
$hcHourCount = @{ }
$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::UtcNow.Addhours(-6))
$WmidtQueryDTf = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::UtcNow)
$WmiNamespace = "ROOT\MicrosoftExchangev2"

$filterblock1 =  "OriginationTime <= '" + $WmidtQueryDTf + "' and OriginationTime >=  '" + $WmidtQueryDT + "' and entrytype = '1028'"
$filterblock2 =  "OriginationTime <= '" + $WmidtQueryDTf + "' and OriginationTime >=  '" + $WmidtQueryDT + "' and entrytype = '1020'"
$filter = $filterblock1 + " or " + $filterblock2
get-wmiobject -class Exchange_MessageTrackingEntry -Namespace $WmiNamespace -ComputerName $servername -filter $filter | ForEach-Object{ 
	$mdate = [System.Management.ManagementDateTimeConverter]::ToDateTime($_.OriginationTime)
	if ($hcHourCount.ContainsKey($mdate.hour)){
		$hcHourCount[$mdate.hour] = [int]$hcHourCount[$mdate.hour]  + 1
	  }
	else{
		$hcHourCount.Add($mdate.hour,1)
	}
}
$valueBlock = ""
$TitleBlock = ""
$lval = 0

$hcHourCount.GetEnumerator() | sort name -descending | foreach-object {
	if ($lval -lt $_.value){$lval = $_.value}
	if ($valueBlock -eq "") {$valueBlock = $_.value.ToString()}
	else {$valueBlock =   $valueBlock + "," + $_.value.ToString()}
	if ($TitleBlock -eq ""){$TitleBlock = $_.key.ToString() + ":00"}
	else {$TitleBlock =  $_.key.ToString() + ":00" + "|" + $TitleBlock}

}

$hcHourCount | get-member
$csString = "http://chart.apis.google.com/chart?cht=bhg&chs=400x250&chd=t:" + $valueBlock + "&chds=0," + ($lval+20)  + "&chxt=x,y&chxr=" + "&chxr=0,0," + ($lval+20) + "&chxl=1:|" + $TitleBlock + "&chco=00ff00&chtt=Message+Volume++Last+6+Hours" 
$form = new-object System.Windows.Forms.form 
$form.Text = "Last 6 Hours Graph"

#add Picture box

$pbox =  new-object System.Windows.Forms.PictureBox
$pbox.Location = new-object System.Drawing.Size(10,10)
$pbox.Size = new-object System.Drawing.Size(400,250)
$pbox.ImageLocation = $csString
$form.Controls.Add($pbox)
$form.Size = new-object System.Drawing.Size(500,350)

$form.topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
