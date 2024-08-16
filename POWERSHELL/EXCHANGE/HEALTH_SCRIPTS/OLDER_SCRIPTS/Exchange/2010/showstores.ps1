[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 


function ExportGrid{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.Showhelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("Storage Group,Store Name,EDB(MB),WhiteSpace(MB),Retained(MB),Deleted(MB)")
	foreach($row in $ssTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString() + "," + $row[5].ToString()) 
	}
	$logfile.Close()
}

}

$snServerName = "servername"
$StoreFileSizes = @{ }
$StoreWhitespace = @{ }
$StoreRetainItemSize = @{ }
$StoreDeletedMailboxSize = @{ }
$hcHourCount = @{ }
$root = [ADSI]'LDAP://RootDSE' 
$cfConfigRootpath = "LDAP://" + $root.ConfigurationNamingContext.tostring()
$configRoot = [ADSI]$cfConfigRootpath 
$searcher = new-object System.DirectoryServices.DirectorySearcher($configRoot)
$searcher.Filter = '(&(objectCategory=msExchExchangeServer)(cn=' + $snServerName  + '))'
$searchres = $searcher.FindOne()
$snServerEntry = New-Object System.DirectoryServices.directoryentry 
$snServerEntry = $searchres.GetDirectoryEntry() 




$adsiServer = [ADSI]('LDAP://' + $snServerEntry.DistinguishedName)
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($adsiServer)
$dfsearcher.Filter = "(objectCategory=msExchPrivateMDB)"
$srSearchResults = $dfsearcher.FindAll()
foreach ($srSearchResult in $srSearchResults){ 
	$msMailStore = $srSearchResult.GetDirectoryEntry()
	$sgStorageGroup = $msMailStore.psbase.Parent
	if ($sgStorageGroup.msExchESEParamBaseName -ne "R00"){
		$edbfile = [WMI]("\\" + $snServerName + "\root\cimv2:CIM_DataFile.Name='" + $msMailStore.msExchEDBFile + "'")
		$StoreFileSize = [math]::round($edbfile.filesize/1048576,0)
		$exStoreName = ($sgStorageGroup.Name.ToString() + "\" + $msMailStore.Name.ToString())	
		$StoreFileSizes.Add($exStoreName,$StoreFileSize)	
	}
}
"Finshed Mailstores"
$dfsearcher.Filter = "(objectCategory=msExchPublicMDB)"
$srSearchResults = $dfsearcher.FindAll()
foreach ($srSearchResult in $srSearchResults){ 
	$msMailStore = $srSearchResult.GetDirectoryEntry()
	$sgStorageGroup = $msMailStore.psbase.Parent
	if ($sgStorageGroup.msExchESEParamBaseName -ne "R00"){
		$edbfile = [WMI]("\\" + $snServerName + "\root\cimv2:CIM_DataFile.Name='" + $msMailStore.msExchEDBFile + "'")
		$StoreFileSize = [math]::round($edbfile.filesize/1048576,0)
		$exStoreName = ($sgStorageGroup.Name.ToString() + "\" + $msMailStore.Name.ToString())	
		$StoreFileSizes.Add($exStoreName,$StoreFileSize)	
	}
}
"Finshed Public Folder Stores"
$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::UtcNow.AddDays(-3))
Get-WmiObject -computer $snServerName -query ("Select * from Win32_NTLogEvent Where Logfile='Application' and Eventcode = '1221' and TimeWritten >='" + $WmidtQueryDT + "'") | sort $_.TimeWritten -Descending |
foreach-object{
	$mbnamearray = $_.message.split("`"")
	$esEndString = $mbnamearray[2].indexof("megabytes ")-6
	if ($StoreWhitespace.Containskey($mbnamearray[1]) -eq $false){
		$StoreWhitespace.Add($mbnamearray[1],$mbnamearray[2].Substring(5,$esEndString))
	}
}
"Finished WhiteSpace"
$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::UtcNow.AddDays(-3))
Get-WmiObject -computer $snServerName -query ("Select * from Win32_NTLogEvent Where Logfile='Application' and Eventcode = '1207' and TimeWritten >='" + $WmidtQueryDT + "'") | sort $_.TimeWritten -Descending |
foreach-object{
	$mbnamearray = $_.message.split("`"")
	$enditem = $mbnamearray[2].Substring($mbnamearray[2].indexof("End:"))
	$esize = $enditem.SubString($enditem.indexof("items")+7,$enditem.indexof(" Kbytes")-($enditem.indexof("items")+7))
	if ($StoreRetainItemSize.Containskey($mbnamearray[1]) -eq $false){
		$StoreRetainItemSize.Add($mbnamearray[1],[math]::round(($esize/1024),0))
	}

}
"Finshed Retained Items"
$WmidtQueryDT = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime([DateTime]::UtcNow.AddDays(-3))
Get-WmiObject -computer $snServerName -query ("Select * from Win32_NTLogEvent Where Logfile='Application' and Eventcode = '9535' and TimeWritten >='" + $WmidtQueryDT + "'") | sort $_.TimeWritten -Descending |
foreach-object{
	$mbnamearray = $_.message.split("`"")
	$retMbs = $mbnamearray[2].Substring($mbnamearray[2].indexof("have been removed."))
	$retMbsSize = $retMbs.Substring(($retMbs.indexof("(")+1),$retMbs.indexof(")")-($retMbs.indexof("(")+1)).Replace(" KB","")
	if ($StoreDeletedMailboxSize.Containskey($mbnamearray[1]) -eq $false){
		$StoreDeletedMailboxSize.Add($mbnamearray[1],[math]::round(($retMbsSize/1024),0))
	}   	

}




"Finished Deleted Items"
$Dataset = New-Object System.Data.DataSet
$ssTable = New-Object System.Data.DataTable
$ssTable.TableName = "TrackingLogs"
$ssTable.Columns.Add("Storage Group")
$ssTable.Columns.Add("Store Name")
$ssTable.Columns.Add("EDB (MB)",[int])
$ssTable.Columns.Add("WhiteSpace (MB)",[int])
$ssTable.Columns.Add("Retained (MB)",[int])
$ssTable.Columns.Add("Deleted (MB)",[int])
$Dataset.tables.add($ssTable)

$i = 0
$TitleBlock =  "Mailboxes" + "|" +  "Public Folders"  + "|"  +  "WhiteSpace" + "|" + "Retained Items" + "|" + "Deleted Mailboxes"
$lval = 0
$mbsize = 0
$pfsize = 0
$wsSize = 0
$risize = 0
$dmsize = 0
$lval = 0
$valueBlockbc = ""
$valueBlockbc1 = ""
$valueBlockbc2 = ""
$valueBlockbc3 = ""
$valueBlockbc4 = ""
$TitleBlockbc1 = ""
$dscale = ""
$StoreFileSizes.GetEnumerator() | sort name -descending | foreach-object {
	$snames = $_.key.split("\")
	if ($StoreDeletedMailboxSize.containskey($_.key)){
		$cmdsize = ([INT]$StoreWhitespace[$_.key] + [INT]$StoreRetainItemSize[$_.key] + [INT]$StoreDeletedMailboxSize[$_.key])
		$mbsize =  + ($_.Value - $cmdsize)
		$dmsize = $dmsize + $StoreDeletedMailboxSize[$_.key]
		$ssTable.Rows.Add($snames[0],$snames[1],[math]::round(($_.Value),2),$StoreWhitespace[$_.key],$StoreRetainItemSize[$_.key],$StoreDeletedMailboxSize[$_.key])
		if ($valueBlockbc -eq ""){$valueBlockbc = $mbsize.ToString() 
					  $valueBlockbc1 = $StoreWhitespace[$_.key].ToString()
					  $valueBlockbc2 = $StoreRetainItemSize[$_.key].ToString()
					  $valueBlockbc3 = $StoreDeletedMailboxSize[$_.key].ToString() }
		else {$valueBlockbc = $valueBlockbc + "," +  $mbsize.ToString() 
 		       $valueBlockbc1 = $valueBlockbc1 + "," + $StoreWhitespace[$_.key].ToString() 
      	               $valueBlockbc2 = $valueBlockbc2 + "," + $StoreRetainItemSize[$_.key].ToString()
		       $valueBlockbc3 = $valueBlockbc3 + "," + $StoreDeletedMailboxSize[$_.key].ToString() 	
		}
	}
	else{
		$pfsize = $pfsize + ($_.Value - ([INT]$StoreWhitespace[$_.key] + [INT]$StoreRetainItemSize[$_.key]))
		$ssTable.Rows.Add($snames[0],$snames[1],[math]::round(($_.Value),2),$StoreWhitespace[$_.key],$StoreRetainItemSize[$_.key],0)
		if ($valueBlockbc -eq ""){$valueBlockbc = $pfsize.ToString() 
					  $valueBlockbc1 = $StoreWhitespace[$_.key].ToString()
					  $valueBlockbc2 = $StoreRetainItemSize[$_.key].ToString()
					  $valueBlockbc3 = "0" }
		else {$valueBlockbc = $valueBlockbc + "," +  $pfsize.ToString()  
		       $valueBlockbc1 = $valueBlockbc1 + "," + $StoreWhitespace[$_.key].ToString()
		       $valueBlockbc2 = $valueBlockbc2 + "," + $StoreRetainItemSize[$_.key].ToString()
		       $valueBlockbc3 = $valueBlockbc3 + ",0"}	
	}
	if ($TitleBlockbc1 -eq ""){$TitleBlockbc1 = $snames[1]}
	else {$TitleBlockbc1 =  $snames[1]  + "|" + $TitleBlockbc1}
	$wsSize = $wsSize + $StoreWhitespace[$_.key]
	$risize = $risize + $StoreRetainItemSize[$_.key]
	if ($dscale -eq "") {$dscale = "0," + $_.Value}
	else {$dscale = $dscale + "|0," + $_.Value} 
	if ($lval -lt $_.Value) {$lval = $_.Value}

	
	
}
$valueBlock = $mbsize.ToString() + "," + $pfsize.ToString()  + "," +  $wssize.ToString()  + "," + $risize.ToString()  + "," + $dmsize.ToString()
$csString = "http://chart.apis.google.com/chart?cht=p3&chs=370x120&chds=0," + $lval + "&chd=t:" + $valueBlock + "&chl=" + $TitleBlock + "&chco=0000ff,00ff00,ff0000,FFFFFF,000000"
$csString 
$csString1 = "http://chart.apis.google.com/chart?cht=bhs&chs=850x350&chd=t:" + $valueBlockbc + "|" + $valueBlockbc1 + "|" + $valueBlockbc2  + "|" + $valueBlockbc3 +"&chds=0," + $lval + "&chxt=x,y&chxr=" + "&chxr=0,0," + ($lval+20) + "&chxl=1:|" + $TitleBlockbc1 + "&chdl=MailStorage|WhiteSpace|RetainedItems|DeletedMailboxes&chco=4d89f9,ff0000,c6d9fd,000000"
$csString1
$form = new-object System.Windows.Forms.form 
$form.Text = "Server Store Size Report"

#add Picture box

$pbox =  new-object System.Windows.Forms.PictureBox
$pbox.Location = new-object System.Drawing.Size(615,10)
$pbox.Size = new-object System.Drawing.Size(380,150)
$pbox.ImageLocation = $csString
$form.Controls.Add($pbox)

$pbox1 =  new-object System.Windows.Forms.PictureBox
$pbox1.Location = new-object System.Drawing.Size(10,340)
$pbox1.Size = new-object System.Drawing.Size(1000,350)
$pbox1.ImageLocation = $csString1
$form.Controls.Add($pbox1)

$exButton = new-object System.Windows.Forms.Button
$exButton.Location = new-object System.Drawing.Size(10,315)
$exButton.Size = new-object System.Drawing.Size(85,20)
$exButton.Text = "Export Grid"
$exButton.Add_Click({ExportGrid})
$form.Controls.Add($exButton)

# Add DataGrid View

$dgDataGrid = new-object System.windows.forms.DataGridView
$dgDataGrid.Location = new-object System.Drawing.Size(10,10) 
$dgDataGrid.size = new-object System.Drawing.Size(600,300) 
$dgDataGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dgDataGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$form.Controls.Add($dgDataGrid)
$dgDataGrid.DataSource = $ssTable

$form.Size = new-object System.Drawing.Size(1000,600)

$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
