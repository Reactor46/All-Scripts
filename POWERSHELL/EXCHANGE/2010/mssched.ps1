$serverName = $args[0]
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
$HistoryDir = "c:\mbsizehistory"

if (!(Test-Path -path $HistoryDir))
{
	New-Item $HistoryDir -type directory
}



function getMailboxSizes(){
$datetime = get-date
$fname = $script:HistoryDir + "\"
$mbcombCollection = @()
$mscombCollection = @()
$mstoresquotas.clear()
$msTable.clear()
$usrquotas.clear()
if ($serverName -eq "ALL Servers"){
	$dbsetting = get-mailboxdatabase 
}
else{
	$dbsetting = get-mailboxdatabase -server $serverName
} 
$dbsetting | ForEach-Object{
	$_.identity
	$_.ProhibitSendReceiveQuota
	if ($qtTypeDrop.SelectedItem -eq $null){
		if ($_.ProhibitSendReceiveQuota.IsUnlimited -ne $true){
			$mstoresquotas.add($_.identity,$_.ProhibitSendReceiveQuota)
		}
	}
	else{
		$soStoreObject = $_
		switch ($qtTypeDrop.SelectedItem.ToString()){
			"Warning" {
					if ($soStoreObject.IssueWarningQuota.IsUnlimited -ne $true){
						$mstoresquotas.add($soStoreObject.identity,$soStoreObject.IssueWarningQuota)
					}
				}
			"Proibit Send" {
					if ($soStoreObject.ProhibitSendQuota.IsUnlimited -ne $true){
						$mstoresquotas.add($soStoreObject.identity,$soStoreObject.ProhibitSendQuota)
					}
				}
			"Proibit Send/Recieve" {
				if ($soStoreObject.ProhibitSendReceiveQuota.IsUnlimited -ne $true){
					$mstoresquotas.add($soStoreObject.identity,$soStoreObject.ProhibitSendReceiveQuota)
				}
			}
		}
	
	}

	
}

$usrquotas = @{ }
if ($serverName -eq "ALL Servers"){
	$mailboxes = get-mailbox -ResultSize Unlimited
}
else{
	$mailboxes = get-mailbox -server $serverName -ResultSize Unlimited
} 
$mailboxes | foreach-object{
	if ($qtTypeDrop.SelectedItem -eq $null){
		if($_.ProhibitSendReceiveQuota.IsUnlimited -ne $true){
			$usrquotas.add($_.ExchangeGuid,$_.ProhibitSendReceiveQuota)
		}
	}
	else {
		$uoUserobject = $_
		switch ($qtTypeDrop.SelectedItem.ToString()){
			"Warning" {
				if($uoUserobject.IssueWarningQuota.isUnlimited -eq $false){
					$usrquotas.add($uoUserobject.ExchangeGuid,$uoUserobject.IssueWarningQuota)
				}
				}
			"Proibit Send" {
				if($uoUserobject.ProhibitSendQuota.isUnlimited -eq $false){
					$usrquotas.add($uoUserobject.ExchangeGuid,$uoUserobject.ProhibitSendQuota)
				}
				}
			"Proibit Send/Recieve" {
				if($uoUserobject.ProhibitSendReceiveQuota.isUnlimited -eq $false){
					$usrquotas.add($uoUserobject.ExchangeGuid,$uoUserobject.ProhibitSendReceiveQuota)
				}
				}
		}
	}
}

$mbServers = get-mailboxserver

if ($mtTypeDrop.SelectedItem -ne $null){
	if ($mtTypeDrop.SelectedItem.ToString() -eq "Disconnected"){
		$fname ="disconnected"
		if ($serverName -eq "ALL Servers"){
			$mbServers | foreach-object{
				$mscombCollection += get-mailboxstatistics -server $_.Name  | Where {$_.DisconnectDate -ne $null}
			}
		}
		else{
			$mscombCollection += get-mailboxstatistics -server $serverName  | Where {$_.DisconnectDate -ne $null}
		} 
		$mscombCollection | ForEach-Object{
		$quQuota = "0"
		if ($usrquotas.ContainsKey($_.MailboxGUID)){
			if ($usrquotas[$_.MailboxGUID].Value -ne $null){
				if ($usrquotas[$_.MailboxGUID].Value.ToMB() -gt 0){
					$quQuota = "{0:P0}" -f ($_.TotalItemSize.Value.ToMB()/$usrquotas[$_.MailboxGUID].Value.ToMB())}
				else{
					$quQuota = "100"
				}
			}
		}
		else{
			if ($mstoresquotas.ContainsKey($_.database)){
				if ($mstoresquotas[$_.database].Value -ne $null){
					$quQuota = "{0:P0}" -f ($_.TotalItemSize.Value.ToMB()/$mstoresquotas[$_.database].Value.ToMB())
			}}
		}
		$icount = 0
		$tisize = 0
		$disize = 0
		if ($_.DisplayName -ne $null){$dname = $_.DisplayName}
		if ($_.ItemCount -ne $null){$icount = $_.ItemCount}
		if ($_.TotalItemSize.Value.ToMB() -ne $null){$tisize = $_.TotalItemSize.Value.ToMB()}
		if ($_.TotalDeletedItemSize.Value.ToMB() -ne $null){$disize = $_.TotalDeletedItemSize.Value.ToMB()}  
		$msTable.Rows.add($dname,$icount,$tisize,$disize,$quQuota.replace("%","").replace(",",""))
		$mbcomb = "" | select Date,ServerName,DisplayName,ItemCount,TotalItemSize,TotalDeletedItemSize
		$mbcomb.Date = $datetime.ToString("yyyyMMdd")
		$mbcomb.ServerName = $serverName
		$mbcomb.DisplayName = $dname
		$mbcomb.ItemCount = $icount
		$mbcomb.TotalItemSize = $tisize
		$mbcomb.TotalDeletedItemSize = $disize
		$mbcombCollection += $mbcomb
		}
	}
	else{	
		if ($serverName -eq "ALL Servers"){
			$fname = $fname  + $datetime.ToString("yyyyMMdd")  + "-ALLSERVERS.csv"
			$mbServers | foreach-object{
				$mscombCollection += get-mailboxstatistics -server $_.Name | Where {$_.DisconnectDate -eq $null}
			}
		}
		else{
			$fname = $fname  + $datetime.ToString("yyyyMMdd")  + "-" + $serverName + ".csv"
			$mscombCollection += get-mailboxstatistics -server $serverName  | Where {$_.DisconnectDate -eq $null}
		} 
		$mscombCollection  | ForEach-Object{
		$quQuota = "0"
		if ($usrquotas.ContainsKey($_.MailboxGUID)){
			if ($usrquotas[$_.MailboxGUID].Value -ne $null){
				if ($usrquotas[$_.MailboxGUID].Value.ToMB() -gt 0){
					$quQuota = "{0:P0}" -f ($_.TotalItemSize.Value.ToMB()/$usrquotas[$_.MailboxGUID].Value.ToMB())}
				else{
					$quQuota = "100"
				}
			}
		}
		else{
			if ($mstoresquotas.ContainsKey($_.database)){
				if ($mstoresquotas[$_.database].Value -ne $null){
				$quQuota = "{0:P0}" -f ($_.TotalItemSize.Value.ToMB()/$mstoresquotas[$_.database].Value.ToMB())}}
		}
		$icount = 0
		$tisize = 0
		$disize = 0
		if ($_.DisplayName -ne $null){$dname = $_.DisplayName}
		if ($_.ItemCount -ne $null){$icount = $_.ItemCount}
		if ($_.TotalItemSize.Value.ToMB() -ne $null){$tisize = $_.TotalItemSize.Value.ToMB()}
		if ($_.TotalDeletedItemSize.Value.ToMB() -ne $null){$disize = $_.TotalDeletedItemSize.Value.ToMB()}    
		$msTable.Rows.add($dname,$icount,$tisize,$disize,$quQuota.replace("%","").replace(",",""))
		$mbcomb = "" | select Date,ServerName,DisplayName,ItemCount,TotalItemSize,TotalDeletedItemSize
		$mbcomb.Date = $datetime.ToString("yyyyMMdd")
		$mbcomb.ServerName = $serverName
		$mbcomb.DisplayName = $dname
		$mbcomb.ItemCount = $icount
		$mbcomb.TotalItemSize = $tisize
		$mbcomb.TotalDeletedItemSize = $disize
		$mbcombCollection += $mbcomb
		}

	}
}
else{
		
		if ($serverName -eq "ALL Servers"){
			$fname = $fname  + $datetime.ToString("yyyyMMdd")  + "-ALLSERVERS.csv"
			$mbServers | foreach-object{
				$mscombCollection += get-mailboxstatistics -server $_.Name
			}
		}
		else{
			$fname = $fname  + $datetime.ToString("yyyyMMdd")  + "-" + $serverName + ".csv"
			$mscombCollection += get-mailboxstatistics -server $serverName  
		} 
		$mscombCollection | ForEach-Object{
		$quQuota = "0"
		if ($usrquotas.ContainsKey($_.MailboxGUID)){
			if ($usrquotas[$_.MailboxGUID].Value -ne $null){
				if ($usrquotas[$_.MailboxGUID].Value.ToMB() -gt 0){
					$quQuota = "{0:P0}" -f ($_.TotalItemSize.Value.ToMB()/$usrquotas[$_.MailboxGUID].Value.ToMB())}
				else{
					$quQuota = "100"
				}
			}
		}
		else{
		if ($mstoresquotas.ContainsKey($_.database)){
				if ($mstoresquotas[$_.database].Value -ne $null){
				$quQuota = "{0:P0}" -f ($_.TotalItemSize.Value.ToMB()/$mstoresquotas[$_.database].Value.ToMB())}}
		}
	        $icount = 0
		$tisize = 0
		$disize = 0
		if ($_.DisplayName -ne $null){$dname = $_.DisplayName}
		if ($_.ItemCount -ne $null){$icount = $_.ItemCount}
		if ($_.TotalItemSize.Value.ToMB() -ne $null){$tisize = $_.TotalItemSize.Value.ToMB()}
		if ($_.TotalDeletedItemSize.Value.ToMB() -ne $null){$disize = $_.TotalDeletedItemSize.Value.ToMB()}    
		$msTable.Rows.add($dname,$icount,$tisize,$disize,$quQuota.replace("%","").replace(",",""))
		$mbcomb = "" | select Date,ServerName,DisplayName,ItemCount,TotalItemSize,TotalDeletedItemSize
		$mbcomb.Date = $datetime.ToString("yyyyMMdd")
		$mbcomb.ServerName = $serverName
		$mbcomb.DisplayName = $dname
		$mbcomb.ItemCount = $icount
		$mbcomb.TotalItemSize = $tisize
		$mbcomb.TotalDeletedItemSize = $disize
		$mbcombCollection += $mbcomb
	}

} 
write-host $fstring 

if ($fname -ne "disconnected") {
	$mbcombCollection | export-csv –encoding "unicode" -noTypeInformation $fname 
}

}

function ShowGrowth(){


$gtTable.clear()
$datetime = get-date
$arArrayList = New-Object System.Collections.ArrayList
dir $script:HistoryDir\*.csv | foreach-object{ 
	$fname = $_.name
	$nmArray = $_.name.split("-")
	if ($nmArray[1].Replace(".csv","") -eq $serverName) {
		[VOID]$arArrayList.Add($nmArray[0])
	}
}
$arArrayList.Sort()
$spoint = $arArrayList[$arArrayList.Count-1]
$oneday = $spoint
$sevenday = $spoint
$onemonth = $spoint
$oneyear = $spoint
foreach ($file in $arArrayList){
	if ($file -gt ($datetime.Adddays(-2).ToString("yyyyMMdd")) -band $file -lt $oneday) {$oneday = $file} 
	if ($file -gt ($datetime.Adddays(-7).ToString("yyyyMMdd")) -band $file -lt $sevenday) {$sevenday = $file} 
	if ($file -gt ($datetime.Adddays(-31).ToString("yyyyMMdd")) -band $file -lt $onemonth) {$onemonth = $file} 
	if ($file -gt ($datetime.Adddays(-256).ToString("yyyyMMdd")) -band $file -lt $oneyear) {$oneyear = $file} 
}
write-host $oneday
write-host $sevenday
write-host $onemonth
write-host $oneyear

$onedaystats = @{ }
$sevendaystats = @{ }
$onemonthsdaystats = @{ }
$oneyearstats = @{ }

Import-Csv ("$script:HistoryDir\" + $oneday + "-" + $serverName + ".csv") | %{ 
	$onedaystats.add($_.DisplayName,$_.TotalItemSize)	
}
Import-Csv ("$script:HistoryDir\" + $sevenday + "-" + $serverName + ".csv") | %{ 
	$sevendaystats.add($_.DisplayName,$_.TotalItemSize)	
}
Import-Csv ("$script:HistoryDir\" + $onemonth + "-" + $serverName + ".csv") | %{ 
	$onemonthsdaystats.add($_.DisplayName,$_.TotalItemSize)	
}
Import-Csv ("$script:HistoryDir\" + $oneyear + "-" + $serverName + ".csv") | %{ 
	$oneyearstats.add($_.DisplayName,$_.TotalItemSize)	
}

foreach($row in $msTable.Rows){
	if ($onedaystats.ContainsKey($row[0].ToString())){
		$ondaysizegrowth = $row[2] - $onedaystats[$row[0].ToString()]
	}
	else{$ondaysizegrowth = 0}
	if ($sevendaystats.ContainsKey($row[0].ToString())){
		$sevendaysizegrowth = $row[2] - $sevendaystats[$row[0].ToString()]}
	else{$sevendaysizegrowth = 0}
	if ($onemonthsdaystats.ContainsKey($row[0].ToString())){
		$onemonthsizegrowth = $row[2] - $onemonthsdaystats[$row[0].ToString()]}
	else{$onemonthsizegrowth = 0}
	if ($oneyearstats.ContainsKey($row[0].ToString())){
		$oneyearsizegrowth = $row[2] - $oneyearstats[$row[0].ToString()]}
	else{$oneyearsizegrowth = 0}
	$gtTable.rows.add($row[0].ToString(),$row[2],$ondaysizegrowth,$sevendaysizegrowth,$onemonthsizegrowth,$oneyearsizegrowth)
	
}
$dgDataGrid.DataSource = $gtTable

}

function GetFolderSizes(){
$fsTable.clear()
$snServername = $serverName
write-host $dgDataGrid.CurrentCell.RowIndex
get-user $msTable.DefaultView[$dgDataGrid.CurrentCell.RowIndex][0] -RecipientType "UserMailbox" | foreach-object{
	$siSIDToSearch = $_
}
write-host $siSIDToSearch.WindowsEmailAddress.ToString()
Get-MailboxFolderStatistics $siSIDToSearch.WindowsEmailAddress.ToString() | ForEach-Object{
	$ficount = 0
	$fisize = 0
	$fsisize = 0
	$fscount = 0
	$fname = $_.Name
	if ($_.FolderSize -ne $null){$fsisize = [math]::round(($_.FolderSize/1mb),2)}
	if ($_.ItemsInFolder -ne $null){$ficount = $_.ItemsInFolder}
	if ($_.ItemsInFolderAndSubfolders -ne $null){$fscount = $_.ItemsInFolderAndSubfolders} 
	if ($_.FolderAndSubfolderSize -ne $null){$fsisize = [math]::round(($_.FolderAndSubfolderSize/1mb),2)}      
	$fsTable.Rows.add($fname,$ficount,$fsisize,$fscount,$fsisize)
}
$dgDataGrid1.DataSource = $fsTable
}

function ExportMBcsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("UserName,# Items,MB Size(MB),DelItems(MB),QuotaUsed")
	foreach($row in $msTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[1].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString()) 
	}
	$logfile.Close()
}
}

function ExportFScsv{

$exFileName = new-object System.Windows.Forms.saveFileDialog
$exFileName.DefaultExt = "csv"
$exFileName.Filter = "csv files (*.csv)|*.csv"
$exFileName.InitialDirectory = "c:\temp"
$exFileName.ShowHelp = $true
$exFileName.ShowDialog()
if ($exFileName.FileName -ne ""){
	$logfile = new-object IO.StreamWriter($exFileName.FileName,$true)
	$logfile.WriteLine("DisplayName,# Items,Folder Size(MB),# Items + Sub,Folder Size + Sub(MB)")
	foreach($row in $fsTable.Rows){
		$logfile.WriteLine("`"" + $row[0].ToString() + "`"," + $row[3].ToString() + "," + $row[2].ToString() + "," + $row[3].ToString() + "," + $row[4].ToString()) 
	}
	$logfile.Close()
}
}

$usrquotas = @{ }
$mstoresquotas = @{ }

$global:LastFolder = ""
# Add DataTable

$Dataset = New-Object System.Data.DataSet
$fsTable = New-Object System.Data.DataTable
$fsTable.TableName = "Folder Sizes"
$fsTable.Columns.Add("DisplayName")
$fsTable.Columns.Add("# Items",[int64])
$fsTable.Columns.Add("Folder Size(MB)",[int64])
$fsTable.Columns.Add("# Items + Sub",[int64])
$fsTable.Columns.Add("Folder Size + Sub(MB)",[int64])
$Dataset.tables.add($fsTable)

$msTable = New-Object System.Data.DataTable
$msTable.TableName = "Mailbox Sizes"
$msTable.Columns.Add("UserName")
$msTable.Columns.Add("# Items",[int64])
$msTable.Columns.Add("MB Size(MB)",[int64])
$msTable.Columns.Add("DelItems(MB)",[int64])
$msTable.Columns.Add("Quota Used",[int64])
$Dataset.tables.add($msTable)


$gtTable = New-Object System.Data.DataTable
$gtTable.TableName = "Mailbox Growth"
$gtTable.Columns.Add("UserName")
$gtTable.Columns.Add("Mailbox Size",[int64])
$gtTable.Columns.Add("1 Day",[int64])
$gtTable.Columns.Add("7 Days",[int64])
$gtTable.Columns.Add("31 Days",[int64])
$gtTable.Columns.Add("1 Year",[int64])
$Dataset.tables.add($gtTable)



getMailboxSizes

