$StoreSizeThreshold1day = 100
$StoreSizeThreshold7day = 300
$StoreSizeThreshold30day = 1000

$WhiteSpaceThreshold1day = 100
$WhiteSpaceThreshold7day = 300
$WhiteSpaceThreshold30day = 1000

$SMTPservername = "servername"
$Toaddress = "user@yourdomain.com"
$fromAddress = "Report@yourdomain.com"

function GetWSforServer($snServerName){

$SgcombCollection = @()
"Querying Server : " + $snServerName

$StoreFileSizes = @{ }
$StoreWhitespace = @{ }
$StoreRetainItemSize = @{ }
$StoreDeletedMailboxSize = @{ }
$sgTotalHash = @{ }
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
$StoreFileSizes.GetEnumerator() | sort name -descending | foreach-object {
	$snames = $_.key.split("\")
	if ($sgTotalHash.containsKey($snames[0])){
		if ($StoreDeletedMailboxSize.containskey($_.key)){
			$sgTotalHash[$snames[0]].MailboxStores = [int]$sgTotalHash[$snames[0]].MailboxStores + 1
			$sgTotalHash[$snames[0]].MSSize = [int]$sgTotalHash[$snames[0]].MSSize + [int]$_.Value 
			$sgTotalHash[$snames[0]].MSWhiteSpace = [int]$sgTotalHash[$snames[0]].MSWhiteSpace + [int]$StoreWhitespace[$_.key].ToString()
			$sgTotalHash[$snames[0]].MSRetainedItems = [int]$sgTotalHash[$snames[0]].RetainedItems + [int]$StoreRetainItemSize[$_.key].ToString()
			$sgTotalHash[$snames[0]].RetainedMB = [int]$sgTotalHash[$snames[0]].RetainedMB + [int]$StoreDeletedMailboxSize[$_.key].ToString()

		}
		else{
			$sgTotalHash[$snames[0]].PublicFolderStores = [int]$sgTotalHash[$snames[0]].MailboxStores + 1
			$sgTotalHash[$snames[0]].PFSize = [int]$sgTotalHash[$snames[0]].MSSize + [int]$_.Value 
			$sgTotalHash[$snames[0]].PFWhiteSpace = [int]$sgTotalHash[$snames[0]].MSWhiteSpace + [int]$StoreWhitespace[$_.key].ToString()
			$sgTotalHash[$snames[0]].PFRetainedItems = [int]$sgTotalHash[$snames[0]].RetainedItems + [int]$StoreRetainItemSize[$_.key].ToString()
	
		}
		
	}
	else{
		$sgtotalObject = "" | select ServerName,Name,MailboxStores,PublicFolderStores,MSSize,PFSize,MSWhiteSpace,PFWhiteSpace,MSRetainedItems,PFRetainedItems,RetainedMB
		$sgtotalObject.ServerName = $snServerName
		$sgtotalObject.Name = $snames[0]
		if ($StoreDeletedMailboxSize.containskey($_.key)){
			$sgtotalObject.MailboxStores = 1
			$sgtotalObject.PublicFolderStores = 0
			$sgtotalObject.MSSize = $_.Value.ToString()  
			$sgtotalObject.PFSize = 0  
			$sgtotalObject.MSWhiteSpace = $StoreWhitespace[$_.key].ToString()
			$sgtotalObject.PFWhiteSpace = 0
			$sgtotalObject.MSRetainedItems = $StoreRetainItemSize[$_.key].ToString()
			$sgtotalObject.PFRetainedItems = 0
			$sgtotalObject.RetainedMB = $StoreDeletedMailboxSize[$_.key].ToString()
		}
		else {
			$sgtotalObject.MailboxStores = 0
			$sgtotalObject.PublicFolderStores = 1
			$sgtotalObject.MSSize = 0
			$sgtotalObject.PFSize = $_.Value.ToString()   
			$sgtotalObject.MSWhiteSpace = 0
			$sgtotalObject.PFWhiteSpace =  $StoreWhitespace[$_.key].ToString()
			$sgtotalObject.MSRetainedItems = 0
			$sgtotalObject.PFRetainedItems = $StoreRetainItemSize[$_.key].ToString()
			$sgtotalObject.RetainedMB = 0
		}
		
		$sgTotalHash.Add($snames[0],$sgtotalObject)
	}

}



foreach ($row in $sgTotalHash.Values){
	$global:rpReport = $global:rpReport + "  <tr>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $snServerName + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $row.Name.ToString() + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $row.MailboxStores.ToString() + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $row.PublicFolderStores.ToString() + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $row.MSSize.ToString() + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $row.PFSize.ToString() + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $row.MSWhiteSpace.ToString() + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $row.PFWhiteSpace.ToString() + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $row.MSRetainedItems.ToString() + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $row.PFRetainedItems.ToString() + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "<td>" + $row.RetainedMB.ToString() + "</td>"  + "`r`n"
	$global:rpReport = $global:rpReport + "</tr>"  + "`r`n"
	$global:ExcombCollection += $row
	if ($onedaystats.Containskey(($snServerName + "\" + $row.Name.ToString()))){
		$onedaystats[($snServerName + "\" + $row.Name.ToString())].MSSize
   		if(([INT]$row.MSSize - [INT]$onedaystats[($snServerName + "\" + $row.Name.ToString())].MSSize) -gt $StoreSizeThreshold1day){
			writeThresholdExecption "One Day MailStore Size" $snServerName $row.Name.ToString() $row.MSSize $onedaystats[($snServerName + "\" + $row.Name.ToString())].MSSize.ToString() 
			}
		if(([INT]$row.PFSize - [INT]$onedaystats[($snServerName + "\" + $row.Name.ToString())].PFSize) -gt $StoreSizeThreshold1day){
			writeThresholdExecption "One Day PublicStore Size" $snServerName $row.Name.ToString() $row.PFSize $onedaystats[($snServerName + "\" + $row.Name.ToString())].PFSize.ToString() 
			}
		if(([INT]$row.MSWhiteSpace - [INT]$onedaystats[($snServerName + "\" + $row.Name.ToString())].MSWhiteSpace) -gt $WhiteSpaceThreshold1day){
			writeThresholdExecption "One Day Mail Store WhiteSpace" $snServerName $row.Name.ToString() $row.MSWhiteSpace $onedaystats[($snServerName + "\" + $row.Name.ToString())].MSWhiteSpace.ToString() 
			}
		if(([INT]$row.PFWhiteSpace - [INT]$onedaystats[($snServerName + "\" + $row.Name.ToString())].PFWhiteSpace) -gt $WhiteSpaceThreshold1day){
			writeThresholdExecption "One Day Public Store WhiteSpace" $snServerName $row.Name.ToString() $row.PFWhiteSpace $onedaystats[($snServerName + "\" + $row.Name.ToString())].PFWhiteSpace.ToString() 
			}
		
	}
	if ($onedaystats.Containskey(($snServerName + "\" + $row.Name.ToString()))){
		$sevendaystats[($snServerName + "\" + $row.Name.ToString())].MSSize
   		if(([INT]$row.MSSize - [INT]$sevendaystats[($snServerName + "\" + $row.Name.ToString())].MSSize) -gt $StoreSizeThreshold7day){
			writeThresholdExecption "Seven Day MailStore Size" $snServerName $row.Name.ToString() $row.MSSize $sevendaystats[($snServerName + "\" + $row.Name.ToString())].MSSize.ToString() 
		}
		if(([INT]$row.PFSize - [INT]$sevendaystats[($snServerName + "\" + $row.Name.ToString())].PFSize) -gt $StoreSizeThreshold7day){
			writeThresholdExecption "Seven Day Public Store Size" $snServerName $row.Name.ToString() $row.PFSize $sevendaystats[($snServerName + "\" + $row.Name.ToString())].PFSize.ToString() 
		}
   		if(([INT]$row.MSWhiteSpace - [INT]$sevendaystats[($snServerName + "\" + $row.Name.ToString())].MSWhiteSpace) -gt $WhiteSpaceThreshold7day){
			writeThresholdExecption "Seven Day Mail Store WhiteSpace" $snServerName $row.Name.ToString() $row.MSWhiteSpace $sevendaystats[($snServerName + "\" + $row.Name.ToString())].MSWhiteSpace.ToString() 
		}
		if(([INT]$row.PFWhiteSpace - [INT]$sevendaystats[($snServerName + "\" + $row.Name.ToString())].PFWhiteSpace) -gt $WhiteSpaceThreshold7day){
			writeThresholdExecption "Seven Day Public Store WhiteSpace" $snServerName $row.Name.ToString() $row.PFWhiteSpace $sevendaystats[($snServerName + "\" + $row.Name.ToString())].PFWhiteSpace.ToString() 
		}
	}
	if ($onemonthsstats.Containskey(($snServerName + "\" + $row.Name.ToString()))){
		$onemonthsstats[($snServerName + "\" + $row.Name.ToString())].MSSize
   		if(([INT]$row.MSSize - [INT]$onemonthsstats[($snServerName + "\" + $row.Name.ToString())].MSSize) -gt $StoreSizeThreshold30day){
			writeThresholdExecption "Thirty Day MailStore Size" $snServerName $row.Name.ToString() $row.MSSize $onemonthsstats[($snServerName + "\" + $row.Name.ToString())].MSSize.ToString() 
		}
		if(([INT]$row.PFSize - [INT]$onemonthsstats[($snServerName + "\" + $row.Name.ToString())].PFSize) -gt $StoreSizeThreshold30day){
			writeThresholdExecption "Thirty Day Public Store Size" $snServerName $row.Name.ToString() $row.PFSize $onemonthsstats[($snServerName + "\" + $row.Name.ToString())].PFSize.ToString() 
		}
   		if(([INT]$row.MSWhiteSpace - [INT]$onemonthsstats[($snServerName + "\" + $row.Name.ToString())].MSWhiteSpace) -gt $WhiteSpaceThreshold30day){
			writeThresholdExecption "Thirty Day MailStore WhiteSpace" $snServerName $row.Name.ToString() $row.MSWhiteSpace $onemonthsstats[($snServerName + "\" + $row.Name.ToString())].MSWhiteSpace.ToString() 
		}
		if(([INT]$row.PFWhiteSpace - [INT]$onemonthsstats[($snServerName + "\" + $row.Name.ToString())].PFWhiteSpace) -gt $WhiteSpaceThreshold30day){
			writeThresholdExecption "Thirty Day Public Store WhiteSpace" $snServerName $row.Name.ToString() $row.PFWhiteSpace $onemonthsstats[($snServerName + "\" + $row.Name.ToString())].PFWhiteSpace.ToString() 
		}
	}
}


}

function writeThresholdExecption([String]$threshold,[String]$tsServerName,[String]$tsSGName,[String]$tscValue,[String]$tspValue) {
$global:expReport = $global:expReport + " <tr>" +"`r`n"
$global:expReport = $global:expReport + "<td>" + $threshold + "</td>" +"`r`n"
$global:expReport = $global:expReport + "<td>" + $tsServerName + "</td>" +"`r`n"
$global:expReport = $global:expReport + "<td>" + $tsSGName + "</td>" +"`r`n"
$global:expReport = $global:expReport + "<td>" + $tscValue + "</td>" +"`r`n"
$global:expReport = $global:expReport + "<td>" + $tspValue + "</td>" +"`r`n"
$global:expReport = $global:expReport + "<td>" + ([INT]$tscValue - [INT]$tspValue) + "</td>" +"`r`n"
$global:expReport = $global:expReport + "</tr>" + "`r`n"

}



$HistoryDir = "c:\sizehistory"
$frun = 0
if (!(Test-Path -path $HistoryDir))
{
	New-Item $HistoryDir -type directory
	$frun = 1
}
$global:rpReport = ""
$global:expReport = ""
$global:rpReport = $global:rpReport + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:10%;`" ><b>ServerName</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:15%;`" ><b>SG Name</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:5%;`" ><b>MS #</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:5%;`" ><b>PFS #</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:10%;`" ><b>MS Size</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:10%;`" ><b>PF Size</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:10%;`" ><b>MSWSpace</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:10%;`" ><b>PFWSpace</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:10%;`" ><b>MSRetItems</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:10%;`" ><b>PFRetItems</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "<td align=`"center`" style=`"width:10%;`" ><b>RetMB</b></td>" +"`r`n"
$global:rpReport = $global:rpReport + "</tr>" + "`r`n"
$global:expReport = $global:expReport + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
$global:expReport = $global:expReport + "<td align=`"center`" style=`"width:15%;`" ><b>Threshold</b></td>" +"`r`n"
$global:expReport = $global:expReport + "<td align=`"center`" style=`"width:15%;`" ><b>ServerName</b></td>" +"`r`n"
$global:expReport = $global:expReport + "<td align=`"center`" style=`"width:15%;`" ><b>SGName</b></td>" +"`r`n"
$global:expReport = $global:expReport + "<td align=`"center`" style=`"width:10%;`" ><b>Current Value</b></td>" +"`r`n"
$global:expReport = $global:expReport + "<td align=`"center`" style=`"width:10%;`" ><b>Previous Value</b></td>" +"`r`n"
$global:expReport = $global:expReport + "<td align=`"center`" style=`"width:10%;`" ><b>Growth</b></td>" +"`r`n"
$global:expReport = $global:expReport + "</tr>" + "`r`n"
$global:ExcombCollection = @()
$datetime = get-date
$onedaystats = @{ }
$sevendaystats = @{ }
$onemonthsstats = @{ }
$oneyearstats = @{ }

if ($frun -eq 0){
	$arArrayList = New-Object System.Collections.ArrayList
	dir $script:HistoryDir\*.csv | foreach-object{ 
		$fname = $_.name
		$nmArray = $_.name.split("-")
		[VOID]$arArrayList.Add($nmArray[0])
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
	}



	Import-Csv ("$script:HistoryDir\" + $oneday + "-wspacerun.csv") | %{ 
		$onedaystats.add(($_.ServerName.ToString() + "\" + $_.Name.ToString()),$_)	
	}
	Import-Csv ("$script:HistoryDir\" + $sevenday + "-wspacerun.csv") | %{ 
		$sevendaystats.add(($_.ServerName.ToString() + "\" + $_.Name.ToString()),$_)	
	}
	Import-Csv ("$script:HistoryDir\" + $onemonth + "-wspacerun.csv") | %{ 
		$onemonthsstats.add(($_.ServerName.ToString() + "\" + $_.Name.ToString()),$_)	
	}
}

Get-MailboxServer | foreach-object{
	GetWSforServer($_.Name.ToString())		
}
$datetime = get-date
$fname = $script:HistoryDir + "\"
$fname = $fname  + $datetime.ToString("yyyyMMdd")  + "-wspacerun.csv"
$global:ExcombCollection | export-csv –encoding "unicode" -noTypeInformation $fname
$global:rpReport = $global:rpReport + "</table>"  + " <BR><BR><H1>Threshold Exceeds</H1><BR> "
$global:rpReport = $global:rpReport + $global:expReport
$global:rpReport = $global:rpReport + "</table>"  + " <BR><BR> "

$SmtpClient = new-object system.net.mail.smtpClient
$SmtpClient.host = $SMTPservername
$MailMessage = new-object System.Net.Mail.MailMessage
$MailMessage.To.Add($ToAddress)
$MailMessage.From = $FromAddress
$MailMessage.Subject = "Storage Group Reports" 
$MailMessage.IsBodyHtml = $TRUE
$MailMessage.body = $global:rpReport
$SMTPClient.Send($MailMessage)
write-host "Mail Sent"
