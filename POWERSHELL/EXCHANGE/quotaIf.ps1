$QuotaIfVal = $args[0]
$Script:rptCollection = @()
get-mailbox -ResultSize unlimited| Get-MailboxStatistics | foreach-object{
	$rptObj = "" | Select MailboxName,TotalSize,QuotaIfPercent,PercentGraph
	$rptObj.MailboxName = $_.DisplayName
	[Int64]$rptObj.TotalSize = ($_.TotalItemSize.ToString() |%{($_.Substring($_.indexof("(")+1,$_.indexof("b")-$_.indexof("(")-2)) -replace(",","")})/1MB

	$rptObj.QuotaIfPercent = 0 	
	if($rptObj.TotalSize -gt 0){
		$rptObj.QuotaIfPercent = [Math]::round((($rptObj.TotalSize/$QuotaIfVal) * 100)) 
	}
	$PercentGraph = ""
	for($intval=0;$intval -lt 100;$intval+=4){
		if($rptObj.QuotaIfPercent -gt $intval){
			$PercentGraph += "▓"
		}
		else{		
			$PercentGraph += "░"
		}
	}
	$rptObj.PercentGraph = $PercentGraph 
	$rptObj | fl
	$Script:rptCollection +=$rptObj	
}
$Script:rptCollection
$Script:rptCollection | Export-Csv -NoTypeInformation -Path c:\temp\QuotaIfReport.csv -Encoding UTF8 
