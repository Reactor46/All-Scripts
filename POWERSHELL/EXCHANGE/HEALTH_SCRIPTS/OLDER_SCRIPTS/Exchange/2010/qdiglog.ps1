Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
$servername = "servername"

$dtQueryDT = [DateTime]::UtcNow.AddHours(-24)
$dtQueryDTf =  [DateTime]::UtcNow
$mIMessageIDArray = @{ }
$htmlCollection = @()

get-agentlog -StartDate $dtQueryDT -EndDate $dtQueryDTf | where { $_.Action -eq "QuarantineMessage"} | foreach-object {
	$valuesobj = "" |  Select-Object DateTime,From,To,Subject,Size
	$valuesobj.DateTime = $_.TimeStamp
	$valuesobj.From = $_.P1FromAddress
	$recpString = ""
	foreach($recp in $_.Recipients){
		if ($recpString -eq ""){
			$recpString = $recp.ToString()}
		else{
			$recpString = $recpString  + ";" + $recp.ToString()
		}
	}	
	$valuesobj.To = $recpString
	$mIMessageIDArray.add($_.MessageID,$valuesobj)
	
}
if ($mIMessageIDArray.Count -gt 0){

Get-MessageTrackingLog -Server $servername -ResultSize Unlimited -Start $dtQueryDT -End $dtQueryDTf -EventId "Receive" | ForEach-Object{
	if ($mIMessageIDArray.Containskey($_.MessageID)){
		$mIMessageIDArray[$_.MessageID].Subject = $_.MessageSubject
		$mIMessageIDArray[$_.MessageID].Size = ($_.TotalBytes/1024).ToString(0.00)
	}


}
$rpReport = $rpReport + "<table><tr bgcolor=`"#95aedc`">"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:15%;`" ><b>Recieved</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>From</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>To</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:40%;`" ><b>Subject</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:5%;`" ><b>Size</b></td>"
$rpReport = $rpReport + "</tr>"

$mIMessageIDArray.GetEnumerator() | sort name -descending | foreach-object { 
	if ($_.Value.From.ToString().length -gt 30){$fromstring = $_.Value.From.ToString().Substring(0,30)}
						    else{$fromstring = $_.Value.From.ToString()}
	if ($_.Value.To.ToString().length -gt 30){$Tostring = $_.Value.To.ToString().Substring(0,30)}
						    else{$Tostring = $_.Value.To.ToString()}
	$rpReport = $rpReport + "  <tr>"  + "  "
	$rpReport = $rpReport + "<td>" + $_.Value.DateTime + "</td>"  + "  "
	$rpReport = $rpReport + "<td>" +  $fromstring + "</td>"  + "  "
	$rpReport = $rpReport + "<td>" + $Tostring + "</td>"  + "  "
	$rpReport = $rpReport + "<td>" + $_.Value.Subject + "</td>"  + "  "
	$rpReport = $rpReport + "<td>" +  $_.Value.Size + "</td>"  + "  "
	$rpReport = $rpReport + "</tr>"  + "  "

}
$rpReport = $rpReport + "</table>"  + "  " 

$SmtpClient = new-object system.net.mail.smtpClient
$SmtpClient.host = $servername
$MailMessage = new-object System.Net.Mail.MailMessage
$MailMessage.To.Add("user@domain.com")
$MailMessage.From = "Digest@domain.com"
$MailMessage.Subject = "Quarantine Digest" 
$MailMessage.IsBodyHtml = $TRUE
$MailMessage.body = $rpReport
$SMTPClient.Send($MailMessage)

}