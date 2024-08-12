$UserName = "user@domain.com"
$Password = "pasdfawrd."
$mbHash = @{ }
$MailboxName = $UserName

$batchSize = 100
$global:oaBoard = ""

$Mailboxes = @()

$cred = New-Object System.Net.NetworkCredential($UserName,$Password)   
$secpassword = ConvertTo-SecureString $Password -AsPlainText -Force
$adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $UserName,$secpassword   

If(Get-PSSession | where-object {$_.ConfigurationName -eq "Microsoft.Exchange"}){
	write-host "Session Exists"
}
else{
	$rpRemotePowershell = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -credential $adminCredential  -Authentication Basic -AllowRedirection 
	$importresults = Import-PSSession $rpRemotePowershell 
} 

get-mailbox -ResultSize unlimited | where-object {$_.HiddenFromAddressListsEnabled -eq $false}  | foreach-object{
	if ($mbHash.ContainsKey($_.WindowsEmailAddress.ToString()) -eq $false){
		$mbHash.Add($_.WindowsEmailAddress.ToString(),$_.DisplayName)
	    
	}
	$emailAddress = $_.WindowsEmailAddress.ToString()
	$Mailboxes += $emailAddress
}

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$service.TraceEnabled = $false

$service.Credentials = $cred
$service.autodiscoverurl($MailboxName,{$true})

function GetMailTips{
$mbMailboxFolderURI = New-Object System.Uri($service.url)
$wrWebRequest = [System.Net.WebRequest]::Create($mbMailboxFolderURI)
$wrWebRequest.KeepAlive = $false;
$wrWebRequest.Headers.Set("Pragma", "no-cache");
$wrWebRequest.Headers.Set("Translate", "f");
$wrWebRequest.Headers.Set("Depth", "0");
$wrWebRequest.ContentType = "text/xml";
$wrWebRequest.ContentLength = $expRequest.Length;
$wrWebRequest.Timeout = 60000;
$wrWebRequest.Method = "POST";
$wrWebRequest.Credentials = $cred
$bqByteQuery = [System.Text.Encoding]::ASCII.GetBytes($expRequest);
$wrWebRequest.ContentLength = $bqByteQuery.Length;
$rsRequestStream = $wrWebRequest.GetRequestStream();
$rsRequestStream.Write($bqByteQuery, 0, $bqByteQuery.Length);
$rsRequestStream.Close();
$wrWebResponse = $wrWebRequest.GetResponse();
$rsResponseStream = $wrWebResponse.GetResponseStream()
$sr = new-object System.IO.StreamReader($rsResponseStream);
$rdResponseDocument = New-Object System.Xml.XmlDocument
$rdResponseDocument.LoadXml($sr.ReadToEnd());
$RecipientNodes = @($rdResponseDocument.getElementsByTagName("t:RecipientAddress"))
$Datanodes = @($rdResponseDocument.getElementsByTagName("t:OutOfOffice"))
for($ic=0;$ic -lt $RecipientNodes.length;$ic++){
	if($Datanodes[$ic].ReplyBody.Message -eq ""){
		
		$global:oaBoard = $global:oaBoard + "<tr>"  + "`r`n"
		$global:oaBoard = $global:oaBoard + "<td>" + $mbHash[$RecipientNodes[$ic].EmailAddress] + "</td>"  + "`r`n"
		$global:oaBoard = $global:oaBoard + "<td bgcolor=`"#41A317`">In the Office</td>"  + "`r`n"
		$global:oaBoard = $global:oaBoard + "<td></td>"  + "`r`n"
		$global:oaBoard = $global:oaBoard + "</tr>"  + "`r`n"
	}
	else{
		$global:oaBoard = $global:oaBoard + "<tr>"  + "`r`n"
		$global:oaBoard = $global:oaBoard + "<td>" + $mbHash[$RecipientNodes[$ic].EmailAddress] + "</td>"  + "`r`n"
		$global:oaBoard = $global:oaBoard + "<td bgcolor=`"#153E7E`">Out of the Office</td>"  + "`r`n"
		$global:oaBoard = $global:oaBoard + "<td>" + $Datanodes[$ic].ReplyBody.Message  + "</td>" + "`r`n"
		$global:oaBoard = $global:oaBoard + "</tr>"  + "`r`n"
	}
}
}


$expHeader = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
<soap:Header><RequestServerVersion Version="Exchange2010_SP1" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" />
</soap:Header>
<soap:Body>
<GetMailTips xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
<SendingAs>
<EmailAddress xmlns="http://schemas.microsoft.com/exchange/services/2006/types">$MailboxName</EmailAddress>
</SendingAs>
<Recipients>
"@

$global:oaBoard = $global:oaBoard + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
$global:oaBoard = $global:oaBoard + "<td align=`"center`" style=`"width=200;`" ><b>User</b></td>" +"`r`n"
$global:oaBoard = $global:oaBoard + "<td align=`"center`" style=`"width=200;`" ><b>Status</b></td>" +"`r`n"
$global:oaBoard = $global:oaBoard + "<td align=`"center`" style=`"width=200;`" ><b>Message</b></td>" +"`r`n"
$global:oaBoard = $global:oaBoard + "</tr>" + "`r`n"

$expRequest = $expHeader

$bCount =0
foreach($mbMailbox in $Mailboxes){
	$bCount++
	$expRequest = $expRequest + "<Mailbox xmlns=`"http://schemas.microsoft.com/exchange/services/2006/types`"><EmailAddress>$mbMailbox</EmailAddress></Mailbox>" 
	if($bCount -eq $batchSize){
		$expRequest = $expRequest + "</Recipients><MailTipsRequested>OutOfOfficeMessage</MailTipsRequested></GetMailTips></soap:Body></soap:Envelope>"
		GetMailTips
		$expRequest = $expHeader
		$bCount = 0
	}
}
if($bCount -ne 0){
	$expRequest = $expRequest + "</Recipients><MailTipsRequested>OutOfOfficeMessage</MailTipsRequested></GetMailTips></soap:Body></soap:Envelope>"
	GetMailTips
}
$global:oaBoard = $global:oaBoard + "</table>"  + "  " 
$global:oaBoard | out-file "c:\offboard.htm"




