$MailboxName = "sender@domain.com"

$Mailboxes = @("user1@domain.com","user2@domain.com")   
  
$cred = New-Object System.Net.NetworkCredential("user1@domain.com","password@#")   

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$service.TraceEnabled = $false

$service.Credentials = $cred
$service.autodiscoverurl($MailboxName,{$true})


$expRequest = @"
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

foreach($mbMailbox in $Mailboxes){
	$expRequest = $expRequest + "<Mailbox xmlns=`"http://schemas.microsoft.com/exchange/services/2006/types`"><EmailAddress>$mbMailbox</EmailAddress></Mailbox>" 
}

$expRequest = $expRequest + "</Recipients><MailTipsRequested>OutOfOfficeMessage</MailTipsRequested></GetMailTips></soap:Body></soap:Envelope>"
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
		$RecipientNodes[$ic].EmailAddress + " : In the Office"
	}
	else{
		$RecipientNodes[$ic].EmailAddress + " : Out of Office"
	}
}


