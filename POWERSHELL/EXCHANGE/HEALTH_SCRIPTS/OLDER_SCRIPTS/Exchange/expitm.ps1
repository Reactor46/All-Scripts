$MailboxName = "mailbox@domain.com"
$fileName = 'c:\exp.fts'
$cred = New-Object System.Net.NetworkCredential("username@domain.com","password")

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$service.TraceEnabled = $false

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind
$service.Credentials = $cred
$service.autodiscoverurl($MailboxName,{$true})

$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)

$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)
$findResults = $service.FindItems($folderid,$view)
$itemid = $findResults.Items[0].Id.Uniqueid

$expRequest = @"
<?xml version="1.0" encoding="utf-8" ?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
      xmlns:xsd="http://www.w3.org/2001/XMLSchema"
      xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
      xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2010_SP1" />
  </soap:Header>
  <soap:Body>
    <m:ExportItems>
      <m:ItemIds>
        <t:ItemId Id="$itemId"/>
      </m:ItemIds>
    </m:ExportItems>
  </soap:Body>
</soap:Envelope>
"@
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
$Datanodes = @($rdResponseDocument.getElementsByTagName("m:Data"))
if ($Datanodes.length -ne 0){
	$Data = [System.Convert]::FromBase64String($Datanodes[0].'#text')
	$fsFileStream = new-object system.io.filestream $fileName, ([io.filemode]::create), ([io.fileaccess]::write), ([io.fileshare]::none)
	$fsFileStream.Write($Data, 0, $Data.Length);
    $fsFileStream.Close();
}

