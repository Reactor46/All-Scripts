## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

$Script:rptCollection = @()
 

###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
if (Test-Path $EWSDLL)
    {
    Import-Module $EWSDLL
    }
else
    {
    "$(get-date -format yyyyMMddHHmmss):"
    "This script requires the EWS Managed API 1.2 or later."
    "Please download and install the current version of the EWS Managed API from"
    "http://go.microsoft.com/fwlink/?LinkId=255472"
    ""
    "Exiting Script."
    exit
    }
	
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1  
  
## Create Exchange Service Object  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
  
## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
  
#Credentials Option 1 using UPN for the windows Account  
$psCred = Get-Credential  
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())  
$service.Credentials = $creds      
  
#Credentials Option 2  
#service.UseDefaultCredentials = $true  
  
## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
  
## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

## end code from http://poshcode.org/624
  
## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
  
#CAS URL Option 1 Autodiscover  
$service.AutodiscoverUrl($MailboxName,{$true})  
"Using CAS Server : " + $Service.url 

function DumpMailTips{	
    param (
	        $Mailboxes = "$( throw 'Folder Path is a mandatory Parameter' )"
		  )
	process{

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

$expRequest = $expRequest + "</Recipients><MailTipsRequested>All</MailTipsRequested></GetMailTips></soap:Body></soap:Envelope>"
$mbMailboxFolderURI = New-Object System.Uri($service.url)
$wrWebRequest = [System.Net.WebRequest]::Create($mbMailboxFolderURI)
$wrWebRequest.KeepAlive = $false;
$wrWebRequest.Headers.Set("Pragma", "no-cache");
$wrWebRequest.Headers.Set("Translate", "f");
$wrWebRequest.Headers.Set("Depth", "0");
$wrWebRequest.ContentType = "text/xml";
$wrWebRequest.ContentLength = $expRequest.Length;
$wrWebRequest.CookieContainer = New-Object System.Net.CookieContainer
$wrWebRequest.Timeout = 60000;
$wrWebRequest.Method = "POST";
$wrWebRequest.Credentials = $creds
$bqByteQuery = [System.Text.Encoding]::ASCII.GetBytes($expRequest);
$wrWebRequest.ContentLength = $bqByteQuery.Length;
$rsRequestStream = $wrWebRequest.GetRequestStream();
$rsRequestStream.Write($bqByteQuery, 0, $bqByteQuery.Length);
$rsRequestStream.Close();
$wrWebResponse = $wrWebRequest.GetResponse();
$rsResponseStream = $wrWebResponse.GetResponseStream()
$sr = new-object System.IO.StreamReader($rsResponseStream);
$rdResponseDocument = New-Object System.Xml.XmlDocument
$rdResponseDocument = New-Object System.Xml.XmlDocument
$rdResponseDocument.LoadXml($sr.ReadToEnd());
$Datanodes = @($rdResponseDocument.getElementsByTagName("m:MailTips"))

foreach($nodeVal in $Datanodes){
	$rptObj = "" | Select RecipientAddress,RecipientTypeDetails,PendingMailTips,OutOfOffice,CustomMailTip,MailboxFull,TotalMemberCount,ExternalMemberCount,MaxMessageSize,DeliveryRestricted,IsModerated
	$rptObj.RecipientAddress = $nodeVal.RecipientAddress.EmailAddress
	$rptObj.RecipientTypeDetails = $Script:emAray[$rptObj.RecipientAddress.ToLower()]
	$rptObj.PendingMailTips = $nodeVal.PendingMailTips."#text"
	$rptObj.OutOfOffice = $nodeVal.OutOfOffice.ReplyBody.Message
	$rptObj.CustomMailTip = $nodeVal.CustomMailTip."#text"
	$rptObj.MailboxFull = $nodeVal.MailboxFull."#text"
	$rptObj.TotalMemberCount = $nodeVal.TotalMemberCount."#text"
	$rptObj.ExternalMemberCount = $nodeVal.ExternalMemberCount."#text"
	$rptObj.MaxMessageSize = $nodeVal.MaxMessageSize."#text"
	$rptObj.DeliveryRestricted = $nodeVal.DeliveryRestricted."#text"
	$rptObj.IsModerated = $nodeVal.IsModerated."#text"
	$Script:rptCollection += $rptObj
}

}
}
$mbscn = @()
$Script:emAray = @{}
$rcps = Get-Recipient -ResultSize Unlimited 
$rcps | ForEach-Object {
	$Script:emAray.Add($_.PrimarySMTPAddress.ToString().ToLower(),$_.RecipientTypeDetails)
	$mbscn += $_.PrimarySMTPAddress.ToString()
	if($mbscn.count -gt 100){
		dumpmailtips -Mailboxes $mbscn
		$mbscn = @()
	}
}
if($mbscn.count -gt 0){
	dumpmailtips -Mailboxes $mbscn
}
$Script:rptCollection | Export-Csv -Path c:\temp\MailTipsDump.csv -NoTypeInformation