
$MailboxName = "user@domain.com"  
       
$cred = New-Object System.Net.NetworkCredential("user@domamin.com","password")   
  
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"  
[void][Reflection.Assembly]::LoadFile($dllpath)   
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)   
$service.TraceEnabled = $false  
  
$service.Credentials = $cred  
$service.autodiscoverurl($MailboxName,{$true})   

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
     
	
	$expRequest = @"
<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
<soap:Header><RequestServerVersion Version="Exchange2010_SP2" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" />
</soap:Header>
<soap:Body>
<GetPasswordExpirationDate xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"><MailboxSmtpAddress>$MailboxName</MailboxSmtpAddress>
</GetPasswordExpirationDate></soap:Body></soap:Envelope>
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
$ExpNodes = @($rdResponseDocument.getElementsByTagName("PasswordExpirationDate"))   
$ExpNodes[0].'#text'

