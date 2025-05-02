$dbxfile = "c:\temp\inbox.dbx"
$dllpath = "C:\temp\psdbxparser.dll"
$MailboxName = "mailbox@exdev.msgdevelop.com"

$casserverName = "exserver"
$userName = "username"
$password = "password"
$domain = "domain"

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

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
$uri=[system.URI] ("https://" + $casserverName + "/ews/exchange.asmx")
$service.Url = $uri
$service.Credentials = New-Object System.Net.NetworkCredential($username,$password,$domain)


$casuri = "https://" + $casserverName  + "/ews/exchange.asmx"
$uri=[system.URI] $casuri
$service.Url = $uri

$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$TargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)


[Reflection.Assembly]::LoadFile($dllpath)
$dbx = new-object psdbxparser.DBX
$mcount = $dbx.Parse($dbxfile)
if ($mcount -gt 0){
	for($iloop=0;$iloop -lt $mcount;$iloop++){
		$msgString = $dbx.Extract($iloop)
		$emUploadEmail = new-object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
		[byte[]]$bdBinaryData1 = [System.Text.Encoding]::ASCII.GetBytes($msgString)
           	$emUploadEmail.MimeContent = new-object Microsoft.Exchange.WebServices.Data.MimeContent("us-ascii", $bdBinaryData1);
	        $PR_Flags = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3591, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
                $emUploadEmail.SetExtendedProperty($PR_Flags,"1")
                if ($msgString.Indexof("message/delivery-status") -eq -1){
                	$emUploadEmail.Save($TargetFolder.id)
			"Uploaded : " + $iloop
		}
		

	}

}