## Get the Mailbox to Access from the 1st command line argument. Enter the SMTP address of the mailbox
$MailboxName = $args[0]

## Load Managed API dll
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"

## Set Exchange Version. Use one of the following:
## Exchange2007_SP1, Exchange2010, Exchange2010_SP1, Exchange2010_SP2, Exchange2013, Exchange2013_SP1
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2

## Create Exchange Service Object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

## Set Credentials to use two options are available Option1 to use explicit credentials or Option 2 use the Default (logged On) credentials

## Credentials Option 1 using UPN for the Windows Account
$psCred = Get-Credential
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())
$service.Credentials = $creds

## Credentials Option 2
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

## End code from http://poshcode.org/624

## Set the URL of the CAS (Client Access Server) to use two options are available to use Autodiscover to find the CAS URL or Hardcode the CAS to use

## CAS URL Option 1 Autodiscover
$service.AutodiscoverUrl($MailboxName,{$true})
"Using CAS Server : " + $Service.url
  
## CAS URL Option 2 Hardcoded
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"
#$service.Url = $uri
  
## Optional section for Exchange Impersonation
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)

## Show Trace. Can be used for troubleshooting errors
#$service.traceenabled = $true

function ConvertId{
	param (
	        $OwaId = "$( throw 'OWAId is a mandatory Parameter' )"
		  )
	process{
	    $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId
	    $aiItem.Mailbox = $MailboxName
	    $aiItem.UniqueId = $OwaId
	    $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::OwaId
	    $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)
		return $convertedId.UniqueId
	}
}

## Get Empty Folders from EMS
get-mailboxfolderstatistics $MailboxName | Where-Object{$_.FolderType -eq "User Created" -band $_.ItemsInFolderAndSubFolders -eq 0} | ForEach-Object{
	# Bind to the Inbox Folder
	"Deleting Folder " + $_.FolderPath
	try{
		Add-Type -AssemblyName System.Web
		$urlEncodedId = [System.Web.HttpUtility]::UrlEncode($_.FolderId.ToString())
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId((Convertid $urlEncodedId))
		#$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId((Convertid $_.FolderId))
		$ewsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		if($ewsFolder.TotalCount -eq 0){
			$ewsFolder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
			"Folder Deleted"
		}
	}
	catch{
	$_.Exception.Message
	}
}