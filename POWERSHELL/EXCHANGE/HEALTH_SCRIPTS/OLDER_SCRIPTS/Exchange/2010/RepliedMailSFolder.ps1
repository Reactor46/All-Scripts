## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]
$SearchFilterName = "Replied To Mail"

## Load Managed API dll  

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
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  
  
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
   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
 
function HexStringToByteArray($HexString)
{
	$ByteArray =  New-Object Byte[] ($HexString.Length/2);
  	for ($i = 0; $i -lt $HexString.Length; $i += 2)
	{
		 $ByteArray[$i/2] = [Convert]::ToByte($HexString.Substring($i, 2), 16)
	} 
 	Return @(,$ByteArray)

}

$PR_LAST_VERB_EXECUTED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x1081, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)  
$sfItemSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_LAST_VERB_EXECUTED,102) 
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
$SearchFolder = new-object Microsoft.Exchange.WebServices.Data.SearchFolder($service);
$rfRootFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)   
$searchFolder.SearchParameters.RootFolderIds.Add($rfRootFolderId);
$searchFolder.SearchParameters.Traversal = [Microsoft.Exchange.WebServices.Data.SearchFolderTraversal]::Deep;
$searchFolder.SearchParameters.SearchFilter = $sfItemSearchFilter
$searchFolder.DisplayName = $SearchFilterName
$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SearchFolders,$MailboxName)   
$sfRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$searchFolder.Save($sfRoot.Id);
$PR_EXTENDED_FOLDER_FLAGS = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x36DA, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
$psPropset.Add($PR_EXTENDED_FOLDER_FLAGS)
$searchFolder.Load($psPropset)
$exVal = $null
if($SearchFolder.TryGetProperty($PR_EXTENDED_FOLDER_FLAGS,[ref]$exVal)){
	 $newVal = "010402001000" + [System.BitConverter]::ToString($exVal).Replace("-","")
	 $byteVal = HexStringToByteArray($newVal)
	 $SearchFolder.SetExtendedProperty($PR_EXTENDED_FOLDER_FLAGS,$byteVal)
	 $searchFolder.Update()
}
"Folder Created"
