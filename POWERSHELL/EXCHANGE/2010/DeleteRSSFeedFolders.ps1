## Get the Mailbox to Access from the 1st commandline argument  
  
$MailboxName = $args[0]  
  
## Load Managed API dll    
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"    
    
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
    
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)   
$PidTagAdditionalRenEntryIdsEx = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x36D9, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)  
$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
$psPropset.Add($PidTagAdditionalRenEntryIdsEx)  
  
# Bind to the NON_IPM_ROOT Root folder    
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)     
$NON_IPM_ROOT = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$psPropset)  
$binVal = $null;  
$AdditionalRenEntryIdsExCol = @{}  
if($NON_IPM_ROOT.TryGetProperty($PidTagAdditionalRenEntryIdsEx,[ref]$binVal)){  
    $hexVal = [System.BitConverter]::ToString($binVal).Replace("-","");  
    ##Parse Binary Value first word is Value type Second word is the Length of the Entry  
    $Sval = 0;  
    while(($Sval+8) -lt $hexVal.Length){  
        $PtypeVal = $hexVal.SubString($Sval,4)  
        $PtypeVal = $PtypeVal.SubString(2,2) + $PtypeVal.SubString(0,2)  
        $Sval +=12;  
        $PropLengthVal = $hexVal.SubString($Sval,4)  
        $PropLengthVal = $PropLengthVal.SubString(2,2) + $PropLengthVal.SubString(0,2)  
        $PropLength = [Convert]::ToInt64($PropLengthVal, 16)  
        $Sval +=4;  
        $ProdIdEntry = $hexVal.SubString($Sval,($PropLength*2))  
        $Sval += ($PropLength*2)  
        #$PtypeVal + " : " + $ProdIdEntry  
        $AdditionalRenEntryIdsExCol.Add($PtypeVal,$ProdIdEntry)   
    }     
}  
  
function ConvertFolderid($hexId){  
    $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId    
    $aiItem.Mailbox = $MailboxName    
    $aiItem.UniqueId = $hexId  
    $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId;    
    return $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EWSId)   
}  

#Define Function to convert String to FolderPath  
function ConvertToString($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++  
    }  
    return $Val1Text  
} 

  
if($AdditionalRenEntryIdsExCol.ContainsKey("8001")){  
    $siId = ConvertFolderid($AdditionalRenEntryIdsExCol["8001"])  
    $RSSFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId($siId.UniqueId.ToString())  
    $RSSFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$RSSFolderID) 
	if($RSSFolder -ne $null){
		"Deleteing : " + $RSSFolder.ChildFolderCount.ToString() + " Feeds"
		$RSSFolder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, $true);
	}
}  