## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

## Load Managed API dll  
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"  
  
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
$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
# Bind to the Inbox Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

#Define ItemView to retrive just 1000 Items    
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)    
$fiItems = $service.FindItems($Inbox.Id,$ivItemView)    
[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
foreach($Item in $fiItems.Items){      
	#Process Item  
	"Processing : " + $Item.Subject
	$dupChk = @{}
	$RegExHtmlLinks = "<a href=\`"(.*?)\`">"
	$matchedItems = [regex]::matches($Item.Body, $RegExHtmlLinks,[system.Text.RegularExpressions.RegexOptions]::Singleline)
	foreach($Match in $matchedItems){
	    $SplitVal = $Match.Value.Split('"')
	    if($SplitVal.Count -gt 0){
	        $ParsedURI=[system.URI]$SplitVal[1]
			if($ParsedURI.Host -eq "meet.lync.com"){
				if(!$dupChk.Contains($ParsedURI.AbsoluteUri)){
					Write-Host -ForegroundColor Green	"LyncURL 	 : " + $ParsedURI.AbsoluteUri
					$dupChk.add($ParsedURI.AbsoluteUri,0)
				}
			} 
			if($ParsedURI.Host -eq "www.dropbox.com"){
				if(!$dupChk.Contains($ParsedURI.AbsoluteUri)){
					Write-Host -ForegroundColor Blue "DropBox 	 : " + $ParsedURI.AbsoluteUri
					$dupChk.add($ParsedURI.AbsoluteUri,0)
				}
			}  
	    }
	}
}    









