## Get the Mailbox to Access from the 1st commandline argument
$MailboxName = $args[0]

$CalendarSetting = $args[1]

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
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013  
  
## Create Exchange Service Object  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
  
## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
  
#Credentials Option 1 using UPN for the windows Account  
$psCred = Get-Credential  
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())  
$service.Credentials = $creds      
#$service.TraceEnabled = $true
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

#Check for existing Item  
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)     
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,"IPM.Configuration.OWA.ViewStateConfiguration")  
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)  
$ivItemView.Traversal =  [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated  
$fiResults = $service.FindItems($folderid,$SfSearchFilter,$ivItemView)  
if($fiResults.Items.Count -eq 0){  
    Write-Host ("No Config Item found, create new Item")  
    $UMConfig = New-Object Microsoft.Exchange.WebServices.Data.UserConfiguration -ArgumentList $service  
    $UMConfig.Save("OWA.ViewStateConfiguration",$folderid)  
}  
else{  
    Write-Host ("Existing Config Item Found");  
}  
#Specify the Root folder where the FAI Item is  
$UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "OWA.ViewStateConfiguration", $folderid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)  
if($UsrConfig.Dictionary.ContainsKey("CalendarViewTypeDesktop")){  
    $UsrConfig.Dictionary["CalendarViewTypeDesktop"] = $CalendarSetting  
}  
else{  
    $UsrConfig.Dictionary.Add("CalendarViewTypeDesktop",$CalendarSetting)  
}
$UsrConfig.Update()  
"Mailbox Updated"