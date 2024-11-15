$MailboxName = $args[0]
$StartDate = $args[1]


## Get the Mailbox to Access from the 1st commandline argument

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
   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
if($StartDate -ne $null){
	$AQSString = "received:>" + $StartDate.ToString("yyyy-MM-dd") 
	$AQSString
}
# Bind to the Inbox Folder
$Script:rptcollection = @{}
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
# Bind to the SentItems Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$MailboxName)   
$SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

function Process-Folder{  
    param (  
            $FolderId = "$( throw 'SMTPAddress is a mandatory Parameter' )",
			$IsSentItems = "$( throw 'SMTPAddress is a mandatory Parameter' )"
          )  
    process{ 
		$cnvItemView = New-Object Microsoft.Exchange.WebServices.Data.ConversationIndexedItemView(1000)
		$cnvs = $null;
		do
		{
			$cnvs = $service.FindConversation($cnvItemView, $FolderId,$AQSString);
			"Number of Conversation returned " + $cnvs.Count
			foreach($cnv in $cnvs){
				$rptobj = $cnv | select Topic,LastDeliveryTime,UniqueSenders,UniqueRecipients,InboxMessageCount,GlobalMessageCount,InboxMessageSize,SentItemsMessageCount,SentItemsMessageSize,ParticipationRate	
				if($Script:rptcollection.Contains($cnv.Id.UniqueId)-eq $false){
					if($IsSentItems){
						$rptobj.SentItemsMessageCount = $cnv.MessageCount
						$rptobj.SentItemsMessageSize = $cnv.Size
						$rptobj.InboxMessageCount = 0
						$rptobj.InboxMessageSize = 0
						$rptobj.LastDeliveryTime = $cnv.LastDeliveryTime
						$rptobj.UniqueSenders = $cnv.GlobalUniqueSenders
						$rptobj.UniqueRecipients = $cnv.GlobalUniqueRecipients
					}
					else{
						$rptobj.InboxMessageCount = $cnv.MessageCount
						$rptobj.InboxMessageSize = $cnv.Size
						$rptobj.SentItemsMessageCount = 0
						$rptobj.SentItemsMessageSize = 0
						$rptobj.LastDeliveryTime = $cnv.LastDeliveryTime
						$rptobj.UniqueSenders = $cnv.GlobalUniqueSenders
						$rptobj.UniqueRecipients = $cnv.GlobalUniqueRecipients
					}
					$Script:rptcollection.Add($cnv.Id.UniqueId,$rptobj)
				}
				else{
					if($IsSentItems){
						$Script:rptcollection[$cnv.Id.UniqueId].SentItemsMessageCount = $cnv.MessageCount
						$Script:rptcollection[$cnv.Id.UniqueId].SentItemsMessageSize = $cnv.Size
					}
					else{
						$Script:rptcollection[$cnv.Id.UniqueId].InboxMessageCount = $cnv.MessageCount
						$Script:rptcollection[$cnv.Id.UniqueId].InboxMessageSize = $cnv.Size
					}
					 
				}				
			
			}
			$cnvItemView.Offset += $cnvs.Count
		}while($cnvs.Count -gt 0)
	}
}
Process-Folder -FolderId $Inbox.Id -IsSentItems $false
Process-Folder -FolderId $SentItems.Id -IsSentItems $true
foreach($value in $Script:rptcollection.Values){
	if($value.GlobalMessageCount -gt 0 -band $value.SentItemsMessageCount -gt 0){
		$value.ParticipationRate = [Math]::round((($value.SentItemsMessageCount/$value.GlobalMessageCount) * 100))
	}
	else{
		$value.ParticipationRate = 0
	}
}
$Script:rptcollection.Values | Where-Object {$_.InboxMessageCount -gt 3} | Sort-Object ParticipationRate -Descending | Export-Csv -NoTypeInformation -Path c:\Temp\$MailboxName-cnvStats.csv
 