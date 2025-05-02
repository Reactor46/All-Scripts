## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]
$DaysBack = $args[1]

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
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 

# Bind to the Inbox Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$Range = [system.DateTime]::Now.AddDays(-$DaysBack).ToString("MM/dd/yyyy") + ".." + [system.DateTime]::Now.ToString("MM/dd/yyyy")   
$AQSString1 = "System.Message.DateReceived:" + $Range   
$AQSString2 = "(" + $AQSString1 + ") AND (isread:false)" 
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)  
$allMail = $Inbox.FindItems($AQSString1,$ivItemView)
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
$unread = $Inbox.FindItems($AQSString2,$ivItemView)
$rptObj = "" | Select MailboxName,TotalCount,Unread,PercentUnread,PercentGraph

#Write-Host ("All Mail " + $allMail.TotalCount) 
#Write-Host ("Unread " + $unread.TotalCount) 
$rptObj.MailboxName = $MailboxName
$rptObj.TotalCount = $allMail.TotalCount
$rptObj.Unread = $unread.TotalCount 
$PercentUnread = [Math]::round((($unread.TotalCount/$allMail.TotalCount) * 100))
$rptObj.PercentUnread = $PercentUnread 
#Write-Host ("Percent Unread " + $PercentUnread)
$ureadGraph = ""
for($intval=0;$intval -lt 100;$intval+=2){
	if($PercentUnread -ge $intval){
		$ureadGraph += "▓"
	}
	else{		
		$ureadGraph += "░"
	}
}
#Write-Host $ureadGraph
$rptObj.PercentGraph = $ureadGraph
$rptObj | fl
