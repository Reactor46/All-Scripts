## Get the Mailbox to Access from the 1st commandline argument

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
  
   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 

$Script:rptCollection = @()

function getFolderItemCounts($MailboxName){
	"Processing Mailbox : " + $MailboxName
	if($service.url -eq $null){
		$service.AutodiscoverUrl($MailboxName,{$true})  
		"Using CAS Server : " + $Service.url 
	}
	# Bind to the Inbox Folder
	$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
	$Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

	$rptObject = "" | Select MailboxName,FolderName,Olderthen2009,ItemCount2009,ItemCount2010,ItemCount2011,ItemCount2012
	$rptObject.MailboxName = $MailboxName
	$rptObject.FolderName = $Folder.DisplayName
	$AQSString = "System.Message.DateReceived:" 

	#Define ItemView to retrive just 1 Item    
	$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)  
	$fiItems = $service.FindItems($Folder.Id,($AQSString + "<01/01/2009"),$ivItemView)
	"Older then 2009 Total Items : " + $fiItems.TotalCount
	$rptObject.Olderthen2009 = $fiItems.TotalCount
	$Range = "01/01/2009..01/01/2010" 
	$fiItems = $service.FindItems($Folder.Id,($AQSString + $Range),$ivItemView)
	"2009 Total Items : " + $fiItems.TotalCount
	$rptObject.ItemCount2009 = $fiItems.TotalCount
	$Range = "01/01/2010..01/01/2011" 
	$fiItems = $service.FindItems($Folder.Id,($AQSString + $Range),$ivItemView)
	"2010 Total Items : " + $fiItems.TotalCount
	$rptObject.ItemCount2010 = $fiItems.TotalCount
	$Range = "01/01/2011..01/01/2012" 
	$fiItems = $service.FindItems($Folder.Id,($AQSString + $Range),$ivItemView)
	"2011 Total Items : " + $fiItems.TotalCount
	$rptObject.ItemCount2011 = $fiItems.TotalCount
	$Range = "01/01/2012..01/01/2013" 
	$fiItems = $service.FindItems($Folder.Id,($AQSString + $Range),$ivItemView)
	$rptObject.ItemCount2012 = $fiItems.TotalCount
	"2012 Total Items : " + $fiItems.TotalCount
	$Script:rptCollection += $rptObject
}

if((Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"}) -eq $null){
	$rpRemotePowershell = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -credential $psCred  -Authentication Basic -AllowRedirection  
	$importresults = Import-PSSession $rpRemotePowershell 
}
Get-Mailbox -ResultSize Unlimited | ForEach-Object{
	getFolderItemCounts($_.PrimarySMTPAddress)
}

$tableStyle = @"
<style>
BODY{background-color:white;}
TABLE{border-width: 1px;
  border-style: solid;
  border-color: black;
  border-collapse: collapse;
}
TH{border-width: 1px;
  padding: 10px;
  border-style: solid;
  border-color: black;
  background-color:#66CCCC
}
TD{border-width: 1px;
  padding: 2px;
  border-style: solid;
  border-color: black;
  background-color:white
}
</style>
"@
  
$body = @"
<p style="font-size:25px;family:calibri;color:#ff9100">
$TableHeader
</p>
"@

$Script:rptCollection | ConvertTo-HTML -head $tableStyle –body $body | Out-File c:\temp\ItemCounts.htm