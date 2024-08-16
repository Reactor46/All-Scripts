$MailboxName = "user@domain.com"
$TimeFrame = (Get-Date).AddDays(-1)
$ReportAddress = "user@domain.com"

$rptcollection = @()
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

function ConvertId($EWSid){  
    $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId    
    $aiItem.Mailbox = $MailboxName    
    $aiItem.UniqueId = $EWSid 
    $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::EWSId;    
    return $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId)   
}  

$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$folderid)
$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000) 
$SfClass = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,"IPM.Note.Microsoft.Approval.Request")
$PidNameApprovalRequestor = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::InternetHeaders,"x-ms-exchange-organization-approval-requestor",[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
$VerbResponse = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common,0x8524,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
$PR_NORMALIZED_SUBJECT = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E1D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);     
$PR_REPORT_TAG = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0031,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);

$SfClass = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,"IPM.Note.Microsoft.Approval.Request")
$Sfgt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $TimeFrame)

$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfCollection.add($Sfgt)
$sfCollection.add($SfClass)

$Propset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$Propset.add($PidNameApprovalRequestor)
$Propset.add($PR_NORMALIZED_SUBJECT) 
$ivItemView.PropertySet = $Propset
#define Table
$rpReport = $rpReport + "<table><tr bgcolor=`"#95aedc`">"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:15%;`" ><b>Recieved</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>From</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>To</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:40%;`" ><b>Subject</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:5%;`" ><b>Size(KB)</b></td>"
$rpReport = $rpReport + "</tr>"
#end 


do{  
    $fiResults = $Inbox.findItems($sfCollection,$ivItemView)
    foreach($Item in $fiResults.Items){ 
	$fromVal = $null
	if($Item.TryGetProperty($PidNameApprovalRequestor,[ref]$fromVal)){
		"From : " + $propval
	}
	$NormalSubject = $null
	[Void]$Item.TryGetProperty($PR_NORMALIZED_SUBJECT,[ref]$NormalSubject)
	"Recieved : " + $Item.DateTimeReceived
	"Subject : " + $NormalSubject
	$Item.Load()
	$Item.attachments[0].Load()
	$rpReport = $rpReport + " <tr>" + " "
	$rpReport = $rpReport + "<td>" + $Item.DateTimeReceived.ToString() + "</td>" + " "
	$rpReport = $rpReport + "<td>" + $fromVal + "</td>" + " "
	$rpReport = $rpReport + "<td>" + $Item.Attachments[0].Item.ToRecipients[0].Address + "</td>" + " "
	$cnvId = ConvertId($Item.Id.UniqueId)
	$rpReport = $rpReport + "<td><a href=`"outlook:" + $cnvId.UniqueId + "`">" + $NormalSubject + "</a></td>" + " "
	$rpReport = $rpReport + "<td>" + [Math]::Round($Item.Size/1KB,2) + "</td>" + " "
	$rpReport = $rpReport + "</tr>" + " "
	
	}  
    $ivItemView.Offset += $fiResults.Items.Count  
}while($fiResults.MoreAvailable -eq $true)

$rpReport = $rpReport + "</table>" + " "
$EmailMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($service) 
#Set the Subject  
$EmailMessage.Subject = "Moderation Pending Approvals"
#Add Recipients  
$EmailMessage.ToRecipients.Add($ReportAddress)
$EmailMessage.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody
$EmailMessage.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML
$EmailMessage.Body.Text = $rpReport
$EmailMessage.SendAndSaveCopy()

