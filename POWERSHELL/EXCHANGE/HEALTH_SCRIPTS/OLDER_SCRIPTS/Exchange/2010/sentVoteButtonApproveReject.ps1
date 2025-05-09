$MailboxName = $args[0]
$SentTo = $args[1]

$VerbSetting = "" | Select DisableReplyAll,DisableReply,DisableForward,DisableReplyToFolder
$VerbSetting.DisableReplyAll = $false
$VerbSetting.DisableReply = $false
$VerbSetting.DisableForward = $false
$VerbSetting.DisableReplyToFolder = $false

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
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 


function Get-VerbStream{  
    param (  
            $VerbSettings = "$( throw 'VerbSettings is a mandatory Parameter' )"  
          )  
    process{  
	
	$Header = "02010600000000000000"
	$ReplyToAllHeader = "055265706C790849504D2E4E6F7465074D657373616765025245050000000000000000"
	$ReplyToAllFooter = "0000000000000002000000660000000200000001000000"
	$ReplyToHeader = "0C5265706C7920746F20416C6C0849504D2E4E6F7465074D657373616765025245050000000000000000"
	$ReplyToFooter = "0000000000000002000000670000000300000002000000"
	$ForwardHeader = "07466F72776172640849504D2E4E6F7465074D657373616765024657050000000000000000"
	$ForwardFooter = "0000000000000002000000680000000400000003000000"
	$ReplyToFolderHeader = "0F5265706C7920746F20466F6C6465720849504D2E506F737404506F737400050000000000000000"
	$ReplyToFolderFooter = "00000000000000020000006C00000008000000"
	$ApproveOption = "0400000007417070726F76650849504D2E4E6F74650007417070726F766500000000000000000001000000020000000200000001000000FFFFFFFF"
	$RejectOption= "040000000652656A6563740849504D2E4E6F7465000652656A65637400000000000000000001000000020000000200000002000000FFFFFFFF"
        $VoteOptionExtras = "0401055200650070006C00790002520045000C5200650070006C007900200074006F00200041006C006C0002520045000746006F007200770061007200640002460057000F5200650070006C007900200074006F00200046006F006C00640065007200000741007000700072006F00760065000741007000700072006F007600650006520065006A0065006300740006520065006A00650063007400"
	if($VerbSetting.DisableReplyAll){
		$DisableReplyAllVal = "00"
	}
	else{
		$DisableReplyAllVal = "01"
	}
	if($VerbSetting.DisableReply){
		$DisableReplyVal = "00"
	}
	else{
		$DisableReplyVal = "01"
	}
	if($VerbSetting.DisableForward){
		$DisableForwardVal = "00"
	}
	else{
		$DisableForwardVal = "01"
	}
	if($VerbSetting.DisableReplyToFolder){
		$DisableReplyToFolderVal = "00"
	}
	else{
		$DisableReplyToFolderVal = "01"
	}
	$VerbValue = $Header +	$ReplyToAllHeader + $DisableReplyAllVal + $ReplyToAllFooter + $ReplyToHeader + $DisableReplyVal +$ReplyToFooter + $ForwardHeader + $DisableForwardVal + $ForwardFooter + $ReplyToFolderHeader + $DisableReplyToFolderVal + $ReplyToFolderFooter + $ApproveOption  + $RejectOption + $VoteOptionExtras
	return $VerbValue
	}
}

function hex2binarray($hexString){
    $i = 0
    [byte[]]$binarray = @()
    while($i -le $hexString.length - 2){
        $strHexBit = ($hexString.substring($i,2))
        $binarray += [byte]([Convert]::ToInt32($strHexBit,16))
        $i = $i + 2
    }
    return ,$binarray
}



$VerbStreamProp = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common,0x8520, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)

$VerbSettingValue = get-VerbStream $VerbSetting

$EmailMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service  
$EmailMessage.Subject = "Message Subject"  
#Add Recipients    
$EmailMessage.ToRecipients.Add($SentTo)  
$EmailMessage.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody  
$EmailMessage.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML  
$EmailMessage.Body.Text = "Body" 
$EmailMessage.SetExtendedProperty($VerbStreamProp,(hex2binarray $VerbSettingValue))
$EmailMessage.SendAndSaveCopy()  


