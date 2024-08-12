## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

#MessageProperties

$VoiceMailSuject = "Voice Mail from .."
$Mp3FileName = "c:\temp\vmMail.mp3"
$duration = 30
$voiceMailFrom = "FirstName LastName"
$callerId = "1234"
$jobTitle = "Tech"
$Company = "CompanyName"
$workNumber = "11111-11111"
$emailAddress = "test@domain.com"
$SipAddress = "sip:test@domain.com"
$ToAddress = "user@domain.com"

$fileInfo = Get-Item $Mp3FileName

$BodyHtml = "<html><head><META HTTP-EQUIV=`"Content-Type`" CONTENT=`"text/html; charset=us-ascii`">"
$BodyHtml += "<style type=`"text/css`"> a:link { color: #3399ff; } a:visited { color: #3366cc; } a:active { color: #ff9900; } </style>"
$BodyHtml += "</head><body><style type=`"text/css`"> a:link { color: #3399ff; } a:visited { color: #3366cc; } a:active { color: #ff9900; } </style>"
$BodyHtml += "<div style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`"><div id=`"UM-call-info`" lang=`"en`">"
$BodyHtml += "<div style=`"font-family: Arial; font-size: 10pt; color:#000066; font-weight: bold;`">You received a voice mail from " + $voiceMailFrom + " at " + $callerId + "</div>"
$BodyHtml += "<br><table border=`"0`" width=`"100%`"><tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
$BodyHtml += "Caller-Id:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
$BodyHtml += "<a style=`"color: #3399ff; `" dir=`"ltr`" href=`"tel:" + $callerId + "`">" + $callerId + "</a></td></tr><tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
$BodyHtml += "Job Title:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
$BodyHtml += $jobTitle + "</td></tr><tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
$BodyHtml += "Company:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
$BodyHtml += $Company + "</td></tr><tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
$BodyHtml += "Work:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
$BodyHtml += "<a style=`"color: #3399ff; `" dir=`"ltr`" href=`"tel:&#43;" + $workNumber + "`">&#43;" + $workNumber + "</a></td></tr><tr><td width=`"12px`"></td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">"
$BodyHtml += "E-mail:</td><td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
$BodyHtml += "<a style=`"color: #3399ff; `" href=`"mailto:" + $emailAddress + "`">" + $voiceMailFrom + "</a></td></tr><tr><td width=`"12px`">"
$BodyHtml += "</td><td width=`"28%`" nowrap=`"`" style=`"font-family: Tahoma; color: #686a6b; font-size:10pt;border-width: 0in;`">IM Address:</td>"
$BodyHtml += "<td width=`"72%`" style=`"font-family: Tahoma; background-color: #ffffff; color: #000000; font-size:10pt;`">"
$BodyHtml += "<a style=`"color: #3399ff; `" href=`"" + $SipAddress + "`">" + $emailAddress + "</a></td></tr></table></div></div></body></html>"



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
  
#$uri=[system.URI] "https://192.168.0.6/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#Extended Prop Definition

$PidTagVoiceMessageAttachmentOrder = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6805, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$PstnCallbackTelephoneNumberVal = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::UnifiedMessaging,"PstnCallbackTelephoneNumber", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$PidTagSIPAddress = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x5FE5, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$PidTagCallIdVal = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6806, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$PidTagSenderTelephoneNumber = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6802, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$PidTagVoiceMessageDuration =  new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6801, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PidTagVoiceMessageSenderName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6803, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
 
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
$EmailMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service  
$EmailMessage.Subject = $VoiceMailSuject
#Add Recipients    
$EmailMessage.ToRecipients.Add($ToAddress)  
$EmailMessage.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody  
$EmailMessage.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML  
$EmailMessage.Body.Text = $BodyHtml
$EmailMessage.ItemClass = "IPM.Note.Microsoft.Voicemail.UM.CA"
$vmAttachment = $EmailMessage.Attachments.AddFileAttachment($Mp3FileName)
$vmAttachment.ContentType = "audio/mp3";
$EmailMessage.SetExtendedProperty($PidTagVoiceMessageAttachmentOrder,$fileInfo.Name)
$EmailMessage.SetExtendedProperty($PstnCallbackTelephoneNumberVal,$workNumber)
$EmailMessage.SetExtendedProperty($PidTagSIPAddress,$SipAddress)
$EmailMessage.SetExtendedProperty($PidTagCallIdVal,$callerId)
$EmailMessage.SetExtendedProperty($PidTagSenderTelephoneNumber,$workNumber)
$EmailMessage.SetExtendedProperty($PidTagVoiceMessageDuration,$duration)
$EmailMessage.SetExtendedProperty($PidTagVoiceMessageSenderName,$voiceMailFrom)
$EmailMessage.SendAndSaveCopy()  