﻿## Get the Mailbox to Access from the 1st commandline argument

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

$tmTemplateEmail = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
$tmTemplateEmail.ItemClass = "IPM.Note.Rules.ReplyTemplate.Microsoft";
$tmTemplateEmail.IsAssociated = $true;
$tmTemplateEmail.Subject = "Recipient of your Email action required";
$htmlBodyString = "Hello,<p>Thanks for your Email we only answer emails enqiures from coperates email domains";
$tmTemplateEmail.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody($htmlBodyString);
$PidTagReplyTemplateId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65C2, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$tmTemplateEmail.SetExtendedProperty($PidTagReplyTemplateId, [System.Guid]::NewGuid().ToByteArray());
$tmTemplateEmail.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox);
$nrNewInboxRule = New-Object Microsoft.Exchange.WebServices.Data.Rule 
$nrNewInboxRule.DisplayName = "Auto Reply Rule";
$nrNewInboxRule.Conditions.ContainsSenderStrings.Add("@yahoo.com")
$nrNewInboxRule.Actions.ServerReplyWithMessage = $tmTemplateEmail.Id;
$nrNewInboxRule.Exceptions.ContainsSubjectStrings.Add("RE:");
$nrNewInboxRule.Exceptions.ContainsSubjectStrings.Add("FW:");
$cnCreateNewRule = New-Object Microsoft.Exchange.WebServices.Data.createRuleOperation[] 1
$cnCreateNewRule[0] = $nrNewInboxRule
$service.UpdateInboxRules($cnCreateNewRule,$true);