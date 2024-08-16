$Script:rptCollection = @()  

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
  
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $SmtpAddress) 

function Process-Mailbox{
	param (
	        $SmtpAddress = "$( throw 'SMTPAddress is a mandatory Parameter' )"
		  )
	process{
		Write-Host ("Processing Mailbox : " + $SmtpAddress)
		$rptObj = "" | select MailboxName,TotalItems,MovedByOutlook,NoSCL,SCLneg1,SCL0,SCL1,SCL2,SCL3,SCL4,SCL5,SCL6,SCL7,SCL8,SCL9
		$rptObj.MovedByOutlook = 0
		$rptObj.NoSCL = 0
		$rptObj.SCLneg1 = 0
		$rptObj.SCL0 = 0
		$rptObj.SCL1 = 0
		$rptObj.SCL2 = 0
		$rptObj.SCL3 = 0
		$rptObj.SCL4 = 0
		$rptObj.SCL5 = 0
		$rptObj.SCL6 = 0
		$rptObj.SCL7 = 0
		$rptObj.SCL8 = 0
		$rptObj.SCL9 = 0
		# Bind to the Junk Email Folder
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail,$SmtpAddress)   
		$JunkEmail = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		$rptObj.MailboxName = $SmtpAddress
		$rptObj.TotalItems = $JunkEmail.TotalCount
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
		$PidTagContentFilterSpamConfidenceLevel = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x4076, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
		$PidLidSpamOriginalFolder = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common,0x859C, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
		$psPropset.Add($PidTagContentFilterSpamConfidenceLevel)
		$psPropset.Add($PidLidSpamOriginalFolder)
		#Define ItemView to retrive just 1000 Items    
		$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
		$ivItemView.PropertySet = $psPropset
		$fiItems = $null    
		do{    
		    $fiItems = $service.FindItems($JunkEmail.Id,$ivItemView)    
		    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
		    foreach($Item in $fiItems.Items){
				$OrigFldVal = $null
				if($Item.TryGetProperty($PidLidSpamOriginalFolder,[ref]$OrigFldVal)){
					$rptObj.MovedByOutlook +=1
				}
				$SCLVal = $null;
				if($Item.TryGetProperty($PidTagContentFilterSpamConfidenceLevel,[ref]$SCLVal)){
					switch($SCLVal){
						{$_ -lt 0}{$rptObj.SCLneg1 += 1}
						0 { $rptObj.SCL0 += 1 }
						1 { $rptObj.SCL1 += 1 }
						2 { $rptObj.SCL2 += 1 }
						3 { $rptObj.SCL3 += 1 }				
						4 { $rptObj.SCL4 += 1 }
						5 { $rptObj.SCL5 += 1 }
						6 { $rptObj.SCL6 += 1 }
						7 { $rptObj.SCL7 += 1 }
						8 { $rptObj.SCL8 += 1 }				
						9 { $rptObj.SCL9 += 1 }
					}
				}
				else{
					$rptObj.NoSCL +=1
				}				         
		    }    
		    $ivItemView.Offset += $fiItems.Items.Count    
		}while($fiItems.MoreAvailable -eq $true) 
		$Script:rptCollection += $rptObj
	}
}

Import-Csv -Path $args[0] | ForEach-Object{
	if($service.url -eq $null){
		$service.AutodiscoverUrl($_.SmtpAddress,{$true}) 
		"Using CAS Server : " + $Service.url 
	}
	Try{
		Process-Mailbox -SmtpAddress $_.SmtpAddress
	}
	catch{
		LogWrite("Error processing Mailbox : " + $_.SmtpAddress + $_.Exception.Message.ToString())
	}
}
$Script:rptCollection | Export-Csv -NoTypeInformation -Path c:\temp\sclReport.csv
