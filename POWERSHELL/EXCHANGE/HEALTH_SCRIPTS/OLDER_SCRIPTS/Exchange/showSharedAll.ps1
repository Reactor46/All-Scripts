## Get the Mailbox to Access from the 1st commandline argument
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
  
#CAS URL Option 1 Autodiscover  

   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 

# Bind to the Inbox Folder
#Define Function to convert String to FolderPath  
function ConvertToString($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++  
    }  
    return $Val1Text  
} 

function Process-Mailbox{
	param (
	        $SmtpAddress = "$( throw 'SMTPAddress is a mandatory Parameter' )"
		  )
	process{

		#Define Extended properties  
		$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
		$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$SmtpAddress)
		#Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
		$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
		#Deep Transval will ensure all folders in the search path are returned  
		$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
		$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
		#Add Properties to the  Property Set  
		$psPropertySet.Add($PR_Folder_Path);  
		$fvFolderView.PropertySet = $psPropertySet;  
		#The Search filter will exclude any Search Folders  
		$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")  
		$fiResult = $null
		#Process the Mailbox Root Folder
		$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::Permissions);
		#The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
		do {  
		    $fiResult = $Service.FindFolders($folderidcnt,$sfSearchFilter,$fvFolderView)  
		    foreach($ffFolder in $fiResult.Folders){  
		        $foldpathval = $null  
		        #Try to get the FolderPath Value and then covert it to a usable String   
		        if ($ffFolder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval))  
		        {  
		            $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
		            $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
		            $hexString = $hexArr -join ''  
		            $hexString = $hexString.Replace("FEFF", "5C00")  
		            $fpath = ConvertToString($hexString)  
		        }  
		        "Processing FolderPath : " + $fpath  
				$ffFolder.Load($psPropset)
				foreach($Permission in  $ffFolder.Permissions){
					if($Permission.UserId.StandardUser -eq [Microsoft.Exchange.WebServices.Data.StandardUser]::Default){
						if($Permission.DisplayPermissionLevel -ne "None"){
							$rptObj = "" | Select Mailbox,FolderPath,FolderType,PermissionLevel
							$rptObj.Mailbox = $SmtpAddress
							$rptObj.FolderPath = $fpath
							$rptObj.FolderType = $ffFolder.FolderClass
							$rptObj.PermissionLevel = $Permission.DisplayPermissionLevel
							$Script:rptCollection += $rptObj
						} 
					}
				}

		    } 
		    $fvFolderView.Offset += $fiResult.Folders.Count
		}while($fiResult.MoreAvailable -eq $true) 
	}
}
Import-Csv -Path $args[0] | ForEach-Object{
	if($service.url -eq $null){
		$service.AutodiscoverUrl($_.SmtpAddress,{$true}) 
		"Using CAS Server : " + $Service.url 
	}  
	Process-Mailbox -SmtpAddress $_.SmtpAddress
}
$Script:rptCollection
