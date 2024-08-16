## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

$KQL = "size>10485760";			 

$SearchableMailboxString = $MailboxName;

## Load Managed API dll  
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"  
  
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013  
  
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

##get folder Paths
#Define Function to convert String to FolderPath  
function ConvertToString($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++  
    }  
    return $Val1Text  
} 


function GetFolderPaths{
	param (
	        $rootFolderId = "$( throw 'rootFolderId is a mandatory Parameter' )",
			$Archive = "$( throw 'Archive is a mandatory Parameter' )"
		  )
	process{
	#Define Extended properties  
	$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
	$folderidcnt = $rootFolderId
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
	        "FolderPath : " + $fpath  
			if($Archive){
				$Script:FolderCache.Add($ffFolder.Id.UniqueId,"\Archive-Mailbox\" + $fpath);
			}
			else{
				$Script:FolderCache.Add($ffFolder.Id.UniqueId,$fpath);
			}
	    } 
	    $fvFolderView.Offset += $fiResult.Folders.Count
	}while($fiResult.MoreAvailable -eq $true)  
	}
}

$Script:FolderCache = New-Object system.collections.hashtable
GetFolderPaths -rootFolderId (new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)) -Archive $false  
GetFolderPaths -rootFolderId (new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot,$MailboxName)) -Archive $true 

$gsMBResponse = $service.GetSearchableMailboxes($SearchableMailboxString, $false);
$gsMBResponse
$msbScope = New-Object  Microsoft.Exchange.WebServices.Data.MailboxSearchScope[] $gsMBResponse.SearchableMailboxes.Length
$mbCount = 0;
foreach ($sbMailbox in $gsMBResponse.SearchableMailboxes)
{
    $msbScope[$mbCount] = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope($sbMailbox.ReferenceId, [Microsoft.Exchange.WebServices.Data.MailboxSearchLocation]::All);
    $mbCount++;
}
$smSearchMailbox = New-Object Microsoft.Exchange.WebServices.Data.SearchMailboxesParameters
$mbq =  New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery($KQL, $msbScope);
$mbqa = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery[] 1
$mbqa[0] = $mbq
$smSearchMailbox.SearchQueries = $mbqa;
$smSearchMailbox.PageSize = 100;
$smSearchMailbox.PageDirection = [Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next;
$smSearchMailbox.PerformDeduplication = $false;           
$smSearchMailbox.ResultType = [Microsoft.Exchange.WebServices.Data.SearchResultType]::PreviewOnly;
$srCol = $service.SearchMailboxes($smSearchMailbox);
$rptCollection = @()

if ($srCol[0].Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success)
{
	Write-Host ("Items Found " + $srCol[0].SearchResult.ItemCount)
    if ($srCol[0].SearchResult.ItemCount -gt 0)
    {                  
        do
        {
            $smSearchMailbox.PageItemReference = $srCol[0].SearchResult.PreviewItems[$srCol[0].SearchResult.PreviewItems.Length - 1].SortValue;
            foreach ($PvItem in $srCol[0].SearchResult.PreviewItems) {
				$rptObj = "" | select FolderPath,DateTimeReceived,Subject,Size
                if($Script:FolderCache.ContainsKey($PvItem.ParentId.UniqueId)){
 					$rptObj.FolderPath = $Script:FolderCache[$PvItem.ParentId.UniqueId]
					$rptObj.DateTimeReceived = $PvItem.ReceivedTime
					$rptObj.Subject = $PvItem.Subject
					$rptObj.Size = $PvItem.Size
				}else{
					$rptObj.DateTimeReceived = $PvItem.ReceivedTime
					$rptObj.Subject = $PvItem.Subject
					$rptObj.Size = $PvItem.Size
				}
				$rptObj
				$rptCollection+=$rptObj
            }                        
            $srCol = $service.SearchMailboxes($smSearchMailbox);
			Write-Host("Items Remaining : " + $srCol[0].SearchResult.ItemCount);
        } while ($srCol[0].SearchResult.ItemCount-gt 0 );
        
    }
    
}
$rptCollection | Export-Csv -NoTypeInformation -Path c:\temp\LaItemReport.csv