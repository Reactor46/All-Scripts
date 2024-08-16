## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]
$AsFolderReport = @()

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
# Bind to the MsgFolderRoot folder  
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
$MsgRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

function ConvertId{    
	param (
	        $HexId = "$( throw 'HexId is a mandatory Parameter' )"
		  )
	process{
	    $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
	    $aiItem.Mailbox = $MailboxName      
	    $aiItem.UniqueId = $HexId   
	    $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId      
	    $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId) 
		return $convertedId.UniqueId
	}
}

function GetFolderPath{
	param (
		$EWSFolder = "$( throw 'Folder is a mandatory Parameter' )"
	)
	process{
		$foldpathval = $null  
		$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
		if ($EWSFolder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval))  
        {  
            $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
            $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
            $hexString = $hexArr -join ''  
	    $hexString = $hexString.Replace("EFBFBE", "5C")  
            $fpath = ConvertToString($hexString) 
	    return $fpath
        }  
	}
}
$fldMappingHash = @{}
#Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)  
#Deep Transval will ensure all folders in the search path are returned  
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;  
#The Search filter will exclude any Search Folders  
$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"ExchangeSyncData")  
$asFolderRoot = $Service.FindFolders($MsgRoot.Id,$sfSearchFilter,$fvFolderView)  
if($asFolderRoot.Folders.Count -eq 1){
	#Define Function to convert String to FolderPath  
	function ConvertToString($ipInputString){  
	    $Val1Text = ""  
	    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
	            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
	            $clInt++  
	    }  
	    return $Val1Text  
	} 
	
	#Define Extended properties  
	$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
	#Define the FolderView used for Export should not be any larger then 1000 folders due to throttling  
	$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
	#Deep Transval will ensure all folders in the search path are returned  
	$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
	$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
	$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
	$CollectionIdProp = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x7C03, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
	$LastModifiedTime = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3008, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
	#Add Properties to the  Property Set  
	$psPropertySet.Add($PR_Folder_Path);  
	$psPropertySet.Add($CollectionIdProp);
	$psPropertySet.Add($LastModifiedTime);	
	$fvFolderView.PropertySet = $psPropertySet;  
	#The Search filter will exclude any Search Folders  
	$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")  
	$fiResult = $null  
	#The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
	do {  
	    $fiResult = $Service.FindFolders($asFolderRoot.Folders[0].Id,$sfSearchFilter,$fvFolderView)  
	    foreach($ffFolder in $fiResult.Folders){ 
			if(!$fldMappingHash.ContainsKey($ffFolder.Id.UniqueId)){
				$fldMappingHash.Add($ffFolder.Id.UniqueId,$ffFolder)
			}
			$asFolderPath = ""
			$asFolderPath = (GetFolderPath -EWSFolder $ffFolder)
			"FolderPath : " + $asFolderPath
			$collectVal = $null
			if($ffFolder.TryGetProperty($CollectionIdProp,[ref]$collectVal)){
				$HexEntryId = [System.BitConverter]::ToString($collectVal).Replace("-","").Substring(2)
				$ewsFolderId = ConvertId -HexId ($HexEntryId.SubString(0,($HexEntryId.Length-2)))
				try{
					$fldReport = "" | Select Mailbox,Device,AsFolderPath,MailboxFolderPath,LastModified
					$fldReport.Mailbox = $MailboxName
					$fldReport.Device = $fldMappingHash[$ffFolder.ParentFolderId.UniqueId].DisplayName
					$fldReport.AsFolderPath = $asFolderPath
					$folderMapId= new-object Microsoft.Exchange.WebServices.Data.FolderId($ewsFolderId)   
					$MappedFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderMapId,$psPropertySet)
					$MappedFolderPath = (GetFolderPath -EWSFolder $MappedFolder)
					$fldReport.MailboxFolderPath = $MappedFolderPath
					$LastModifiedVal = $null
					if($ffFolder.TryGetProperty($LastModifiedTime,[ref]$LastModifiedVal)){
						Write-Host ("Last-Modified : " +  $LastModifiedVal.ToLocalTime().ToString())
						$fldReport.LastModified = $LastModifiedVal.ToLocalTime().ToString()
					}
					Write-Host $MappedFolderPath
					$AsFolderReport += $fldReport
				}
				catch{
					
				}
				$ewsFolderId
			}
			#Process folder here
	    } 
	    $fvFolderView.Offset += $fiResult.Folders.Count
	}while($fiResult.MoreAvailable -eq $true)  
	
	
}
$reportFile = "c:\temp\$MailboxName-asFolders.csv"
$AsFolderReport | Export-Csv -NoTypeInformation -Path $reportFile
Write-Host ("Report wrtten to " + $reportFile)
