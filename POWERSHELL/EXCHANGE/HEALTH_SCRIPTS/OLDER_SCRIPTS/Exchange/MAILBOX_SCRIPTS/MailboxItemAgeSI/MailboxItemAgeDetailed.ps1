	## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

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

#Define ItemView to retrive just 1000 Items    
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
$ivItemView.PropertySet = $psPropset
$TotalSize = 0
$TotalItemCount = 0


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
$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)  
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
$fldhash = @{}
$minYear = (Get-Date).Year
$maxYear = (Get-Date).Year
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
		
		#Define ItemView to retrive just 1000 Items    
		$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
		$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
		$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
		$ivItemView.PropertySet = $psPropset
		
		$rptHash = @{}		
		$fiItems = $null    
		do{    
		    $fiItems = $service.FindItems($ffFolder.Id,$ivItemView)    
		    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
		    foreach($Item in $fiItems.Items){
				$dateVal = $null
				if($Item.TryGetProperty([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[ref]$dateVal )-eq $false){
					$dateVal = $Item.DateTimeCreated
				}
				if($rptHash.ContainsKey($dateVal.Year)){
					$rptHash[$dateVal.Year].TotalNumber += 1
					$rptHash[$dateVal.Year].TotalSize += [Int64]$Item.Size
				}
				else{
					$rptObj = "" | Select TotalNumber,TotalSize
					$rptObj.TotalNumber = 1
					$rptObj.TotalSize = [Int64]$Item.Size
					$rptHash.add($dateVal.Year,$rptObj)
					if($dateVal.Year -lt $minYear){$minYear = $dateVal.Year}
				}
		    }    
		    $ivItemView.Offset += $fiItems.Items.Count    
		}while($fiItems.MoreAvailable -eq $true)		
		$fldhash.add($fpath,$rptHash)
    } 
    $fvFolderView.Offset += $fiResult.Folders.Count
}while($fiResult.MoreAvailable -eq $true)

$rptCollection = @()
foreach($key in $fldhash.Keys){
	$fldReport = New-Object System.Object
	$fldReport | Add-Member -type NoteProperty -name FolderPath -value $key
	$fldReport | Add-Member -type NoteProperty -name TotalCount -value $key
	$fldReport | Add-Member -type NoteProperty -name TotalSize -value $key
	for($currYear = $minYear;$currYear -le $maxYear;$currYear++){
		$fldReport | Add-Member -type NoteProperty -Name ($currYear.ToString() + "_Count") -Value 0
		$fldReport | Add-Member -type NoteProperty -Name ($currYear.ToString() + "_Size") -Value 0
	}
	$itemHash = $fldhash[$key]
	$totalCount = 0
	$TotalSize = 0
	foreach($itmKey in $itemHash.Keys){
		$totalCount += $itemHash[$itmKey].TotalNumber
		$TotalSize += [Math]::Round(($itemHash[$itmKey].TotalSize/1MB),2)
		$fldReport.($itmKey.ToString() +"_Count") = $itemHash[$itmKey].TotalNumber
		$fldReport.($itmKey.ToString() +"_Size") = [Math]::Round(($itemHash[$itmKey].TotalSize/1MB),2)
	}
	$fldReport.TotalCount = $totalCount
	$fldReport.TotalSize = $TotalSize
	$rptCollection += $fldReport 
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

$rptCollection | ConvertTo-HTML -head $tableStyle –body $body | Out-File c:\Temp\MailboxAgeReport.htm
