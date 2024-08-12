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
#$service.AutodiscoverUrl($MailboxName,{$true})  
#"Using CAS Server : " + $Service.url   
   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 


function Process-Mailbox{
	param (
	        $SmtpAddress = "$( throw 'SMTPAddress is a mandatory Parameter' )"
		  )
	process{
	$rptObj = "" | select MailboxName,TotalItem,TotalItemSize,TotalItemsNoAttach,TotalItemsNoAttachSize,TotalItemsAttach,TotalItemsAttachSize,TotalFileAttachments,TotalFileAttachmentsSize,TotalItemAttachments,TotalItemAttachmentsSize,LargestAttachmentSize,LargestAttachmentName
	$rptObj.MailboxName = $SmtpAddress
	$rptObj.TotalItem = 0
	$rptObj.TotalItemSize = [Int64]0
	$rptObj.TotalItemsNoAttach = 0
	$rptObj.TotalItemsNoAttachSize = [Int64]0
	$rptObj.TotalItemsAttach = 0
	$rptObj.TotalItemsAttachSize = [Int64]0
	$rptObj.TotalFileAttachments = 0
	$rptObj.TotalFileAttachmentsSize  = [Int64]0
	$rptObj.TotalItemAttachments = 0
	$rptObj.TotalItemAttachmentsSize  = [Int64]0
	$rptObj.LargestAttachmentSize = [Int64]0
	$rptObj.LargestAttachmentName = ""
	"Processing Mailbox : " + $SmtpAddress
	
	#check Anchor header for Exchange 2013/Office365
	if($service.HttpHeaders.ContainsKey("X-AnchorMailbox")){
		$service.HttpHeaders["X-AnchorMailbox"] = $SmtpAddress
	}else{
		$service.HttpHeaders.Add("X-AnchorMailbox", $SmtpAddress);
	}
	"AnchorMailbox : " + $service.HttpHeaders["X-AnchorMailbox"]
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
	        
			$totalItemCnt = 1
			if($ffFolder.TotalCount -ne $null){
				$totalItemCnt = $ffFolder.TotalCount
				"Processing FolderPath : " + $fpath  + " Item Count " + $totalItemCnt
			}
			else{
				"Processing FolderPath : " + $fpath
			}
			if($totalItemCnt -gt 0){
				#Define ItemView to retrive just 1000 Items    
				$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
				$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
				$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
				$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
				$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
				$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Attachments)
				$psPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::HasAttachments)
				$fipsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)  
				$ivItemView.PropertySet = $fipsPropset		
				$fiItems = $null    
				do{    
				    $fiItems = $service.FindItems($ffFolder.Id,$ivItemView) 
					if($fiItems.Items.Count -gt 0){
					    [Void]$service.LoadPropertiesForItems($fiItems,$psPropset) 
						"processing : " + $fiItems.Items.Count + " Items"
					    foreach($Item in $fiItems.Items){
							$rptObj.TotalItem +=1
							$rptObj.TotalItemSize += [Int64]$Item.Size
							if($Item.Attachments.Count -gt 0){
								$rptObj.TotalItemsAttach +=1
								$rptObj.TotalItemsAttachSize += [Int64]$Item.Size
								foreach($Attachment in $Item.Attachments){							
									if($Attachment -is [Microsoft.Exchange.WebServices.Data.FileAttachment]){
										$rptObj.TotalFileAttachments +=1
										$rptObj.TotalFileAttachmentsSize += $Attachment.Size
										$attachSize = [Math]::Round($Attachment.Size/1MB,2)
										if($attachSize -gt $rptobj.LargestAttachmentSize){
											$rptobj.LargestAttachmentSize = $attachSize
											$rptobj.LargestAttachmentName = $Attachment.Name
										}
									}
									else{
										$rptObj.TotalItemAttachments +=1
										$rptObj.TotalItemAttachmentsSize += $Attachment.Size
									}
								}
							}
							else{
								$rptObj.TotalItemsNoAttach +=1
								$rptObj.TotalItemsNoAttachSize += [Int64]$Item.Size
							}
					    }
					}    
				    $ivItemView.Offset += $fiItems.Items.Count    
				}while($fiItems.MoreAvailable -eq $true)
			}
		} 
	    $fvFolderView.Offset += $fiResult.Folders.Count
	}while($fiResult.MoreAvailable -eq $true)
	#convert Sizes to MB

	if($rptObj.TotalItemSize -ne 0){
		$rptObj.TotalItemSize = [Math]::Round($rptObj.TotalItemSize/1MB)
	}
	if($rptObj.TotalItemsNoAttachSize -ne 0){
		$rptObj.TotalItemsNoAttachSize = [Math]::Round($rptObj.TotalItemsNoAttachSize/1MB)
	}
	if($rptObj.TotalItemsAttachSize -ne 0){
		$rptObj.TotalItemsAttachSize = [Math]::Round($rptObj.TotalItemsAttachSize/1MB)
	}
	if($rptObj.TotalFileAttachmentsSize -ne 0){
		$rptObj.TotalFileAttachmentsSize = [Math]::Round($rptObj.TotalFileAttachmentsSize/1MB)
	}
	if($rptObj.TotalItemAttachmentsSize -ne 0){
		$rptObj.TotalItemAttachmentsSize = [Math]::Round($rptObj.TotalItemAttachmentsSize/1MB)
	}
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
$Script:rptCollection | Export-Csv -NoTypeInformation -Path c:\temp\mbAttachReport.csv