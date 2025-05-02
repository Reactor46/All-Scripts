## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

## Load Managed API dll  
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"  
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 
  
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  
  
## Create Exchange Service Object  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
  
## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
  
#Credentials Option 1 using UPN for the windows Account  
$psCred = Get-Credential  
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())  
$service.Credentials = $creds      
$service.EnableScpLookup = $false
  
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
#Define Extended properties  
$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)  
# Bind to the Contacts Folder

$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderidcnt)

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
#

$Treeinfo = @{}
$TNRoot = new-object System.Windows.Forms.TreeNode("Root")
$TNRoot.Name = "Mailbox"
$TNRoot.Text = "Mailbox - " + $MailboxName
#The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox  
do {  
    $fiResult = $Service.FindFolders($folderidcnt,$sfSearchFilter,$fvFolderView)  
    foreach($ffFolder in $fiResult.Folders){   
		#Process folder here
		$TNChild = new-object System.Windows.Forms.TreeNode($ffFolder.DisplayName.ToString())
		$TNChild.Name = $ffFolder.DisplayName.ToString()
		$TNChild.Text = $ffFolder.DisplayName.ToString()
		$TNChild.tag = $ffFolder.Id.UniqueId.ToString()
		if ($ffFolder.ParentFolderId.UniqueId -eq $rfRootFolder.Id.UniqueId ){
			$ffFolder.DisplayName
			[void]$TNRoot.Nodes.Add($TNChild) 
			$Treeinfo.Add($ffFolder.Id.UniqueId.ToString(),$TNChild)
		}
		else{
			$pfFolder = $Treeinfo[$ffFolder.ParentFolderId.UniqueId.ToString()]
			[void]$pfFolder.Nodes.Add($TNChild)
			if ($Treeinfo.ContainsKey($ffFolder.Id.UniqueId.ToString()) -eq $false){
				$Treeinfo.Add($ffFolder.Id.UniqueId.ToString(),$TNChild)
			}
		}
    } 
    $fvFolderView.Offset += $fiResult.Folders.Count
}while($fiResult.MoreAvailable -eq $true)  
$Script:clickedFolder = $null
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Folder Select Form"
$objForm.Size = New-Object System.Drawing.Size(600,600) 
$objForm.StartPosition = "CenterScreen"
$tvTreView1 = new-object System.Windows.Forms.TreeView
$tvTreView1.Location = new-object System.Drawing.Size(1,1) 
$tvTreView1.add_DoubleClick({
	$Script:clickedFolder = $this.SelectedNode.tag
	$objForm.Close()
})
$tvTreView1.size = new-object System.Drawing.Size(580,580) 
$tvTreView1.Anchor = "Top,left,Bottom"
[void]$tvTreView1.Nodes.Add($TNRoot) 
$objForm.controls.add($tvTreView1)
$objForm.ShowDialog()

$clickedfolderid = new-object Microsoft.Exchange.WebServices.Data.FolderId($Script:clickedFolder)   

$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems,$MailboxName)   
$DuplicatesFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder -ArgumentList $service
$DuplicatesFolder.DisplayName = "DuplicateItems-Deduped-" + (Get-Date).ToString("yyyy-MM-dd-hh-mm-ss")
$DuplicatesFolder.Save($folderid)

#Define ItemView to retrive just 1000 Items
$PidTagSearchKey = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x300B, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
$psPropset.add($PidTagSearchKey)

$dupHash = @{}

#Create Collection for Move Batch
$Itemids = @()
$script:allChoice = $false

$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
$ivItemView.PropertySet = $psPropset
$fiItems = $null    
do{    
    $fiItems = $service.FindItems($clickedfolderid,$ivItemView)    
    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
    foreach($Item in $fiItems.Items){
		$PropVal =  $null
		if($Item.TryGetProperty($PidTagSearchKey,[ref]$PropVal)){
			$SearchString = [System.BitConverter]::ToString($PropVal).Replace("-","")
			if($dupHash.ContainsKey($SearchString)){
				#Check the recivedDate if availible
				if($Item.DateTimeReceived -ne $null){
					if($Item.DateTimeReceived -eq $dupHash[$SearchString]){
						if($script:allChoice -eq $false){
							$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""
							$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No",""
							$all = new-Object System.Management.Automation.Host.ChoiceDescription "&All","";
							$choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no,$all)
							$message = "Duplicated Detected : Subject " + $Item.Subject + " : Received-" + $dupHash[$SearchString] + " : Created-" + $Item.DateTimeCreated
							$result = $Host.UI.PromptForChoice($caption,$message,$choices,0)
							if($result -eq 0) {						
								$Itemids += $Item						
							}
							else{
								if($result -eq 2){
									$script:allChoice = $true
									$Itemids += $Item
								}
							}
						}
						else{
							$Itemids += $Item
						}
					}					
				}else{
					"Duplicate Found : " + $Item.Subject
					$Itemids += $Item
				}
			}
			else{
				"Procesing Item " + $Item.Subject
				if($Item.DateTimeReceived -ne $null){
					$dupHash.add($SearchString,$Item.DateTimeReceived)
				}
				else{
					$dupHash.add($SearchString,"")
				}
			}
		}        
    }    
    $ivItemView.Offset += $fiItems.Items.Count    
}while($fiItems.MoreAvailable -eq $true) 

#Total Items Processed Varible
$nmbProcessed = 0
if($Itemids.Count -gt 0){
	write-host ("Move " + $Itemids.Count + " Items")
	#Create Collection for Move Batch
	$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
	$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.ItemId" -as "Type")
	$BatchItemids = [Activator]::CreateInstance($type)
	#Varible to Track BatchSize
	$batchSize = 0
	foreach($iiID in $Itemids){
		$nmbProcessed++
		$BatchItemids.Add($iiID.Id)
		if($iiID.Size -ne $null){
			$batchSize += $iiID.Size
		}
		#if BatchCount greator then 50 or larger the 10 MB Move Batch
		if($BatchItemids.Count -eq 50 -bor $batchSize -gt (10*1MB)){
			$Result = $null
			$Result = $service.MoveItems($BatchItemids,$DuplicatesFolder.Id)
			[INT]$collectionCount = 0
			[INT]$Rcount = 0  
			[INT]$Errcount = 0
			$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
			$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.ItemId" -as "Type")
			#Define Collection to Retry Move For faild Items
			if($Result -ne $null){
			    foreach ($res in $Result){ 
			        if ($res.Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success){  
			            $Rcount++  
			        } 
					else{
						$Errcount++
					}
					$collectionCount++
			    }  
			}
			else{
				Write-Host -foregroundcolor red ("Move Result Null Exception")
			}
			Write-host ($Rcount.ToString() + " Items Moved Successfully " + "Total Processed " + $nmbProcessed + " Total Folder Items " + $Itemids.Count) 
			if($Errcount -gt 0){
				Write-Host -foregroundcolor red ($Errcount.ToString() + " Error failed Moved")
			}
			$BatchItemids.Clear()
			$batchSize = 0
		}
	}
	if($BatchItemids.Count -gt 0){
		$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
		$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.ItemId" -as "Type")
		$RetryBatchItemids = [Activator]::CreateInstance($type)
		$Result = $service.MoveItems($BatchItemids,$DuplicatesFolder.Id) 
		[INT]$Rcount = 0  
		[INT]$Errcount = 0 
		foreach ($res in $Result){  
		    if ($res.Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success){  
		        $Rcount++  
		    }  
			else{
				$Errcount++
			}
		}  
		Write-host ($Rcount.ToString() + " Items Moved Successfully")

		if($Errcount -gt 0){
			Write-Host -foregroundcolor red ($Errcount.ToString() + " Error failed Moved")
		}
	}
}
$DuplicatesFolder.Load()
if($DuplicatesFolder.TotalCount -eq 0){
	$DuplicatesFolder.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
}