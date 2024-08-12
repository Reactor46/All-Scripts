$RunMode = Read-Host "Enter Mode for Script (Report or Delete)"
switch($RunMode){
	"Report" {$runokay = $true
			  "Report Mode Detected"
			}
	"Delete" {$runokay = $true
			 "Delete Mode Detected"
			 }
	default {$runOkay = $false
			 "Invalid Reponse you need to type Report or Delete"
			 }
}
if($runokay){
## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = (Read-Host "Enter mailbox to run it against(email address)" )

$rptcollection = @()

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
#Define Attachment Contet To Search for

$AttachmentContent = ".pdf"
$AQSString = "System.Message.AttachmentContents:$AttachmentContent"

#Define ItemView to retrive just 1000 Items    
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
$fiItems = $null    
do{    
    $fiItems = $service.FindItems($clickedfolderid,$AQSString,$ivItemView)    
    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
    foreach($Item in $fiItems.Items){  
			$Item.Load()
			
			$attachtoDelete = @()
			foreach ($Attachment in $Item.Attachments)
			{
                if($Attachment -is [Microsoft.Exchange.WebServices.Data.FileAttachment]){  
					if($Attachment.Name.Contains(".pdf")){
						if($RunMode -eq "Report"){
							$rptObj = "" | Select DateTimeReceived,Subject,AttachmentName,Size
							$rptObj.DateTimeReceived = $Item.DateTimeReceived
							$rptObj.Subject = $Item.Subject
							$rptObj.AttachmentName = $Attachment.Name						
							$rptObj.Size = $Attachment.Size
							$rptcollection += $rptObj
						}
						$Attachment.Name
						if($RunMode -eq "Delete"){
							$attachtoDelete += $Attachment
						}
					}
                }  
			}
			$updateItem = $false
			foreach($AttachmentToDelete in $attachtoDelete){
				if($RunMode -eq "Delete"){
					$Item.Attachments.Remove($AttachmentToDelete)
					$updateItem = $true
				}
			}
			if($updateItem){$Item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverWrite)};
    }    
    $ivItemView.Offset += $fiItems.Items.Count    
}while($fiItems.MoreAvailable -eq $true)
if($RunMode -eq "Report"){
	$rptcollection | Export-Csv -Path c:\temp\$MailboxName-AttachmentReport.csv -NoTypeInformation
}

}