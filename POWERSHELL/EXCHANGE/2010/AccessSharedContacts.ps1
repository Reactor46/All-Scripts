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

#PropDefs 
$pidTagStoreEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(4091, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$PidTagNormalizedSubject = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E1D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String); 
$PidTagWlinkType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6849, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PidTagWlinkFlags = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PidTagWlinkOrdinal = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684B, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$PidTagWlinkFolderType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684F, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$PidTagWlinkSection = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6852, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PidTagWlinkGroupHeaderID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6842, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$PidTagWlinkSaveStamp = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6847, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PidTagWlinkGroupName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6851, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$PidTagWlinkStoreEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684E, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$PidTagWlinkGroupClsid = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6850, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$PidTagWlinkEntryId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684C, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$PidTagWlinkRecordKey = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x684D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$PidTagWlinkCalendarColor = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6853, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PidTagWlinkAddressBookEID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6854,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$PidTagWlinkROGroupType = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6892,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PidTagWlinkAddressBookStoreEID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x6891,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)


$SharedFolders = @{}  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
Write-Host ("Getting CommonVeiwFolder")
#Get CommonViewFolder
$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"Common Views") 
$findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView) 
if ($findFolderResults.TotalCount -gt 0){ 
	$ExistingShortCut = $false
	$cvCommonViewsFolder = $findFolderResults.Folders[0]
	#Define ItemView to retrive just 1000 Items    
	#Find Items that are unread
	$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
	$psPropset.add($PidTagWlinkStoreEntryId)
	$psPropset.add($PidTagWlinkFolderType)
	$cntSearch = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PidTagWlinkGroupName, "Shared Contacts");
	$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)   
	$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
	$ivItemView.PropertySet = $psPropset
	$fiItems = $service.FindItems($cvCommonViewsFolder.Id,$cntSearch,$ivItemView)    
    foreach($Item in $fiItems.Items){
		$idVal = $null
		if($Item.TryGetProperty($PidTagWlinkStoreEntryId,[ref]$idVal)){
			Write-Host("Processing " + $Item.Subject)
			    $ssStoreID = $idVal;
                $leLegDnStart = 0;
                $lnLegDN = "";
                for ($ssArraynum=($ssStoreID.Length - 2);$ssArraynum -ne 0; $ssArraynum--)
                        {
                            if ($ssStoreID[$ssArraynum] -eq 0)
                            {
                                $leLegDnStart = $ssArraynum;
                                $lnLegDN = [System.Text.ASCIIEncoding]::ASCII.GetString($ssStoreID, $leLegDnStart + 1, ($ssStoreID.Length - ($leLegDnStart + 2)));
                                $ssArraynum = 1;
                            }
                        }
						Write-Host($lnLegDN)
                       	$ncCol = $service.ResolveName($lnLegDN, [Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryOnly, $true);
                        if ($ncCol.Count -gt 0)
                        {
                            try
                            {
                                $SharedContactsId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $ncCol[0].Mailbox.Address);
                                $SharedContactFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $SharedContactsId);
                                $SharedFolders.Add($ncCol[0].Mailbox.Address, $SharedContactFolder);
                            }
                            catch  {
                                Write-Host "Error getting Shared Folder"
                            }
                        }
			
		}						      
	}
}
if($SharedFolders.Keys.Count -ne 0){
	foreach($mbFolder in $SharedFolders.Keys){
		#Define ItemView to retrive just 1000 Items    
		$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
		$fiItems = $null    
		do{    
		    $fiItems = $service.FindItems($SharedFolders[$mbFolder].Id,$ivItemView)    
		    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
		    foreach($Item in $fiItems.Items){      
				Write-Host ("Mailbox : " + $mbFolder)
				Write-Host ("Contact : " + $Item.Subject + " : " + $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1])
		    }    
		    $ivItemView.Offset += $fiItems.Items.Count    
		}while($fiItems.MoreAvailable -eq $true) 
	}
}






