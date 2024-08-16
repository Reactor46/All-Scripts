## Get the Mailbox to Access from the 1st commandline argument
$TargetCalendarMailbox = $args[1]
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

function GetAutoDiscoverSettings{
	param (
	        $adEmailAddress = "$( throw 'emailaddress is a mandatory Parameter' )",
			$Credentials = "$( throw 'Credentials is a mandatory Parameter' )"
		  )
	process{
		$adService = New-Object Microsoft.Exchange.WebServices.AutoDiscover.AutodiscoverService($ExchangeVersion);
		$adService.Credentials = $Credentials
		$adService.EnableScpLookup = $false;
		$adService.RedirectionUrlValidationCallback = {$true}
		$UserSettings = new-object Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[] 3
		$UserSettings[0] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDN
		$UserSettings[1] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::InternalRpcClientServer
		$UserSettings[2] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDisplayName
		$adResponse = $adService.GetUserSettings($adEmailAddress , $UserSettings);
		return $adResponse
	}
}
function GetAddressBookId{
	param (
	        $AutoDiscoverSettings = "$( throw 'AutoDiscoverSettings is a mandatory Parameter' )"
		  )
	process{
		$userdnString = $AutoDiscoverSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDN]
		$userdnHexChar = $userdnString.ToCharArray();
		foreach ($element in $userdnHexChar) {$userdnStringHex = $userdnStringHex + [System.String]::Format("{0:X}", [System.Convert]::ToUInt32($element))}
		$Provider = "00000000DCA740C8C042101AB4B908002B2FE1820100000000000000"
		$userdnStringHex = $Provider + $userdnStringHex + "00"
		return $userdnStringHex
	}
}
function GetStoreId{
	param (
	        $AutoDiscoverSettings = "$( throw 'AutoDiscoverSettings is a mandatory Parameter' )"
		  )
	process{
		$userdnString = $AutoDiscoverSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDN]
		$userdnHexChar = $userdnString.ToCharArray();
		foreach ($element in $userdnHexChar) {$userdnStringHex = $userdnStringHex + [System.String]::Format("{0:X}", [System.Convert]::ToUInt32($element))}	
		$serverNameString = $AutoDiscoverSettings.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::InternalRpcClientServer]
		$serverNameHexChar = $serverNameString.ToCharArray();
		foreach ($element in $serverNameHexChar) {$serverNameStringHex = $serverNameStringHex + [System.String]::Format("{0:X}", [System.Convert]::ToUInt32($element))}
		$flags = "00000000"
		$ProviderUID = "38A1BB1005E5101AA1BB08002B2A56C2"
		$versionFlag = "0000"
		$DLLFileName = "454D534D44422E444C4C00000000"
		$WrappedFlags = "00000000"
		$WrappedProviderUID = "1B55FA20AA6611CD9BC800AA002FC45A"
		$WrappedType = "0C000000"
		$StoredIdStringHex = $flags + $ProviderUID + $versionFlag + $DLLFileName + $WrappedFlags + $WrappedProviderUID + $WrappedType + $serverNameStringHex + "00" + $userdnStringHex + "00"
		return $StoredIdStringHex
	}
}


function hex2binarray($hexString){
    $i = 0
    [byte[]]$binarray = @()
    while($i -le $hexString.length - 2){
        $strHexBit = ($hexString.substring($i,2))
        $binarray += [byte]([Convert]::ToInt32($strHexBit,16))
        $i = $i + 2
    }
    return ,$binarray
}
function ConvertId($EWSid){    
    $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
    $aiItem.Mailbox = $MailboxName      
    $aiItem.UniqueId = $EWSid   
    $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::EWSId;      
    return $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::StoreId)     
} 

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


#Get the TargetUsers Calendar
# Bind to the Calendar Folder
$fldPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
$fldPropset.Add($pidTagStoreEntryId);
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$TargetCalendarMailbox)   
$TargetCalendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$fldPropset)
#Check for existing ShortCut for TargetMailbox
#Get AddressBook Id for TargetUser
Write-Host ("Getting Autodiscover Settings Target")
Write-Host ("Getting Autodiscover Settings Mailbox")
$adset = GetAutoDiscoverSettings -adEmailAddress $MailboxName -Credentials $creds
$storeID = ""
if($adset -is [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponse]){
	Write-Host ("Get StoreId")
	$storeID = GetStoreId -AutoDiscoverSettings $adset
}
$adset = $null
$abTargetABEntryId = ""
$adset = GetAutoDiscoverSettings -adEmailAddress $TargetCalendarMailbox -Credentials $creds
if($adset -is [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverResponse]){
	Write-Host ("Get AB Id")
	$abTargetABEntryId = GetAddressBookId -AutoDiscoverSettings $adset
	$SharedUserDisplayName =  $adset.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::UserDisplayName]
}
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
	$psPropset.add($PidTagWlinkAddressBookEID)
	$psPropset.add($PidTagWlinkFolderType)
	$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)   
	$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
	$ivItemView.PropertySet = $psPropset
	$fiItems = $service.FindItems($cvCommonViewsFolder.Id,$ivItemView)    
    foreach($Item in $fiItems.Items){
		$aeidVal = $null
		if($Item.TryGetProperty($PidTagWlinkAddressBookEID,[ref]$aeidVal)){
				$fldType = $null
				if($Item.TryGetProperty($PidTagWlinkFolderType,[ref]$fldType)){
					if([System.BitConverter]::ToString($fldType).Replace("-","") -eq "0278060000000000C000000000000046"){
						if([System.BitConverter]::ToString($aeidVal).Replace("-","") -eq $abTargetABEntryId){
							$ExistingShortCut = $true
							Write-Host "Found existing Shortcut"
							###$Item.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
						}
					}
				}
			}						      
	}
	if($ExistingShortCut -eq $false){
		If($storeID.length -gt 5 -band $abTargetABEntryId.length -gt 5){
			$objWunderBarLink = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service  
			$objWunderBarLink.Subject = $SharedUserDisplayName  
			$objWunderBarLink.ItemClass = "IPM.Microsoft.WunderBar.Link"  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkAddressBookEID,(hex2binarray $abTargetABEntryId))  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkAddressBookStoreEID,(hex2binarray $storeID))  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkCalendarColor,-1)
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkFlags,0)
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkGroupName,"Shared Calendars")
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkFolderType,(hex2binarray "0278060000000000C000000000000046"))  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkGroupClsid,(hex2binarray "B9F0060000000000C000000000000046"))  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkROGroupType,-1)
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkSection,3)  
			$objWunderBarLink.SetExtendedProperty($PidTagWlinkType,2)  
			$objWunderBarLink.IsAssociated = $true
			$objWunderBarLink.Save($findFolderResults.Folders[0].Id)
			Write-Host ("ShortCut Created for - " + $SharedUserDisplayName)
		}
		else{
			Write-Host ("Error with Id's")
		}
	}
}



