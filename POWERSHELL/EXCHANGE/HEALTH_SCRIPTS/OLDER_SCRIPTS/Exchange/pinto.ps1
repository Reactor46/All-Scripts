$MailboxName = $args[0]
$MessageText = $args[1]
$MessageSubject = "Pin To Note Subject"

$expire = [DateTime]::Now.AddDays(-1)
if (Test-Path pifile.txt){
	$pmailid = gc pifile.txt -totalcount 1
}
else{
	$pmailid = "3B"
}
if ($pmailid -eq $null){$pmailid = "3B"}
$pmailid
function ConvertToByteArray($newString){ 

$byteArray = New-Object byte[] ($newString.length/2)
for($icnt=0;$icnt-lt$newString.length;$icnt=$icnt+2){
	if ($icnt -eq 0){$byteArray[$icnt] = [Convert]::ToInt32($newString.SubString($icnt,2).tolower(), 16)}
	else{$byteArray[($icnt/2)] = [Convert]::ToInt32($newString.SubString($icnt,2).tolower(), 16)}
}
return $byteArray

}

function CovertToHex($BinArray1){
	$hexOut = ""
	for($bc=0;$bc -lt $BinArray1.Length;$bc++){
		$nhex = [Convert]::ToString($BinArray1[$bc], 16)
		if ($nhex.length -eq 1){$nhex = "0" + $nhex}
		$hexOut = $hexOut + $nhex
	}
	return $hexOut 
}



$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$PrSearchKey = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(12299, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
$PR_Client_Submit_Time = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(57, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime);
$PR_Delivery_Time = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3590, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime);
$PR_Flags = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3591, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

$Propset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$Propset.Add($PR_Client_Submit_Time)
$Propset.Add($PR_Delivery_Time)
$Propset.Add($PrSearchKey)
$Propset.Add($PR_Flags)
$Propset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

$Sfir = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PrSearchKey, [Convert]::ToBase64String((ConvertToByteArray($pmailId))))
$Sflt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $expire)

$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)

$sfCollection.add($Sfir)
$sfCollection.add($Sflt)


$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(1)

$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$frFolderResult = $InboxFolder.FindItems($sfCollection,$view)
if ($frFolderResult.Items.Count -ne 0){
	"Item Found"
	$frFolderResult.Items[0].load($Propset)
	$frFolderResult.Items[0].Body.Text
	$frFolderResult.Items[0].ExtendedProperties[0].Value = [DateTime]::Now
	$frFolderResult.Items[0].ExtendedProperties[1].Value = [DateTime]::Now
	$frFolderResult.Items[0].Body.Text = $frFolderResult.Items[0].Body.Text + $MessageText  + "`r`n" 
	$frFolderResult.Items[0].isread = $false
	$frFolderResult.Items[0].Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
}
else{
	"Item Not found"
	$message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
	$message.From = $MailboxName
	$message.ToRecipients.Add($MailboxName)
	$message.Subject = $MessageSubject
	$message.Body = $MessageText  + "`r`n" 
	$message.SetExtendedProperty($PR_Flags,"1")
	$message.isread = $false
	$message.save($InboxFolder.Id)
	$message.load($Propset)
	$message.ExtendedProperties[0].Value = [DateTime]::Now
	$message.ExtendedProperties[1].Value = [DateTime]::Now
	$message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
	CovertToHex($message.ExtendedProperties[2].Value) | out-file pifile.txt	
}