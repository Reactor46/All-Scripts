$MailboxName = "mailbox@domain.com"

$sendAlertTo = "sendto@domain.com"
$sendAlertFrom = "report@domain.com"
$SMTPServer = "smtpservername"


$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())


$rptCollection = @()


## Define Extended Properties

$PR_DELETED_ON = new-object  Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26255, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
$PR_DELETED_MSG_COUNT = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26176, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PR_DELETED_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26267, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long)
$PR_DELETED_FOLDER_COUNT = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26177, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PR_Sender_Name = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26177, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)

## End Define Extended Properties
## Define Property Sets
## Folder Set

$fpsFolderPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$fpsFolderPropertySet.add($PR_DELETED_ON)
$fpsFolderPropertySet.add($PR_DELETED_MSG_COUNT)
$fpsFolderPropertySet.add($PR_DELETED_MESSAGE_SIZE_EXTENDED)
$fpsFolderPropertySet.add($PR_DELETED_FOLDER_COUNT)

## Item Set

$ipsItemPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
$ipsItemPropertySet.add($PR_DELETED_ON)
$ipsItemPropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
$ipsItemPropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
$ipsItemPropertySet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From)
# End Set 

$rfRootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)
$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$rfRootFolderID)
$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000);
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$fvFolderView.PropertySet = $fpsFolderPropertySet
# $service.traceenabled = $true
$ffResponse = $rfRootFolder.FindFolders($fvFolderView);
foreach ($ffFolder in $ffResponse.Folders){
	$dcDeleteItemCount = $null
	$fptProptest = $ffFolder.TryGetProperty($PR_DELETED_MSG_COUNT, [ref]$dcDeleteItemCount) 
	if($fptProptest){
		if ($dcDeleteItemCount -ne 0){
			$ffFolder.DisplayName +  " - Number Items Deleted :" + $dcDeleteItemCount
			$bcBatchCount = 0;
			$bcBatchSize = 1000
			$ivItemView = new-object Microsoft.Exchange.WebServices.Data.ItemView($bcBatchSize, $bcBatchCount)
			$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::SoftDeleted
			$ivItemView.PropertySet = $ipsItemPropertySet
			$service.traceenabled = $false
			while (($fiFindItems = $ffFolder.FindItems($ivItemView)).Items.Count -gt 0)
			{
				foreach ($item in $fiFindItems.Items)
				{
					$lnum ++
					write-progress "Processing message" $lnum
					$delon = $null
					$ptProptest = $item.TryGetProperty($PR_DELETED_ON, [ref]$delon) 
					$Itemobj = "" | select Type,DeletedOn,From,Subject,Size
					$Itemobj.DeletedOn = $delon
					$Itemobj.From = $item.From.Name
					$Itemobj.Subject = $item.Subject
					$Itemobj.Size = $item.Size
					$Itemobj.Type = "Item"
					$rptCollection += $Itemobj
				}
				$bcBatchCount += $fiFindItems.Items.Count
				$ivItemView = new-object Microsoft.Exchange.WebServices.Data.ItemView($bcBatchSize, $bcBatchCount)
				$ivItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::SoftDeleted
				$ivItemView.PropertySet = $ipsItemPropertySet
			}
		}
	}
	$dcDeletedFolderCount = $null
	$fptProptest = $ffFolder.TryGetProperty($PR_DELETED_FOLDER_COUNT, [ref]$dcDeletedFolderCount) 
	if($fptProptest){
		if ($dcDeletedFolderCount -ne 0){
			$ffFolder.DisplayName + " - Number folders Deleted :" + $dcDeletedFolderCount		
			$fvFolderView1 = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000);
			$fvFolderView1.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::SoftDeleted
			$fvFolderView1.PropertySet = $fpsFolderPropertySet
			$ffResponse2 = $ffFolder.FindFolders($fvFolderView1)
			
			foreach ($ffDelFolder in $ffResponse2.Folders){
				$dcDeletedSize = $null
				$fptProptest = $ffDelFolder.TryGetProperty($PR_DELETED_MESSAGE_SIZE_EXTENDED, [ref]$dcDeletedSize) 
				$Deletedon = $null
				$ptProptest = $ffDelFolder.TryGetProperty($PR_DELETED_ON, [ref]$Deletedon) 
				$Itemobj = "" | select Type,DeletedOn,From,Subject,Size
				$Itemobj.DeletedOn = $Deletedon
				$Itemobj.Subject = $ffDelFolder.DisplayName
				$Itemobj.Size = $dcDeletedSize
				$Itemobj.Type = "Folder"
				$rptCollection += $Itemobj

			}
		}
	}

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
  


$SmtpClient = new-object system.net.mail.smtpClient
$SmtpClient.host =  $SMTPServer
$MailMessage = new-object System.Net.Mail.MailMessage
$MailMessage.To.Add($sendAlertTo)
$MailMessage.From = $sendAlertFrom
$MailMessage.Subject = "Dumpster Report for " +  $MailboxName
$MailMessage.IsBodyHtml = $TRUE
$MailMessage.body = $rptCollection | ConvertTo-HTML -head $tableStyle –body $body 
$SMTPClient.Send($MailMessage)
