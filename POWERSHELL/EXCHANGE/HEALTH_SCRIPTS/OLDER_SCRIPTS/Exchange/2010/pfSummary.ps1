$MailDate = [system.DateTime]::Now.AddDays(-1)
$pfFolder = "/Folder-offroot/Folder1"
$SenderName = "targetAddress@domain.com"

$sendAlertTo = "Sendalerttoaddress@domain.com"
$sendAlertFrom = "Sendalertfromaddress@domain.com"

$SMTPServer = "HubServerName"

Function FindTargetFolder($FolderPath){
	$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
	$pfArray = $FolderPath.Split("/")
	for ($lint = 1; $lint -lt $pfArray.Length; $lint++) {
		$pfArray[$lint]
		$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
		$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$pfArray[$lint])
                $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView)
		if ($findFolderResults.TotalCount -gt 0){
			foreach($folder in $findFolderResults.Folders){
				$tfTargetFolder = $folder				
			}
		}
		else{
			"Error Folder Not Found"
			$tfTargetFolder = $null
			break
		}	
	}
	$Global:findFolder = $tfTargetFolder
}


$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

FindTargetFolder($pfFolder)
$PublicFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$Global:findFolder.Id.UniqueId)

$Sfir = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender,$SenderName)
$Sflt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $MailDate)

$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfCollection.add($Sfir)
$sfCollection.add($Sflt)


$rpReport = $rpReport + "<table><tr bgcolor=`"#95aedc`">"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:15%;`" ><b>Recieved</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>From</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>To</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:40%;`" ><b>Subject</b></td>"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:5%;`" ><b>Size</b></td>"
$rpReport = $rpReport + "</tr>"

$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$fiFindItemsResult = $PublicFolder.FindItems($sfCollection,$view)
foreach($Item in $fiFindItemsResult.items){
	$aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternatePublicFolderItemId
	$aiItem.FolderId = $PublicFolder.Id
	$aiItem.ItemId = $Item.Id
	$aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId;
	$CasServer = $service.Url.Host.ToString()
	$openType = "ae=Item";
	$owaid = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::OwaId)
	$Item.load()
	$OWAURL = "https://" + $CasServer + "/owa/?" + $openType + "&t=" + $Item.ItemClass + "&id=" + $owaid.ItemId
	$rpReport = $rpReport + "  <tr>"  + "  "
	$rpReport = $rpReport + "<td>" + $Item.DateTimeReceived.ToString() + "</td>"  + "  "
	$rpReport = $rpReport + "<td>" + $Item.From.Address + "</td>"  + "  "
	$rpReport = $rpReport + "<td>" + $Item.ToRecipients[0].Address + "</td>"  + "  "
	$rpReport = $rpReport + "<td><a href=`"" + $OWAURL + "`">" + $Item.Subject + "</a></td>"  + "  "
	$rpReport = $rpReport + "<td>" + $Item.Size + "</td>"  + "  "
	$rpReport = $rpReport + "</tr>"  + "  "
	 
}

$rpReport = $rpReport + "</table>"  + "  " 
$SmtpClient = new-object system.net.mail.smtpClient
$SmtpClient.host =  $SMTPServer
$MailMessage = new-object System.Net.Mail.MailMessage
$MailMessage.To.Add($sendAlertTo)
$MailMessage.From = $sendAlertFrom
$MailMessage.Subject = "Summary Email" 
$MailMessage.IsBodyHtml = $TRUE
$MailMessage.body = $rpReport
$SMTPClient.Send($MailMessage)





