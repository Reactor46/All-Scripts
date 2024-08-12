$MailboxName = "user@domain.com"
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$service.TraceEnabled = $false

$service.Credentials = New-Object System.Net.NetworkCredential("user@domain.com","password")
$service.AutodiscoverUrl($MailboxName ,{$true})

$fldArray = new-object Microsoft.Exchange.WebServices.Data.FolderId[] 1
$Inboxid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$fldArray[0] = $Inboxid
$stmsubscription = $service.SubscribeToStreamingNotifications($fldArray, [Microsoft.Exchange.WebServices.Data.EventType]::NewMail)
$stmConnection = new-object Microsoft.Exchange.WebServices.Data.StreamingSubscriptionConnection($service, 30);
$stmConnection.AddSubscription($stmsubscription)
Register-ObjectEvent -inputObject $stmConnection -eventName "OnNotificationEvent" -Action {
	foreach($notEvent in $event.SourceEventArgs.Events){	
		[String]$itmId = $notEvent.ItemId.UniqueId.ToString()
		$message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($event.MessageData,$itmId)
		"Subject : " + $message.Subject + " " + (Get-Date) | Out-File c:\temp\log2.txt -Append 
	} 
} -MessageData $service
Register-ObjectEvent -inputObject $stmConnection -eventName "OnDisconnect" -Action {$event.MessageData.Open()} -MessageData $stmConnection
$stmConnection.Open()
