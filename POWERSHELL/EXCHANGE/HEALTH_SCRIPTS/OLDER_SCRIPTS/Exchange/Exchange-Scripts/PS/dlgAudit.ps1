## EWS Managed API Connect Module Script written by Glen Scales
## Requires the EWS Managed API and Powershell V2.0 or greator
$rptArray = New-Object System.Collections.ArrayList
## Load Managed API dll
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"

## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1

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

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}



function CovertBitValue($String){
	$numItempattern = '(?=\().*(?=bytes)'
	$matchedItemsNumber = [regex]::matches($String, $numItempattern) 
	$Mb = [INT64]$matchedItemsNumber[0].Value.Replace("(","").Replace(",","")
	return [math]::round($Mb/1048576,0)
}


Get-Mailbox -ResultSize Unlimited | ForEach-Object{ 
	$MailboxName = $_.PrimarySMTPAddress.ToString()
	"Processing Mailbox : " + $MailboxName
	if($service.url -eq $null){
		## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use

		#CAS URL Option 1 Autodiscover
		$service.AutodiscoverUrl($MailboxName,{$true})
		"Using CAS Server : " + $Service.url 
		
		#CAS URL Option 2 Hardcoded
		#$uri=[system.URI] "https://casservername/ews/exchange.asmx"
		#$service.Url = $uri  
	}

	$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
    $delegates = $service.getdelegates($MailboxName,$true)  
	foreach($Delegate in $delegates.DelegateUserResponses){  
	    $rptObj = "" | select Delegate,Mailbox,Inbox,Calendar,Contacts,Tasks,Notes,Journal,MeetingMessages,ViewPrivateItems  
	    $rptObj.Mailbox = $MailboxName
		$rptObj.Delegate = $Delegate.DelegateUser.UserId.PrimarySmtpAddress  
	    $rptObj.Inbox = $Delegate.DelegateUser.Permissions.InboxFolderPermissionLevel  
	    $rptObj.Calendar = $Delegate.DelegateUser.Permissions.CalendarFolderPermissionLevel  
	    $rptObj.Contacts = $Delegate.DelegateUser.Permissions.ContactsFolderPermissionLevel  
	    $rptObj.Tasks = $Delegate.DelegateUser.Permissions.TasksFolderPermissionLevel  
	    $rptObj.Notes = $Delegate.DelegateUser.Permissions.NotesFolderPermissionLevel  
	    $rptObj.Journal = $Delegate.DelegateUser.Permissions.JournalFolderPermissionLevel  
	    $rptObj.ViewPrivateItems = $Delegate.DelegateUser.ViewPrivateItems  
	    $rptObj.MeetingMessages = $Delegate.DelegateUser.ReceiveCopiesOfMeetingMessages 
		[Void]$rptArray.Add($rptObj)
	      
	}  
}
$rptArray | Sort-Object Delegate | Export-Csv -Path c:\temp\ReverseDelegateReport.csv -NoTypeInformation
