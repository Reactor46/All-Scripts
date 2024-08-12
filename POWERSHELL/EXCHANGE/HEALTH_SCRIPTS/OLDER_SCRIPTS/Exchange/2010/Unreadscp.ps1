## EWS Managed API Connect Module Script written by Glen Scales
## Requires the EWS Managed API and Powershell V2.0 or greator
$RptCollection = @()
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
$MailboxName = $_.PrimarySMTPAddress
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


$rptObj = "" | select  MailboxName,Mailboxsize,LastLogon,LastLogonAccount,Last6MonthsTotal,Last6MonthsUnread,LastMailRecieved,Last6MonthsSent,LastMailSent
$rptObj.MailboxName = $MailboxName

$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)


$Range = [system.DateTime]::Now.AddMonths(-6).ToString("MM/dd/yyyy") + ".." + [system.DateTime]::Now.ToString("MM/dd/yyyy") 
$AQSString1 = "System.Message.DateReceived:" + $Range 
$AQSString2 = $AQSString1 + " and unread:true"
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$MailboxName)   
$SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)

$MailboxStats = Get-MailboxStatistics $MailboxName
$ts = CovertBitValue($MailboxStats.TotalItemSize.ToString())
"Total Size : " + $MailboxStats.TotalItemSize
$rptObj.MailboxSize = $ts
"Last Logon Time : " + $MailboxStats.LastLogonTime
$rptObj.LastLogon = $MailboxStats.LastLogonTime
"Last Logon Account : " + $MailboxStats.LastLoggedOnUserAccount
$rptObj.LastLogonAccount = $MailboxStats.LastLoggedOnUserAccount
$fiResults = $Inbox.findItems($AQSString1,$ivItemView)
$rptObj.Last6MonthsTotal = $fiResults.TotalCount
"Last 6 Months : " + $fiResults.TotalCount
if($fiResults.TotalCount -gt 0){
	"Last Mail Recieved : " + $fiResults.Items[0].DateTimeReceived
	$rptObj.LastMailRecieved = $fiResults.Items[0].DateTimeReceived
}
$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)
$fiResults = $Inbox.findItems($AQSString2,$ivItemView)
"Last 6 Months Unread : " + $fiResults.TotalCount
$rptObj.Last6MonthsUnread = $fiResults.TotalCount
$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)
$fiResults = $SentItems.findItems($AQSString1,$ivItemView)
"Last 6 Months Sent : " + $fiResults.TotalCount
$rptObj.Last6MonthsSent = $fiResults.TotalCount
if($fiResults.TotalCount -gt 0){
	"Last Mail Sent Date : " + $fiResults.Items[0].DateTimeSent
	$rptObj.LastMailSent = $fiResults.Items[0].DateTimeSent
}
$RptCollection +=$rptObj
}
$RptCollection | Export-Csv -NoTypeInformation  c:\temp\unreadReport.csv