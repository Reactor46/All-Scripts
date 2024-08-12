$MailboxName = "user@domain.com"

$sendAlertTo = "alertto@domaim.com"
$sendAlertFrom = "from@domain.com"
$SMTPServer = "smtpservername"

$rptCollection2 = @()

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)

$subjecttoSearch  = "I hate Spam"
$AQSQuery = "Received:this week AND subject:`"" + $subjecttoSearch + "`""
$MailDate = [system.DateTime]::Now.AddDays(-7) 

 
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$rfRootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)
$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$rfRootFolderID)
$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000);
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$fvFolderView.PropertySet = $Propset
$ffResponse = $rfRootFolder.FindFolders($fvFolderView);
$PR_TRANSPORT_MESSAGE_HEADERS = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(125, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$Propset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$Propset.add($PR_TRANSPORT_MESSAGE_HEADERS)

foreach ($ffFolder in $ffResponse.Folders){
	"Checking " + $ffFolder.DisplayName
	$ivview = new-object Microsoft.Exchange.WebServices.Data.ItemView(20000)
	$ivview.propertyset =  $Propset
	$frFolderResult = $ffFolder.FindItems($AQSQuery,$ivview)
	foreach ($miMailItems in $frFolderResult.Items){
        #Doublic Check Exact Subect match
        if ($miMailItems.Subject -eq $subjecttoSearch){
		   $ItemReport = "" | select FolderName,Recieved,From,Subject,Headers
		   $Headers = $null	           
		   $miMailItems.TryGetProperty($PR_TRANSPORT_MESSAGE_HEADERS, [ref]$Headers)
		   $ItemReport.FolderName = $ffFolder.DisplayName	
		   $ItemReport.Recieved = $miMailItems.DateTimeReceived
                   $ItemReport.From = $miMailItems.From.Name
		   $ItemReport.Subject = $miMailItems.Subject
		   $ItemReport.Headers =  $Headers
		   $rptCollection2 += $ItemReport

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
$MailMessage.Subject = "Subject Search Report for " +  $MailboxName
$MailMessage.IsBodyHtml = $TRUE
$MailMessage.body = $rptCollection2 | ConvertTo-HTML -head $tableStyle –body $body 
$SMTPClient.Send($MailMessage)
	

