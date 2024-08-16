$MailboxName = "user@domain.com"

$twitterusername = "userName"
$twitterpassword = "password"


function updateTwiterStatus($PostString){

	$tweetstring = [String]::Format("status={0}", $PostString )
	[System.Net.ServicePointManager]::Expect100Continue = $false
	$request = [System.Net.WebRequest]::Create("http://twitter.com/statuses/update.xml")
	$request.Credentials = new-object System.Net.NetworkCredential($twitterusername,$twitterpassword)  
	$request.Method = "POST"
	$request.ContentType = "application/x-www-form-urlencoded"        
	$formdata = [System.Text.Encoding]::UTF8.GetBytes($tweetstring)
	$request.ContentLength = $formdata.Length
	$requestStream = $request.GetRequestStream()
	$requestStream.Write($formdata, 0, $formdata.Length)
	$requestStream.Close()
	$response = $request.GetResponse()
	$reader = new-object System.IO.StreamReader($response.GetResponseStream())
	$returnvalue = $reader.ReadToEnd()
	$reader.Close()  

}

function GetTwitterStatus(){
	[System.Net.ServicePointManager]::Expect100Continue = $false
	$request = [System.Net.WebRequest]::Create("http://twitter.com/users/show/" + $twitterusername)
	$request.Credentials = new-object System.Net.NetworkCredential($twitterusername,$twitterpassword)  
	$request.Method = "GET"
	$request.ContentType = "application/x-www-form-urlencoded"        
	$response = $request.GetResponse()
	$ResponseStream = $response.GetResponseStream()
	$ResponseXmlDoc = new-object System.Xml.XmlDocument
	$ResponseXmlDoc.Load($ResponseStream)
	$StatusNodes = @($ResponseXmlDoc.getElementsByTagName("status"))
	$returnStatus = $StatusNodes[0].text	
	$ResponseXmlDoc.Save("c:\dd.xml")
	return $returnStatus.ToString()
}

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)
$CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($service,$folderid)

$cvCalendarview = new-object Microsoft.Exchange.WebServices.Data.CalendarView
$cvCalendarview.StartDate = [System.DateTime]::Now
$cvCalendarview.EndDate = [System.DateTime]::Now.AddDays(14)
$cvCalendarview.MaxItemsReturned = 200;
$cvCalendarview.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

$upset = 0
$statusup = 0

$frCalendarResult = $CalendarFolder.FindAppointments($cvCalendarview)

foreach ($apApointment in $frCalendarResult.Items){

	if ($apApointment.Sensitivity -ne [Microsoft.Exchange.WebServices.Data.Sensitivity]::Private -band $apApointment.IsReminderSet -eq $true){
		        $ReminderSendTime = $apApointment.Start.AddMinutes(-$apApointment.ReminderMinutesBeforeStart)
                        if ($ReminderSendTime -ge [System.DateTime]::Now -band $ReminderSendTime -le [System.DateTime]::Now.AddMinutes(15))
                        {
                            	$twitString = $apApointment.Subject.ToString()
				if ($twitString.Length -gt 140){$twitString = $twitString.Substring(0,110)} 
				$twitString = $twitString   + " Starts : " + $apApointment.Start.ToString("yyyy-MM-dd HH:mm:ss") + " "
				[String]$currentStatus = GetTwitterStatus
				$twitString = $twitString.Substring(0,$twitString.length-1)
				write-host $twitString.ToLower().ToString()
				write-host $currentStatus.ToLower().ToString()
				if ($currentStatus.ToLower().ToString() -ne $twitString.ToLower().ToString()){	
		   	        if ($statusup -ne 3){$statusup = 1}
		 			$upset = 1
				}
				else{
					$statusup = 3
				}
                        }

	}

}
if ($statusup -eq 1){
	updateTwiterStatus($twitString)
	"Twitter Status Updated"
}
else {"Nothing changed since last update"}