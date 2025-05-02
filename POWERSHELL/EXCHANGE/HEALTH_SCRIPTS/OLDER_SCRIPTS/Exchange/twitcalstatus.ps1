$twitterusername = "username"
$twitterpassword = "password"
$ExchangeServername = "servername"
$emEmailAddress = "twittest@domain.com"
$userName = "twittest"
$password = "Password"
$domain = "domain"

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


$casURL = "https://" + $ExchangeServername + "/EWS/Exchange.asmx"
[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")
$ewc = new-object EWSUtil.EWSConnection($emEmailAddress,$false, $userName, $password,$domain, $casURL)
$drDuration = new-object EWSUtil.EWS.Duration
$drDuration.StartTime = [DateTime]::UtcNow
$drDuration.EndTime = [DateTime]::UtcNow.AddMinutes(15)
$upset = 0
$statusup = 0
[EWSUtil.EWS.DistinguishedFolderIdType] $dType = new-object EWSUtil.EWS.DistinguishedFolderIdType
$dType.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::calendar
$qfolder = new-object EWSUtil.QueryFolder($ewc,$dType, $drDuration,$false,$true)
foreach ($item in $qfolder.fiFolderItems){
	if ($item.Location -ne $null){
		$twitString = $item.Subject.ToString() + " " + $item.Location.ToString()
	}
	else{
		$twitString = $item.Subject.ToString() 
	}
	if ($twitString.Length -gt 140){$twitString = $twitString.Substring(0,139)} 
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
if ($statusup -eq 1){
	updateTwiterStatus($twitString)
	"Twitter Status Updated"
}
else {"Nothing changed since last update"}

