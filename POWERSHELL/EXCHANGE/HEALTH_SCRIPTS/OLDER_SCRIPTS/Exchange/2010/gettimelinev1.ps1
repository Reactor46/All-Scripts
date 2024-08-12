$twitterusername = "username"
$twitterpassword = "password"
$ExchangeServername = "servername"
$emEmailAddress = "twittest@domain.com"
$userName = "twittest"
$password = "Password"
$domain = "domain"

$casURL = "https://" + $ExchangeServername + "/EWS/Exchange.asmx"
[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")

$statushash = @{ }

function convertTwitterTime($ptim){
	$year = $ptim.Substring(26,4)
	$month = $ptim.Substring(4,3)
	$day = $ptim.Substring(0,3)
	$dayNum = $ptim.Substring(8,2)
	$time = $ptim.Substring(11,8)
	$combTime = $day + " " + $dayNum + " " + $month + " " + $year + " " + $time + " GMT"
	$convertedTime = [DateTime]::Parse($combTime)
	return $convertedTime.ToLocalTime()
}


$TempDir = "c:\Temp"
if (!(Test-Path -path $TempDir))
{
	New-Item $TempDir -type directory
}
$StatusFile = $TempDir + "\twitStatus.xml"
if (!(Test-Path -path $StatusFile))
{
	"Non Status File First Run ?"
}
else{
	"Status Found"
	[xml]$StatusXmlDoc = Get-Content $StatusFile
	foreach($status in $StatusXmlDoc.statuses.status){
		$statushash.Add($status.id.ToString(),1)
	}
	
}
$ewc = new-object EWSUtil.EWSConnection($emEmailAddress,$false, $userName, $password,$domain, $casURL)
[System.Net.ServicePointManager]::Expect100Continue = $false
$request = [System.Net.WebRequest]::Create("http://twitter.com/statuses/friends_timeline.xml")
$request.Credentials = new-object System.Net.NetworkCredential($twitterusername,$twitterpassword)  
$request.Method = "GET"
$request.ContentType = "application/x-www-form-urlencoded"        
$response = $request.GetResponse()
$ResponseStream = $response.GetResponseStream()
$ResponseXmlDoc = new-object System.Xml.XmlDocument
$ResponseXmlDoc.Load($ResponseStream)
$StatusNodes = @($ResponseXmlDoc.getElementsByTagName("status"))
for($snodes=0;$snodes -lt $StatusNodes.Length;$snodes++){
	if ($statushash.Containskey($StatusNodes[$snodes].id) -eq $false){
		if ($twitterusername -ne $StatusNodes[$snodes].user.screen_name){
			[EWSUtil.EWS.DistinguishedFolderIdType] $dType = new-object EWSUtil.EWS.DistinguishedFolderIdType
			$dType.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox
			$time = convertTwitterTime($StatusNodes[$snodes].created_at)
			$ewc.CreateTwitMail($dType,$time,$StatusNodes[$snodes].user.Name.Replace(" ","") + "@twitterexdev.com",$StatusNodes[$snodes].user.Name,$StatusNodes[$snodes].text)
		}
	}
}
$ResponseXmlDoc.Save($TempDir + "\twitStatus.xml")