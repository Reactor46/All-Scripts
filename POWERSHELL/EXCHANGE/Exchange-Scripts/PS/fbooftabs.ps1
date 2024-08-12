$tbRptSourceHeader=@'
<!DOCTYPE html>
<html lang="en">
<head>
<style>
body
{
	font-family: "Segoe UI", arial, helvetica, freesans, sans-serif;
	font-size: 90%;
	color: #333;
	background-color: #e5eaff;
	margin: 10px;
	z-index: 0;
}

h1
{
	font-size: 1.5em;
	font-weight: normal;
	margin: 0;
}

h2
{
	font-size: 1.3em;
	font-weight: normal;
	margin: 2em 0 0 0;
}

p
{
	margin: 0.6em 0;
}

p.tabnav
{
	font-size: 1.1em;
	text-transform: uppercase;
	text-align: right;
}

p.tabnav a
{
	text-decoration: none;
	color: #999;
}

article.tabs
{
	position: relative;
	display: block;
	width: 80em;
	height: 30em;
	margin: 2em auto;
}

article.tabs section
{
	position: absolute;
	display: block;
	top: 1.8em;
	left: 0;
	height: 42em;
	padding: 10px 20px;
	background-color: #ddd;
	border-radius: 5px;
	box-shadow: 0 3px 3px rgba(0,0,0,0.1);
	z-index: 0;
}

article.tabs section:first-child
{
	z-index: 1;
}

article.tabs section h2
{
	position: absolute;
	font-size: 1em;
	font-weight: normal;
	width: 120px;
	height: 1.8em;
	top: -1.8em;
	left: 10px;
	padding: 0;
	margin: 0;
	color: #999;
	background-color: #ddd;
	border-radius: 5px 5px 0 0;
}
'@ 

$styleFooter=@'
article.tabs section h2 a
{
	display: block;
	width: 100%;
	line-height: 1.8em;
	text-align: center;
	text-decoration: none;
	color: inherit;
	outline: 0 none;
}

article.tabs section,
article.tabs section h2
{
	-webkit-transition: all 500ms ease;
	-moz-transition: all 500ms ease;
	-ms-transition: all 500ms ease;
	-o-transition: all 500ms ease;
	transition: all 500ms ease;
}

article.tabs section:target,
article.tabs section:target h2
{
	color: #333;
	background-color: #fff;
	z-index: 2;
}
</Style>
<meta charset="UTF-8" />
<title>Tabbed FreeBusy-Out of Office Board</title>
<!--[if lt IE 9]>
<script src="http://html5shiv.googlecode.com/svn/trunk/html5.js"></script>
<![endif]-->
</head>
<body>
<article class="tabs">
'@
## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

## Load Managed API dll  
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"  
  
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  
  
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
  
## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

## end code from http://poshcode.org/624
  
## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
  
#CAS URL Option 1 Autodiscover  
$service.AutodiscoverUrl($MailboxName,{$true})  
"Using CAS Server : " + $Service.url   
   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 



$fbGroup = $service.ExpandGroup($args[1]);
$StartTime = [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00"))
$EndTime = $StartTime.AddDays(1)

$displayStartTime =  [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 08:30"))
$tmValHash = @{ }
$tidx = 0

for($vsStartTime=[DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00"));$vsStartTime -lt [DateTime]::Parse([DateTime]::Now.ToString("yyyy-MM-dd 0:00")).AddDays(1);$vsStartTime = $vsStartTime.AddMinutes(30)){
	$tmValHash.add($vsStartTime.ToString("HH:mm"),$tidx)	
	$tidx++
}
  
$drDuration = new-object Microsoft.Exchange.WebServices.Data.TimeWindow($StartTime,$EndTime)  
$AvailabilityOptions = new-object Microsoft.Exchange.WebServices.Data.AvailabilityOptions  
$AvailabilityOptions.RequestedFreeBusyView = [Microsoft.Exchange.WebServices.Data.FreeBusyViewType]::DetailedMerged  
 
$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.AttendeeInfo" -as "Type")
$Attendeesbatch = [Activator]::CreateInstance($type) 
$mbrRequest = ""
foreach ($mbr in $fbGroup.Members){
	$Attendee = new-object Microsoft.Exchange.WebServices.Data.AttendeeInfo($mbr.Address)
	$mbrRequest = $mbrRequest + "<Mailbox xmlns=`"http://schemas.microsoft.com/exchange/services/2006/types`"><EmailAddress>" + $mbr.Address + "</EmailAddress></Mailbox>" 
	$Attendeesbatch.add($Attendee)  
}

$attendeeOOFHash = @{}
##Get OOF Status
$expHeader = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
<soap:Header><RequestServerVersion Version="Exchange2010_SP1" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" />
</soap:Header>
<soap:Body>
<GetMailTips xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
<SendingAs>
<EmailAddress xmlns="http://schemas.microsoft.com/exchange/services/2006/types">$MailboxName</EmailAddress>
</SendingAs>
<Recipients>
"@    
    

$expRequest = $expHeader + $mbrRequest + "</Recipients><MailTipsRequested>OutOfOfficeMessage</MailTipsRequested></GetMailTips></soap:Body></soap:Envelope>"

$mbMailboxFolderURI = New-Object System.Uri($service.url)
$wrWebRequest = [System.Net.WebRequest]::Create($mbMailboxFolderURI)
$wrWebRequest.CookieContainer =  New-Object System.Net.CookieContainer 
$wrWebRequest.KeepAlive = $false;
$wrWebRequest.Headers.Set("Pragma", "no-cache");
$wrWebRequest.Headers.Set("Translate", "f");
$wrWebRequest.Headers.Set("Depth", "0");
$wrWebRequest.ContentType = "text/xml";
$wrWebRequest.ContentLength = $expRequest.Length;
$wrWebRequest.Timeout = 60000;
$wrWebRequest.Method = "POST";
$wrWebRequest.Credentials = $creds
$bqByteQuery = [System.Text.Encoding]::ASCII.GetBytes($expRequest);
$wrWebRequest.ContentLength = $bqByteQuery.Length;
$rsRequestStream = $wrWebRequest.GetRequestStream();
$rsRequestStream.Write($bqByteQuery, 0, $bqByteQuery.Length);
$rsRequestStream.Close();
$wrWebResponse = $wrWebRequest.GetResponse();
$rsResponseStream = $wrWebResponse.GetResponseStream()
$sr = new-object System.IO.StreamReader($rsResponseStream);
$rdResponseDocument = New-Object System.Xml.XmlDocument
$rdResponseDocument.LoadXml($sr.ReadToEnd());
$RecipientNodes = @($rdResponseDocument.getElementsByTagName("t:RecipientAddress"))
$Datanodes = @($rdResponseDocument.getElementsByTagName("t:OutOfOffice"))
for($ic=0;$ic -lt $RecipientNodes.length;$ic++){
	if($Datanodes[$ic].ReplyBody.Message -eq ""){
		$attendeeOOFHash.add($Attendeesbatch[$ic].SmtpAddress,"In the Office")
	}
	else{
		$attendeeOOFHash.add($Attendeesbatch[$ic].SmtpAddress,"Out of the Office")
	}
}
### End OOF

$rptOutput = $tbRptSourceHeader

$tabNumber = 1
$taboffset = 10
$SectionReport = ""

$atndCnt = 0  
$fbType = [Microsoft.Exchange.WebServices.Data.AvailabilityData]::FreeBusy
$availresponse = $service.GetUserAvailability($Attendeesbatch,$drDuration,$fbType,$AvailabilityOptions)  
foreach($avail in $availresponse.AttendeesAvailability){  
	if($tabNumber -eq 1){
		
	}
	else{
		$taboffset+=122
		$rptOutput = $rptOutput + "article.tabs section:nth-child(" + $tabNumber +  ") h2 { 	left: " + ($taboffset) + "px;}`r`n" 
		
	}
	$SectionReport = $SectionReport + "<section id=`"tab" + $tabNumber + "`">`r`n"
	$SectionReport = $SectionReport + "<h2><a href=`"#tab" + $tabNumber + "`">" + $Attendeesbatch[$atndCnt].SmtpAddress.SubString(0,10) + "</a></h2>`r`n"
	$SectionReport = $SectionReport + "<p>User : " + $Attendeesbatch[$atndCnt].SmtpAddress + "_________________________________________________________________________________________</p>`r`n"
	$SectionReport = $SectionReport + "<p>OOF Status : " + $attendeeOOFHash[$Attendeesbatch[$atndCnt].SmtpAddress] + "</p>`r`n"
	$SectionReport = $SectionReport + "<p>Number of Calendar Events : " + $avail.CalendarEvents.Count + "</p>`r`n"
	$SectionReport = $SectionReport + "<table><tr>" +"`r`n"
	$SectionReport = $SectionReport + "<td align=`"center`" style=`"width=200;`" ><b>Time</b></td>" +"`r`n"
	$SectionReport = $SectionReport + "<td align=`"center`" style=`"width=200;`" ><b>Status</b></td>" +"`r`n"
	$SectionReport = $SectionReport + "<td align=`"center`" style=`"width=200;`" ><b>Meetings</b></td>" +"`r`n"
	$SectionReport = $SectionReport + "</tr>"

	
	$tabNumber++
	
	""
	"User : " + $Attendeesbatch[$atndCnt].SmtpAddress
	"OOF Status : " + $attendeeOOFHash[$Attendeesbatch[$atndCnt].SmtpAddress]
	"Number of Calender Events : " + $avail.CalendarEvents.Count

	$fbcnt = 0;
	for($stime = $displayStartTime;$stime -lt $displayStartTime.AddHours(10);$stime = $stime.AddMinutes(30)){
		$title = ""
		if ($avail.MergedFreeBusyStatus[$tmValHash[$stime.ToString("HH:mm")]] -eq "Busy" -bor $avail.MergedFreeBusyStatus[$tmValHash[$stime.ToString("HH:mm")]] -eq "OOF"){
			if ($avail.CalendarEvents.Count -ne 0){
				for($ci=0;$ci -lt $avail.CalendarEvents.Count;$ci++){
					if ($avail.CalendarEvents[$ci].StartTime -ge $stime -band $stime -le $avail.CalendarEvents[$ci].EndTime ){
						if($avail.CalendarEvents[$ci].Details.IsPrivate -eq $False){
							$subject = ""
							$location = ""
							if ($avail.CalendarEvents[$ci].Details.Subject -ne $null){
								$subject = $avail.CalendarEvents[$ci].Details.Subject.ToString()
							}
							if ($avail.CalendarEvents[$ci].Details.Location -ne $null){
								$location = $avail.CalendarEvents[$ci].Details.Location.ToString()
							}
							$title = $title + "`"" + $subject + " " + $location + "`" "
						}
					}
				}
			}
		}
		$tbClr = "bgcolor=`"#41A317`""
		if($avail.MergedFreeBusyStatus[$tmValHash[$stime.ToString("HH:mm")]] -eq "Busy"){
			$tbClr = "bgcolor=`"#153E7E`""
		}
		$SectionReport = $SectionReport + "<tr>" +"`r`n"
		$SectionReport = $SectionReport + "<td align=`"center`" style=`"width=200;`" ><b>" + $stime.ToString("HH:mm") + " </b></td>" +"`r`n"
		$SectionReport = $SectionReport + "<td $tbClr align=`"center`" style=`"width=200;`" ><b>" + $avail.MergedFreeBusyStatus[$tmValHash[$stime.ToString("HH:mm")]] + "</b></td>" +"`r`n"
		$SectionReport = $SectionReport + "<td align=`"center`" style=`"width=200;`" ><b>" + $title + "</b></td>" +"`r`n"
		$SectionReport = $SectionReport + "</tr>"	
		$stime.ToString("HH:mm") + " : " +  $avail.MergedFreeBusyStatus[$tmValHash[$stime.ToString("HH:mm")]] + " : " + $title
		$fbcnt++
	}
	$SectionReport = $SectionReport + "</table>"
	$SectionReport = $SectionReport + "</section>`r`n"
	$atndCnt++
} 
$rptOutput = $rptOutput + $styleFooter + $SectionReport + "</article></body></html>"
$rptOutput | Out-File c:\temp\taboutput.htm


