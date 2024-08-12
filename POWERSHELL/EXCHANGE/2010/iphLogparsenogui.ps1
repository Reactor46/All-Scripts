get-childitem c:\inetpub\logs\logfiles\w3svc1\*.* | where-object{$_.LastWriteTime -gt (get-date).addhours(-6)} | foreach-object{
  $fname = $_.FullName
}

$sendAlertTo = "mailbox@example.com"
$sendAlertFrom = "mailbox@example.com"
$SMTPServer = "smtp.example.com"

$appleCollection = @()
$appleverhash = @{ }
$hndSethash = @{ }

function addtoIOShash($inputvar){
	$inparray = $inputvar.Split(',')
	$v1 = $inparray[1].Substring(0,1)
	$v2 = $inparray[1].Substring(1,1)
	$v3 = $inparray[1].Substring(2,($inparray[1].Length-2))
	$useragent = "{0:D2}" -f (([int][char]$v2)-64) 
	$apverobj = "" | select IOSVersion,IOSReleaseDate,ActiveSyncUserAgent,AppleBuildCode
	$apverobj.IOSVersion = $inparray[0]
	$apverobj.IOSReleaseDate = $inparray[2]
	$apverobj.AppleBuildCode = $inparray[1]
	$apverobj.ActiveSyncUserAgent = $v1 + $useragent + "." + $v3
	$appleverhash.add($apverobj.ActiveSyncUserAgent,$apverobj)

}



#IOSVersion,Deviceid,ReleaseDate
addtoIOShash("1,1A543,Jun-07")
addtoIOShash("1.0.1,1C25,Jul-07")
addtoIOShash("1.0.2,1C28,Aug-07")
addtoIOShash("1.1,3A100,Sep-07")
addtoIOShash("1.1,3A101,Sep-07")
addtoIOShash("1.1.1,3A109,Sep-07")
addtoIOShash("1.1.1,3A110,Sep-07")
addtoIOShash("1.1.2 ,3B48,Nov-07")
addtoIOShash("1.1.3,4A93,Jan-08")
addtoIOShash("1.1.4,4A102,Feb-08")
addtoIOShash("1.1.5,4B1,Jul-08")
addtoIOShash("2,5A347,Jul-08")
addtoIOShash("2.0.1,5B108,Aug-08")
addtoIOShash("2.0.2,5C1,Aug-08")
addtoIOShash("2.1,5F136,Sep-08")
addtoIOShash("2.1,5F137,Sep-08")
addtoIOShash("2.1,5F138,Sep-08")
addtoIOShash("2.2,5G77,Nov-08")
addtoIOShash("2.2.1,5H11,Jan-09")
addtoIOShash("3,7A341,Jun-09")
addtoIOShash("3.0.1,7A400,Jul-09")
addtoIOShash("3.1,7C144,Sep-09")
addtoIOShash("3.1,7C145,Sep-09")
addtoIOShash("3.1,7C146,Sep-09")
addtoIOShash("3.1.2,7D11,Oct-09")
addtoIOShash("3.1.3,7E18,Feb-09")
addtoIOShash("3.2,7B367,Apr-10")
addtoIOShash("3.2.1,7B405,Jul-10")
addtoIOShash("4.0,8A293,Jun-10")
addtoIOShash("4.0.1,8A306,Jul-10")

$hndSethash.add("Apple-iPhone","IPhone")
$hndSethash.add("Apple-iPhone1C2","IPhone 3G")
$hndSethash.add("Apple-iPhone2C1","IPhone 3GS")
$hndSethash.add("Apple-iPhone3C1","IPhone 4")
$hndSethash.add("Apple-iPad","IPad")
$hndSethash.add("Apple-iPod","IPod Touch")

$CurrentDate = Get-Date





$mbcombCollection = @()
$FldHash = @{}
$usHash = @{} 
$fieldsline = (Get-Content $fname)[3]
$fldarray = $fieldsline.Split(" ")
$fnum = -1
foreach ($fld in $fldarray){
	$FldHash.add($fld,$fnum)
	$fnum++
}

get-content $fname | Where-Object -FilterScript { $_ -ilike “*&DeviceType=*” } | %{ 
	$lnum ++
	write-progress "Scanning Line" $lnum
	if ($lnum -eq $rnma){ Write-Progress -Activity "Read Lines" -Status $lnum
		$rnma = $rnma + 1000
	}
	$linarr = $_.split(" ")
	$uid = $linarr[$FldHash["cs-username"]] + $linarr[$FldHash["cs(User-Agent)"]]
	if ($linarr[$FldHash["cs-username"]].length -gt 2){
		if ($usHash.Containskey($uid) -eq $false){
			$usrobj = "" | select UserName,UserAgent,Iphone,IphoneType,IOSVersion,IOSReleaseDate,AppleBuildCode
			$usrobj.UserName = $linarr[$FldHash["cs-username"]]
			$usrobj.UserAgent = $linarr[$FldHash["cs(User-Agent)"]]
			if ($usrobj.UserAgent -match "apple"){
				$apcodearray = $usrobj.UserAgent.split("/")
				$usrobj.Iphone = "Yes"
				$usrobj.IphoneType = $hndSethash[$apcodearray[0]]
				$usrobj.IOSVersion = $appleverhash[$apcodearray[1]].IOSVersion
				$usrobj.IOSReleaseDate = $appleverhash[$apcodearray[1]].IOSReleaseDate
				$usrobj.AppleBuildCode = $appleverhash[$apcodearray[1]].AppleBuildCode
			}
			else{
				$usrobj.Iphone = "No"
			}
			$usHash.add($uid,$usrobj)
			$mbcombCollection += $usrobj
		
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
$MailMessage.Subject = "iPhone Registrration Report" 
$MailMessage.IsBodyHtml = $TRUE
$MailMessage.body = $mbcombCollection | ConvertTo-HTML -head $tableStyle –body $body 
$SMTPClient.Send($MailMessage)
  